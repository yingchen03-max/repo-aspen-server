import os
import json
import pandas as pd
import win32com.client 
import pythoncom
from typing import Dict, List, Any, Optional
import time
from datetime import datetime
from collections import deque
import traceback
import uuid
import tempfile
from pathlib import Path

# 导入配置
try:
    from config import (
        HOST, PORT, DEBUG,
        BASE_DIR, TEMPLATE_DIR, OUTPUT_DIR, RESULT_DIR, CONFIG_DIR,
        DEFAULT_TEMPLATE, SSL_CERT_FILE, SSL_KEY_FILE, SSL_CERT_PATH, SSL_KEY_PATH,
        SCHEMA_DIR,
        print_config, validate_config
    )
    CONFIG_AVAILABLE = True
except ImportError:
    # 如果配置文件不存在，使用默认值
    CONFIG_AVAILABLE = False
    from pathlib import Path
    HOST = "0.0.0.0"
    PORT = int(os.getenv("ASPEN_SIMULATOR_PORT"))
    DEBUG = True
    BASE_DIR = Path("D:/aspen")
    TEMPLATE_DIR = BASE_DIR / "orgfile"
    OUTPUT_DIR = BASE_DIR / "bkpfile"
    RESULT_DIR = BASE_DIR / "resultfile"
    CONFIG_DIR = BASE_DIR / "configfile"
    DEFAULT_TEMPLATE = TEMPLATE_DIR / "test.bkp"
    SSL_CERT_FILE = "ssl/cert.pem"
    SSL_KEY_FILE = "ssl/key.pem"
    SSL_CERT_PATH = Path(__file__).parent / SSL_CERT_FILE
    SSL_KEY_PATH = Path(__file__).parent / SSL_KEY_FILE
    SCHEMA_DIR = Path(__file__).parent.parent / "schema"

from ssl_utils import ensure_server_certificate

# dotenv 相关导入改为可选
try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
    load_dotenv()
except ImportError:
    DOTENV_AVAILABLE = False
    def load_dotenv():
        pass  # 空函数，不做任何操作

# Flask 相关导入改为可选（只有在作为 Flask 应用运行时才需要）
try:
    from flask import Flask, request, jsonify
    FLASK_AVAILABLE = True
except ImportError:
    FLASK_AVAILABLE = False
    Flask = None
    request = None
    jsonify = None

# 只有在 Flask 可用时才创建 app
if FLASK_AVAILABLE:
    app = Flask(__name__)
else:
    app = None

# 全局变量存储控制面板消息
control_panel_messages = deque(maxlen=1000)  # 限制最多存储1000条消息

class AspenSimulationManager:
    def __init__(self, aspen_executable_path: str = None):
        """
        初始化Aspen Plus模拟管理器

        Args:
            aspen_executable_path: Aspen Plus可执行文件路径(可选)
        """
        try:
            pythoncom.CoInitialize()
            self.aspen = win32com.client.Dispatch("Apwn.Document")
            
            print("成功连接到Aspen Plus")
            # 连接事件处理器
            self.aspen_events = win32com.client.WithEvents(self.aspen, AspenEvents)
        except Exception as e:
            print(f"无法连接到Aspen Plus: {e}")
            if aspen_executable_path and os.path.exists(aspen_executable_path):
                os.startfile(aspen_executable_path)
                # 等待Aspen启动
                time.sleep(5)
                self.aspen = win32com.client.Dispatch("Apwn.Document")
            else:
                raise Exception("无法启动Aspen Plus，请检查安装")

        # 添加获取控制面板消息的方法
    def get_control_panel_messages(self) -> str:
        """获取控制面板消息"""
        if hasattr(self, 'aspen_events'):
            return self.aspen_events.get_current_session_messages_as_string()
        return ""

    def create_new_simulation(self, template_path: str = None):
        """
        创建新的模拟文件

        Args:
            template_path: 模板文件路径(可选)
        """
        try:
            if template_path and os.path.exists(template_path):
                self.aspen.InitFromArchive2(template_path)
            else:
                self.aspen.InitFromArchive2("")  # 空模拟
                #self.aspen.InitNew2()
            print("成功创建新模拟")
            #self.aspen.Visible = True
        except Exception as e:
            print(f"创建模拟失败: {e}")
            raise

    def load_json_config(self, config_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        加载JSON配置数据

        Args:
            config_data: JSON配置数据字典

        Returns:
            JSON配置字典
        """
        print("成功加载JSON配置数据")
        return config_data

    def get_child_nodes(self, parent_path: str) -> List[str]:
        """获取指定父节点下的所有子节点名称"""
        try:
            parent_node = self.aspen.Tree.FindNode(parent_path)
            if parent_node and parent_node.Elements.Count > 0:
                return [child.Name for child in parent_node.Elements]
            else:
                return []
        except Exception as e:
            print(f"获取 {parent_path} 子节点时出错: {e}")
            return []

    def safe_get_node_value(self, node_path: str) -> Any:
        """安全获取节点值"""
        try:
            node = self.aspen.Tree.FindNode(node_path)
            if node:
                return node.Value
            return None
        except Exception as e:
            print(f"获取节点 {node_path} 值时出错: {e}")
            return None

    def safe_set_node_value(self, node_path: str, value: Any) -> bool:
        """安全设置节点值"""
        try:
            node = self.aspen.Tree.FindNode(node_path)
            if node:
                node.Value = value
                return True
            return False
        except Exception as e:
            print(f"设置节点 {node_path} 值时出错: {e}")
            return False

    def safe_get_node_units(self, node_path: str, default: Any = None) -> Any:
        """安全获取节点单位，避免节点不存在时抛出异常"""
        try:
            node = self.aspen.Tree.FindNode(node_path)
            if node:
                return node.UnitString
            else:
                return default
        except Exception as e:
            print(f"获取节点 {node_path} 单位时出错: {e}")
            return default

    def convert_unitstr(self, s):
        conversion_map = {
            "bar": 5,
            "C": 4,
            "K": 1,  # 开尔文温度单位
            "mol/l": 4,
            "kmol/hr": 3,
            "cum/hr": 7,  # 体积流量单位：立方米每小时
            "cum": 1,  # 体积单位：立方米
            "kPa": 10, # 压力单位
            "kg/hr": 3, # 质量流量单位：千克每小时
            "kg/sec": 1,  # 质量流量单位：千克每秒
            "kg": 3,
            "atm": 3,
            "kW": 14, # 负荷单位
            "Gcal/hr": 18,  # 负荷单位
            "kcal/hr-sqm-K": 5,  # 传热系数单位
            "cal/sec-sqcm-K": 3,  # 传热系数单位
            "Watt/sqm-K": 1,  # 传热系数单位
            "sqm": 1,  # 面积单位：平方米
            "kg/cum": 1, #颗粒密度单位
            "gm/cc": 3, #颗粒密度单位
            "gm/ml": 6, #颗粒密度单位
            "lb/bbl": 6, #颗粒密度单位
            "lb/cuft": 6, #颗粒密度单位
            "lb/gal": 4, #颗粒密度单位
            "cal/sec": 3, #负荷单位
            "cal/mol": 3,  # 能量单位：卡每摩尔
            "Btu/lbmol": 2,  # 能量单位：英热单位每磅摩尔
            "MMkcal/hr": 7,  # 负荷单位：百万千卡每小时
            "Gcal/hr": 18,  # 负荷单位：千兆卡每小时
            "MPa": 20,
            "N/sqm": 1,  # 压力单位：牛顿每平方米（帕斯卡）
            "mm": 7,  # 长度单位：毫米
            "l": 3,  # 体积单位：升
            "": 0,
        }
        if s in conversion_map:
            return conversion_map[s]
        else:
            raise ValueError(f"无法转换字符串 '{s}'，未找到对应的转换规则")

    def add_if_not_empty(self, data_dict, node, value_key, unit_key=None, basis_key=None):
        """如果值不为空，则将其添加到字典中"""
        if node is None:
            return
        if value_key in data_dict and unit_key in data_dict and data_dict[
            value_key] is not None and basis_key is not None:
            node.SetValueUnitAndBasis(data_dict[value_key], self.convert_unitstr(data_dict[unit_key]),
                                      data_dict[basis_key])
        elif value_key in data_dict and unit_key in data_dict and data_dict[value_key] is not None:
            node.SetValueAndUnit(data_dict[value_key], self.convert_unitstr(data_dict[unit_key]))
        elif value_key in data_dict and data_dict[value_key] is not None and unit_key is None:
            node.Value = data_dict[value_key]

    def write_config_to_aspen(self, config: Dict[str, Any]):
        """
        将所有配置写入Aspen模拟文件
        """
        print("开始将配置写入Aspen模拟文件...")
        self.write_setup_to_aspen(config)
        self.write_components_to_aspen(config)
        self.write_property_methods_to_aspen(config)
        self.write_blocks_to_aspen(config)
        self.write_stream_to_aspen(config)
        self.write_block_connections_to_aspen(config)
        self.write_stream_data_to_aspen(config)
        self.write_reactions_data_to_aspen(config)
        self.write_convergence_data_to_aspen(config)
        self.write_design_specs_data_to_aspen(config)
        self.write_blocks_Mixer_data_to_aspen(config)
        self.write_blocks_Valve_data_to_aspen(config)
        self.write_blocks_Compr_data_to_aspen(config)
        self.write_blocks_Heater_data_to_aspen(config)
        self.write_blocks_Pump_data_to_aspen(config)
        self.write_blocks_RStoic_data_to_aspen(config)
        self.write_blocks_RPlug_data_to_aspen(config)
        self.write_blocks_Flash2_data_to_aspen(config)
        self.write_blocks_Flash3_data_to_aspen(config)
        self.write_blocks_Decanter_data_to_aspen(config)
        self.write_blocks_Sep_data_to_aspen(config)
        self.write_blocks_Sep2_data_to_aspen(config)
        self.write_blocks_RadFrac_data_to_aspen(config)
        self.write_blocks_DSTWU_data_to_aspen(config)
        self.write_blocks_Distl_data_to_aspen(config)
        self.write_blocks_Dupl_data_to_aspen(config)
        self.write_blocks_Extract_data_to_aspen(config)
        self.write_blocks_FSplit_data_to_aspen(config)
        self.write_blocks_HeatX_data_to_aspen(config)
        self.write_blocks_MCompr_data_to_aspen(config)
        self.write_blocks_RCSTR_data_to_aspen(config)
        print("所有数据提取完成")

    def write_setup_to_aspen(self, config: Dict[str, Any]):
        """
        将设置的配置写入Aspen模拟文件
        """
        try:
            sim_options = config.get("setup", {}).get("sim_options", {})
            ENERGY_BAL_NODE = self.aspen.Tree.FindNode(r"\Data\Setup\Sim-Options\Input\ENERGY_BAL")
            self.add_if_not_empty(sim_options, ENERGY_BAL_NODE, "energy_bal_value")
            print("成功添加setup")
        except Exception as e:
            print(f"在添加setup时出错: {e}")
            raise
    def write_components_to_aspen(self, config: Dict[str, Any]):
        """
        将配置写入Aspen模拟文件
        """
        try:
            # 添加组分
            try:
                aname1_node = self.aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\ANAME1")
                casn_node = self.aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\CASN")
                actual_count = 0
                for i, component in enumerate(config.get('components', [])):
                    aname1_node.Elements.InsertRow(0, 0)
                    aname1_node.Elements.LabelNode(0, 0)[0].Value = component['cid']
                    aname1_node.Elements(0).Value = component['name']
                    casn_node.Elements(0).Value = component['cas_number']
                    actual_count += 1
                    print(f"添加组分成功:{component['name']}")
                if len(config.get('components', [])) > 0 and actual_count == 0:
                    print(f"WARNING: {len(config['components'])} 个组分定义未被写入，请检查 components 字段。")
                print(f"成功添加组分 ({actual_count}/{len(config.get('components', []))})")
            except Exception as e:
                print(f"在添加组分时出错: {e}")
                raise

            # 处理亨利组分
            try:
                henry_components = config.get('henry_components', {})
                if henry_components:
                    print("开始设置亨利组分...")
                    # 确保Henry-Comps目录存在
                    henry_comps_path = r"\Data\Components\Henry-Comps"
                    henry_comps_node = self.aspen.Tree.FindNode(henry_comps_path)
                    if not henry_comps_node:
                        # 如果目录不存在，可能需要创建
                        components_node = self.aspen.Tree.FindNode(r"\Data\Components")
                        components_node.Elements.Add("Henry-Comps")
                    # 遍历所有Henry组分集
                    for henry_set, hc_data in henry_components.items():
                        # 创建或获取Henry组分集
                        henry_set_path = fr"{henry_comps_path}\{henry_set}"
                        henry_set_node = self.aspen.Tree.FindNode(henry_set_path)
                        if not henry_set_node:
                            henry_comps_node.Elements.Add(henry_set)
                        # 确保Input和CID目录存在
                        cid_path = fr"{henry_set_path}\Input\CID"
                        cid_node_path = self.aspen.Tree.FindNode(cid_path)
                        if not cid_path:
                            print("目录不存在...")
                        # 添加组分
                        for i, component in enumerate(hc_data.get('components', [])):
                            # 创建CID节点
                            cid_node_path.Elements.InsertRow(0, 0)
                            # 设置CID节点的值
                            cid_node_path.Elements(0).Value = component.get('formula', '')
                    print(f"成功设置 {len(henry_components)} 个Henry组分集")
            except Exception as e:
                print(f"在处理亨利组分时出错: {e}")
            # print("components配置已成功写入Aspen模拟文件")
        except Exception as e:
            print(f"写入components配置时出错: {e}")
            raise
    def write_property_methods_to_aspen(self, config: Dict[str, Any]):
        """
        将配置写入Aspen模拟文件
        """
        # 添加物性方法
        try:
            property_methods_node = self.aspen.Tree.FindNode(r"\Data\Properties\Property Methods")
            # 找到基本的物性方法
            basis_method = None
            for i, method_data in enumerate(config.get('property_methods', [])):
                if method_data.get('is_basis_method', True):
                    basis_method = method_data['method_name']
                    GBASEOPSET_node = self.aspen.Tree.FindNode(r"\Data\Properties\Specifications\Input\GBASEOPSET")
                    GBASEOPSET_node.Value = basis_method
                    GOPSETNAME_node = self.aspen.Tree.FindNode(r"\Data\Properties\Specifications\Input\GOPSETNAME")
                    GOPSETNAME_node.Value = basis_method
                    GPPROCTYPE_node = self.aspen.Tree.FindNode(r"\Data\Properties\Specifications\Input\GPPROCTYPE")
                    GPPROCTYPE_node.Value = "ALL"
                print(f"成功设置property_methods: {basis_method}")
        except Exception as e:
            print(f"在设置property_methods时出错: {e}")
            raise
    def write_blocks_to_aspen(self, config: Dict[str, Any]):
        """
        将配置写入Aspen模拟文件
        """
        # 添加模块blocks
        try:
            blocks_node = self.aspen.Tree.FindNode(r"\Data\Blocks")
            for i, blocks in enumerate(config.get('blocks', [])):
                print(f"开始添加blocks:{blocks['name']}!{blocks['type']}")
                blocks_node.Elements.Add(f"{blocks['name']}!{blocks['type']}")
                print(f"添加blocks成功:{blocks['name']}!{blocks['type']}")
            print("成功添加blocks")
        except Exception as e:
            print(f"在添加blocks时出错: {e}")
            raise
    def write_stream_to_aspen(self, config: Dict[str, Any]):
        """
        将配置写入Aspen模拟文件
        """
        # 添加物流streams
        try:
            streams_node = self.aspen.Tree.FindNode(r"\Data\Streams")
            for i, streams in enumerate(config.get('streams', [])):
                streams_node.Elements.Add(f"{streams}")
                print(f"添加streams成功: {streams}")
            print("成功添加streams")
        except Exception as e:
            print(f"在添加streams时出错: {e}")
            raise
    def write_block_connections_to_aspen(self, config: Dict[str, Any]):
        """
        将配置写入Aspen模拟文件
        """
        # 添加连接
        try:
            blocks_node = self.aspen.Tree.FindNode(r"\Data\Blocks")
            for block_name, connection_data in config.get('block_connections', {}).items():
                for streams, type in connection_data.items():
                    #sengwu 测试开始
                    #blocks_node.Elements(block_name).Elements("Ports").Elements(type).Elements.Add(streams) 源代码
                    try:
                        print("Block_Connections: ", block_name, streams, type)
                        blocks_node.Elements(block_name).Elements("Ports").Elements(type).Elements.Add(streams)
                    except Exception as e:
                        print(f"在添加连接 {block_name} - {streams} ({type}) 时出错: {e}，跳过该连接")
                        continue
                    #sengwu 测试结束
            print("成功添加block_connections")
        except Exception as e:
            print(f"在添加block_connections时出错: {e}")
            raise
    def write_stream_data_to_aspen(self, config: Dict[str, Any]):
        """
        将stream_data配置写入Aspen模拟文件
        """
        try:
            for stream, stream_data_detail in config.get('stream_data', {}).items():
                MIXED_SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\MIXED_SPEC\MIXED")
                self.add_if_not_empty(stream_data_detail, MIXED_SPEC_NODE, "MIXED_SPEC")
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\PRES\MIXED")
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\TEMP\MIXED")
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\VFRAC\MIXED")
                if stream_data_detail["MIXED_SPEC"] == "TP":
                    self.add_if_not_empty(stream_data_detail["pressure"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                    self.add_if_not_empty(stream_data_detail["temperature"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                elif stream_data_detail["MIXED_SPEC"] == "TV":
                    self.add_if_not_empty(stream_data_detail["temperature"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                    self.add_if_not_empty(stream_data_detail["vfrac"], VFRAC_NODE, "VFRAC_VALUE")
                elif stream_data_detail["MIXED_SPEC"] == "PV":
                    self.add_if_not_empty(stream_data_detail["pressure"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                    self.add_if_not_empty(stream_data_detail["vfrac"], VFRAC_NODE, "VFRAC_VALUE")
                if "flow" in stream_data_detail:
                    flow_nodes = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\FLOW\MIXED") # 规定-组分流量
                    FLOWBASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\FLOWBASE\MIXED")  # 规定-总流量-基准
                    TOTFLOW_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\TOTFLOW\MIXED")  # 规定-总流量
                    BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Streams\{stream}\Input\BASIS\MIXED")  # 规定-组成-基准
                    self.add_if_not_empty(stream_data_detail["flow"], FLOWBASE_NODE, "FLOWBASE")
                    self.add_if_not_empty(stream_data_detail["flow"], TOTFLOW_NODE, "TOTFLOW_VALUE", "TOTFLOW_UNITS","FLOWBASE")
                    self.add_if_not_empty(stream_data_detail["flow"], BASIS_NODE, "BASIS")
                    for i, components in enumerate(config.get('components', [])):
                        comp = components['cid']
                        if comp in stream_data_detail["flow"]:
                            # comp_flow = stream_data_detail["flow"][comp]
                            # flow_nodes.Elements(comp).Value = comp_flow['FLOW_VALUE']
                            self.add_if_not_empty(stream_data_detail["flow"][comp], flow_nodes.Elements(comp), "FLOW_VALUE", "FLOW_UNITS","FLOW_BASIS")
                print(f"成功添加{stream}的stream_data")
            print("成功添加stream_data")
        except Exception as e:
            print(f"在添加stream_data时出错: {e}")
            raise
    def write_reactions_data_to_aspen(self, config: Dict[str, Any]):
        """
        将reactions_data配置写入Aspen模拟文件
        """
        try:
            for reaction, reactions_data in config.get('reactions', {}).items():
                # 1. 创建反应节点（如果不存在）
                REAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions")
                if not REAC_NODE:
                    print(f"未找到反应节点路径 \\Data\\Reactions\\Reactions")
                    continue
                
                reaction_type = reactions_data.get('type', 'POWERLAW')
                composite_string = f"{reaction}!{reaction_type}"
                
                try:
                    # 检查反应节点是否已存在
                    existing_node = REAC_NODE.Elements(reaction)
                    print(f"反应节点 '{reaction}' 已存在，跳过创建")
                except:
                    # 节点不存在，创建新节点
                    try:
                        REAC_NODE.Elements.Add(composite_string)
                        print(f"成功创建反应节点 '{reaction}' ({reaction_type})")
                        time.sleep(0.3)  # 等待节点创建完成
                    except Exception as e:
                        print(f"创建反应节点失败: {e}")
                        continue
                
                # 2. 获取反应节点和输入节点
                reaction_node = REAC_NODE.Elements(reaction)
                input_node = reaction_node.Elements("Input")
                reactype_node = input_node.Elements("REACTYPE")
                coef_node = input_node.Elements("COEF")  # 反应物系数节点
                coef1_node = input_node.Elements("COEF1")  # 产物系数节点
                
                # 3. 处理 REAC_DATA 数组
                reac_data_list = reactions_data.get('REAC_DATA', [])
                if not reac_data_list:
                    print(f"⚠ 警告: 反应 '{reaction}' 未提供 REAC_DATA 数据")
                    continue
                
                for reac_data in reac_data_list:
                    REAC_ID = reac_data.get('REAC_ID')
                    if not REAC_ID:
                        print(f"⚠ 警告: 反应数据中缺少 REAC_ID")
                        continue
                    
                    # 3.1 添加反应编号到 REACTYPE 节点
                    try:
                        # 插入新反应编号
                        reactype_node.Elements.InsertRow(0, 0)
                        reactype_node.Elements.LabelNode(0, 0)[0].Value = REAC_ID
                        print(f"  ✓ 添加反应编号 {REAC_ID}")
                        
                        # 设置反应类型（REACTYPE）
                        REACTYPE = reac_data.get('REACTYPE')
                        if REACTYPE:
                            REACTYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\REACTYPE\{REAC_ID}")
                            if REACTYPE_NODE:
                                REACTYPE_NODE.Value = REACTYPE
                                print(f"  ✓ 设置 REACTYPE: {REACTYPE}")
                    except Exception as e:
                        print(f"  ✗ 添加反应编号失败: {e}")
                        continue
                    
                    # 3.2 添加反应物（COEF_DATA）
                    COEF_DATA = reac_data.get('COEF_DATA', {})
                    if COEF_DATA:
                        try:
                            COEF_MIX_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\COEF\{REAC_ID}")
                            if not COEF_MIX_NODE:
                                print(f"  ✗ 无法获取反应编号 {REAC_ID} 的 COEF 节点")
                            else:
                                for comp_name, coef_value in COEF_DATA.items():
                                    if coef_value is None:
                                        continue
                                    try:
                                        COEF_MIX_NODE.Elements.InsertRow(0, 0)
                                        COEF_MIX_NODE.Elements.LabelNode(0, 0)[0].Value = comp_name
                                        print(f"    ✓ 插入反应物组分 {comp_name}")
                                    
                                        # 设置反应物系数
                                        COEF_VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\COEF\{REAC_ID}\{comp_name}\MIXED")
                                        if COEF_VALUE_NODE:
                                            COEF_VALUE_NODE.Value = coef_value
                                            print(f"      ✓ 设置系数: {coef_value}")
                                    except Exception as e:
                                        print(f"    ✗ 添加反应物 {comp_name} 失败: {e}")
                        except Exception as e:
                            print(f"  ✗ 处理反应物数据失败: {e}")
                    
                    # 3.3 添加产物（COEF1_DATA）
                    COEF1_DATA = reac_data.get('COEF1_DATA', {})
                    if COEF1_DATA:
                        try:
                            COEF1_MIX_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\COEF1\{REAC_ID}")
                            if not COEF1_MIX_NODE:
                                print(f"  ✗ 无法获取反应编号 {REAC_ID} 的 COEF1 节点")
                            else:
                                for comp_name, coef1_value in COEF1_DATA.items():
                                    if coef1_value is None:
                                        continue
                                    try:

                                        # 插入产物组分
                                        COEF1_MIX_NODE.Elements.InsertRow(0, 0)
                                        COEF1_MIX_NODE.Elements.LabelNode(0, 0)[0].Value = comp_name
                                        print(f"    ✓ 插入产物组分 {comp_name}")
                                    
                                        # 设置产物系数
                                        COEF1_VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\COEF1\{REAC_ID}\{comp_name}\MIXED")
                                        if COEF1_VALUE_NODE:
                                            COEF1_VALUE_NODE.Value = coef1_value
                                            print(f"      ✓ 设置系数: {coef1_value}")
                                    except Exception as e:
                                        print(f"    ✗ 添加产物 {comp_name} 失败: {e}")
                        except Exception as e:
                            print(f"  ✗ 处理产物数据失败: {e}")
                    
                    # 3.4 根据反应类型设置参数
                    REACTYPE = reac_data.get('REACTYPE')
                    
                    # 处理所有反应类型都可能存在的通用参数
                    # PHASE（相态）- EQUIL和KINETIC类型都需要
                    if 'PHASE' in reac_data and reac_data.get('PHASE'):
                        try:
                            PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\PHASE\{REAC_ID}")
                            if PHASE_NODE:
                                PHASE_NODE.Value = reac_data['PHASE']
                                print(f"  ✓ 设置 PHASE: {reac_data['PHASE']}")
                        except Exception as e:
                            print(f"  ✗ 设置 PHASE 失败: {e}")
                    
                    # R_D_RBASIS（速率基准）- EQUIL和KINETIC类型都需要
                    if 'R_D_RBASIS' in reac_data and reac_data.get('R_D_RBASIS'):
                        try:
                            R_D_RBASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\R_D_RBASIS\{REAC_ID}")
                            if R_D_RBASIS_NODE:
                                R_D_RBASIS_NODE.Value = reac_data['R_D_RBASIS']
                                print(f"  ✓ 设置 R_D_RBASIS: {reac_data['R_D_RBASIS']}")
                        except Exception as e:
                            print(f"  ✗ 设置 R_D_RBASIS 失败: {e}")
                    
                    # KINETIC 类型反应的动力学参数（仅在JSON中存在时设置）
                    if REACTYPE == 'KINETIC':
                        # PRE_EXP（指前因子）
                        if 'PRE_EXP' in reac_data and reac_data.get('PRE_EXP') is not None:
                            try:
                                PRE_EXP_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\PRE_EXP\{REAC_ID}")
                                if PRE_EXP_NODE:
                                    PRE_EXP_NODE.Value = reac_data['PRE_EXP']
                                    print(f"  ✓ 设置 PRE_EXP: {reac_data['PRE_EXP']}")
                            except Exception as e:
                                print(f"  ✗ 设置 PRE_EXP 失败: {e}")
                        
                        # T_EXP（温度指数）
                        if 'T_EXP' in reac_data and reac_data.get('T_EXP') is not None:
                            try:
                                T_EXP_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\T_EXP\{REAC_ID}")
                                if T_EXP_NODE:
                                    T_EXP_NODE.Value = reac_data['T_EXP']
                                    print(f"  ✓ 设置 T_EXP: {reac_data['T_EXP']}")
                            except Exception as e:
                                print(f"  ✗ 设置 T_EXP 失败: {e}")
                        
                        # ACT_ENERGY（活化能，有单位）
                        if 'ACT_ENERGY_VALUE' in reac_data and reac_data.get('ACT_ENERGY_VALUE') is not None:
                            try:
                                ACT_ENERGY_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\ACT_ENERGY\{REAC_ID}")
                                if ACT_ENERGY_NODE:
                                    ACT_ENERGY_VALUE = reac_data.get('ACT_ENERGY_VALUE')
                                    ACT_ENERGY_UNITS = reac_data.get('ACT_ENERGY_UNITS')
                                    if ACT_ENERGY_UNITS:
                                        ACT_ENERGY_NODE.SetValueAndUnit(ACT_ENERGY_VALUE, self.convert_unitstr(ACT_ENERGY_UNITS))
                                        print(f"  ✓ 设置 ACT_ENERGY: {ACT_ENERGY_VALUE} (单位: {ACT_ENERGY_UNITS})")
                                    else:
                                        ACT_ENERGY_NODE.Value = ACT_ENERGY_VALUE
                                        print(f"  ✓ 设置 ACT_ENERGY: {ACT_ENERGY_VALUE}")
                            except Exception as e:
                                print(f"  ✗ 设置 ACT_ENERGY 失败: {e}")
                    
                    # CONV 类型反应的参数（仅在JSON中存在时设置）
                    elif REACTYPE == 'CONV':
                        # KEY_CID（关键组分ID）
                        if 'KEY_CID' in reac_data and reac_data.get('KEY_CID'):
                            try:
                                KEY_CID_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\KEY_CID\{REAC_ID}")
                                if KEY_CID_NODE:
                                    KEY_CID_NODE.Value = reac_data['KEY_CID']
                                    print(f"  ✓ 设置 KEY_CID: {reac_data['KEY_CID']}")
                            except Exception as e:
                                print(f"  ✗ 设置 KEY_CID 失败: {e}")
                        
                        # CONV_A
                        if 'CONV_A' in reac_data and reac_data.get('CONV_A') is not None:
                            try:
                                CONV_A_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\CONV_A\{REAC_ID}")
                                if CONV_A_NODE:
                                    CONV_A_NODE.Value = reac_data['CONV_A']
                                    print(f"  ✓ 设置 CONV_A: {reac_data['CONV_A']}")
                            except Exception as e:
                                print(f"  ✗ 设置 CONV_A 失败: {e}")
                        
                        # CONV_B
                        if 'CONV_B' in reac_data and reac_data.get('CONV_B') is not None:
                            try:
                                CONV_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\CONV_B\{REAC_ID}")
                                if CONV_B_NODE:
                                    CONV_B_NODE.Value = reac_data['CONV_B']
                                    print(f"  ✓ 设置 CONV_B: {reac_data['CONV_B']}")
                            except Exception as e:
                                print(f"  ✗ 设置 CONV_B 失败: {e}")
                        
                        # CONV_C
                        if 'CONV_C' in reac_data and reac_data.get('CONV_C') is not None:
                            try:
                                CONV_C_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\CONV_C\{REAC_ID}")
                                if CONV_C_NODE:
                                    CONV_C_NODE.Value = reac_data['CONV_C']
                                    print(f"  ✓ 设置 CONV_C: {reac_data['CONV_C']}")
                            except Exception as e:
                                print(f"  ✗ 设置 CONV_C 失败: {e}")
                        
                        # CONV_D
                        if 'CONV_D' in reac_data and reac_data.get('CONV_D') is not None:
                            try:
                                CONV_D_NODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{reaction}\Input\CONV_D\{REAC_ID}")
                                if CONV_D_NODE:
                                    CONV_D_NODE.Value = reac_data['CONV_D']
                                    print(f"  ✓ 设置 CONV_D: {reac_data['CONV_D']}")
                            except Exception as e:
                                print(f"  ✗ 设置 CONV_D 失败: {e}")
            
            print(f"成功添加reactions_data")
        except Exception as e:
            print(f"在添加reactions_data时出错: {e}")
            raise
    def write_convergence_data_to_aspen(self, config: Dict[str, Any]):
        """
        将convergence_data配置写入Aspen模拟文件
        """
        try:
            conv_options = config.get("convergence", {}).get("conv_options", {})
            # 默认值 - 撕裂收敛
            TOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TOL")
            TRACE_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TRACE")
            TRACEOPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TRACEOPT")
            COMPS_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\COMPS")
            STATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\STATE")
            FLASH_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\FLASH")
            UPDATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\UPDATE")
            VARITERHIST_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\VARITERHIST")
            self.add_if_not_empty(conv_options, TOL_NODE, "tol")
            self.add_if_not_empty(conv_options, TRACE_NODE, "trace")
            self.add_if_not_empty(conv_options, TRACEOPT_NODE, "traceopt")
            self.add_if_not_empty(conv_options, COMPS_NODE, "comps")
            self.add_if_not_empty(conv_options, STATE_NODE, "state")
            self.add_if_not_empty(conv_options, FLASH_NODE, "flash")
            self.add_if_not_empty(conv_options, UPDATE_NODE, "update")
            self.add_if_not_empty(conv_options, VARITERHIST_NODE, "variterhist")
            # 默认方法
            TEAR_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TEAR_METHOD")  # 收敛-选项-默认方法
            SPEC_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SPEC_METHOD")
            MSPEC_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\MSPEC_METHOD")
            COMB_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\COMB_METHOD")
            OPT_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\OPT_METHOD")
            self.add_if_not_empty(conv_options, TEAR_METHOD_NODE, "tear_method")
            self.add_if_not_empty(conv_options, SPEC_METHOD_NODE, "spec_method")
            self.add_if_not_empty(conv_options, MSPEC_METHOD_NODE, "mspec_method")
            self.add_if_not_empty(conv_options, COMB_METHOD_NODE, "comb_method")
            self.add_if_not_empty(conv_options, OPT_METHOD_NODE, "opt_method")
            # 顺序确定
            SPEC_LOOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SPEC_LOOP")
            USER_LOOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\USER_LOOP")
            TEAR_WEIGHT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TEAR_WEIGHT")
            LOOP_WEIGHT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\LOOP_WEIGHT")
            AFFECT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\AFFECT")
            CHECKSEQ_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\CHECKSEQ")
            TEAR_VAR_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\TEAR_VAR")
            self.add_if_not_empty(conv_options, SPEC_LOOP_NODE, "spec_loop")
            self.add_if_not_empty(conv_options, USER_LOOP_NODE, "user_loop")
            self.add_if_not_empty(conv_options, TEAR_WEIGHT_NODE, "tear_weight")
            self.add_if_not_empty(conv_options, LOOP_WEIGHT_NODE, "loop_weight")
            self.add_if_not_empty(conv_options, AFFECT_NODE, "affect")
            self.add_if_not_empty(conv_options, CHECKSEQ_NODE, "checkseq")
            self.add_if_not_empty(conv_options, TEAR_VAR_NODE, "tear_var")
            # 方法 - Wegstein
            WEG_MAXIT_NOD = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\WEG_MAXIT")  # 收敛-选项-迭代次数
            WEG_WAIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\WEG_WAIT")
            ACCELERATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\ACCELERATE")
            NACCELERATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NACCELERATE")
            WEG_QMIN_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\WEG_QMIN")
            WEG_QMAX_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\WEG_QMAX")
            self.add_if_not_empty(conv_options, WEG_MAXIT_NOD, "weg_maxit")
            self.add_if_not_empty(conv_options, WEG_WAIT_NODE, "weg_wait")
            self.add_if_not_empty(conv_options, ACCELERATE_NODE, "accelerate")
            self.add_if_not_empty(conv_options, NACCELERATE_NODE, "naccelerate")
            self.add_if_not_empty(conv_options, WEG_QMIN_NODE, "weg_qmin")
            self.add_if_not_empty(conv_options, WEG_QMAX_NODE, "weg_qmax")
            # 方法 - 直接
            DIR_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\DIR_MAXIT")
            self.add_if_not_empty(conv_options, DIR_MAXIT_NODE, "dir_maxit")
            # 方法 - 正割
            SEC_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SEC_MAXIT")
            STEP_SIZ_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\STEP_SIZ")
            SEC_XTOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SEC_XTOL")
            XFINAL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\XFINAL")
            BRACKET_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\BRACKET")
            STOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\STOP")
            self.add_if_not_empty(conv_options, SEC_MAXIT_NODE, "sec_maxit")
            self.add_if_not_empty(conv_options, STEP_SIZ_NODE, "step_siz")
            self.add_if_not_empty(conv_options, SEC_XTOL_NODE, "sec_xtol")
            self.add_if_not_empty(conv_options, XFINAL_NODE, "xfinal")
            self.add_if_not_empty(conv_options, BRACKET_NODE, "bracket")
            self.add_if_not_empty(conv_options, STOP_NODE, "stop")
            # 方法 - Broyden
            BR_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\BR_MAXIT")
            BR_XTOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\BR_XTOL")
            BR_WAIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\BR_WAIT")
            self.add_if_not_empty(conv_options, BR_MAXIT_NODE, "br_maxit")
            self.add_if_not_empty(conv_options, BR_XTOL_NODE, "br_xtol")
            self.add_if_not_empty(conv_options, BR_WAIT_NODE, "br_wait")
            # 方法 - Newton
            NEW_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NEW_MAXIT")
            NEW_MAXPASS_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NEW_MAXPASS")
            NEW_WAIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NEW_WAIT")
            NEW_XTOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NEW_XTOL")
            OPT_N_JAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\OPT_N_JAC")
            RED_FACTOR_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\RED_FACTOR")
            REINIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\REINIT")
            self.add_if_not_empty(conv_options, NEW_MAXIT_NODE, "new_maxit")
            self.add_if_not_empty(conv_options, NEW_MAXPASS_NODE, "new_maxpass")
            self.add_if_not_empty(conv_options, NEW_WAIT_NODE, "new_wait")
            self.add_if_not_empty(conv_options, NEW_XTOL_NODE, "new_xtol")
            self.add_if_not_empty(conv_options, OPT_N_JAC_NODE, "opt_n_jac")
            self.add_if_not_empty(conv_options, RED_FACTOR_NODE, "red_factor")
            self.add_if_not_empty(conv_options, REINIT_NODE, "reinit")
            # 方法 - SQP
            SQP_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_MAXIT")
            SQP_MAXPASS_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_MAXPASS")
            CONST_ITER_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\CONST_ITER")
            MAXLSPASS_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\MAXLSPASS")
            NLIMIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NLIMIT")
            SQP_TOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_TOL")
            SQP_WAIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_WAIT")
            SQP_QMIN_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_QMIN")
            SQP_QMAX_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\SQP_QMAX")
            self.add_if_not_empty(conv_options, SQP_MAXIT_NODE, "sqp_maxit")
            self.add_if_not_empty(conv_options, SQP_MAXPASS_NODE, "sqp_maxpass")
            self.add_if_not_empty(conv_options, CONST_ITER_NODE, "const_iter")
            self.add_if_not_empty(conv_options, MAXLSPASS_NODE, "maxlspass")
            self.add_if_not_empty(conv_options, NLIMIT_NODE, "nlimit")
            self.add_if_not_empty(conv_options, SQP_TOL_NODE, "sqp_tol")
            self.add_if_not_empty(conv_options, SQP_WAIT_NODE, "sqp_wait")
            self.add_if_not_empty(conv_options, SQP_QMIN_NODE, "sqp_qmin")
            self.add_if_not_empty(conv_options, SQP_QMAX_NODE, "sqp_qmax")
            # 方法 - BOBYQA
            BOBY_MAXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\BOBY_MAXIT")
            NCONDITIONS_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\NCONDITIONS")
            INIT_REGION_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\INIT_REGION")
            FINAL_REGION_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\FINAL_REGION")
            INITPREF_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\INITPREF")
            PREFGROWI_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\PREFGROWI")
            PREFGROWF_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\PREFGROWF")
            EQPENTYP_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\EQPENTYP")
            INEQPENTYP_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\INEQPENTYP")
            PENSCL_NODE = self.aspen.Tree.FindNode(fr"\Data\Convergence\Conv-Options\Input\PENSCL")
            self.add_if_not_empty(conv_options, BOBY_MAXIT_NODE, "boby_maxit")
            self.add_if_not_empty(conv_options, NCONDITIONS_NODE, "nconditions")
            self.add_if_not_empty(conv_options, INIT_REGION_NODE, "init_region")
            self.add_if_not_empty(conv_options, FINAL_REGION_NODE, "final_region")
            self.add_if_not_empty(conv_options, INITPREF_NODE, "initpref")
            self.add_if_not_empty(conv_options, PREFGROWI_NODE, "prefgrowi")
            self.add_if_not_empty(conv_options, PREFGROWF_NODE, "prefgrowf")
            self.add_if_not_empty(conv_options, EQPENTYP_NODE, "eqpentyp")
            self.add_if_not_empty(conv_options, INEQPENTYP_NODE, "ineqpentyp")
            self.add_if_not_empty(conv_options, PENSCL_NODE, "penscl")
            #TEAR_COMPS_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Tear\Input\COMPS")
            TEAR_TOL_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Tear\Input\TOL")
            # 撕裂数据
            tear_data = config.get("convergence", {}).get("tear_data", [])
            for i, tear_streams in enumerate(tear_data):
                tear_stream_name = tear_streams["tear_stream_name"]
                TEAR_TOL_NODES.Elements.InsertRow(0, 0)
                TEAR_TOL_NODES.Elements.LabelNode(0, 0)[0].Value = tear_stream_name
                TEAR_TOL_NODES.Elements(0).Value = tear_streams["tear_stream_tol"]
            # # 计算顺序数据
            # seq_data = config.get("convergence", {}).get("seq_data", [])
            # SEQ_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Sequence")  # 收敛-序列
            # for i, seq in enumerate(seq_data):
            #     seq_name = seq["sep_name"]
            #     sep_type = seq["sep_type"] # 无需添加
            #     SEQ_NODES.Elements.Add(seq_name)
            #     BLOCK_ID_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Sequence\{seq_name}\Input\BLOCK_ID")  # 序列-计算顺序-模块
            #     BLOCK_TYPE_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Sequence\{seq_name}\Input\BLOCK_TYPE")  # 序列-计算顺序-模块
            #     calc_seq_data = seq["calc_seq"]
            #     for num, calc_seq in enumerate(calc_seq_data):
            #         calc_seq_num = calc_seq["seq"]
            #         block_id = calc_seq["block_id"]
            #         block_type = calc_seq["block_type"]
            #         print(block_id)
            #         BLOCK_TYPE_NODES.Elements.InsertRow(0, num)
            #         BLOCK_TYPE_NODES.Elements(num).Value = block_type
            #         BLOCK_ID_NODES.Elements(num).Value = block_id
            # # 收敛-收敛数据
            # conv_data = config.get("convergence", {}).get("conv_data", [])
            # CONV_NODES = self.aspen.Tree.FindNode(fr"\Data\Convergence\Convergence")  # 收敛节点
            # for i, conv in enumerate(conv_data):
            #     conv_name = conv["conv_name"]
            #     CONV_NODES.Elements.Add(conv_name)
            print(f"成功添加convergence_data")
        except Exception as e:
            print(f"在添加convergence_data时出错: {e}")
            raise
    def write_design_specs_data_to_aspen(self, config: Dict[str, Any]):
        """
        将设计规定配置写入Aspen模拟文件
        """
        try:
            # 获取设计规定配置
            design_specs_config = config.get('design_specs', {})
            for spec_name, spec_data in design_specs_config.items():
                print(f"开始写入设计规定: {spec_name}")
                Design_Spec_NODE = self.aspen.Tree.FindNode(fr"\Data\Flowsheeting Options\Design-Spec")
                Design_Spec_NODE.Elements.Add(spec_name)
                base_path = fr"\Data\Flowsheeting Options\Design-Spec\{spec_name}\Input"
                fvn_variable_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_VARIABLE")

                # 2. 写入采样变量 (FVN_*系列)
                sampled_var = spec_data.get("sampled_variables", [])
                for i, sampled_var_data in enumerate(sampled_var):
                    sampled_var_name = sampled_var_data["variable_name"]
                    fvn_variable_node.Elements.InsertRow(0, 0)
                    fvn_variable_node.Elements.LabelNode(0, 0)[0].Value = sampled_var_name
                    # 写入采样变量引用参数（模型工具，物性参数，反应暂不支持）
                    opt_categ_node = self.aspen.Tree.FindNode(fr"{base_path}\OPT_CATEG\{sampled_var_name}") #类别
                    self.add_if_not_empty(sampled_var_data, opt_categ_node, f"opt_categ")
                    variable_type_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_VARTYPE\{sampled_var_name}") #类型
                    block_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_BLOCK\{sampled_var_name}") #模块
                    variable_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_VARIABLE\{sampled_var_name}") #变量
                    sentence_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_SENTENCE\{sampled_var_name}") #语句
                    units_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_UOM\{sampled_var_name}") #单位
                    stream_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_STREAM\{sampled_var_name}") #流股
                    substream_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_SUBS\{sampled_var_name}") #子流股
                    component_node = self.aspen.Tree.FindNode(fr"{base_path}\FVN_COMPONEN\{sampled_var_name}") #组分
                    # fvn_params = ["variable_type", "stream", "block", "variable", "component", "substream", "variable_type", "units", "sentence"]
                    fvn_params_node = [
                        (variable_type_node, "variable_type"),
                        (block_node, "block"),
                        (variable_node, "variable"),
                        (stream_node, "stream"),
                        (substream_node, "substream"),
                        (component_node, "component"),
                        (sentence_node, "sentence"),
                        (units_node, "units")
                    ]
                    for node, key in fvn_params_node:
                        if key in sampled_var_data and node is not None:
                            self.add_if_not_empty(sampled_var_data, node, f"{key}")
                            # self.add_if_not_empty(sampled_var_data, opt_categ_node, f"opt_categ")
                            # self.add_if_not_empty(sampled_var_data, variable_type_node, f"variable_type")
                            # self.add_if_not_empty(sampled_var_data, block_node, f"block")
                            # self.add_if_not_empty(sampled_var_data, variable_node, f"variable")
                            # self.add_if_not_empty(sampled_var_data, sentence_node, f"sentence")
                            # self.add_if_not_empty(sampled_var_data, units_node, f"units")
                            # self.add_if_not_empty(sampled_var_data, stream_node, f"stream")
                            # self.add_if_not_empty(sampled_var_data, substream_node, f"substream")
                            # self.add_if_not_empty(sampled_var_data, component_node, f"component")

                # 3. 写入目标函数配置
                objective_function = spec_data.get("objective_function", {})
                expr1_node = self.aspen.Tree.FindNode(fr"{base_path}\EXPR1")
                tol_node = self.aspen.Tree.FindNode(fr"{base_path}\TOL")
                expr2_node = self.aspen.Tree.FindNode(fr"{base_path}\EXPR2")
                self.add_if_not_empty(objective_function, expr1_node, f"EXPR1")
                self.add_if_not_empty(objective_function, tol_node, f"TOL")
                self.add_if_not_empty(objective_function, expr2_node, f"EXPR2")

                # 4. 写入操纵变量 (VARY_*系列)
                manipulated_variables = spec_data.get("manipulated_variables", [])
                for i, manipulated_var_data in enumerate(manipulated_variables):
                    variable_type_node = self.aspen.Tree.FindNode(fr"{base_path}\VARY_VARTYPE")
                    block_node = self.aspen.Tree.FindNode(fr"{base_path}\VARYBLOCK")
                    variable_name_node = self.aspen.Tree.FindNode(fr"{base_path}\VARYVARIABLE")
                    sentence_node = self.aspen.Tree.FindNode(fr"{base_path}\VARYSENTENCE")
                    units_node = self.aspen.Tree.FindNode(fr"{base_path}\VARYUOM")
                    self.add_if_not_empty(manipulated_var_data, variable_type_node, f"variable_type")
                    self.add_if_not_empty(manipulated_var_data, block_node, f"block")
                    self.add_if_not_empty(manipulated_var_data, variable_name_node, f"variable_name")
                    self.add_if_not_empty(manipulated_var_data, sentence_node, f"sentence")
                    self.add_if_not_empty(manipulated_var_data, units_node, f"units")
                    # 写入VARYLINE1-4
                    for line_num in range(1, 5):
                        line_key = f"line{line_num}"
                        if line_key in manipulated_var_data:
                            line_value = manipulated_var_data[line_key]
                            node_name = f"VARYLINE{line_num}"
                            node = self.aspen.Tree.FindNode(fr"{base_path}\{node_name}")
                            node.Value = line_value

                # 4. 写入操纵变量限制
                bounds = spec_data.get("bounds", {})
                upper_node = self.aspen.Tree.FindNode(fr"{base_path}\UPPER") #上界
                lower_node = self.aspen.Tree.FindNode(fr"{base_path}\LOWER") #下界
                step_size_node = self.aspen.Tree.FindNode(fr"{base_path}\STEP_SIZE") #步长
                max_step_size_node = self.aspen.Tree.FindNode(fr"{base_path}\MAX_STEP_SIZ") #最大步长
                self.add_if_not_empty(bounds, lower_node, f"LOWER")
                self.add_if_not_empty(bounds, upper_node, f"UPPER")
                self.add_if_not_empty(bounds, step_size_node, f"STEP_SIZE")
                self.add_if_not_empty(bounds, max_step_size_node, f"MAX_STEP_SIZ")




                #
                # # 5. 写入边界和步长设置
                # bounds = spec_data.get("bounds", {})
                #
                # # 写入下界
                # if "LOWER" in bounds:
                #     lower_value = bounds["LOWER"]
                #     lower_node = self.aspen.Tree.FindNode(fr"{base_path}\LOWER")
                #     if lower_node is not None and lower_value is not None:
                #         lower_node.Value = lower_value
                #         print(f"  写入LOWER: {lower_value}")
                #
                # # 写入上界
                # if "UPPER" in bounds:
                #     upper_value = bounds["UPPER"]
                #     upper_node = self.aspen.Tree.FindNode(fr"{base_path}\UPPER")
                #     if upper_node is not None and upper_value is not None:
                #         upper_node.Value = upper_value
                #         print(f"  写入UPPER: {upper_value}")
                #
                # # 写入步长
                # if "STEP_SIZE" in bounds:
                #     step_size_value = bounds["STEP_SIZE"]
                #     step_size_node = self.aspen.Tree.FindNode(fr"{base_path}\STEP_SIZE")
                #     if step_size_node is not None and step_size_value is not None:
                #         step_size_node.Value = step_size_value
                #         print(f"  写入STEP_SIZE: {step_size_value}")
                #
                # # 写入最大步长
                # if "MAX_STEP_SIZ" in bounds:
                #     max_step_size_value = bounds["MAX_STEP_SIZ"]
                #     max_step_size_node = self.aspen.Tree.FindNode(fr"{base_path}\MAX_STEP_SIZ")
                #     if max_step_size_node is not None and max_step_size_value is not None:
                #         max_step_size_node.Value = max_step_size_value
                #         print(f"  写入MAX_STEP_SIZ: {max_step_size_value}")
                #
                # # 写入阈值
                # if "THRESHOLD" in bounds:
                #     threshold_value = bounds["THRESHOLD"]
                #     threshold_node = self.aspen.Tree.FindNode(fr"{base_path}\THRESHOLD")
                #     if threshold_node is not None and threshold_value is not None:
                #         threshold_node.Value = threshold_value
                #         print(f"  写入THRESHOLD: {threshold_value}")

                print(f"  设计规定 '{spec_name}' 写入完成")

            print("所有设计规定配置写入完成")

        except Exception as e:
            print(f"写入设计规定配置时出错: {e}")
            import traceback
            traceback.print_exc()
            raise
    def write_blocks_Mixer_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Mixer_data配置写入Aspen模拟文件
        """
        try:
            for block, Mixer_data in config.get('blocks_Mixer_data', {}).items():
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 闪蒸选项-压力
                T_EST_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\T_EST")  # 闪蒸选项-温度估值
                MIXIT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MIXIT")  # 闪蒸选项-最大迭代次数
                TOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TOL")  # 闪蒸选项-容许误差
                self.add_if_not_empty(Mixer_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(Mixer_data["SPEC_DATA"], T_EST_NODE, "T_EST_VALUE", "T_EST_UNITS")
                self.add_if_not_empty(Mixer_data["SPEC_DATA"], MIXIT_NODE, "MIXIT")
                self.add_if_not_empty(Mixer_data["SPEC_DATA"], TOL_NODE, "TOL", )
            print(f"成功添加blocks_Mixer_data")
        except Exception as e:
            print(f"在添加blocks_Mixer_data时出错: {e}")
            raise
    def write_blocks_Valve_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Valve_data配置写入Aspen模拟文件
        """
        try:
            for block, Valve_data in config.get('blocks_Valve_data', {}).items():
                MODE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MODE")  # 作业-计算类型
                self.add_if_not_empty(Valve_data["JOB_DATA"], MODE_NODE, "MODE")
                if Valve_data["JOB_DATA"]["MODE"] == "ADIAB-FLASH":  # 当前只抽取指定出口压力下绝热闪蒸，可自行添加
                    P_OUT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\P_OUT")  # 作业-压力规范-出口压力
                    NPHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NPHASE")  # 作业-闪蒸选项-有效相态
                    FLASH_MAXIT_NODE = self.aspen.Tree.FindNode(
                        fr"\Data\Blocks\{block}\Input\FLASH_MAXIT")  # 作业-闪蒸选项-最大迭代次数
                    FLASH_TOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FLASH_TOL")  # 作业-闪蒸选项-容许误差
                    self.add_if_not_empty(Valve_data["JOB_DATA"], P_OUT_NODE, "P_OUT_VALUE", "P_OUT_UNITS")
                    self.add_if_not_empty(Valve_data["JOB_DATA"], NPHASE_NODE, "NPHASE")
                    self.add_if_not_empty(Valve_data["JOB_DATA"], FLASH_MAXIT_NODE, "FLASH_MAXIT")
                    self.add_if_not_empty(Valve_data["JOB_DATA"], FLASH_TOL_NODE, "FLASH_TOL", )
            print(f"成功添加blocks_Value_data")
        except Exception as e:
            print(f"在添加blocks_Value_data时出错: {e}")
            raise
    def write_blocks_Compr_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Compr_data配置写入Aspen模拟文件
        """
        try:
            for block, Compr_data in config.get('blocks_Compr_data', {}).items():
                MODEL_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MODEL_TYPE")  # 规定-模型
                TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TYPE")  # 规定-类型
                OPT_SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_SPEC")  # 规定-出口规范
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-排放压力
                # UTILITY_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UTILITY_ID")  # 公用工程--暂不添加
                self.add_if_not_empty(Compr_data["SPEC_DATA"], MODEL_TYPE_NODE, "MODEL_TYPE")
                self.add_if_not_empty(Compr_data["SPEC_DATA"], TYPE_NODE, "TYPE", )
                self.add_if_not_empty(Compr_data["SPEC_DATA"], OPT_SPEC_NODE, "OPT_SPEC")
                self.add_if_not_empty(Compr_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                # self.add_if_not_empty(Compr_data["SPEC_DATA"], UTILITY_ID_NODE, "UTILITY_ID")
            print(f"成功添加blocks_Compr_data")
        except Exception as e:
            print(f"在添加blocks_Compr_data时出错: {e}")
            raise
    def write_blocks_Heater_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Heater_data配置写入Aspen模拟文件
        """
        try:
            for block, Heater_data in config.get('blocks_Heater_data', {}).items():
                SPEC_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")  # 规定-温度
                DELT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DELT")  # 规定-温度变化
                DEGSUP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DEGSUP")  # 规定-过热度
                DEGSUB_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DEGSUB")  # 规定-过冷度
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VFRAC")  # 规定-汽相分率
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-压力
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")  # 规定-负载
                # UTILITY_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UTILITY_ID")  # 公用工程--暂不添加
                self.add_if_not_empty(Heater_data["SPEC_DATA"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], DELT_NODE, "DELT_VALUE", "DELT_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], DEGSUP_NODE, "DEGSUP_VALUE", "DEGSUP_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], DEGSUB_NODE, "DEGSUB_VALUE", "DEGSUB_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], VFRAC_NODE, "VFRAC_VALUE")
                self.add_if_not_empty(Heater_data["SPEC_DATA"], SPEC_OPT_NODE, "SPEC_OPT")
                # self.add_if_not_empty(Heater_data["SPEC_DATA"], UTILITY_ID_NODE, "UTILITY_ID")
            print(f"成功添加blocks_Heater_data")
        except Exception as e:
            print(f"在添加blocks_Heater_data时出错: {e}")
            raise
    def write_blocks_Pump_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Pump_data配置写入Aspen模拟文件
        """
        try:
            for block, Pump_data in config.get('blocks_Pump_data', {}).items():
                PUMP_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PUMP_TYPE")  # 规定-模型
                OPT_SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_SPEC")  # 规定-出口规范
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-排放压力
                # UTILITY_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UTILITY_ID")  # 公用工程--暂不添加
                self.add_if_not_empty(Pump_data["SPEC_DATA"], PUMP_TYPE_NODE, "PUMP_TYPE")
                self.add_if_not_empty(Pump_data["SPEC_DATA"], OPT_SPEC_NODE, "OPT_SPEC")
                self.add_if_not_empty(Pump_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                # self.add_if_not_empty(Pump_data["SPEC_DATA"], UTILITY_ID_NODE, "UTILITY_ID")
            print(f"成功添加blocks_Pump_data")
        except Exception as e:
            print(f"在添加blocks_Pump_data时出错: {e}")
            raise
    def write_blocks_RStoic_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_RStoic_data配置写入Aspen模拟文件
        """
        try:
            for block, RStoic_data in config.get('blocks_RStoic_data', {}).items():
                # 规定提取
                SPEC_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")  # 规定-温度
                DELT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DELT")  # 规定-温度变化
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VFRAC")  # 规定-汽相分率
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-压力
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")  # 规定-负载
                PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PHASE")  # 规定-有效相态
                # UTILITY_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UTILITY_ID")  # 公用工程
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], SPEC_OPT_NODE, "SPEC_OPT")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], DELT_NODE, "DELT_VALUE", "DELT_UNITS")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], VFRAC_NODE, "VFRAC_VALUE")
                self.add_if_not_empty(RStoic_data["SPEC_DATA"], PHASE_NODE, "PHASE_VALUE")
                # self.add_if_not_empty(RStoic_data["SPEC_DATA"], UTILITY_ID_NODE, "UTILITY_ID")
                # 反应提取
                SERIES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SERIES")  # 反应-反应连续发生
                self.add_if_not_empty(RStoic_data["REAC_DATA"], SERIES, "SERIES")
                KEY_SSID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\KEY_SSID")  # 反应-反应编号
                CONV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CONV") # 反应-转化率
                KEY_CID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\KEY_CID")  # 反应-组分转化率
                OPT_EXT_CONV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_EXT_CONV")  # 反应-规范类型
                EXTENT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\EXTENT")  # 反应-摩尔反应进度
                COEF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COEF")  # 反应-化学计量-反应物
                COEF1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COEF1")  # 反应-化学计量-反应物
                for i, reac_data in enumerate(RStoic_data["REAC_DATA"]["REAC"]):
                    KEY_SSID_NODE.Elements.InsertRow(0, 0)
                    CONV_NODE.Elements.InsertRow(0, 0)
                    KEY_CID_NODE.Elements.InsertRow(0, 0)
                    OPT_EXT_CONV_NODE.Elements.InsertRow(0, 0)
                    EXTENT_NODE.Elements.InsertRow(0, 0)
                    COEF_NODE.Elements.InsertRow(0, 0)
                    COEF1_NODE.Elements.InsertRow(0, 0)
                    reac_id = reac_data["KEY_SSID"]
                    KEY_SSID_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    CONV_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    KEY_CID_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    OPT_EXT_CONV_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    EXTENT_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    COEF_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    COEF1_NODE.Elements.LabelNode(0, 0)[0].Value = reac_id
                    CONV = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CONV\{reac_id}")  # 反应-转化率
                    KEY_CID = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\KEY_CID\{reac_id}")  # 反应-组分转化率
                    OPT_EXT_CONV = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_EXT_CONV\{reac_id}")  # 反应-规范类型
                    EXTENT = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\EXTENT\{reac_id}")  # 反应-摩尔反应进度
                    self.add_if_not_empty(reac_data, CONV, "CONV")
                    self.add_if_not_empty(reac_data, KEY_CID, "KEY_CID")
                    self.add_if_not_empty(reac_data, OPT_EXT_CONV, "OPT_EXT_CONV")
                    self.add_if_not_empty(reac_data, EXTENT, "EXTENT")
                    COEF_MIX_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COEF\{reac_id}")  # 反应-化学计量-反应物
                    for cofe_mix, cofe_value in reac_data.get('COEF_DATA', {}).items():
                        COEF_MIX_NODE.Elements.InsertRow(0, 0)
                        COEF_MIX_NODE.Elements.LabelNode(0, 0)[0].Value = cofe_mix
                        COEF_MIX_NODE.Elements(0, 0).Value = cofe_value
                    COEF1_MIX_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COEF1\{reac_id}")  # 反应-化学计量-反应物
                    for cofe1_mix, cofe1_value in reac_data.get('COEF1_DATA', {}).items():
                        COEF1_MIX_NODE.Elements.InsertRow(0, 0)
                        COEF1_MIX_NODE.Elements.LabelNode(0, 0)[0].Value = cofe1_mix
                        COEF1_MIX_NODE.Elements(0, 0).Value = cofe1_value
            print(f"成功添加blocks_RStoic_data")
        except Exception as e:
            print(f"在添加blocks_RStoic_data时出错: {e}")
            raise
    def write_blocks_RPlug_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_RPlug_data配置写入Aspen模拟文件
        """
        try:
            for block, RPlug_data in config.get('blocks_RPlug_data', {}).items():
                # 添加规定
                TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TYPE")  # 规定-反应器类型
                OPT_TSPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_TSPEC")  # 规定-操作条件
                self.add_if_not_empty(RPlug_data["SPEC_DATA"], TYPE_NODE, "TYPE")
                self.add_if_not_empty(RPlug_data["SPEC_DATA"], OPT_TSPEC_NODE, "OPT_TSPEC")
                # 使用 .get() 方法安全访问 OPT_TSPEC，避免 KeyError
                opt_tspec = RPlug_data["SPEC_DATA"].get("OPT_TSPEC")
                if opt_tspec == "CONST-TEMP":
                    REAC_TEMP_NODE = self.aspen.Tree.FindNode(
                        fr"\Data\Blocks\{block}\Input\REAC_TEMP")  # 规定-反应器类型-操作条件-指定反应器温度
                    self.add_if_not_empty(RPlug_data["SPEC_DATA"], REAC_TEMP_NODE, "REAC_TEMP")
                if opt_tspec == "TEMP-PROF":
                    SPEC_TEMP_NODE = self.aspen.Tree.FindNode(
                        fr"\Data\Blocks\{block}\Input\SPEC_TEMP")  # 规定-反应器类型-操作条件-温度分布-温度
                    SPEC_TEMP_SUBNODES = self.get_child_nodes(
                        fr"\Data\Blocks\{block}\Input\SPEC_TEMP")  # 规定-反应器类型-操作条件-温度分布-温度
                    for i, SPEC_TEMP in enumerate(SPEC_TEMP_SUBNODES):
                        SPEC_TEMP_NODE.Elements.InsertRow(0, i)
                        SPEC_TEMP_NODE.Elements.Elements(i).SetValueAndUnit(
                            RPlug_data["SPEC_DATA"][SPEC_TEMP]["SPEC_TEMP_VALUE"],
                            self.convert_unitstr(RPlug_data["SPEC_DATA"][SPEC_TEMP]["SPEC_TEMP_UNITS"]))
                # 添加配置
                CHK_NTUBE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CHK_NTUBE")  # 配置-多管反应器
                NTUBE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NTUBE")  # 配置-多管反应器-管数
                LENGTH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LENGTH")  # 配置-反应器维度-长度
                DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DIAM")  # 配置-反应器维度-直径
                PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PHASE")  # 配置-有效相-工艺流股
                self.add_if_not_empty(RPlug_data["CONFIG_DATA"], CHK_NTUBE_NODE, "CHK_NTUBE")
                self.add_if_not_empty(RPlug_data["CONFIG_DATA"], LENGTH_NODE, "LENGTH")
                self.add_if_not_empty(RPlug_data["CONFIG_DATA"], DIAM_NODE, "DIAM")
                self.add_if_not_empty(RPlug_data["CONFIG_DATA"], PHASE_NODE, "PHASE")
                self.add_if_not_empty(RPlug_data["CONFIG_DATA"], NTUBE_NODE, "NTUBE")
                # 添加反应
                REACSYS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\REACSYS")  # 反应-反应体系
                self.add_if_not_empty(RPlug_data["REAC_DATA"], REACSYS_NODE, "REACSYS")
                RXN_ID_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RXN_ID")  # 反应-所选反应集
                for RXN_ID, RXN_ID_DATA in RPlug_data["REAC_DATA"].get('RXN_ID', {}).items():
                    RXN_ID_NODES.Elements.InsertRow(0, 0)
                    RXN_ID_NODES.Elements(0).Value = RXN_ID_DATA
                # 添加压力
                PRES_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 压力-进口压力
                OPT_PDROP_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_PDROP ")  # 压力-通过反应器的压降
                PDROP_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PDROP ")  # 压力-压降-工艺流股
                ROUGHNESS_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\ROUGHNESS ")  # 压力-摩擦关联式-粗糙度
                DP_FCOR_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DP_FCOR")  # 压力-摩擦关联式-压降关联式
                DP_MULT_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DP_MULT")  # 压力-摩擦关联式-压降比例因子
                self.add_if_not_empty(RPlug_data["PRES_DATA"], PRES_NODES, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(RPlug_data["PRES_DATA"], OPT_PDROP_NODES, "OPT_PDROP")
                self.add_if_not_empty(RPlug_data["PRES_DATA"], PDROP_NODES, "PDROP_VALUE", "PDROP_UNITS")
                self.add_if_not_empty(RPlug_data["PRES_DATA"], ROUGHNESS_NODES, "ROUGHNESS_VALUE", "ROUGHNESS_UNITS")
                self.add_if_not_empty(RPlug_data["PRES_DATA"], DP_FCOR_NODES, "DP_FCOR")
                self.add_if_not_empty(RPlug_data["PRES_DATA"], DP_MULT_NODES, "DP_MULT")
                # 添加催化剂
                CAT_PRESENT_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CAT_PRESENT")  # 催化剂-反应器内的催化剂
                IGN_CAT_VOL_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\IGN_CAT_VOL")  # 催化剂-忽略催化器体积
                BED_VOIDAGE_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BED_VOIDAGE")  # 催化剂-规定-床空隙率
                CAT_RHO_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CAT_RHO")  # 催化剂-规定-颗粒密度
                CATWT_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CATWT")  # 催化剂-规定-催化剂装填
                self.add_if_not_empty(RPlug_data["CAT_DATA"], CAT_PRESENT_NODES, "CAT_PRESENT")
                self.add_if_not_empty(RPlug_data["CAT_DATA"], IGN_CAT_VOL_NODES, "IGN_CAT_VOL")
                self.add_if_not_empty(RPlug_data["CAT_DATA"], BED_VOIDAGE_NODES, "BED_VOIDAGE")
                self.add_if_not_empty(RPlug_data["CAT_DATA"], CAT_RHO_NODES, "CAT_RHO_VALUE", "CAT_RHO_UNITS")
                self.add_if_not_empty(RPlug_data["CAT_DATA"], CATWT_NODES, "CATWT_VALUE", "CATWT_UNITS")
            print(f"成功添加blocks_RPlug_data_data")
        except Exception as e:
            print(f"在添加blocks_RPlug_data_data时出错: {e}")
            raise
    def write_blocks_Flash2_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Flash2_data配置写入Aspen模拟文件
        """
        try:
            for block, Flash2_data in config.get('blocks_Flash2_data', {}).items():
                SPEC_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")  # 规定-温度
                DELT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DELT")  # 规定-温度变化
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VFRAC")  # 规定-汽相分率
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-压力
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")  # 规定-负载
                # UTILITY_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UTILITY_ID")  # 公用工程(放规定一起)
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], DELT_NODE, "DELT_VALUE", "DELT_UNITS")
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], VFRAC_NODE, "VFRAC_VALUE")
                # self.add_if_not_empty(Flash2_data["SPEC_DATA"], UTILITY_ID_NODE, "UTILITY_ID")
                self.add_if_not_empty(Flash2_data["SPEC_DATA"], SPEC_OPT_NODE, "SPEC_OPT")
            print(f"成功添加blocks_Flash2_data")
        except Exception as e:
            print(f"在添加blocks_Flash2_data时出错: {e}")
            raise
    def write_blocks_Flash3_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Flash3_data配置写入Aspen模拟文件
        """
        try:
            for block, Flash3_data in config.get('blocks_Flash3_data', {}).items():
                SPEC_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")  # 规定-温度
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-压力
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")  # 规定-负载
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VFRAC")  # 规定-汽相分率
                L2_COMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L2_COMP")  # 规定-第二液相的关键组分
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], VFRAC_NODE, "VFRAC_VALUE")
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], SPEC_OPT_NODE, "SPEC_OPT")
                self.add_if_not_empty(Flash3_data["SPEC_DATA"], L2_COMP_NODE, "L2_COMP")
            print(f"成功添加blocks_Flash3_data")
        except Exception as e:
            print(f"在添加blocks_Flash3_data时出错: {e}")
            raise
    def write_blocks_Decanter_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Decanter_data配置写入Aspen模拟文件
        """
        try:
            for block, Decanter_data in config.get('blocks_Decanter_data', {}).items():
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")  # 规定-倾析器规范-温度
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")  # 规定-倾析器规范-压力
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")  # 规定-倾析器规范-负荷
                self.add_if_not_empty(Decanter_data["SPEC_DATA"], TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                self.add_if_not_empty(Decanter_data["SPEC_DATA"], PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                self.add_if_not_empty(Decanter_data["SPEC_DATA"], DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                L2_COMPS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L2_COMPS")
                L2_CUTOFF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L2_CUTOFF")  # 规定-第二液相的组分摩尔分率
                L2_COMPS = Decanter_data["SPEC_DATA"]["L2_COMPS"]
                for num, comps in enumerate(L2_COMPS):
                    L2_COMPS_NODE.Elements.InsertRow(0, num)
                    L2_COMPS_NODE.Elements(num).Value = comps
                self.add_if_not_empty(Decanter_data["SPEC_DATA"], L2_CUTOFF_NODE, "L2_CUTOFF")
            print(f"成功添加blocks_Decanter_data")
        except Exception as e:
            print(f"在添加blocks_Decanter_data时出错: {e}")
            raise
    def write_blocks_Sep_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Sep_data配置写入Aspen模拟文件
        """
        try:
            for block, Sep_data in config.get('blocks_Sep_data', {}).items():
                for FLOW, FLOW_DATA in Sep_data.get('SPEC_DATA', {}).items():
                    for i, COMP_DATA in enumerate(FLOW_DATA):
                        FLOWBASIS_NODE = self.aspen.Tree.FindNode(
                            fr"\Data\Blocks\{block}\Input\FLOWBASIS\{FLOW}\MIXED\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-基准
                        FRACS_NODE = self.aspen.Tree.FindNode(
                            fr"\Data\Blocks\{block}\Input\FRACS\{FLOW}\MIXED\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-规定-分流分率
                        FLOWS_NODE = self.aspen.Tree.FindNode(
                            fr"\Data\Blocks\{block}\Input\FLOWS\{FLOW}\MIXED\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-规定-流量
                        self.add_if_not_empty(COMP_DATA, FLOWBASIS_NODE, "FLOWBASIS_VALUE")
                        self.add_if_not_empty(COMP_DATA, FRACS_NODE, "FRACS")
                        self.add_if_not_empty(COMP_DATA, FLOWS_NODE, "FLOWS")
            print(f"成功添加blocks_Sep_data")
        except Exception as e:
            print(f"在添加blocks_Sep_data时出错: {e}")
            raise
    def write_blocks_Sep2_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Sep2_data配置写入Aspen模拟文件
        """
        try:
            for block, Sep2_data in config.get('blocks_Sep2_data', {}).items():
                for FLOW, FLOW_DATA in Sep2_data.get('SPEC_DATA', {}).items():
                    for i, COMP_DATA in enumerate(FLOW_DATA):
                        FLOWBASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FLOWBASIS\MIXED\{FLOW}\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-基准
                        FRACS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FRACS\MIXED\{FLOW}\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-规定-分流分率
                        FLOWS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FLOWS\MIXED\{FLOW}\{COMP_DATA['COMP_ID']}")  # 规定-出口流股条件-规定-流量
                        self.add_if_not_empty(COMP_DATA, FLOWBASIS_NODE, "FLOWBASIS_VALUE")
                        self.add_if_not_empty(COMP_DATA, FRACS_NODE, "FRACS")
                        self.add_if_not_empty(COMP_DATA, FLOWS_NODE, "FLOWS")
            print(f"成功添加blocks_Sep2_data")
        except Exception as e:
            print(f"在添加blocks_Sep2_data时出错: {e}")
            raise
    def write_blocks_RadFrac_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_RadFrac_data配置写入Aspen模拟文件
        """
        try:
            for block, RadFrac_data in config.get('blocks_RadFrac_data', {}).items():
                # 添加 Unit Set 参数
                UNIT_SET_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\Unit Set")
                if UNIT_SET_NODE:
                    UNIT_SET_NODE.Value = "MET"
                
                # 添加配置
                CALC_MODE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CALC_MODE")  # 配置-计算类型
                NSTAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSTAGE")  # 配置-塔板数
                CONDENSER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CONDENSER")  # 配置-冷凝器
                REBOILER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\REBOILER")  # 配置-再沸器
                NO_PHASE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NO_PHASE")  # 配置-有效相态
                BLKOPFREWAT = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BLKOPFREWAT")  # 配置-有效相态
                CONV_METH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CONV_METH")  # 配置-收敛
                BASIS_RR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_RR")  # 配置-操作规范-回流比
                RR_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RR_BASIS")  # 配置-操作规范-回流比
                BASIS_L1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_L1")  # 配置-操作规范-回流速率
                L1_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L1_BASIS")  # 配置-操作规范-回流速率
                BASIS_D_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_D")  # 配置-操作规范-馏出物流率
                D_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\D_BASIS")  # 配置-操作规范-馏出物流率
                BASIS_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_B")  # 配置-操作规范-塔底物流率
                B_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B_BASIS")  # 配置-操作规范-塔底物流率
                BASIS_VN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_VN")  # 配置-操作规范-再沸蒸汽流速
                VN_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VN_BASIS")  # 配置-操作规范-再沸蒸汽流速
                BASIS_BR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_BR")  # 配置-操作规范-再沸比
                BR_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BR_BASIS")  # 配置-操作规范-再沸比
                Q1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\Q1")  # 配置-操作规范-冷凝器负荷
                QN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\QN")  # 配置-操作规范-再沸器负荷
                DF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\D:F")  # 配置-操作规范-馏出物进料比
                DF_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\D:F_BASIS")  # 配置-操作规范-馏出物进料比-单位
                BF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B:F")  # 配置-操作规范-馏出物进料比
                BF_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B:F_BASIS")  # 配置-操作规范-馏出物进料比-单位
                # RW_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RW")  # 配置-自由水回流比
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], CALC_MODE_NODE, "CALC_MODE")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], NSTAGE_NODE, "NSTAGE")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], CONDENSER_NODE, "CONDENSER")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], REBOILER_NODE, "REBOILER")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], NO_PHASE, "NO_PHASE")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], BLKOPFREWAT, "BLKOPFREWAT")
                self.add_if_not_empty(RadFrac_data["CONFIG_DATA"], CONV_METH_NODE, "CONV_METH")
                for i, OP_SPEC_DATA in enumerate(RadFrac_data["CONFIG_DATA"]["OP_SPEC"]):
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_RR_NODE, "BASIS_RR_VALUE", None, "BASIS_RR_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, RR_BASIS_NODE, "BASIS_RR_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_L1_NODE, "BASIS_L1_VALUE", "BASIS_L1_UNITS","BASIS_L1_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, L1_BASIS_NODE, "BASIS_L1_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_D_NODE, "BASIS_D_VALUE", "BASIS_D_UNITS", "BASIS_D_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, D_BASIS_NODE, "BASIS_D_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_B_NODE, "BASIS_B_VALUE", "BASIS_B_UNITS", "BASIS_B_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, B_BASIS_NODE, "BASIS_B_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_VN_NODE, "BASIS_VN_VALUE", "BASIS_VN_UNITS","BASIS_VN_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, VN_BASIS_NODE, "BASIS_VN_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BASIS_BR_NODE, "BASIS_BR_VALUE", None, "BASIS_BR_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, DF_NODE, "DF_VALUE", None, "DF_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, DF_BASIS_NODE, "DF_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BF_NODE, "BF_VALUE", None, "BF_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BF_BASIS_NODE, "BF_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, BR_BASIS_NODE, "BASIS_BR_BASIS")
                    self.add_if_not_empty(OP_SPEC_DATA, Q1_NODE, "Q1_VALUE", "Q1_UNITS")
                    self.add_if_not_empty(OP_SPEC_DATA, QN_NODE, "QN_VALUE", "QN_UNITS")
                for i, FEED_DATA in enumerate(RadFrac_data["FEED_STAGE_DATA"]):
                    FEED_STAGE = FEED_DATA["FEED_STAGE"]
                    FEED_CONVEN_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_CONVEN\{FEED_STAGE}")  # 流股-进料流股-常规
                    FEED_STAGE_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_STAGE\{FEED_STAGE}")  # 流股-进料流股-塔板
                    FEED_CONVEN_NODES.Value = FEED_DATA["FEED_CONVEN"]
                    FEED_STAGE_NODES.Value = FEED_DATA["FEED_STAGE_VALUE"]
                for i, PROD_DATA in enumerate(RadFrac_data["PROD_STAGE_DATA"]):
                    PROD_STAGE = PROD_DATA["PROD_STAGE"]
                    PROD_PHASE_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_PHASE\{PROD_STAGE}")  # 流股-产品流股-相态
                    PROD_STAGE_NODES = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_STAGE\{PROD_STAGE}")  # 流股-产品流股-塔板
                    PROD_PHASE_NODES.Value = PROD_DATA["PROD_PHASE"]
                    PROD_STAGE_NODES.Value = PROD_DATA["PROD_STAGE_VALUE"]
                # 添加压力
                VIEW_PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VIEW_PRES")  # 压力-查看
                if RadFrac_data["PRES_DATA"]["VIEW_PRES"] == "TOP/BOTTOM": # 压力-查看-塔顶/塔底
                    VIEW_PRES_NODE.Value = "TOP/BOTTOM"
                    PRES1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES1")  # 压力-查看-塔板1压力
                    OPT_PRES_TOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_PRES_TOP")  # 压力-查看-塔板2压力-选项
                    PRES2_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES2")  # 压力-查看-塔板2压力
                    DP_COND_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DP_COND")  # 压力-查看-塔板2压力-冷凝器压降
                    OPT_PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_PRES")  # 压力-查看-塔其余部分压降
                    DP_STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DP_STAGE")  # 压力-查看-塔其余部分压降-塔板压降
                    DP_COL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DP_COL")  # 压力-查看-塔其余部分压降-塔压降
                    for i, STAGE_PRES_DATA in enumerate(RadFrac_data["PRES_DATA"]["STAGE_PRES"]):  # 压力-查看-塔其余部分压降-塔压降
                        self.add_if_not_empty(STAGE_PRES_DATA, PRES1_NODE, "PRES1_VALUE", "PRES1_UNITS")
                        self.add_if_not_empty(STAGE_PRES_DATA, OPT_PRES_TOP_NODE, "OPT_PRES_TOP")
                        self.add_if_not_empty(STAGE_PRES_DATA, PRES2_NODE, "PRES2_VALUE", "PRES2_UNITS")
                        self.add_if_not_empty(STAGE_PRES_DATA, DP_COND_NODE, "DP_COND_VALUE", "DP_COND_UNITS")
                        self.add_if_not_empty(STAGE_PRES_DATA, OPT_PRES_NODE, "OPT_PRES")
                        self.add_if_not_empty(STAGE_PRES_DATA, DP_STAGE_NODE, "DP_STAGE_VALUE", "DP_STAGE_UNITS")
                        self.add_if_not_empty(STAGE_PRES_DATA, DP_COL_NODE, "DP_COL_VALUE", "DP_COL_UNITS")
                if RadFrac_data["PRES_DATA"]["VIEW_PRES"] == "PROFILE":  # 压力-查看-压力分布
                    VIEW_PRES_NODE.Value = "PROFILE"
                    for i, STAGE_PRES_DATA in enumerate(RadFrac_data["PRES_DATA"]["STAGE_PRES"]):
                        PRES_STAGE = STAGE_PRES_DATA["PRES_STAGE"]
                        STAGE_PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\STAGE_PRES")
                        STAGE_PRES_NODE.Elements.InsertRow(0, 0)
                        STAGE_PRES_NODE.Elements.LabelNode(0, 0)[0].Value = PRES_STAGE
                        self.add_if_not_empty(STAGE_PRES_DATA, STAGE_PRES_NODE.Elements(0), "PRES_VALUE", "PRES_UNITS")
                    # if view_pres == "PDROP":  # 压力-查看-塔段压降  暂未实现
                # 添加冷凝器
                if "CONDENSER_DATA" in RadFrac_data:
                    OPT_COND_SPC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_COND_SPC")  # 冷凝器-冷凝器规范
                    T1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\T1")  # 冷凝器-冷凝器规范-温度
                    BASIS_RDV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_RDV")  # 冷凝器-冷凝器规范-馏出物汽相分率
                    SC_TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SC_TEMP")  # 冷凝器-冷凝器规范-过冷规范-过冷温度
                    SC_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SC_OPTION")  # 冷凝器-冷凝器规范
                    self.add_if_not_empty(RadFrac_data['CONDENSER_DATA'], OPT_COND_SPC_NODE, "OPT_COND_SPC")
                    self.add_if_not_empty(RadFrac_data['CONDENSER_DATA'], T1_NODE, "T1_VALUE", "T1_UNITS")
                    self.add_if_not_empty(RadFrac_data['CONDENSER_DATA'], BASIS_RDV_NODE, "BASIS_RDV_VALUE", None, "BASIS_RDV_BASIS")
                    self.add_if_not_empty(RadFrac_data['CONDENSER_DATA'], SC_TEMP_NODE, "SC_TEMP_VALUE", "SC_TEMP_UNITS")
                    self.add_if_not_empty(RadFrac_data['CONDENSER_DATA'], SC_OPTION_NODE, "SC_OPTION")
                # 添加设计规定
                if "DESIGN_SPEC_DATA" in RadFrac_data:
                    DESIGN_SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Subobjects\Design Specs")
                    base_node = fr"\Data\Blocks\{block}\Subobjects\Design Specs"
                    for design_spec_data in RadFrac_data["DESIGN_SPEC_DATA"]:
                        design_spec_id = design_spec_data["SPEC_ID"]
                        DESIGN_SPEC_NODE.Elements.Add(design_spec_id)
                        # 按照正确顺序查找和设置节点 - 使用 Subobjects 路径
                        SPEC_DESCRIP_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_DESCRIP\{design_spec_id}")
                        SPEC_TYPE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_TYPE\{design_spec_id}")
                        VALUE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\VALUE\{design_spec_id}")
                        OPT_SPC_STR_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_STR\{design_spec_id}")
                        SPEC_STAGE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_STAGE\{design_spec_id}")
                        SPEC_PHASE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_PHASE\{design_spec_id}")
                        SP_DEC_STRM_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SP_DEC_STRM\{design_spec_id}")
                        # [旧代码 - 使用 Input 路径] 注释掉
                        # SPEC_DESCRIP_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\SPEC_DESCRIP\{design_spec_id}")
                        # SPEC_TYPE_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\SPEC_TYPE\{design_spec_id}")
                        # VALUE_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\VALUE\{design_spec_id}")
                        # OPT_SPC_STR_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\OPT_SPC_STR\{design_spec_id}")
                        # SPEC_STAGE_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\SPEC_STAGE\{design_spec_id}")
                        # SPEC_PHASE_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\SPEC_PHASE\{design_spec_id}")
                        # SP_DEC_STRM_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\SP_DEC_STRM\{design_spec_id}")
                        # 按照正确顺序设置参数：1. SPEC_DESCRIP, 2. SPEC_TYPE, 3. VALUE, 4. OPT_SPC_STR, 5. SPEC_STAGE, 6. SPEC_PHASE, 7. SP_DEC_STRM
                        self.add_if_not_empty(design_spec_data, SPEC_DESCRIP_NODE, "SPEC_DESCRIP")
                        self.add_if_not_empty(design_spec_data, SPEC_TYPE_NODE, "SPEC_TYPE_VALUE")
                        self.add_if_not_empty(design_spec_data, VALUE_NODE, "SPEC_VALUE", "SPEC_VALUE_UNITS")
                        self.add_if_not_empty(design_spec_data, OPT_SPC_STR_NODE, "OPT_SPC_STR_VALUE")
                        self.add_if_not_empty(design_spec_data, SPEC_STAGE_NODE, "SPEC_STAGE")
                        self.add_if_not_empty(design_spec_data, SPEC_PHASE_NODE, "SPEC_PHASE")
                        self.add_if_not_empty(design_spec_data, SP_DEC_STRM_NODE, "SP_DEC_STRM")
                        COMPS_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_COMPS\{design_spec_id}")
                        for i, comp in enumerate(design_spec_data["COMP_DATA"]):
                            COMPS_NODE.Elements.InsertRow(0, 0)
                            COMPS_NODE.Elements(0, 0).Value = comp
                        SPEC_STREAMS_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_STREAMS\{design_spec_id}")
                        for i, spec_stream in enumerate(design_spec_data["SPEC_STREAMS"]):
                            SPEC_STREAMS_NODE.Elements.InsertRow(0, 0)
                            SPEC_STREAMS_NODE.Elements(0, 0).Value = spec_stream
                        
                        # 写入新增的设计规范单个值参数
                        OPT_SPC_RAT_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_RAT\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, OPT_SPC_RAT_NODE, "OPT_SPC_RAT")
                        
                        BASE_PHASE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\BASE_PHASE\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, BASE_PHASE_NODE, "BASE_PHASE")
                        
                        BASE_STAGE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\BASE_STAGE\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, BASE_STAGE_NODE, "BASE_STAGE")
                        
                        OPT_SPC_PRP1_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_PRP1\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, OPT_SPC_PRP1_NODE, "OPT_SPC_PRP1")
                        
                        PROPERTY_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\PROPERTY\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, PROPERTY_NODE, "PROPERTY")
                        
                        OPT_SPC_PRP2_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_PRP2\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, OPT_SPC_PRP2_NODE, "OPT_SPC_PRP2")
                        
                        BASE_PROPERT_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{design_spec_id}\Input\BASE_PROPERT\{design_spec_id}")
                        self.add_if_not_empty(design_spec_data, BASE_PROPERT_NODE, "BASE_PROPERT")
                # 添加设计变化
                if "VARY_DATA" in RadFrac_data:
                    VARY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Subobjects\Vary")
                    base_node = fr"\Data\Blocks\{block}\Subobjects\Vary"
                    for vary_data in RadFrac_data["VARY_DATA"]:
                        vary_id = vary_data["VARY_ID"]
                        VARY_NODE.Elements.Add(vary_id)
                        # 按照正确顺序查找和设置节点 - 使用 Subobjects 路径
                        VARY_DESCRIP_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{vary_id}\Input\VARY_DESCRIP\{vary_id}")
                        VARTYPE_NODE = self.aspen.Tree.FindNode(fr"{base_node}\{vary_id}\Input\VARTYPE\{vary_id}")
                        LB_NODE = self.aspen.Tree.FindNode(fr"{base_node}\{vary_id}\Input\LB\{vary_id}")
                        UB_NODE = self.aspen.Tree.FindNode(fr"{base_node}\{vary_id}\Input\UB\{vary_id}")
                        STEP_NODE = self.aspen.Tree.FindNode(fr"{base_node}\{vary_id}\Input\STEP\{vary_id}")
                        VALUE_NODE = self.aspen.Tree.FindNode(fr"{base_node}\{vary_id}\Input\VALUE\{vary_id}")
                        # [旧代码 - 使用 Input 路径] 注释掉
                        # VARY_DESCRIP_NODE = self.aspen.Tree.FindNode(
                        #     fr"\Data\Blocks\{block}\Input\VARY_DESCRIP\{vary_id}")
                        # VARTYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VARTYPE\{vary_id}")
                        # LB_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LB\{vary_id}")
                        # UB_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UB\{vary_id}")
                        # STEP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\STEP\{vary_id}")
                        # VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VALUE\{vary_id}")
                        # 按照正确顺序设置参数：1. VARY_DESCRIP, 2. VARTYPE, 3. LB, 4. UB, 5. STEP
                        self.add_if_not_empty(vary_data, VARY_DESCRIP_NODE, "VARY_DESCRIP")
                        self.add_if_not_empty(vary_data, VARTYPE_NODE, "VARTYPE_VALUE")
                        # 下限 / 上限：直接使用单位写入
                        self.add_if_not_empty(vary_data, LB_NODE, "LB_VALUE", "LB_UNITS")
                        self.add_if_not_empty(vary_data, UB_NODE, "UB_VALUE", "UB_UNITS")
                        self.add_if_not_empty(vary_data, STEP_NODE, "STEP_VALUE")
                        # VALUE 仍按原方式，使用单位写入
                        self.add_if_not_empty(vary_data, VALUE_NODE, "VARY_VALUE", "VARY_VALUE_UNITS")
                        if vary_data["COMP_DATA"] != []:
                            COMPS_NODE = self.aspen.Tree.FindNode(
                                fr"{base_node}\{vary_id}\Input\VARY_COMPS\{vary_id}")
                            for i, comp in enumerate(vary_data["COMP_DATA"]):
                                COMPS_NODE.Elements.InsertRow(0, 0)
                                COMPS_NODE.Elements(0, 0).Value = comp
                        
                        # 写入新增的变化单个值参数
                        VARY_STAGE_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE\{vary_id}")
                        self.add_if_not_empty(vary_data, VARY_STAGE_NODE, "VARY_STAGE")
                        
                        VARY_STREAM_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{vary_id}\Input\VARY_STREAM\{vary_id}")
                        self.add_if_not_empty(vary_data, VARY_STREAM_NODE, "VARY_STREAM")
                        
                        VARY_STAGE1_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE1\{vary_id}")
                        self.add_if_not_empty(vary_data, VARY_STAGE1_NODE, "VARY_STAGE1")
                        
                        VARY_STAGE2_NODE = self.aspen.Tree.FindNode(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE2\{vary_id}")
                        self.add_if_not_empty(vary_data, VARY_STAGE2_NODE, "VARY_STAGE2")
            print(f"成功添加blocks_RadFrac_data")
        except Exception as e:
            print(f"在添加blocks_RadFrac_data时出错: {e}")
            raise
    def write_blocks_DSTWU_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_DSTWU_data配置写入Aspen模拟文件
        DSTWU: Distillation-Shortcut Waton-Underwood (精馏快捷计算)
        """
        try:
            for block, DSTWU_data in config.get('blocks_DSTWU_data', {}).items():
                spec_data = DSTWU_data.get("SPEC_DATA", {})
                
                # 塔规范参数
                OPT_NTRR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_NTRR")  # 塔规范-选择RR或NSTAGE
                self.add_if_not_empty(spec_data, OPT_NTRR_NODE, "OPT_NTRR")
                
                # 根据OPT_NTRR的值选择设置RR或NSTAGE
                if "OPT_NTRR" in spec_data and spec_data["OPT_NTRR"] == "RR":
                    RR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RR")  # 塔规范-回流比
                    self.add_if_not_empty(spec_data, RR_NODE, "RR")
                elif "OPT_NTRR" in spec_data and spec_data["OPT_NTRR"] == "NSTAGE":
                    NSTAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSTAGE")  # 塔规范-塔板数
                    self.add_if_not_empty(spec_data, NSTAGE_NODE, "NSTAGE")
                
                # 压力
                PTOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PTOP")  # 压力-塔顶压力
                self.add_if_not_empty(spec_data, PTOP_NODE, "PTOP", "PTOP_UNITS")
                
                PBOT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PBOT")  # 压力-塔底压力
                self.add_if_not_empty(spec_data, PBOT_NODE, "PBOT", "PBOT_UNITS")
                
                # 冷凝器规范
                OPT_RDV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_RDV")  # 冷凝器规范-选择LIQUID/VAPOR/VAPLIQ
                self.add_if_not_empty(spec_data, OPT_RDV_NODE, "OPT_RDV")
                
                # 当OPT_RDV为VAPLIQ时才设置RDV
                if "OPT_RDV" in spec_data and spec_data["OPT_RDV"] == "VAPLIQ":
                    RDV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RDV")  # 冷凝器规范-汽相分率
                    self.add_if_not_empty(spec_data, RDV_NODE, "RDV")
                
                # 关键组分回收率
                LIGHTKEY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LIGHTKEY")  # 关键组分-轻关键组分
                self.add_if_not_empty(spec_data, LIGHTKEY_NODE, "LIGHTKEY")
                
                RECOVH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RECOVH")  # 关键组分-重关键组分回收率
                self.add_if_not_empty(spec_data, RECOVH_NODE, "RECOVH")
                
                HEAVYKEY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HEAVYKEY")  # 关键组分-重关键组分
                self.add_if_not_empty(spec_data, HEAVYKEY_NODE, "HEAVYKEY")
                
                RECOVL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RECOVL")  # 关键组分-轻关键组分回收率
                self.add_if_not_empty(spec_data, RECOVL_NODE, "RECOVL")
                
            print(f"成功添加blocks_DSTWU_data")
        except Exception as e:
            print(f"在添加blocks_DSTWU_data时出错: {e}")
            raise
    def write_blocks_Distl_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Distl_data配置写入Aspen模拟文件
        Distl: Distillation Column (精馏塔)
        """
        try:
            for block, Distl_data in config.get('blocks_Distl_data', {}).items():
                spec_data = Distl_data.get("SPEC_DATA", {})
                
                # 塔板数和进料位置
                NSTAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSTAGE")  # 塔板数
                self.add_if_not_empty(spec_data, NSTAGE_NODE, "NSTAGE")
                
                FEED_LOC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_LOC")  # 进料塔板数
                self.add_if_not_empty(spec_data, FEED_LOC_NODE, "FEED_LOC")
                
                # 回流比
                RR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RR")  # 回流比
                self.add_if_not_empty(spec_data, RR_NODE, "RR")
                
                # 馏出物与进料摩尔比
                D_F_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\D_F")  # 馏出物与进料摩尔比
                self.add_if_not_empty(spec_data, D_F_NODE, "D_F")
                
                # 冷凝器类型
                COND_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COND_TYPE")  # 冷凝器类型
                self.add_if_not_empty(spec_data, COND_TYPE_NODE, "COND_TYPE")
                
                # 压力
                PTOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PTOP")  # 冷凝器压力
                self.add_if_not_empty(spec_data, PTOP_NODE, "PTOP", "PTOP_UNITS")
                
                PBOT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PBOT")  # 再沸器压力
                self.add_if_not_empty(spec_data, PBOT_NODE, "PBOT", "PBOT_UNITS")
                
            print(f"成功添加blocks_Distl_data")
        except Exception as e:
            print(f"在添加blocks_Distl_data时出错: {e}")
            raise
    def write_blocks_Dupl_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Dupl_data配置写入Aspen模拟文件
        Dupl: Duplicate (复制/重复单元)
        """
        try:
            for block, Dupl_data in config.get('blocks_Dupl_data', {}).items():
                spec_data = Dupl_data.get("SPEC_DATA", {})
                
                # 物性方法集名称
                OPSETNAME_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPSETNAME")
                self.add_if_not_empty(spec_data, OPSETNAME_NODE, "OPSETNAME")
                
                # 化学计算
                CHEMISTRY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CHEMISTRY")
                self.add_if_not_empty(spec_data, CHEMISTRY_NODE, "CHEMISTRY")
                
                # 真实组分
                TRUE_COMPS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TRUE_COMPS")
                self.add_if_not_empty(spec_data, TRUE_COMPS_NODE, "TRUE_COMPS")
                
                # 自由水物性方法集
                FRWATEROPSET_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FRWATEROPSET")
                self.add_if_not_empty(spec_data, FRWATEROPSET_NODE, "FRWATEROPSET")
                
                # 可溶性水（整数，需要特殊处理）
                SOLU_WATER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SOLU_WATER")
                if "SOLU_WATER" in spec_data and spec_data["SOLU_WATER"] is not None:
                    SOLU_WATER_NODE.Value = int(spec_data["SOLU_WATER"])
                
                # Henry组分
                HENRY_COMPS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HENRY_COMPS")
                self.add_if_not_empty(spec_data, HENRY_COMPS_NODE, "HENRY_COMPS")
                
            print(f"成功添加blocks_Dupl_data")
        except Exception as e:
            print(f"在添加blocks_Dupl_data时出错: {e}")
            raise
    def write_blocks_Extract_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_Extract_data配置写入Aspen模拟文件
        Extract: Extraction Column (萃取塔)
        """
        try:
            for block, Extract_data in config.get('blocks_Extract_data', {}).items():
                spec_data = Extract_data.get("SPEC_DATA", {})
                
                # 1. 塔设定
                NSTAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSTAGE")  # 塔板数
                self.add_if_not_empty(spec_data, NSTAGE_NODE, "NSTAGE")
                
                OPT_THERMAL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_THERMAL")  # 热力学选项
                self.add_if_not_empty(spec_data, OPT_THERMAL_NODE, "OPT_THERMAL")
                
                # 根据 OPT_THERMAL 的值设置不同的参数
                if "OPT_THERMAL" in spec_data and spec_data["OPT_THERMAL"] == "TEMP":
                    # 设置 TSPEC_TEMP（动态塔板节点）
                    if "TSPEC_TEMP" in spec_data and spec_data["TSPEC_TEMP"]:
                        TSPEC_TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TSPEC_TEMP")
                        for stage_num, temp_data in spec_data["TSPEC_TEMP"].items():
                            # 创建动态节点（参考 RadFrac 的 STAGE_PRES 模式）
                            TSPEC_TEMP_NODE.Elements.InsertRow(0, 0)
                            TSPEC_TEMP_NODE.Elements.LabelNode(0, 0)[0].Value = stage_num
                            # 设置值和单位
                            self.add_if_not_empty(temp_data, TSPEC_TEMP_NODE.Elements(0), "TSPEC_TEMP_VALUE", "TSPEC_TEMP_UNITS")
                
                elif "OPT_THERMAL" in spec_data and spec_data["OPT_THERMAL"] == "DUTY":
                    # 设置 HEATER_DUTY（动态塔板节点）
                    if "HEATER_DUTY" in spec_data and spec_data["HEATER_DUTY"]:
                        HEATER_DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HEATER_DUTY")
                        for stage_num, duty_data in spec_data["HEATER_DUTY"].items():
                            # 创建动态节点（参考 RadFrac 的 STAGE_PRES 模式）
                            HEATER_DUTY_NODE.Elements.InsertRow(0, 0)
                            HEATER_DUTY_NODE.Elements.LabelNode(0, 0)[0].Value = stage_num
                            # 设置值和单位
                            self.add_if_not_empty(duty_data, HEATER_DUTY_NODE.Elements(0), "HEATER_DUTY_VALUE", "HEATER_DUTY_UNITS")
                
                # 2. 关键组分
                # 设置 COMP1_LIST（参考 Decanter 的 L2_COMPS 模式，不使用 LabelNode）
                if "COMP1_LIST" in spec_data and spec_data["COMP1_LIST"]:
                    COMP1_LIST_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMP1_LIST")
                    # 如果 COMP1_LIST 是字典格式（支持不连续索引）
                    if isinstance(spec_data["COMP1_LIST"], dict):
                        # 将字典转换为列表，按索引排序
                        sorted_items = sorted(spec_data["COMP1_LIST"].items(), key=lambda x: int(x[0].replace("#", "")) if x[0].replace("#", "").isdigit() else 0)
                        for num, (comp1_index, comp1_value) in enumerate(sorted_items):
                            if comp1_value is not None and comp1_value != "":
                                # 使用 InsertRow 创建节点（参考 Decanter 的 L2_COMPS 模式）
                                COMP1_LIST_NODE.Elements.InsertRow(0, num)
                                COMP1_LIST_NODE.Elements(num).Value = comp1_value
                    # 如果 COMP1_LIST 是数组格式（向后兼容）
                    elif isinstance(spec_data["COMP1_LIST"], list):
                        for num, comp1_value in enumerate(spec_data["COMP1_LIST"]):
                            if comp1_value is not None and comp1_value != "":
                                COMP1_LIST_NODE.Elements.InsertRow(0, num)
                                COMP1_LIST_NODE.Elements(num).Value = comp1_value
                
                # 设置 COMP2_LIST（参考 Decanter 的 L2_COMPS 模式，不使用 LabelNode）
                if "COMP2_LIST" in spec_data and spec_data["COMP2_LIST"]:
                    COMP2_LIST_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMP2_LIST")
                    # 如果 COMP2_LIST 是字典格式（支持不连续索引）
                    if isinstance(spec_data["COMP2_LIST"], dict):
                        # 将字典转换为列表，按索引排序
                        sorted_items = sorted(spec_data["COMP2_LIST"].items(), key=lambda x: int(x[0].replace("#", "")) if x[0].replace("#", "").isdigit() else 0)
                        for num, (comp2_index, comp2_value) in enumerate(sorted_items):
                            if comp2_value is not None and comp2_value != "":
                                # 使用 InsertRow 创建节点（参考 Decanter 的 L2_COMPS 模式）
                                COMP2_LIST_NODE.Elements.InsertRow(0, num)
                                COMP2_LIST_NODE.Elements(num).Value = comp2_value
                    # 如果 COMP2_LIST 是数组格式（向后兼容）
                    elif isinstance(spec_data["COMP2_LIST"], list):
                        for num, comp2_value in enumerate(spec_data["COMP2_LIST"]):
                            if comp2_value is not None and comp2_value != "":
                                COMP2_LIST_NODE.Elements.InsertRow(0, num)
                                COMP2_LIST_NODE.Elements(num).Value = comp2_value
                
                # 3. 压力
                # 设置 STAGE_PRES（动态塔板节点）
                if "STAGE_PRES" in spec_data and spec_data["STAGE_PRES"]:
                    STAGE_PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\STAGE_PRES")
                    for stage_num, pres_data in spec_data["STAGE_PRES"].items():
                        # 创建动态节点（参考 RadFrac 的 STAGE_PRES 模式）
                        STAGE_PRES_NODE.Elements.InsertRow(0, 0)
                        STAGE_PRES_NODE.Elements.LabelNode(0, 0)[0].Value = stage_num
                        # 设置值和单位
                        self.add_if_not_empty(pres_data, STAGE_PRES_NODE.Elements(0), "STAGE_PRES_VALUE", "STAGE_PRES_UNITS")
                
            print(f"成功添加blocks_Extract_data")
        except Exception as e:
            print(f"在添加blocks_Extract_data时出错: {e}")
            raise
    def write_blocks_FSplit_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_FSplit_data配置写入Aspen模拟文件
        FSplit: Flow Splitter (分流器)
        """
        try:
            for block, FSplit_data in config.get('blocks_FSplit_data', {}).items():
                spec_data = FSplit_data.get("SPEC_DATA", {})
                
                # 1. 参数列表：单位: 0 表示无单位，单位: 3 表示需要单位
                # (参数名, 值键名, 单位键名, 是否有单位)
                param_list = [
                    ("BASIS_C_LIM", "BASIS_C_LIM_VALUE", "BASIS_C_LIM_UNITS", True),  # 单位: 3
                    ("BASIS_FLOW", "BASIS_FLOW_VALUE", "BASIS_FLOW_UNITS", True),  # 单位: 3
                    ("BASIS_KEYNO", "BASIS_KEYNO_VALUE", None, False),  # 单位: 0
                    ("BASIS_LIMIT", "BASIS_LIMIT_VALUE", "BASIS_LIMIT_UNITS", True),  # 单位: 3
                    ("C_LIM_BASIS", "C_LIM_BASIS_VALUE", None, False),  # 单位: 0
                    ("DUTY", "DUTY_VALUE", "DUTY_UNITS", True),  # 单位: 3
                    ("FLOW_BASIS", "FLOW_BASIS_VALUE", None, False),  # 单位: 0
                    ("FRAC", "FRAC_VALUE", None, False),  # 单位: 0
                    ("LIMIT_BASIS", "LIMIT_BASIS_VALUE", None, False),  # 单位: 0
                    ("ORDER", "ORDER_VALUE", None, False),  # 单位: 0
                    ("POWER", "POWER_VALUE", "POWER_UNITS", True),  # 单位: 3
                    ("R_FRAC", "R_FRAC_VALUE", None, False),  # 单位: 0
                    ("VOL_C_LIM", "VOL_C_LIM_VALUE", "VOL_C_LIM_UNITS", True),  # 单位: 3
                    ("VOL_FLOW", "VOL_FLOW_VALUE", "VOL_FLOW_UNITS", True),  # 单位: 3
                    ("VOL_LIMIT", "VOL_LIMIT_VALUE", "VOL_LIMIT_UNITS", True),  # 单位: 3
                ]
                
                for param_name, value_key, units_key, has_units in param_list:
                    if param_name in spec_data and spec_data[param_name]:
                        # 遍历所有子节点（如 S1, S2, PRODUCT1 等）
                        for subnode, param_data in spec_data[param_name].items():
                            PARAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\{param_name}\{subnode}")
                            if has_units:
                                # 有单位的参数
                                self.add_if_not_empty(param_data, PARAM_NODE, value_key, units_key)
                            else:
                                # 无单位的参数
                                self.add_if_not_empty(param_data, PARAM_NODE, value_key)
                
                # 2. COMPS (无单位，只有值) - 最后添加
                # COMPS 结构：COMPS/1/MIXED/#0
                # 其中 1 是子节点（comp_subnode），MIXED 需要创建，#0 需要创建并赋值
                if "COMPS" in spec_data and spec_data["COMPS"]:
                    COMPS_BASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMPS")
                    for comp_subnode, comp_data in spec_data["COMPS"].items():
                        # 找到或获取 COMPS/comp_subnode 节点（应该已存在，由 BASIS_KEYNO 自动创建）
                        comp_subnode_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMPS\{comp_subnode}")
                        
                        if comp_subnode_node and "MIXED" in comp_data:
                            # 尝试找到 MIXED 节点，如果不存在则创建
                            MIXED_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMPS\{comp_subnode}\MIXED")
                            if not MIXED_NODE:
                                # 如果 MIXED 节点不存在，尝试使用 InsertRow 创建
                                try:
                                    # 使用 InsertRow 创建节点（参考 Extract 的 TSPEC_TEMP 模式）
                                    comp_subnode_node.Elements.InsertRow(0, 0)
                                    # 设置节点标签为 "MIXED"
                                    comp_subnode_node.Elements.LabelNode(0, 0)[0].Value = "MIXED"
                                    # 重新查找节点
                                    MIXED_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COMPS\{comp_subnode}\MIXED")
                                    if not MIXED_NODE:
                                        # 如果仍然找不到，尝试直接访问创建的元素
                                        MIXED_NODE = comp_subnode_node.Elements(0)
                                except Exception as e:
                                    print(f"创建 MIXED 节点失败: {e}")
                                    # 如果 InsertRow 也失败，可能需要先设置某个属性来触发节点创建
                                    continue
                            
                            # 处理 MIXED 下的叶子节点（#0, #1 等）
                            if MIXED_NODE and comp_data["MIXED"]:
                                # 将字典的键（如 "#0", "#1"）转换为数字索引
                                sorted_items = sorted(comp_data["MIXED"].items(), 
                                                    key=lambda x: int(x[0].replace("#", "")) if x[0].replace("#", "").isdigit() else 0)
                                
                                for num, (leaf_node_name, comp_value) in enumerate(sorted_items):
                                    if comp_value is not None and comp_value != "":
                                        try:
                                            # 使用 InsertRow 创建叶子节点（参考 Decanter 的 L2_COMPS 模式）
                                            MIXED_NODE.Elements.InsertRow(0, num)
                                            # 设置节点标签（如果需要）
                                            try:
                                                MIXED_NODE.Elements.LabelNode(0, num)[0].Value = leaf_node_name
                                            except:
                                                pass  # 如果 LabelNode 失败，继续尝试直接赋值
                                            # 赋值
                                            MIXED_NODE.Elements(num).Value = comp_value
                                        except Exception as e:
                                            print(f"创建或设置 COMPS/{comp_subnode}/MIXED/{leaf_node_name} 失败: {e}")
                                            continue
                
            print(f"成功添加blocks_FSplit_data")
        except Exception as e:
            print(f"在添加blocks_FSplit_data时出错: {e}")
            raise
    def write_blocks_HeatX_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_HeatX_data配置写入Aspen模拟文件
        HeatX: Heat Exchanger (换热器)
        """
        try:
            for block, HeatX_data in config.get('blocks_HeatX_data', {}).items():
                spec_data = HeatX_data.get("SPEC_DATA", {})
                
                # 按照指定顺序添加参数
                # 1. MODE (无单位)
                MODE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MODE")
                self.add_if_not_empty(spec_data, MODE_NODE, "MODE")
                
                # 2. HSHELL_TUBE (无单位)
                HSHELL_TUBE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HSHELL_TUBE")
                self.add_if_not_empty(spec_data, HSHELL_TUBE_NODE, "HSHELL_TUBE")
                
                # 3. TYPE (无单位)
                TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TYPE")
                self.add_if_not_empty(spec_data, TYPE_NODE, "TYPE")
                
                # 4. PROGRAM_MODE (无单位)
                PROGRAM_MODE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROGRAM_MODE")
                self.add_if_not_empty(spec_data, PROGRAM_MODE_NODE, "PROGRAM_MODE")
                
                # 5. SPEC (无单位)
                SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC")
                self.add_if_not_empty(spec_data, SPEC_NODE, "SPEC")
                
                # 6. VALUE (有单位)
                VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VALUE")
                self.add_if_not_empty(spec_data, VALUE_NODE, "VALUE_VALUE")
                
                # 7. AREA (有单位)
                AREA_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\AREA")
                self.add_if_not_empty(spec_data, AREA_NODE, "AREA_VALUE", "AREA_UNITS")
                
                # 8. UA (有单位)
                UA_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UA")
                self.add_if_not_empty(spec_data, UA_NODE, "UA_VALUE", "UA_UNITS")
                
                # 9. MIN_TAPP (有单位)
                MIN_TAPP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MIN_TAPP")
                self.add_if_not_empty(spec_data, MIN_TAPP_NODE, "MIN_TAPP_VALUE", "MIN_TAPP_UNITS")
                
                # 10. FT_MIN (无单位)
                FT_MIN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FT_MIN")
                self.add_if_not_empty(spec_data, FT_MIN_NODE, "FT_MIN")
                
                # 11. F_OPTION (无单位)
                F_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\F_OPTION")
                self.add_if_not_empty(spec_data, F_OPTION_NODE, "F_OPTION")
                
                # 12. LMTD_CORRECT (无单位)
                LMTD_CORRECT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LMTD_CORRECT")
                self.add_if_not_empty(spec_data, LMTD_CORRECT_NODE, "LMTD_CORRECT")
                
                # 13. SIDE_VAR (无单位)
                SIDE_VAR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SIDE_VAR")
                self.add_if_not_empty(spec_data, SIDE_VAR_NODE, "SIDE_VAR")
                
                # 14. CDP_OPTION (无单位)
                CDP_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CDP_OPTION")
                self.add_if_not_empty(spec_data, CDP_OPTION_NODE, "CDP_OPTION")
                
                # 15. PRES_COLD (有单位)
                PRES_COLD_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES_COLD")
                self.add_if_not_empty(spec_data, PRES_COLD_NODE, "PRES_COLD_VALUE", "PRES_COLD_UNITS")
                
                # 16. CMAX_DP (无单位)
                CMAX_DP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CMAX_DP")
                self.add_if_not_empty(spec_data, CMAX_DP_NODE, "CMAX_DP")
                
                # 17. CDP_SCALE (无单位)
                CDP_SCALE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CDP_SCALE")
                self.add_if_not_empty(spec_data, CDP_SCALE_NODE, "CDP_SCALE")
                
                # 18. TUBE_DP_FCOR (无单位)
                TUBE_DP_FCOR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_DP_FCOR")
                self.add_if_not_empty(spec_data, TUBE_DP_FCOR_NODE, "TUBE_DP_FCOR")
                
                # 19. TUBE_DP_HCOR (无单位)
                TUBE_DP_HCOR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_DP_HCOR")
                self.add_if_not_empty(spec_data, TUBE_DP_HCOR_NODE, "TUBE_DP_HCOR")
                
                # 20. TUBE_DP_PROF (无单位)
                TUBE_DP_PROF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_DP_PROF")
                self.add_if_not_empty(spec_data, TUBE_DP_PROF_NODE, "TUBE_DP_PROF")
                
                # 21. P_UPDATE (无单位)
                P_UPDATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\P_UPDATE")
                self.add_if_not_empty(spec_data, P_UPDATE_NODE, "P_UPDATE")
                
                # 22. U_OPTION (无单位)
                U_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U_OPTION")
                self.add_if_not_empty(spec_data, U_OPTION_NODE, "U_OPTION")
                
                # 23. U (有单位)
                U_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U")
                self.add_if_not_empty(spec_data, U_NODE, "U_VALUE", "U_UNITS")
                
                # 24. B_B (有单位)
                B_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B_B")
                self.add_if_not_empty(spec_data, B_B_NODE, "B_B_VALUE", "B_B_UNITS")
                
                # 25. B_L (有单位)
                B_L_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B_L")
                self.add_if_not_empty(spec_data, B_L_NODE, "B_L_VALUE", "B_L_UNITS")
                
                # 26. B_V (有单位)
                B_V_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\B_V")
                self.add_if_not_empty(spec_data, B_V_NODE, "B_V_VALUE", "B_V_UNITS")
                
                # 27. L_B (有单位)
                L_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L_B")
                self.add_if_not_empty(spec_data, L_B_NODE, "L_B_VALUE", "L_B_UNITS")
                
                # 28. L_L (有单位)
                L_L_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L_L")
                self.add_if_not_empty(spec_data, L_L_NODE, "L_L_VALUE", "L_L_UNITS")
                
                # 29. L_V (有单位)
                L_V_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\L_V")
                self.add_if_not_empty(spec_data, L_V_NODE, "L_V_VALUE", "L_V_UNITS")
                
                # 30. V_B (有单位)
                V_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\V_B")
                self.add_if_not_empty(spec_data, V_B_NODE, "V_B_VALUE", "V_B_UNITS")
                
                # 31. V_L (有单位)
                V_L_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\V_L")
                self.add_if_not_empty(spec_data, V_L_NODE, "V_L_VALUE", "V_L_UNITS")
                
                # 32. V_V (有单位)
                V_V_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\V_V")
                self.add_if_not_empty(spec_data, V_V_NODE, "V_V_VALUE", "V_V_UNITS")
                
                # 33. U_REF_SIDE (无单位)
                U_REF_SIDE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U_REF_SIDE")
                self.add_if_not_empty(spec_data, U_REF_SIDE_NODE, "U_REF_SIDE")
                
                # 34. UFLOW_BASIS (无单位)
                UFLOW_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\UFLOW_BASIS")
                self.add_if_not_empty(spec_data, UFLOW_BASIS_NODE, "UFLOW_BASIS")
                
                # 35. BASIS_UFLOW (有单位)
                BASIS_UFLOW_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_UFLOW")
                self.add_if_not_empty(spec_data, BASIS_UFLOW_NODE, "BASIS_UFLOW_VALUE", "BASIS_UFLOW_UNITS")
                
                # 36. U_REF_VALUE (有单位)
                U_REF_VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U_REF_VALUE")
                self.add_if_not_empty(spec_data, U_REF_VALUE_NODE, "U_REF_VALUE_VALUE", "U_REF_VALUE_UNITS")
                
                # 37. U_EXPONENT (无单位)
                U_EXPONENT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U_EXPONENT")
                self.add_if_not_empty(spec_data, U_EXPONENT_NODE, "U_EXPONENT")
                
                # 38. U_SCALE (无单位)
                U_SCALE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\U_SCALE")
                self.add_if_not_empty(spec_data, U_SCALE_NODE, "U_SCALE")
                
                # 39. CH_OPTION (无单位)
                CH_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_OPTION")
                self.add_if_not_empty(spec_data, CH_OPTION_NODE, "CH_OPTION")
                
                # 40. CH (有单位)
                CH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH")
                self.add_if_not_empty(spec_data, CH_NODE, "CH_VALUE", "CH_UNITS")
                
                # 41. CH_B (有单位)
                CH_B_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_B")
                self.add_if_not_empty(spec_data, CH_B_NODE, "CH_B_VALUE", "CH_B_UNITS")
                
                # 42. CH_L (有单位)
                CH_L_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_L")
                self.add_if_not_empty(spec_data, CH_L_NODE, "CH_L_VALUE", "CH_L_UNITS")
                
                # 43. CH_V (有单位)
                CH_V_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_V")
                self.add_if_not_empty(spec_data, CH_V_NODE, "CH_V_VALUE", "CH_V_UNITS")
                
                # 44. CHFLOW_BASIS (无单位)
                CHFLOW_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CHFLOW_BASIS")
                self.add_if_not_empty(spec_data, CHFLOW_BASIS_NODE, "CHFLOW_BASIS")
                
                # 45. CH_EXPONENT (无单位)
                CH_EXPONENT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_EXPONENT")
                self.add_if_not_empty(spec_data, CH_EXPONENT_NODE, "CH_EXPONENT")
                
                # 46. BASIS_CHFLOW (有单位)
                BASIS_CHFLOW_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BASIS_CHFLOW")
                self.add_if_not_empty(spec_data, BASIS_CHFLOW_NODE, "BASIS_CHFLOW_VALUE", "BASIS_CHFLOW_UNITS")
                
                # 47. CH_REF_VALUE (有单位)
                CH_REF_VALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CH_REF_VALUE")
                self.add_if_not_empty(spec_data, CH_REF_VALUE_NODE, "CH_REF_VALUE_VALUE", "CH_REF_VALUE_UNITS")
                
                # 48. TEMA_TYPE (无单位)
                TEMA_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMA_TYPE")
                self.add_if_not_empty(spec_data, TEMA_TYPE_NODE, "TEMA_TYPE")
                
                # 49. TUBE_NPASS (无单位)
                TUBE_NPASS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_NPASS")
                self.add_if_not_empty(spec_data, TUBE_NPASS_NODE, "TUBE_NPASS")
                
                # 50. ORIENTATION (无单位)
                ORIENTATION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\ORIENTATION")
                self.add_if_not_empty(spec_data, ORIENTATION_NODE, "ORIENTATION")
                
                # 51. NSEAL_STRIP (无单位)
                NSEAL_STRIP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSEAL_STRIP")
                self.add_if_not_empty(spec_data, NSEAL_STRIP_NODE, "NSEAL_STRIP")
                
                # 52. TUBE_FLOW (无单位)
                TUBE_FLOW_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_FLOW")
                self.add_if_not_empty(spec_data, TUBE_FLOW_NODE, "TUBE_FLOW")
                
                # 53. SHELL_BND_SP (有单位)
                SHELL_BND_SP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SHELL_BND_SP")
                self.add_if_not_empty(spec_data, SHELL_BND_SP_NODE, "SHELL_BND_SP_VALUE", "SHELL_BND_SP_UNITS")
                
                # 54. SHELL_DIAM (有单位)
                SHELL_DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SHELL_DIAM")
                self.add_if_not_empty(spec_data, SHELL_DIAM_NODE, "SHELL_DIAM_VALUE", "SHELL_DIAM_UNITS")
                
                # 55. SHELL_NPAR (无单位)
                SHELL_NPAR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SHELL_NPAR")
                self.add_if_not_empty(spec_data, SHELL_NPAR_NODE, "SHELL_NPAR")
                
                # 56. SHELL_NSER (无单位)
                SHELL_NSER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SHELL_NSER")
                self.add_if_not_empty(spec_data, SHELL_NSER_NODE, "SHELL_NSER")
                
                # 57. TUBE_TYPE (无单位)
                TUBE_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_TYPE")
                self.add_if_not_empty(spec_data, TUBE_TYPE_NODE, "TUBE_TYPE")
                
                # 58. TOTAL_NUMBER (无单位)
                TOTAL_NUMBER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TOTAL_NUMBER")
                self.add_if_not_empty(spec_data, TOTAL_NUMBER_NODE, "TOTAL_NUMBER")
                
                # 59. PATTERN (无单位)
                PATTERN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PATTERN")
                self.add_if_not_empty(spec_data, PATTERN_NODE, "PATTERN")
                
                # 60. MATERIAL (无单位)
                MATERIAL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MATERIAL")
                self.add_if_not_empty(spec_data, MATERIAL_NODE, "MATERIAL")
                
                # 61. LENGTH (有单位)
                LENGTH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LENGTH")
                self.add_if_not_empty(spec_data, LENGTH_NODE, "LENGTH_VALUE", "LENGTH_UNITS")
                
                # 62. PITCH (有单位)
                PITCH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PITCH")
                self.add_if_not_empty(spec_data, PITCH_NODE, "PITCH_VALUE", "PITCH_UNITS")
                
                # 63. TCOND (有单位)
                TCOND_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TCOND")
                self.add_if_not_empty(spec_data, TCOND_NODE, "TCOND_VALUE", "TCOND_UNITS")
                
                # 64. OUTSIDE_DIAM (有单位)
                OUTSIDE_DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OUTSIDE_DIAM")
                self.add_if_not_empty(spec_data, OUTSIDE_DIAM_NODE, "OUTSIDE_DIAM_VALUE", "OUTSIDE_DIAM_UNITS")
                
                # 65. WALL_THICK (有单位)
                WALL_THICK_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\WALL_THICK")
                self.add_if_not_empty(spec_data, WALL_THICK_NODE, "WALL_THICK_VALUE", "WALL_THICK_UNITS")
                
                # 66. OPT_FHEIGHT (无单位)
                OPT_FHEIGHT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_FHEIGHT")
                self.add_if_not_empty(spec_data, OPT_FHEIGHT_NODE, "OPT_FHEIGHT")
                
                # 67. HEIGHT (有单位)
                HEIGHT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HEIGHT")
                self.add_if_not_empty(spec_data, HEIGHT_NODE, "HEIGHT_VALUE", "HEIGHT_UNITS")
                
                # 68. ROOT_DIAM (有单位)
                ROOT_DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\ROOT_DIAM")
                self.add_if_not_empty(spec_data, ROOT_DIAM_NODE, "ROOT_DIAM_VALUE", "ROOT_DIAM_UNITS")
                
                # 69. OPT_FSPACING (无单位)
                OPT_FSPACING_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_FSPACING")
                self.add_if_not_empty(spec_data, OPT_FSPACING_NODE, "OPT_FSPACING")
                
                # 70. NPER_LENGTH (有单位)
                NPER_LENGTH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NPER_LENGTH")
                self.add_if_not_empty(spec_data, NPER_LENGTH_NODE, "NPER_LENGTH_VALUE", "NPER_LENGTH_UNITS")
                
                # 71. THICKNESS (有单位)
                THICKNESS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\THICKNESS")
                self.add_if_not_empty(spec_data, THICKNESS_NODE, "THICKNESS_VALUE", "THICKNESS_UNITS")
                
                # 72. AREA_RATIO (无单位)
                AREA_RATIO_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\AREA_RATIO")
                self.add_if_not_empty(spec_data, AREA_RATIO_NODE, "AREA_RATIO")
                
                # 73. EFFICIENCY (无单位)
                EFFICIENCY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\EFFICIENCY")
                self.add_if_not_empty(spec_data, EFFICIENCY_NODE, "EFFICIENCY")
                
                # 74. BAFFLE_TYPE (无单位)
                BAFFLE_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BAFFLE_TYPE")
                self.add_if_not_empty(spec_data, BAFFLE_TYPE_NODE, "BAFFLE_TYPE")
                
                # 75. NSEG_BAFFLE (无单位) - 只添加一次
                NSEG_BAFFLE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSEG_BAFFLE")
                self.add_if_not_empty(spec_data, NSEG_BAFFLE_NODE, "NSEG_BAFFLE")
                
                # 76. RING_INDIAM (有单位)
                RING_INDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RING_INDIAM")
                self.add_if_not_empty(spec_data, RING_INDIAM_NODE, "RING_INDIAM_VALUE", "RING_INDIAM_UNITS")
                
                # 77. RING_OUTDIAM (有单位)
                RING_OUTDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RING_OUTDIAM")
                self.add_if_not_empty(spec_data, RING_OUTDIAM_NODE, "RING_OUTDIAM_VALUE", "RING_OUTDIAM_UNITS")
                
                # 78. ROD_DIAM (有单位)
                ROD_DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\ROD_DIAM")
                self.add_if_not_empty(spec_data, ROD_DIAM_NODE, "ROD_DIAM_VALUE", "ROD_DIAM_UNITS")
                
                # 79. ROD_LENGTH (有单位)
                ROD_LENGTH_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\ROD_LENGTH")
                self.add_if_not_empty(spec_data, ROD_LENGTH_NODE, "ROD_LENGTH_VALUE", "ROD_LENGTH_UNITS")
                
                # 80. BAFFLE_CUT (无单位)
                BAFFLE_CUT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\BAFFLE_CUT")
                self.add_if_not_empty(spec_data, BAFFLE_CUT_NODE, "BAFFLE_CUT")
                
                # 81. IN_BFL_SP (有单位)
                IN_BFL_SP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\IN_BFL_SP")
                self.add_if_not_empty(spec_data, IN_BFL_SP_NODE, "IN_BFL_SP_VALUE", "IN_BFL_SP_UNITS")
                
                # 82. SHELL_BFL_SP (有单位)
                SHELL_BFL_SP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SHELL_BFL_SP")
                self.add_if_not_empty(spec_data, SHELL_BFL_SP_NODE, "SHELL_BFL_SP_VALUE", "SHELL_BFL_SP_UNITS")
                
                # 83. SMID_BFL_SP (有单位)
                SMID_BFL_SP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SMID_BFL_SP")
                self.add_if_not_empty(spec_data, SMID_BFL_SP_NODE, "SMID_BFL_SP_VALUE", "SMID_BFL_SP_UNITS")
                
                # 84. TUBES_IN_WIN (无单位)
                TUBES_IN_WIN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBES_IN_WIN")
                self.add_if_not_empty(spec_data, TUBES_IN_WIN_NODE, "TUBES_IN_WIN")
                
                # 85. TUBE_BFL_SP (有单位)
                TUBE_BFL_SP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TUBE_BFL_SP")
                self.add_if_not_empty(spec_data, TUBE_BFL_SP_NODE, "TUBE_BFL_SP_VALUE", "TUBE_BFL_SP_UNITS")
                
                # 86. SNOZ_INDIAM (有单位)
                SNOZ_INDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SNOZ_INDIAM")
                self.add_if_not_empty(spec_data, SNOZ_INDIAM_NODE, "SNOZ_INDIAM_VALUE", "SNOZ_INDIAM_UNITS")
                
                # 87. SNOZ_OUTDIAM (有单位)
                SNOZ_OUTDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SNOZ_OUTDIAM")
                self.add_if_not_empty(spec_data, SNOZ_OUTDIAM_NODE, "SNOZ_OUTDIAM_VALUE", "SNOZ_OUTDIAM_UNITS")
                
                # 88. TNOZ_INDIAM (有单位)
                TNOZ_INDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TNOZ_INDIAM")
                self.add_if_not_empty(spec_data, TNOZ_INDIAM_NODE, "TNOZ_INDIAM_VALUE", "TNOZ_INDIAM_UNITS")
                
                # 89. TNOZ_OUTDIAM (有单位)
                TNOZ_OUTDIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TNOZ_OUTDIAM")
                self.add_if_not_empty(spec_data, TNOZ_OUTDIAM_NODE, "TNOZ_OUTDIAM_VALUE", "TNOZ_OUTDIAM_UNITS")
                
                # 其他不在列表中的参数（放在最后）
                # NUM_SHELLS (无单位)
                NUM_SHELLS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NUM_SHELLS")
                self.add_if_not_empty(spec_data, NUM_SHELLS_NODE, "NUM_SHELLS")
                
                # SPECUN (无单位)
                SPECUN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPECUN")
                self.add_if_not_empty(spec_data, SPECUN_NODE, "SPECUN")
                
                # PRES_HOT (有单位)
                PRES_HOT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES_HOT")
                self.add_if_not_empty(spec_data, PRES_HOT_NODE, "PRES_HOT_VALUE", "PRES_HOT_UNITS")
                
                # SCUT_INTVLS (无单位)
                SCUT_INTVLS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SCUT_INTVLS")
                self.add_if_not_empty(spec_data, SCUT_INTVLS_NODE, "SCUT_INTVLS")
                
                # MIN_FLS_PTS (无单位)
                MIN_FLS_PTS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MIN_FLS_PTS")
                self.add_if_not_empty(spec_data, MIN_FLS_PTS_NODE, "MIN_FLS_PTS")
                
                # MAX_NSHELLS (无单位)
                MAX_NSHELLS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MAX_NSHELLS")
                self.add_if_not_empty(spec_data, MAX_NSHELLS_NODE, "MAX_NSHELLS")
                
                # MIN_HRC_PTS (无单位)
                MIN_HRC_PTS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MIN_HRC_PTS")
                self.add_if_not_empty(spec_data, MIN_HRC_PTS_NODE, "MIN_HRC_PTS")
                
                # HDP_OPTION (无单位)
                HDP_OPTION_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HDP_OPTION")
                self.add_if_not_empty(spec_data, HDP_OPTION_NODE, "HDP_OPTION")
                
                # HDP_SCALE (无单位)
                HDP_SCALE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HDP_SCALE")
                self.add_if_not_empty(spec_data, HDP_SCALE_NODE, "HDP_SCALE")
                
                # HMAX_DP (无单位)
                HMAX_DP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HMAX_DP")
                self.add_if_not_empty(spec_data, HMAX_DP_NODE, "HMAX_DP")
                
                # CDPPARM (无单位)
                CDPPARM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CDPPARM")
                self.add_if_not_empty(spec_data, CDPPARM_NODE, "CDPPARM")
                
                # HDPPARM (无单位)
                HDPPARM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HDPPARM")
                self.add_if_not_empty(spec_data, HDPPARM_NODE, "HDPPARM")
                
                # HDPPARMOP (无单位)
                HDPPARMOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HDPPARMOP")
                self.add_if_not_empty(spec_data, HDPPARMOP_NODE, "HDPPARMOP")
                
                # CDPPARMOP (无单位)
                CDPPARMOP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CDPPARMOP")
                self.add_if_not_empty(spec_data, CDPPARMOP_NODE, "CDPPARMOP")
                
            print(f"成功添加blocks_HeatX_data")
        except Exception as e:
            print(f"在添加blocks_HeatX_data时出错: {e}")
            raise
    def write_blocks_MCompr_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_MCompr_data配置写入Aspen模拟文件
        MCompr: Multi-Stage Compressor (多级压缩机)
        """
        try:
            for block, MCompr_data in config.get('blocks_MCompr_data', {}).items():
                spec_data = MCompr_data.get("SPEC_DATA", {})
                
                # 按照指定顺序添加参数
                # 1. NSTAGE (无单位)
                NSTAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NSTAGE")
                self.add_if_not_empty(spec_data, NSTAGE_NODE, "NSTAGE")
                
                # 2. PROD_STAGE (只设置子节点的值)
                if "PROD_STAGE" in spec_data and spec_data["PROD_STAGE"]:
                    PROD_STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_STAGE")
                    for prod_stage_data in spec_data["PROD_STAGE"]:
                        PROD_STAGE = prod_stage_data.get("PROD_STAGE")  # 动态流股名称
                        PROD_STREAM_VALUE = prod_stage_data.get("PROD_STREAM_VALUE")  # 子节点的值
                        
                        # 设置子节点的值
                        if PROD_STAGE and PROD_STREAM_VALUE:
                            # 先检查子节点是否已存在
                            STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_STAGE\{PROD_STAGE}")
                            if not STAGE_NODE:
                                # 节点不存在，创建子节点
                                row_count = PROD_STAGE_NODE.Elements.Count
                                PROD_STAGE_NODE.Elements.InsertRow(0, row_count)
                                PROD_STAGE_NODE.Elements.SetLabel(0, row_count, False, PROD_STAGE)
                            # 设置子节点的值
                            PROD_STREAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_STAGE\{PROD_STAGE}")
                            if PROD_STREAM_NODE:
                                PROD_STREAM_NODE.Value = PROD_STREAM_VALUE
                
                # 3. TYPE (无单位)
                TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TYPE")
                self.add_if_not_empty(spec_data, TYPE_NODE, "TYPE")
                
                # 4. OPT_SPEC (无单位)
                OPT_SPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_SPEC")
                self.add_if_not_empty(spec_data, OPT_SPEC_NODE, "OPT_SPEC")
                
                # 5. PRES (有单位，单位: 10)
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")
                self.add_if_not_empty(spec_data, PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                
                # 6. TYPE_STG (无单位)
                TYPE_STG_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TYPE_STG")
                self.add_if_not_empty(spec_data, TYPE_STG_NODE, "TYPE_STG")
                
                # 7. CALC_SPEED (无单位)
                CALC_SPEED_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CALC_SPEED")
                self.add_if_not_empty(spec_data, CALC_SPEED_NODE, "CALC_SPEED")
                
                # 8. GPSA_BASIS (无单位)
                GPSA_BASIS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\GPSA_BASIS")
                self.add_if_not_empty(spec_data, GPSA_BASIS_NODE, "GPSA_BASIS")
                
                # 9. CPR_METHOD (无单位)
                CPR_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CPR_METHOD")
                self.add_if_not_empty(spec_data, CPR_METHOD_NODE, "CPR_METHOD")
                
                # 10. FEED_STAGE (只设置子节点的值)
                if "FEED_STAGE" in spec_data and spec_data["FEED_STAGE"]:
                    FEED_STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_STAGE")
                    for feed_stage_data in spec_data["FEED_STAGE"]:
                        FEED_STAGE = feed_stage_data.get("FEED_STAGE")  # 动态流股名称
                        FEED_STREAM_VALUE = feed_stage_data.get("FEED_STREAM_VALUE")  # 子节点的值
                        
                        # 设置子节点的值
                        if FEED_STAGE and FEED_STREAM_VALUE:
                            # 先检查子节点是否已存在
                            STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_STAGE\{FEED_STAGE}")
                            if not STAGE_NODE:
                                # 节点不存在，创建子节点
                                row_count = FEED_STAGE_NODE.Elements.Count
                                FEED_STAGE_NODE.Elements.InsertRow(0, row_count)
                                FEED_STAGE_NODE.Elements.SetLabel(0, row_count, False, FEED_STAGE)
                            # 设置子节点的值
                            FEED_STREAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\FEED_STAGE\{FEED_STAGE}")
                            if FEED_STREAM_NODE:
                                FEED_STREAM_NODE.Value = FEED_STREAM_VALUE
                
                # 11. GLOBAL (只设置子节点的值)
                if "GLOBAL" in spec_data and spec_data["GLOBAL"]:
                    GLOBAL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\GLOBAL")
                    for global_name, global_data in spec_data["GLOBAL"].items():
                        PROD_STREAM_VALUE = global_data.get("PROD_STREAM_VALUE")  # 子节点的值
                        
                        # 设置子节点的值
                        if PROD_STREAM_VALUE:
                            # 先检查子节点是否已存在
                            STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\GLOBAL\{global_name}")
                            if not STAGE_NODE:
                                # 节点不存在，创建子节点
                                row_count = GLOBAL_NODE.Elements.Count
                                GLOBAL_NODE.Elements.InsertRow(0, row_count)
                                GLOBAL_NODE.Elements.SetLabel(0, row_count, False, global_name)
                            # 设置子节点的值
                            PROD_STREAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\GLOBAL\{global_name}")
                            if PROD_STREAM_NODE:
                                PROD_STREAM_NODE.Value = PROD_STREAM_VALUE
                
                # 12. PROD_PHASE (只设置子节点的值)
                if "PROD_PHASE" in spec_data and spec_data["PROD_PHASE"]:
                    PROD_PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_PHASE")
                    for prod_phase_data in spec_data["PROD_PHASE"]:
                        PROD_PHASE = prod_phase_data.get("PROD_PHASE")  # 动态流股名称
                        PROD_STREAM_VALUE = prod_phase_data.get("PROD_STREAM_VALUE")  # 子节点的值
                        
                        # 设置子节点的值
                        if PROD_PHASE and PROD_STREAM_VALUE:
                            # 先检查子节点是否已存在
                            STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_PHASE\{PROD_PHASE}")
                            if not STAGE_NODE:
                                # 节点不存在，创建子节点
                                row_count = PROD_PHASE_NODE.Elements.Count
                                PROD_PHASE_NODE.Elements.InsertRow(0, row_count)
                                PROD_PHASE_NODE.Elements.SetLabel(0, row_count, False, PROD_PHASE)
                            # 设置子节点的值
                            PROD_STREAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PROD_PHASE\{PROD_PHASE}")
                            if PROD_STREAM_NODE:
                                PROD_STREAM_NODE.Value = PROD_STREAM_VALUE
                
                # 13. TEMP (有单位，单位: 4)
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")
                self.add_if_not_empty(spec_data, TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                
                # 14-32. 按顺序添加带stage_num的参数（只需要在CLFR下创建节点，其他会自动生成）
                # 先收集所有需要的 stage_num 值（从所有相关参数中提取）
                stage_num_set = set()
                stage_param_names = ["CLFR", "CL_TEMP", "COOLER_UTL", "C_S_PRES", "DELP", "DUTY", "MEFF", 
                                     "OPT_CLFR", "OPT_CLSPEC", "OPT_CSPEC", "OPT_TEMP", "PDROP", "PEFF", 
                                     "POWER", "PRATIO", "SEFF", "SPECS_UTL", "TEMP", "TRATIO"]
                
                for param_name in stage_param_names:
                    if param_name in spec_data and spec_data[param_name]:
                        # 如果参数值是字典，提取所有的键（stage_num）
                        if isinstance(spec_data[param_name], dict):
                            stage_num_set.update(spec_data[param_name].keys())
                
                # 对于每个 stage_num，先在 CLFR 节点下创建节点
                CLFR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CLFR")
                if CLFR_NODE:
                    for stage_num in sorted(stage_num_set, key=lambda x: int(x) if x.isdigit() else 0):  # 排序确保顺序一致
                        STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CLFR\{stage_num}")
                        if not STAGE_NODE:
                            # 节点不存在，创建节点
                            row_count = CLFR_NODE.Elements.Count
                            CLFR_NODE.Elements.InsertRow(0, row_count)
                            CLFR_NODE.Elements.SetLabel(0, row_count, False, stage_num)
                
                # 然后按顺序处理所有参数，对每个 stage_num 都进行处理
                for stage_num in sorted(stage_num_set, key=lambda x: int(x) if x.isdigit() else 0):
                    # 14. CLFR\{stage_num} (无单位)
                    if "CLFR" in spec_data and spec_data["CLFR"] and stage_num in spec_data["CLFR"]:
                        STAGE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CLFR\{stage_num}")
                        if STAGE_NODE:
                            STAGE_NODE.Value = spec_data["CLFR"][stage_num]
                    
                    # 14. CL_TEMP\{stage_num} (有单位，单位: 4)
                    if "CL_TEMP" in spec_data and spec_data["CL_TEMP"] and stage_num in spec_data["CL_TEMP"]:
                        CL_TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CL_TEMP\{stage_num}")
                        if CL_TEMP_NODE:
                            cl_temp_data = spec_data["CL_TEMP"][stage_num]
                            self.add_if_not_empty(cl_temp_data, CL_TEMP_NODE, "CL_TEMP_VALUE", "CL_TEMP_UNITS")
                    
                    # 15. COOLER_UTL\{stage_num} (无单位)
                    if "COOLER_UTL" in spec_data and spec_data["COOLER_UTL"] and stage_num in spec_data["COOLER_UTL"]:
                        COOLER_UTL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\COOLER_UTL\{stage_num}")
                        if COOLER_UTL_NODE:
                            COOLER_UTL_NODE.Value = spec_data["COOLER_UTL"][stage_num]
                    
                    # 16. C_S_PRES\{stage_num} (有单位，单位: 10)
                    if "C_S_PRES" in spec_data and spec_data["C_S_PRES"] and stage_num in spec_data["C_S_PRES"]:
                        C_S_PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\C_S_PRES\{stage_num}")
                        if C_S_PRES_NODE:
                            c_s_pres_data = spec_data["C_S_PRES"][stage_num]
                            self.add_if_not_empty(c_s_pres_data, C_S_PRES_NODE, "C_S_PRES_VALUE", "C_S_PRES_UNITS")
                    
                    # 17. DELP\{stage_num} (有单位，单位: 10)
                    if "DELP" in spec_data and spec_data["DELP"] and stage_num in spec_data["DELP"]:
                        DELP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DELP\{stage_num}")
                        if DELP_NODE:
                            delp_data = spec_data["DELP"][stage_num]
                            self.add_if_not_empty(delp_data, DELP_NODE, "DELP_VALUE", "DELP_UNITS")
                    
                    # 18. DUTY\{stage_num} (有单位，单位: 18)
                    if "DUTY" in spec_data and spec_data["DUTY"] and stage_num in spec_data["DUTY"]:
                        DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY\{stage_num}")
                        if DUTY_NODE:
                            duty_data = spec_data["DUTY"][stage_num]
                            self.add_if_not_empty(duty_data, DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                    
                    # 19. MEFF\{stage_num} (无单位)
                    if "MEFF" in spec_data and spec_data["MEFF"] and stage_num in spec_data["MEFF"]:
                        MEFF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\MEFF\{stage_num}")
                        if MEFF_NODE:
                            MEFF_NODE.Value = spec_data["MEFF"][stage_num]
                    
                    # 20. OPT_CLFR\{stage_num} (无单位)
                    if "OPT_CLFR" in spec_data and spec_data["OPT_CLFR"] and stage_num in spec_data["OPT_CLFR"]:
                        OPT_CLFR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_CLFR\{stage_num}")
                        if OPT_CLFR_NODE:
                            OPT_CLFR_NODE.Value = spec_data["OPT_CLFR"][stage_num]
                    
                    # 21. OPT_CLSPEC\{stage_num} (无单位)
                    if "OPT_CLSPEC" in spec_data and spec_data["OPT_CLSPEC"] and stage_num in spec_data["OPT_CLSPEC"]:
                        OPT_CLSPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_CLSPEC\{stage_num}")
                        if OPT_CLSPEC_NODE:
                            OPT_CLSPEC_NODE.Value = spec_data["OPT_CLSPEC"][stage_num]
                    
                    # 22. OPT_CSPEC\{stage_num} (无单位)
                    if "OPT_CSPEC" in spec_data and spec_data["OPT_CSPEC"] and stage_num in spec_data["OPT_CSPEC"]:
                        OPT_CSPEC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_CSPEC\{stage_num}")
                        if OPT_CSPEC_NODE:
                            OPT_CSPEC_NODE.Value = spec_data["OPT_CSPEC"][stage_num]
                    
                    # 23. OPT_TEMP\{stage_num} (无单位)
                    if "OPT_TEMP" in spec_data and spec_data["OPT_TEMP"] and stage_num in spec_data["OPT_TEMP"]:
                        OPT_TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_TEMP\{stage_num}")
                        if OPT_TEMP_NODE:
                            OPT_TEMP_NODE.Value = spec_data["OPT_TEMP"][stage_num]
                    
                    # 24. PDROP\{stage_num} (有单位，单位: 10)
                    if "PDROP" in spec_data and spec_data["PDROP"] and stage_num in spec_data["PDROP"]:
                        PDROP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PDROP\{stage_num}")
                        if PDROP_NODE:
                            pdrop_data = spec_data["PDROP"][stage_num]
                            if isinstance(pdrop_data, dict):
                                self.add_if_not_empty(pdrop_data, PDROP_NODE, "PDROP_VALUE", "PDROP_UNITS")
                            else:
                                PDROP_NODE.Value = pdrop_data
                    
                    # 25. PEFF\{stage_num} (无单位)
                    if "PEFF" in spec_data and spec_data["PEFF"] and stage_num in spec_data["PEFF"]:
                        PEFF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PEFF\{stage_num}")
                        if PEFF_NODE:
                            PEFF_NODE.Value = spec_data["PEFF"][stage_num]
                    
                    # 26. POWER\{stage_num} (有单位，单位: 3)
                    if "POWER" in spec_data and spec_data["POWER"] and stage_num in spec_data["POWER"]:
                        POWER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\POWER\{stage_num}")
                        if POWER_NODE:
                            power_data = spec_data["POWER"][stage_num]
                            if isinstance(power_data, dict):
                                self.add_if_not_empty(power_data, POWER_NODE, "POWER_VALUE", "POWER_UNITS")
                            else:
                                POWER_NODE.Value = power_data
                    
                    # 27. PRATIO\{stage_num} (无单位)
                    if "PRATIO" in spec_data and spec_data["PRATIO"] and stage_num in spec_data["PRATIO"]:
                        PRATIO_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRATIO\{stage_num}")
                        if PRATIO_NODE:
                            PRATIO_NODE.Value = spec_data["PRATIO"][stage_num]
                    
                    # 28. SEFF\{stage_num} (无单位)
                    if "SEFF" in spec_data and spec_data["SEFF"] and stage_num in spec_data["SEFF"]:
                        SEFF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SEFF\{stage_num}")
                        if SEFF_NODE:
                            SEFF_NODE.Value = spec_data["SEFF"][stage_num]
                    
                    # 29. SPECS_UTL\{stage_num} (无单位)
                    if "SPECS_UTL" in spec_data and spec_data["SPECS_UTL"] and stage_num in spec_data["SPECS_UTL"]:
                        SPECS_UTL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPECS_UTL\{stage_num}")
                        if SPECS_UTL_NODE:
                            SPECS_UTL_NODE.Value = spec_data["SPECS_UTL"][stage_num]
                    
                    # 31. TEMP\{stage_num} (有单位，单位: 4)
                    if "TEMP" in spec_data and spec_data["TEMP"] and stage_num in spec_data["TEMP"]:
                        TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP\{stage_num}")
                        if TEMP_NODE:
                            temp_data = spec_data["TEMP"][stage_num]
                            self.add_if_not_empty(temp_data, TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                    
                    # 32. TRATIO\{stage_num} (无单位)
                    if "TRATIO" in spec_data and spec_data["TRATIO"] and stage_num in spec_data["TRATIO"]:
                        TRATIO_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TRATIO\{stage_num}")
                        if TRATIO_NODE:
                            TRATIO_NODE.Value = spec_data["TRATIO"][stage_num]
                
            print(f"成功添加blocks_MCompr_data")
        except Exception as e:
            print(f"在添加blocks_MCompr_data时出错: {e}")
            raise
    def write_blocks_RCSTR_data_to_aspen(self, config: Dict[str, Any]):
        """
        将blocks_RCSTR_data配置写入Aspen模拟文件
        RCSTR: Continuous Stirred-Tank Reactor (连续搅拌釜式反应器)
        """
        try:
            for block, RCSTR_data in config.get('blocks_RCSTR_data', {}).items():
                spec_data = RCSTR_data.get("SPEC_DATA", {})
                
                # 按照指定顺序添加参数
                # 1. HTRANMODE (无单位)
                HTRANMODE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\HTRANMODE")
                self.add_if_not_empty(spec_data, HTRANMODE_NODE, "HTRANMODE")
                
                # 2. PRES (有单位)
                PRES_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PRES")
                self.add_if_not_empty(spec_data, PRES_NODE, "PRES_VALUE", "PRES_UNITS")
                
                # 3. SPEC_OPT (无单位)
                SPEC_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_OPT")
                self.add_if_not_empty(spec_data, SPEC_OPT_NODE, "SPEC_OPT")
                
                # 4. NPHASE (无单位)
                NPHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\NPHASE")
                self.add_if_not_empty(spec_data, NPHASE_NODE, "NPHASE")
                
                # 5. TEMP (有单位)
                TEMP_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\TEMP")
                self.add_if_not_empty(spec_data, TEMP_NODE, "TEMP_VALUE", "TEMP_UNITS")
                
                # 6. DUTY (有单位)
                DUTY_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\DUTY")
                self.add_if_not_empty(spec_data, DUTY_NODE, "DUTY_VALUE", "DUTY_UNITS")
                
                # 7. VFRAC (无单位)
                VFRAC_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VFRAC")
                self.add_if_not_empty(spec_data, VFRAC_NODE, "VFRAC")
                
                # 8. SPEC_TYPE (无单位) - 移到 PHASE 之前，避免参数依赖问题
                SPEC_TYPE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_TYPE")
                self.add_if_not_empty(spec_data, SPEC_TYPE_NODE, "SPEC_TYPE")
                
                # 9. SPEC_PHASE (无单位)
                SPEC_PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SPEC_PHASE")
                self.add_if_not_empty(spec_data, SPEC_PHASE_NODE, "SPEC_PHASE")
                
                # 10. REACT_VOL (有单位)
                REACT_VOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\REACT_VOL")
                self.add_if_not_empty(spec_data, REACT_VOL_NODE, "REACT_VOL_VALUE", "REACT_VOL_UNITS")
                
                # 11. REACT_VOL_FR (无单位)
                REACT_VOL_FR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\REACT_VOL_FR")
                self.add_if_not_empty(spec_data, REACT_VOL_FR_NODE, "REACT_VOL_FR")
                
                # 12. PH_RES_TIME (有单位)
                PH_RES_TIME_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PH_RES_TIME")
                self.add_if_not_empty(spec_data, PH_RES_TIME_NODE, "PH_RES_TIME_VALUE", "PH_RES_TIME_UNITS")
                
                # 13. PHASE (无单位)
                PHASE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\PHASE")
                self.add_if_not_empty(spec_data, PHASE_NODE, "PHASE")
                
                # 14. VOL (有单位)
                VOL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\VOL")
                self.add_if_not_empty(spec_data, VOL_NODE, "VOL_VALUE", "VOL_UNITS")
                
                # 15. RES_TIME (有单位)
                RES_TIME_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RES_TIME")
                self.add_if_not_empty(spec_data, RES_TIME_NODE, "RES_TIME_VALUE", "RES_TIME_UNITS")
                
                # 16. CHK_MASSTR (无单位)
                CHK_MASSTR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CHK_MASSTR")
                self.add_if_not_empty(spec_data, CHK_MASSTR_NODE, "CHK_MASSTR")
                
                # 17. REACSYS (无单位)
                REACSYS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\REACSYS")
                self.add_if_not_empty(spec_data, REACSYS_NODE, "REACSYS")
                
                # 18. RXN_ID (动态节点列表，无单位)
                if "RXN_ID" in spec_data and spec_data["RXN_ID"]:
                    RXN_ID_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RXN_ID")
                    if RXN_ID_NODE:
                        for RXN_ID, RXN_ID_VALUE in spec_data["RXN_ID"].items():
                            # 检查节点是否已存在
                            EXISTING_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\RXN_ID\{RXN_ID}")
                            if not EXISTING_NODE:
                                # 节点不存在，创建节点（参考 RPlug 的方式）
                                RXN_ID_NODE.Elements.InsertRow(0, 0)
                                RXN_ID_NODE.Elements(0).Value = RXN_ID_VALUE
                            else:
                                # 节点已存在，直接设置值
                                EXISTING_NODE.Value = RXN_ID_VALUE
                
                # 19. SUBBYPASS (有单位)
                SUBBYPASS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SUBBYPASS")
                self.add_if_not_empty(spec_data, SUBBYPASS_NODE, "SUBBYPASS_VALUE", "SUBBYPASS_UNITS")
                
                # 20. CRYSTSYS (无单位)
                CRYSTSYS_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CRYSTSYS")
                self.add_if_not_empty(spec_data, CRYSTSYS_NODE, "CRYSTSYS")
                
                # 21. LOWER (有单位)
                LOWER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\LOWER")
                self.add_if_not_empty(spec_data, LOWER_NODE, "LOWER_VALUE", "LOWER_UNITS")
                
                # 22. SUB_RRSBN (有单位)
                SUB_RRSBN_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SUB_RRSBN")
                self.add_if_not_empty(spec_data, SUB_RRSBN_NODE, "SUB_RRSBN_VALUE", "SUB_RRSBN_UNITS")
                
                # 23. SUB_STDDEV (有单位)
                SUB_STDDEV_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\SUB_STDDEV")
                self.add_if_not_empty(spec_data, SUB_STDDEV_NODE, "SUB_STDDEV_VALUE", "SUB_STDDEV_UNITS")
                
                # 24. S_OPT (有单位)
                S_OPT_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\S_OPT")
                self.add_if_not_empty(spec_data, S_OPT_NODE, "S_OPT_VALUE", "S_OPT_UNITS")
                
                # 25. USER_SLOWER (有单位)
                USER_SLOWER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\USER_SLOWER")
                self.add_if_not_empty(spec_data, USER_SLOWER_NODE, "USER_SLOWER_VALUE", "USER_SLOWER_UNITS")
                
                # 26. USER_SVALUE (有单位)
                USER_SVALUE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\USER_SVALUE")
                self.add_if_not_empty(spec_data, USER_SVALUE_NODE, "USER_SVALUE_VALUE", "USER_SVALUE_UNITS")
                
                # 27. AGITATOR (无单位)
                AGITATOR_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\AGITATOR")
                self.add_if_not_empty(spec_data, AGITATOR_NODE, "AGITATOR")
                
                # 28. AGITRATE (有单位)
                AGITRATE_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\AGITRATE")
                self.add_if_not_empty(spec_data, AGITRATE_NODE, "AGITRATE_VALUE", "AGITRATE_UNITS")
                
                # 29. IMPELLR_DIAM (有单位)
                IMPELLR_DIAM_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\IMPELLR_DIAM")
                self.add_if_not_empty(spec_data, IMPELLR_DIAM_NODE, "IMPELLR_DIAM_VALUE", "IMPELLR_DIAM_UNITS")
                
                # 30. POWERNUMBER (无单位)
                POWERNUMBER_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\POWERNUMBER")
                self.add_if_not_empty(spec_data, POWERNUMBER_NODE, "POWERNUMBER")
                
                # 31. OPT_PSD (无单位)
                OPT_PSD_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_PSD")
                self.add_if_not_empty(spec_data, OPT_PSD_NODE, "OPT_PSD")
                
                # 32. CONST_METHOD (无单位)
                CONST_METHOD_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\CONST_METHOD")
                self.add_if_not_empty(spec_data, CONST_METHOD_NODE, "CONST_METHOD")
                
                # 33. OPT_SUBPSD (无单位)
                OPT_SUBPSD_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_SUBPSD")
                self.add_if_not_empty(spec_data, OPT_SUBPSD_NODE, "OPT_SUBPSD")
                
                # 34. OPT_OVERALL (无单位)
                OPT_OVERALL_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block}\Input\OPT_OVERALL")
                self.add_if_not_empty(spec_data, OPT_OVERALL_NODE, "OPT_OVERALL")
                
            print(f"成功添加blocks_RCSTR_data")
        except Exception as e:
            print(f"在添加blocks_RCSTR_data时出错: {e}")
            raise

    def run_simulation(self):
        """运行模拟并保存结果到CSV文件"""
        # 运行模拟
        try:
            print("开始运行模拟...")
            self.aspen.Engine.Run2()
            print("模拟运行完成")
        except Exception as e:
            print(f"模拟运行失败: {e}")


    def check_convergence(self):
        """检查模拟是否收敛"""
        try:
            # 收敛节点待调测
            # Conv_node = self.aspen.Tree.FindNode(fr"\Data\Results Summary\Conv-Sum\TEAR-SUMMARY\Output") #收敛结果节点
            # CVSTAT_node = self.aspen.Tree.FindNode(fr"\Data\Results Summary\Conv-Sum\TEAR-SUMMARY\Output\Output\CVSTAT") #结果-收敛状态
            # BLK_node = self.aspen.Tree.FindNode("\Data\Results Summary\Run-Status\Output\BLKSTAT")
            # convstat_node = self.aspen.Tree.FindNode("\Data\Convergence\Convergence\$OLVER01\Output\BLKSTAT") #收敛-收敛状态
            # self.aspen.Tree.FindNode("\Data\Convergence\Convergence\$OLVER01\Output\ERR_TOL2\30")
            # self.aspen.Tree.FindNode("\Data\Convergence\Conv-Options\Input\WEG_MAXIT")
            # self.aspen.Tree.FindNode("\Data\Convergence\Conv-Options\Input\WEG_QMIN")
            # self.aspen.Tree.FindNode("\Data\Convergence\Conv-Options\Input\WEG_QMAX")
            # self.aspen.Tree.FindNode("\Data\Convergence\Conv-Options\Input\TEAR_METHOD")
            # 获取收敛状态
            conv_status_node = self.aspen.Tree.FindNode(r"\Data\Results Summary\Conv-Sum\Output\STREAMID\1")
            conv_status = conv_status_node.Value

            if conv_status == "RECYCLE":
                print("模拟已收敛")
                return True
            else:
                print(f"模拟未收敛，状态: {conv_status}")
                return False

        except Exception as e:
            print(f"检查收敛状态时出错: {e}")
            return False

    def get_all_simulation_results(self, excel_filename, config: Dict[str, Any]):
        # 使用传入的文件名
        # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # excel_filename = RESULT_DIR / f"aspen_result_export_{timestamp}.xlsx"
        
        # 确保目录存在
        result_dir = os.path.dirname(excel_filename)
        if not os.path.exists(result_dir):
            os.makedirs(result_dir, exist_ok=True)
            print(f"创建结果目录: {result_dir}")

        # 创建一个Excel写入器
        with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
            # 1. 首先处理流结果，保存到"Stream Summary"工作表
            table_node = self.aspen.Tree.FindNode(fr"\Data\Results Summary\Stream-Sum\Stream-Sum\Table")

            row_count = table_node.Elements.RowCount(0)
            col_count = table_node.Elements.RowCount(1)

            # 获取列名称
            col_names = []
            for j in range(col_count):
                try:
                    col_name = table_node.Elements.LabelNode(1, j)[0].Value
                    col_names.append(col_name)
                except:
                    col_names.append(f"Col_{j + 1}")

            # 准备数据
            rows_list = []
            row_names = []

            for i in range(row_count):
                try:
                    # 获取行名称
                    row_name = table_node.Elements.LabelNode(0, i)[0].Value
                    row_names.append(row_name)

                    # 获取行数据
                    row_data = {}
                    for j in range(col_count):
                        try:
                            cell_value = table_node.Elements(i, j).Value
                            row_data[col_names[j]] = cell_value if cell_value is not None else "N/A"
                        except:
                            row_data[col_names[j]] = "N/A"

                    rows_list.append(row_data)
                except Exception as e:
                    print(f"处理第 {i + 1} 行时出错: {e}")

            # 创建DataFrame并保存到工作表
            if rows_list:
                df_stream = pd.DataFrame(rows_list, index=row_names)
                df_stream.to_excel(writer, sheet_name='Stream Summary')

            # 2. 处理每个block的结果，为每个block创建单独的工作表
            for i, block in enumerate(config.get('blocks', [])):
                block_name = block['name']
                if block['type'] == "DSTWU":
                    # 收集DSTWU block的所有结果
                    block_results = {}
                    # 最小回流比
                    min_reflux = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MIN_REFLUX")
                    block_results['MIN_REFLUX'] = min_reflux
                    # 实际回流比
                    act_reflux = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\ACT_REFLUX")
                    block_results['ACT_REFLUX'] = act_reflux
                    # 最小塔板数
                    min_stages = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MIN_STAGES")
                    block_results['MIN_STAGES'] = min_stages
                    # 实际塔板数
                    act_stages = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\ACT_STAGES")
                    block_results['ACT_STAGES'] = act_stages
                    # 进料塔板
                    feed_locatn = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FEED_LOCATN")
                    block_results['FEED_LOCATN'] = feed_locatn
                    # 进料上方实际塔板数
                    rect_stages = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\RECT_STAGES")
                    block_results['RECT_STAGES'] = rect_stages
                    # 冷凝器热负荷
                    cond_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results['COND_DUTY'] = cond_duty
                    # 冷凝器热负荷单位
                    cond_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results['COND_DUTY_UNITS'] = cond_duty_units
                    # 再沸器热负荷
                    reb_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results['REB_DUTY'] = reb_duty
                    # 再沸器热负荷单位
                    reb_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results['REB_DUTY_UNITS'] = reb_duty_units
                    # 馏出物温度
                    distil_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\DISTIL_TEMP")
                    block_results['DISTIL_TEMP'] = distil_temp
                    # 馏出物温度单位
                    distil_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\DISTIL_TEMP")
                    block_results['DISTIL_TEMP_UNITS'] = distil_temp_units
                    # 塔底物温度
                    bottom_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP'] = bottom_temp
                    # 塔底物温度单位
                    bottom_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP_UNITS'] = bottom_temp_units
                    # 馏出物进料比率
                    dist_vs_feed = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\DIST_VS_FED")
                    block_results['DIST_VS_FEED'] = dist_vs_feed

                    # 将block结果转换为DataFrame
                    # 转换为列格式：参数名称作为一列，值作为另一列
                    df_block = pd.DataFrame(list(block_results.items()), columns=['Parameter', 'Value'])

                    # 保存到以block名称命名的工作表
                    # 确保工作表名称有效（Excel工作表名称有长度和字符限制）
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Distl":
                    # 收集Distl block的所有结果
                    block_results = {}
                    # 冷凝器负荷
                    cond_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results['COND_DUTY'] = cond_duty
                    # 冷凝器负荷单位
                    cond_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results['COND_DUTY_UNITS'] = cond_duty_units
                    # 再沸器负荷
                    reb_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results['REB_DUTY'] = reb_duty
                    # 再沸器负荷单位
                    reb_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results['REB_DUTY_UNITS'] = reb_duty_units
                    # 进料塔板温度
                    feed_tray_t = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FEED_TRAY_T")
                    block_results['FEED_TRAY_T'] = feed_tray_t
                    # 进料塔板温度单位
                    feed_tray_t_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\FEED_TRAY_T")
                    block_results['FEED_TRAY_T_UNITS'] = feed_tray_t_units
                    # 顶端塔板温度
                    top_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results['TOP_TEMP'] = top_temp
                    # 顶端塔板温度单位
                    top_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results['TOP_TEMP_UNITS'] = top_temp_units
                    # 底端塔板温度
                    bottom_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP'] = bottom_temp
                    # 底端塔板温度单位
                    bottom_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP_UNITS'] = bottom_temp_units
                    # 进料质量
                    feed_quality = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FEED_QUALITY")
                    block_results['FEED_QUALITY'] = feed_quality

                    # 将block结果转换为DataFrame
                    # 转换为列格式：参数名称作为一列，值作为另一列
                    df_block = pd.DataFrame(list(block_results.items()), columns=['Parameter', 'Value'])

                    # 保存到以block名称命名的工作表
                    # 确保工作表名称有效（Excel工作表名称有长度和字符限制）
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Extract":
                    # 收集Extract block的所有结果
                    block_results = {}
                    # 顶端塔板温度
                    top_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results['TOP_TEMP'] = top_temp
                    # 顶端塔板温度单位
                    top_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results['TOP_TEMP_UNITS'] = top_temp_units
                    # 顶端塔板第一液相流量
                    top_l1flow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOP_L1FLOW")
                    block_results['TOP_L1FLOW'] = top_l1flow
                    # 顶端塔板第一液相流量单位
                    top_l1flow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOP_L1FLOW")
                    block_results['TOP_L1FLOW_UNITS'] = top_l1flow_units
                    # 顶端塔板第二液相流量
                    top_l2flow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOP_L2FLOW")
                    block_results['TOP_L2FLOW'] = top_l2flow
                    # 顶端塔板第二液相流量单位
                    top_l2flow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOP_L2FLOW")
                    block_results['TOP_L2FLOW_UNITS'] = top_l2flow_units
                    # 底端塔板温度
                    bottom_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP'] = bottom_temp
                    # 底端塔板温度单位
                    bottom_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results['BOTTOM_TEMP_UNITS'] = bottom_temp_units
                    # 底端塔板第一液相流量
                    bot_l1flow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOT_L1FLOW")
                    block_results['BOT_L1FLOW'] = bot_l1flow
                    # 底端塔板第一液相流量单位
                    bot_l1flow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOT_L1FLOW")
                    block_results['BOT_L1FLOW_UNITS'] = bot_l1flow_units
                    # 底端塔板第二液相流量
                    bot_l2flow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOT_L2FLOW")
                    block_results['BOT_L2FLOW'] = bot_l2flow
                    # 底端塔板第二液相流量单位
                    bot_l2flow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOT_L2FLOW")
                    block_results['BOT_L2FLOW_UNITS'] = bot_l2flow_units

                    # 将block结果转换为DataFrame
                    # 转换为列格式：参数名称作为一列，值作为另一列
                    df_block = pd.DataFrame(list(block_results.items()), columns=['Parameter', 'Value'])

                    # 保存到以block名称命名的工作表
                    # 确保工作表名称有效（Excel工作表名称有长度和字符限制）
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "FSplit":
                    # 收集FSplit block的所有结果
                    block_results = {}
                    
                    # 动态获取输出流股列表
                    output_streams = []
                    try:
                        # 首先尝试从Aspen Plus树结构中获取
                        ports_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block_name}\Ports\P(OUT)")
                        if ports_node and ports_node.Elements.Count > 0:
                            output_streams = [child.Name for child in ports_node.Elements]
                    except Exception as e:
                        print(f"从Ports节点获取FSplit设备 {block_name} 的输出流股时出错: {e}")
                    
                    # 如果无法从Ports获取，尝试从配置中获取
                    if not output_streams:
                        try:
                            if 'block_connections' in config:
                                block_conns = config['block_connections'].get(block_name, {})
                                # 查找P(OUT)端口的流股
                                for stream, port_type in block_conns.items():
                                    if port_type == "P(OUT)":
                                        output_streams.append(stream)
                        except Exception as e:
                            print(f"从配置获取FSplit设备 {block_name} 的输出流股时出错: {e}")
                    
                    # 如果仍然没有找到输出流股，尝试从STREAMFRAC节点获取所有子节点
                    if not output_streams:
                        try:
                            streamfrac_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block_name}\Output\STREAMFRAC")
                            if streamfrac_node and streamfrac_node.Elements.Count > 0:
                                output_streams = [child.Name for child in streamfrac_node.Elements]
                        except Exception as e:
                            print(f"从STREAMFRAC节点获取FSplit设备 {block_name} 的输出流股时出错: {e}")
                    
                    # 如果还是没有找到，使用默认的PRODUCT1/2/3
                    if not output_streams:
                        output_streams = ["PRODUCT1", "PRODUCT2", "PRODUCT3"]
                        print(f"警告：无法获取FSplit设备 {block_name} 的输出流股，使用默认流股名称")
                    
                    # 按照顺序提取每个输出流股的STREAMFRAC和STREAM_ORDER
                    for stream_name in output_streams:
                        # STREAMFRAC
                        try:
                            streamfrac_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block_name}\Output\STREAMFRAC\{stream_name}")
                            if streamfrac_node:
                                streamfrac_value = streamfrac_node.Value
                                block_results[f'STREAMFRAC_{stream_name}'] = streamfrac_value
                        except Exception as e:
                            print(f"获取STREAMFRAC_{stream_name}时出错: {e}")
                        
                        # STREAM_ORDER
                        try:
                            stream_order_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block_name}\Output\STREAM_ORDER\{stream_name}")
                            if stream_order_node:
                                stream_order_value = stream_order_node.Value
                                block_results[f'STREAM_ORDER_{stream_name}'] = stream_order_value
                        except Exception as e:
                            print(f"获取STREAM_ORDER_{stream_name}时出错: {e}")

                    # 将block结果转换为DataFrame
                    # 转换为列格式：参数名称作为一列，值作为另一列
                    df_block = pd.DataFrame(list(block_results.items()), columns=['Parameter', 'Value'])

                    # 保存到以block名称命名的工作表
                    # 确保工作表名称有效（Excel工作表名称有长度和字符限制）
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Flash3":
                    # 收集Flash3 block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 气相分率（摩尔）
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 气相分率（质量）
                    mvfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MVFRAC")
                    block_results["MVFRAC"] = mvfrac
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 压降
                    pdrop = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP"] = pdrop
                    pdrop_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP_UNITS"] = pdrop_units
                    # 相平衡（动态子节点，无单位）
                    eq_params = ["F", "X1", "X2", "Y", "K1", "K2"]
                    for eq_param in eq_params:
                        subnodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\{eq_param}")
                        for subnode in subnodes:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\{eq_param}\{subnode}")
                            block_results[f"{eq_param}_{subnode}"] = value

                    # 将block结果转换为DataFrame
                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "MCompr":
                    # 收集MCompr block的所有结果
                    block_results = {}
                    # 出口压力
                    b_pres2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES2")
                    block_results["B_PRES2"] = b_pres2
                    b_pres2_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES2")
                    block_results["B_PRES2_UNITS"] = b_pres2_units
                    # 总功
                    qcalc2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC2")
                    block_results["QCALC2"] = qcalc2
                    qcalc2_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC2")
                    block_results["QCALC2_UNITS"] = qcalc2_units
                    # 总冷却负荷
                    duty_out = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\DUTY_OUT")
                    block_results["DUTY_OUT"] = duty_out
                    duty_out_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\DUTY_OUT")
                    block_results["DUTY_OUT_UNITS"] = duty_out_units
                    # 净功要求
                    wnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET"] = wnet
                    wnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET_UNITS"] = wnet_units
                    # 净冷却负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units

                    # 分布（动态编号）
                    idx_nodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    if not idx_nodes:
                        idx_nodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    idx_list = sorted(idx_nodes, key=lambda x: int(x) if str(x).isdigit() else 0)

                    for idx in idx_list:
                        # B_TEMP\{idx}
                        b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP\{idx}")
                        block_results[f"B_TEMP_{idx}"] = b_temp
                        b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP\{idx}")
                        block_results[f"B_TEMP_{idx}_UNITS"] = b_temp_units
                        # B_PRES\{idx}
                        b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES\{idx}")
                        block_results[f"B_PRES_{idx}"] = b_pres
                        b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES\{idx}")
                        block_results[f"B_PRES_{idx}_UNITS"] = b_pres_units
                        # PRES_RATIO\{idx}（无单位）
                        pres_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PRES_RATIO\{idx}")
                        block_results[f"PRES_RATIO_{idx}"] = pres_ratio
                        # IND_POWER\{idx}
                        ind_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\IND_POWER\{idx}")
                        block_results[f"IND_POWER_{idx}"] = ind_power
                        ind_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\IND_POWER\{idx}")
                        block_results[f"IND_POWER_{idx}_UNITS"] = ind_power_units
                        # BRAKE_POWER\{idx}
                        brake_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER\{idx}")
                        block_results[f"BRAKE_POWER_{idx}"] = brake_power
                        brake_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER\{idx}")
                        block_results[f"BRAKE_POWER_{idx}_UNITS"] = brake_power_units
                        # HEAD_CAL\{idx}
                        head_cal = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL\{idx}")
                        block_results[f"HEAD_CAL_{idx}"] = head_cal
                        head_cal_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL\{idx}")
                        block_results[f"HEAD_CAL_{idx}_UNITS"] = head_cal_units
                        # VFLOW\{idx}
                        vflow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VFLOW\{idx}")
                        block_results[f"VFLOW_{idx}"] = vflow
                        vflow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VFLOW\{idx}")
                        block_results[f"VFLOW_{idx}_UNITS"] = vflow_units

                    # 冷却器（动态编号）
                    cool_idx_nodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\COOL_TEMP")
                    if not cool_idx_nodes:
                        cool_idx_nodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\COOL_PRES")
                    cool_idx_list = sorted(cool_idx_nodes, key=lambda x: int(x) if str(x).isdigit() else 0)

                    for idx in cool_idx_list:
                        # COOL_TEMP\{idx}
                        cool_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COOL_TEMP\{idx}")
                        block_results[f"COOL_TEMP_{idx}"] = cool_temp
                        cool_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COOL_TEMP\{idx}")
                        block_results[f"COOL_TEMP_{idx}_UNITS"] = cool_temp_units
                        # COOL_PRES\{idx}
                        cool_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COOL_PRES\{idx}")
                        block_results[f"COOL_PRES_{idx}"] = cool_pres
                        cool_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COOL_PRES\{idx}")
                        block_results[f"COOL_PRES_{idx}_UNITS"] = cool_pres_units
                        # QCALC\{idx}
                        qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC\{idx}")
                        block_results[f"QCALC_{idx}"] = qcalc
                        qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC\{idx}")
                        block_results[f"QCALC_{idx}_UNITS"] = qcalc_units
                        # B_VFRAC\{idx}（无单位）
                        b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC\{idx}")
                        block_results[f"B_VFRAC_{idx}"] = b_vfrac

                    # 将block结果转换为DataFrame
                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "RCSTR":
                    # 收集RCSTR block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 出口汽相分率
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净热负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units
                    # 反应器体积
                    ot_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\OT_VOL")
                    block_results["OT_VOL"] = ot_vol
                    ot_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\OT_VOL")
                    block_results["OT_VOL_UNITS"] = ot_vol_units
                    # 汽相体积
                    vap_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VAP_VOL")
                    block_results["VAP_VOL"] = vap_vol
                    vap_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VAP_VOL")
                    block_results["VAP_VOL_UNITS"] = vap_vol_units
                    # 液相体积
                    liq_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_VOL")
                    block_results["LIQ_VOL"] = liq_vol
                    liq_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\LIQ_VOL")
                    block_results["LIQ_VOL_UNITS"] = liq_vol_units
                    # 第一液相体积
                    liq1_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ1_VOL")
                    block_results["LIQ1_VOL"] = liq1_vol
                    liq1_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\LIQ1_VOL")
                    block_results["LIQ1_VOL_UNITS"] = liq1_vol_units
                    # 盐相体积
                    salt_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\SALT_VOL")
                    block_results["SALT_VOL"] = salt_vol
                    salt_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\SALT_VOL")
                    block_results["SALT_VOL_UNITS"] = salt_vol_units
                    # 凝固相体积
                    cond_vol = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COND_VOL")
                    block_results["COND_VOL"] = cond_vol
                    cond_vol_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COND_VOL")
                    block_results["COND_VOL_UNITS"] = cond_vol_units
                    # 反应器停留时间
                    tot_res_time = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOT_RES_TIME")
                    block_results["TOT_RES_TIME"] = tot_res_time
                    tot_res_time_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOT_RES_TIME")
                    block_results["TOT_RES_TIME_UNITS"] = tot_res_time_units
                    # 汽相停留时间
                    vap_res_time = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VAP_RES_TIME")
                    block_results["VAP_RES_TIME"] = vap_res_time
                    vap_res_time_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VAP_RES_TIME")
                    block_results["VAP_RES_TIME_UNITS"] = vap_res_time_units
                    # 凝固相停留时间
                    cond_res_tim = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COND_RES_TIM")
                    block_results["COND_RES_TIM"] = cond_res_tim
                    cond_res_tim_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COND_RES_TIM")
                    block_results["COND_RES_TIM_UNITS"] = cond_res_tim_units

                    # 将block结果转换为DataFrame
                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Mixer":
                    # 收集Mixer block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 汽相分率
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 压降
                    pdrop = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP"] = pdrop
                    pdrop_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP_UNITS"] = pdrop_units

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Valve":
                    # 收集Valve block的所有结果
                    block_results = {}
                    # 阻塞状态
                    chok_stat = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CHOK_STAT")
                    block_results["CHOK_STAT"] = chok_stat
                    # 出口压力
                    p_out_out = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\P_OUT_OUT")
                    block_results["P_OUT_OUT"] = p_out_out
                    p_out_out_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\P_OUT_OUT")
                    block_results["P_OUT_OUT_UNITS"] = p_out_out_units
                    # 压降
                    valve_dp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VALVE_DP")
                    block_results["VALVE_DP"] = valve_dp
                    valve_dp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VALVE_DP")
                    block_results["VALVE_DP_UNITS"] = valve_dp_units
                    # 阻塞出口压力
                    choke_pout = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CHOKE_POUT")
                    block_results["CHOKE_POUT"] = choke_pout
                    choke_pout_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\CHOKE_POUT")
                    block_results["CHOKE_POUT_UNITS"] = choke_pout_units
                    # 出口温度
                    tcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TCALC")
                    block_results["TCALC"] = tcalc
                    tcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TCALC")
                    block_results["TCALC_UNITS"] = tcalc_units
                    # 出口汽相分率
                    vcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VCALC")
                    block_results["VCALC"] = vcalc
                    # 阀流量系数
                    flow_coef = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FLOW_COEF")
                    block_results["FLOW_COEF"] = flow_coef
                    # 阀开度%
                    valve_posn = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VALVE_POSN")
                    block_results["VALVE_POSN"] = valve_posn
                    # 汽蚀指数
                    cav_indx = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CAV_INDX")
                    block_results["CAV_INDX"] = cav_indx
                    # 压降比率因子
                    pdrop_fac2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP_FAC2")
                    block_results["PDROP_FAC2"] = pdrop_fac2
                    # 压力回复因子
                    prrec_fac2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PRREC_FAC2")
                    block_results["PRREC_FAC2"] = prrec_fac2
                    # 管部件几何因子
                    pipe_fit_fac2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PIPE_FIT_FAC2")
                    block_results["PIPE_FIT_FAC2"] = pipe_fit_fac2

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Compr":
                    # 收集Compr block的所有结果
                    block_results = {}
                    # 压缩机模型
                    comptype = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPTYPE")
                    block_results["COMPTYPE"] = comptype
                    # 相计算
                    kode = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\KODE")
                    block_results["KODE"] = kode
                    # 指示马力
                    ind_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\IND_POWER")
                    block_results["IND_POWER"] = ind_power
                    ind_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\IND_POWER")
                    block_results["IND_POWER_UNITS"] = ind_power_units
                    # 制动马力
                    brake_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER")
                    block_results["BRAKE_POWER"] = brake_power
                    brake_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER")
                    block_results["BRAKE_POWER_UNITS"] = brake_power_units
                    # 净功要求
                    wnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET"] = wnet
                    wnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET_UNITS"] = wnet_units
                    # 功率损耗
                    power_loss = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\POWER_LOSS")
                    block_results["POWER_LOSS"] = power_loss
                    power_loss_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\POWER_LOSS")
                    block_results["POWER_LOSS_UNITS"] = power_loss_units
                    # 效率
                    epc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EPC")
                    block_results["EPC"] = epc
                    # 机械效率
                    eff_mech = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EFF_MECH")
                    block_results["EFF_MECH"] = eff_mech
                    # 出口压力
                    poc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\POC")
                    block_results["POC"] = poc
                    poc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\POC")
                    block_results["POC_UNITS"] = poc_units
                    # 出口温度
                    toc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOC")
                    block_results["TOC"] = toc
                    toc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOC")
                    block_results["TOC_UNITS"] = toc_units
                    # 等熵出口温度
                    tos = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOS")
                    block_results["TOS"] = tos
                    tos_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOS")
                    block_results["TOS_UNITS"] = tos_units
                    # 汽相分率
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 位移
                    dis = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\DIS")
                    block_results["DIS"] = dis
                    # 体积效率
                    ev = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EV")
                    block_results["EV"] = ev
                    # 产生的压头
                    head_cal = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL")
                    block_results["HEAD_CAL"] = head_cal
                    head_cal_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL")
                    block_results["HEAD_CAL_UNITS"] = head_cal_units
                    # 等熵功率要求
                    power_isen = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\POWER_ISEN")
                    block_results["POWER_ISEN"] = power_isen
                    power_isen_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\POWER_ISEN")
                    block_results["POWER_ISEN_UNITS"] = power_isen_units
                    # 进口热容比
                    in_cpr = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\IN_CPR")
                    block_results["IN_CPR"] = in_cpr
                    # 体积流率
                    feed_vflow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FEED_VFLOW")
                    block_results["FEED_VFLOW"] = feed_vflow
                    feed_vflow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\FEED_VFLOW")
                    block_results["FEED_VFLOW_UNITS"] = feed_vflow_units
                    vflow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VFLOW")
                    block_results["VFLOW"] = vflow
                    vflow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VFLOW")
                    block_results["VFLOW_UNITS"] = vflow_units
                    # 压缩因子
                    z_in = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\Z_IN")
                    block_results["Z_IN"] = z_in
                    z_in_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\Z_IN")
                    block_results["Z_IN_UNITS"] = z_in_units
                    z_out = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\Z_OUT")
                    block_results["Z_OUT"] = z_out
                    z_out_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\Z_OUT")
                    block_results["Z_OUT_UNITS"] = z_out_units
                    # 平均体积指数
                    exp_v_isen = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EXP_V_ISEN")
                    block_results["EXP_V_ISEN"] = exp_v_isen
                    exp_v_act = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EXP_V_ACT")
                    block_results["EXP_V_ACT"] = exp_v_act
                    # 平均温度指数
                    exp_t_isen = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EXP_T_ISEN")
                    block_results["EXP_T_ISEN"] = exp_t_isen
                    exp_t_act = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EXP_T_ACT")
                    block_results["EXP_T_ACT"] = exp_t_act

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Heater":
                    # 收集Heater block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 汽相分率
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 压降关联式参数
                    cor_pdrp_fac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COR_PDRP_FAC")
                    block_results["COR_PDRP_FAC"] = cor_pdrp_fac
                    # 压降
                    pdrop = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP"] = pdrop
                    pdrop_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP_UNITS"] = pdrop_units
                    # 相平衡（动态子节点，无单位，按你给的 X/Y）
                    eq_params = ["F", "X", "Y", "B_K"]
                    for eq_param in eq_params:
                        subnodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\{eq_param}")
                        for subnode in subnodes:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\{eq_param}\{subnode}")
                            block_results[f"{eq_param}_{subnode}"] = value

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Pump":
                    # 收集Pump block的所有结果
                    block_results = {}
                    # 流体功率
                    fluid_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\FLUID_POWER")
                    block_results["FLUID_POWER"] = fluid_power
                    fluid_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\FLUID_POWER")
                    block_results["FLUID_POWER_UNITS"] = fluid_power_units
                    # 制动功率
                    brake_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER")
                    block_results["BRAKE_POWER"] = brake_power
                    brake_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BRAKE_POWER")
                    block_results["BRAKE_POWER_UNITS"] = brake_power_units
                    # 电
                    elec_power = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\ELEC_POWER")
                    block_results["ELEC_POWER"] = elec_power
                    elec_power_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\ELEC_POWER")
                    block_results["ELEC_POWER_UNITS"] = elec_power_units
                    # 体积流率
                    vflow = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\VFLOW")
                    block_results["VFLOW"] = vflow
                    vflow_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\VFLOW")
                    block_results["VFLOW_UNITS"] = vflow_units
                    # 压力变化
                    pdrp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDRP")
                    block_results["PDRP"] = pdrp
                    pdrp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDRP")
                    block_results["PDRP_UNITS"] = pdrp_units
                    # 可用NPSH
                    npsh_avail = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\NPSH-AVAIL")
                    block_results["NPSH-AVAIL"] = npsh_avail
                    npsh_avail_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\NPSH-AVAIL")
                    block_results["NPSH-AVAIL_UNITS"] = npsh_avail_units
                    # NPSH要求
                    npsh_req = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\NPSH_REQ")
                    block_results["NPSH_REQ"] = npsh_req
                    npsh_req_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\NPSH_REQ")
                    block_results["NPSH_REQ_UNITS"] = npsh_req_units
                    # 产生的压头
                    head_cal = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL")
                    block_results["HEAD_CAL"] = head_cal
                    head_cal_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HEAD_CAL")
                    block_results["HEAD_CAL_UNITS"] = head_cal_units
                    # 采用的泵效率
                    ceff = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CEFF")
                    block_results["CEFF"] = ceff
                    # 净功要求
                    wnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET"] = wnet
                    wnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\WNET")
                    block_results["WNET_UNITS"] = wnet_units
                    # 出口压力
                    poc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\POC")
                    block_results["POC"] = poc
                    poc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\POC")
                    block_results["POC_UNITS"] = poc_units
                    # 出口温度
                    toc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOC")
                    block_results["TOC"] = toc
                    toc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOC")
                    block_results["TOC_UNITS"] = toc_units

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "RStoic":
                    # 收集RStoic block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净热负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units
                    # 汽相分率
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 相平衡（动态子节点，无单位）
                    eq_params = ["F", "X", "Y", "K"]
                    for eq_param in eq_params:
                        subnodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\{eq_param}")
                        for subnode in subnodes:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\{eq_param}\{subnode}")
                            block_results[f"{eq_param}_{subnode}"] = value
                    # 反应（动态编号）
                    rxn_idx_nodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\RXNID")
                    rxn_idx_list = sorted(rxn_idx_nodes, key=lambda x: int(x) if str(x).isdigit() else 0)
                    for idx in rxn_idx_list:
                        rxnid = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\RXNID\{idx}")
                        block_results[f"RXNID_{idx}"] = rxnid
                        extent_out = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\EXTENT_OUT\{idx}")
                        block_results[f"EXTENT_OUT_{idx}"] = extent_out
                        extent_out_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\EXTENT_OUT\{idx}")
                        block_results[f"EXTENT_OUT_{idx}_UNITS"] = extent_out_units
                        refid = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\REFID\{idx}")
                        block_results[f"REFID_{idx}"] = refid
                        reac_stoi = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\REAC_STOI\{idx}")
                        block_results[f"REAC_STOI_{idx}"] = reac_stoi

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "RPlug":
                    # 收集RPlug block的所有结果
                    block_results = {}
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 最小
                    tmin = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TMIN")
                    block_results["TMIN"] = tmin
                    tmin_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TMIN")
                    block_results["TMIN_UNITS"] = tmin_units
                    # 最大
                    tmax = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TMAX")
                    block_results["TMAX"] = tmax
                    tmax_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TMAX")
                    block_results["TMAX_UNITS"] = tmax_units
                    # 停留时间
                    res_time = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\RES_TIME")
                    block_results["RES_TIME"] = res_time
                    res_time_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\RES_TIME")
                    block_results["RES_TIME_UNITS"] = res_time_units
                    # 温度
                    coolant_tin = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COOLANT_TIN")
                    block_results["COOLANT_TIN"] = coolant_tin
                    coolant_tin_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COOLANT_TIN")
                    block_results["COOLANT_TIN_UNITS"] = coolant_tin_units
                    # 汽相分率
                    coolant_vin = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COOLANT_VIN")
                    block_results["COOLANT_VIN"] = coolant_vin

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Flash2":
                    # 收集Flash2 block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 汽相分率（摩尔）
                    b_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_VFRAC")
                    block_results["B_VFRAC"] = b_vfrac
                    # 汽相分率（质量）
                    mvfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MVFRAC")
                    block_results["MVFRAC"] = mvfrac
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净负荷
                    qnet = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET"] = qnet
                    qnet_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET")
                    block_results["QNET_UNITS"] = qnet_units
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 压降
                    pdrop = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP"] = pdrop
                    pdrop_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP_UNITS"] = pdrop_units
                    # 相平衡（动态子节点，无单位，按你给的 X/Y）
                    eq_params = ["F", "X", "Y", "K"]
                    for eq_param in eq_params:
                        subnodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\{eq_param}")
                        for subnode in subnodes:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\{eq_param}\{subnode}")
                            block_results[f"{eq_param}_{subnode}"] = value

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Decanter":
                    # 收集Decanter block的所有结果
                    block_results = {}
                    # 出口温度
                    b_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP"] = b_temp
                    b_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_TEMP")
                    block_results["B_TEMP_UNITS"] = b_temp_units
                    # 出口压力
                    b_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES"] = b_pres
                    b_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\B_PRES")
                    block_results["B_PRES_UNITS"] = b_pres_units
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 净负荷
                    qnet2 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QNET2")
                    block_results["QNET2"] = qnet2
                    qnet2_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QNET2")
                    block_results["QNET2_UNITS"] = qnet2_units
                    # 第一液相/全液相
                    liq_ratio = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\LIQ_RATIO")
                    block_results["LIQ_RATIO"] = liq_ratio
                    # 压降
                    pdrop = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP"] = pdrop
                    pdrop_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\PDROP")
                    block_results["PDROP_UNITS"] = pdrop_units
                    # 相平衡（动态子节点，无单位）
                    eq_params = ["F", "X1", "X2", "B_K"]
                    for eq_param in eq_params:
                        subnodes = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\{eq_param}")
                        for subnode in subnodes:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\{eq_param}\{subnode}")
                            block_results[f"{eq_param}_{subnode}"] = value

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Sep":
                    # 收集Sep block的所有结果
                    block_results = {}
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 子流股（动态获取组分，OVERHEAD/BOT 叶子，无单位）
                    comps = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED")
                    for comp in comps:
                        comp_value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}")
                        block_results[f"COMPFRAC_MIXED_{comp}"] = comp_value
                        overhead = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}\OVERHEAD")
                        block_results[f"COMPFRAC_MIXED_{comp}_OVERHEAD"] = overhead
                        bot = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}\BOT")
                        block_results[f"COMPFRAC_MIXED_{comp}_BOT"] = bot

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "Sep2":
                    # 收集Sep2 block的所有结果
                    block_results = {}
                    # 热负荷
                    qcalc = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC"] = qcalc
                    qcalc_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\QCALC")
                    block_results["QCALC_UNITS"] = qcalc_units
                    # 子流股（动态获取组分，OVERHEAD/BOT 叶子，无单位）
                    comps = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED")
                    for comp in comps:
                        comp_value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}")
                        block_results[f"COMPFRAC_MIXED_{comp}"] = comp_value
                        overhead = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}\OVERHEAD")
                        block_results[f"COMPFRAC_MIXED_{comp}_OVERHEAD"] = overhead
                        bot = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COMPFRAC\MIXED\{comp}\BOT")
                        block_results[f"COMPFRAC_MIXED_{comp}_BOT"] = bot

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "RadFrac":
                    # 收集RadFrac block的所有结果
                    block_results = {}
                    # 冷凝器/顶端塔板性能：温度
                    top_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results["TOP_TEMP"] = top_temp
                    top_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\TOP_TEMP")
                    block_results["TOP_TEMP_UNITS"] = top_temp_units
                    # 过冷温度
                    sctemp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\SCTEMP")
                    block_results["SCTEMP"] = sctemp
                    sctemp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\SCTEMP")
                    block_results["SCTEMP_UNITS"] = sctemp_units
                    # 热负荷
                    cond_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results["COND_DUTY"] = cond_duty
                    cond_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COND_DUTY")
                    block_results["COND_DUTY_UNITS"] = cond_duty_units
                    # 过冷负荷
                    scduty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\SCDUTY")
                    block_results["SCDUTY"] = scduty
                    scduty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\SCDUTY")
                    block_results["SCDUTY_UNITS"] = scduty_units
                    # 馏出物流率
                    mole_d = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_D")
                    block_results["MOLE_D"] = mole_d
                    mole_d_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_D")
                    block_results["MOLE_D_UNITS"] = mole_d_units
                    # 回流速率
                    mole_l1 = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_L1")
                    block_results["MOLE_L1"] = mole_l1
                    mole_l1_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_L1")
                    block_results["MOLE_L1_UNITS"] = mole_l1_units
                    # 回流比
                    mole_rr = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_RR")
                    block_results["MOLE_RR"] = mole_rr
                    # 自由水馏出物流率
                    mole_dw = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_DW")
                    block_results["MOLE_DW"] = mole_dw
                    mole_dw_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_DW")
                    block_results["MOLE_DW_UNITS"] = mole_dw_units
                    # 自由水回流比
                    rw = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\RW")
                    block_results["RW"] = rw
                    rw_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\RW")
                    block_results["RW_UNITS"] = rw_units
                    # 馏出物进料比
                    mole_dfr = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_DFR")
                    block_results["MOLE_DFR"] = mole_dfr
                    mole_dfr_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_DFR")
                    block_results["MOLE_DFR_UNITS"] = mole_dfr_units
                    # 再沸器/底端塔板性能：温度
                    bottom_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results["BOTTOM_TEMP"] = bottom_temp
                    bottom_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\BOTTOM_TEMP")
                    block_results["BOTTOM_TEMP_UNITS"] = bottom_temp_units
                    # 热负荷
                    reb_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results["REB_DUTY"] = reb_duty
                    reb_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\REB_DUTY")
                    block_results["REB_DUTY_UNITS"] = reb_duty_units
                    # 塔底物流率
                    mole_b = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_B")
                    block_results["MOLE_B"] = mole_b
                    mole_b_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_B")
                    block_results["MOLE_B_UNITS"] = mole_b_units
                    # 再沸蒸汽流率
                    mole_vn = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_VN")
                    block_results["MOLE_VN"] = mole_vn
                    mole_vn_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\MOLE_VN")
                    block_results["MOLE_VN_UNITS"] = mole_vn_units
                    # 再沸比
                    mole_br = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_BR")
                    block_results["MOLE_BR"] = mole_br
                    # 塔底采出与进料比
                    mole_bfr = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MOLE_BFR")
                    block_results["MOLE_BFR"] = mole_bfr
                    # 切割分率（MASS_CONC 下所有叶子节点动态获取，无单位）
                    comps = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\MASS_CONC")
                    for comp in comps:
                        streams = self.get_child_nodes(fr"\Data\Blocks\{block_name}\Output\MASS_CONC\{comp}")
                        for stream in streams:
                            value = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\MASS_CONC\{comp}\{stream}")
                            block_results[f"MASS_CONC_{comp}_{stream}"] = value

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                elif block['type'] == "HeatX":
                    # 收集HeatX block的所有结果
                    block_results = {}
                    # 热结果：计算模型
                    calc_model = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CALC_MODEL")
                    block_results["CALC_MODEL"] = calc_model
                    # 热流股
                    hotin = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOTIN")
                    block_results["HOTIN"] = hotin
                    hotout = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOTOUT")
                    block_results["HOTOUT"] = hotout
                    # 温度
                    hotint = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOTINT")
                    block_results["HOTINT"] = hotint
                    hotint_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HOTINT")
                    block_results["HOTINT_UNITS"] = hotint_units
                    hot_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOT_TEMP")
                    block_results["HOT_TEMP"] = hot_temp
                    hot_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HOT_TEMP")
                    block_results["HOT_TEMP_UNITS"] = hot_temp_units
                    # 压力
                    hotinp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOTINP")
                    block_results["HOTINP"] = hotinp
                    hotinp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HOTINP")
                    block_results["HOTINP_UNITS"] = hotinp_units
                    hot_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOT_PRES")
                    block_results["HOT_PRES"] = hot_pres
                    hot_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HOT_PRES")
                    block_results["HOT_PRES_UNITS"] = hot_pres_units
                    # 汽相分率
                    hotinvf = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOTINVF")
                    block_results["HOTINVF"] = hotinvf
                    hot_vfrac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOT_VFRAC")
                    block_results["HOT_VFRAC"] = hot_vfrac
                    # 第一液相/全液相
                    hin_l1frac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HIN_L1FRAC")
                    block_results["HIN_L1FRAC"] = hin_l1frac
                    hout_l1frac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HOUT_L1FRAC")
                    block_results["HOUT_L1FRAC"] = hout_l1frac
                    # 冷流股
                    coldin = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLDIN")
                    block_results["COLDIN"] = coldin
                    coldout = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLDOUT")
                    block_results["COLDOUT"] = coldout
                    # 温度
                    coldint = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLDINT")
                    block_results["COLDINT"] = coldint
                    coldint_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COLDINT")
                    block_results["COLDINT_UNITS"] = coldint_units
                    cold_temp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLD_TEMP")
                    block_results["COLD_TEMP"] = cold_temp
                    cold_temp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COLD_TEMP")
                    block_results["COLD_TEMP_UNITS"] = cold_temp_units
                    # 压力
                    coldinp = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLDINP")
                    block_results["COLDINP"] = coldinp
                    coldinp_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COLDINP")
                    block_results["COLDINP_UNITS"] = coldinp_units
                    cold_pres = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLD_PRES")
                    block_results["COLD_PRES"] = cold_pres
                    cold_pres_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\COLD_PRES")
                    block_results["COLD_PRES_UNITS"] = cold_pres_units
                    # 汽相分率
                    coldinvf = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLDINVF")
                    block_results["COLDINVF"] = coldinvf
                    cold_frac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COLD_FRAC")
                    block_results["COLD_FRAC"] = cold_frac
                    # 第一液相/全液相
                    cin_l1frac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\CIN_L1FRAC")
                    block_results["CIN_L1FRAC"] = cin_l1frac
                    cout_l1frac = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\COUT_L1FRAC")
                    block_results["COUT_L1FRAC"] = cout_l1frac
                    # 热负荷
                    hx_duty = self.safe_get_node_value(fr"\Data\Blocks\{block_name}\Output\HX_DUTY")
                    block_results["HX_DUTY"] = hx_duty
                    hx_duty_units = self.safe_get_node_units(fr"\Data\Blocks\{block_name}\Output\HX_DUTY")
                    block_results["HX_DUTY_UNITS"] = hx_duty_units

                    df_block = pd.DataFrame(list(block_results.items()), columns=["Parameter", "Value"])
                    sheet_name = block_name + "_result"
                    df_block.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"Block '{block_name}' 的结果已保存到工作表 '{sheet_name}'")

                # 可以添加其他block类型的处理
                # elif block['type'] == "RADFRAC":
                #     # 处理RADFRAC类型的block
                #     pass

        print(f"所有数据已保存到Excel文件: {os.path.abspath(excel_filename)}")
        result_path = os.path.abspath(excel_filename)
        return result_path

    def save_simulation(self, file_path: str):
        """
        保存模拟文件

        Args:
            file_path: 保存路径
        """
        try:
            self.aspen.SaveAs(file_path)
            print(f"模拟文件已保存到: {file_path}")
        except Exception as e:
            print(f"保存模拟文件失败: {e}")
            raise

    def close_simulation(self):
        """关闭模拟"""
        try:
            self.aspen.Close()
            print("模拟已关闭")
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"关闭模拟时出错: {e}")
            raise

def analyze_aspen_error(error_detail):
    """
    分析Aspen模拟配置写入错误返回的错误信息，判断错误类型
    """
    # 定义错误类型映射字典列表
    error_type_mappings = [
        {
            "keyword": "write_components_to_aspen",
            "error_message": "components配置写入错误"
        },
        {
            "keyword": "write_property_methods_to_aspen",
            "error_message": "property_methods配置写入错误"
        },
        {
            "keyword": "write_blocks_to_aspen",
            "error_message": "blocks配置写入错误"
        },
        {
            "keyword": "write_stream_to_aspen",
            "error_message": "stream配置写入错误"
        },
        {
            "keyword": "write_block_connections_to_aspen",
            "error_message": "block_connections配置写入错误"
        },
        {
            "keyword": "write_stream_data_to_aspen",
            "error_message": "stream_data配置写入错误"
        },
        {
            "keyword": "write_reactions_data_to_aspen",
            "error_message": "reactions_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Mixer_data_to_aspen",
            "error_message": "blocks_Mixer_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Valve_data_to_aspen",
            "error_message": "blocks_Valve_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Compr_data_to_aspen",
            "error_message": "blocks_Compr_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Heater_data_to_aspen",
            "error_message": "blocks_Heater_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Pump_data_to_aspen",
            "error_message": "blocks_Pump_data配置写入错误"
        },
        {
            "keyword": "write_blocks_RStoic_data_to_aspen",
            "error_message": "blocks_RStoic_data配置写入错误"
        },
        {
            "keyword": "write_blocks_RPlug_data_to_aspen",
            "error_message": "blocks_RPlug_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Flash2_data_to_aspen",
            "error_message": "blocks_Flash2_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Flash3_data_to_aspen",
            "error_message": "blocks_Flash3_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Sep_data_to_aspen",
            "error_message": "blocks_Sep_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Sep2_data_to_aspen",
            "error_message": "blocks_Sep2_data配置写入错误"
        },
        {
            "keyword": "write_blocks_RadFrac_data_to_aspen",
            "error_message": "blocks_RadFrac_data配置写入错误"
        },
        {
            "keyword": "write_blocks_DSTWU_data_to_aspen",
            "error_message": "blocks_DSTWU_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Distl_data_to_aspen",
            "error_message": "blocks_Distl_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Dupl_data_to_aspen",
            "error_message": "blocks_Dupl_data配置写入错误"
        },
        {
            "keyword": "write_blocks_Extract_data_to_aspen",
            "error_message": "blocks_Extract_data配置写入错误"
        },
        {
            "keyword": "write_blocks_FSplit_data_to_aspen",
            "error_message": "blocks_FSplit_data配置写入错误"
        },
        {
            "keyword": "write_blocks_HeatX_data_to_aspen",
            "error_message": "blocks_HeatX_data配置写入错误"
        },
        {
            "keyword": "write_blocks_MCompr_data_to_aspen",
            "error_message": "blocks_MCompr_data配置写入错误"
        },
        {
            "keyword": "write_blocks_RCSTR_data_to_aspen",
            "error_message": "blocks_RCSTR_data配置写入错误"
        }
    ]
    for error_map in error_type_mappings:
        if error_map["keyword"] in error_detail:
            return error_map["error_message"]

    # 如果没有匹配到已知错误类型
    return "未知配置写入错误"
class AspenEvents:
    def __init__(self):
        self.messages = []  # 存储所有控制面板消息
        self.current_session_messages = []  # 存储本次会话的消息
    def OnControlPanelMessage(self, clear, msg):
        if clear:
            print("控制面板已清空")
        else:
            print(f"控制面板消息: {msg}")
            # 存储消息
            self.messages.append(msg)
            self.current_session_messages.append(msg)
            # 可以在这里添加自定义处理逻辑
            self.process_control_panel_message(msg)

    def OnDialogSuppressed(self, msg, result):
        print(f"对话框被抑制: {msg}, 默认结果: {result}")

    def OnGUIClosing(self):
        print("ASPEN GUI正在关闭")
    def process_control_panel_message(self, message):
        """处理控制面板消息的自定义逻辑"""
        # 例如：记录到文件
        try:
            os.makedirs("../aspenlog", exist_ok=True)
            message_file = f"../aspenlog/aspen_control_panel.log"
            with open(message_file, "a", encoding='utf-8') as f:
                f.write(f"{datetime.now().isoformat()}: {message}\n")
        except Exception as e:
            print(f"写入日志文件失败: {e}")

    def get_current_session_messages(self):
        """获取本次会话的所有控制面板消息"""
        return self.current_session_messages

    def get_current_session_messages_as_string(self):
        """获取本次会话的所有控制面板消息，作为字符串"""
        return "\n".join(self.current_session_messages)

    def get_all_messages(self):
        """获取所有控制面板消息"""
        return self.messages


@app.route('/get-aspen-result', methods=['GET'])
def get_result():
    """
    读取本地Excel结果文件，自动读取所有工作表，用于大模型调用后进行结果分析

    请求参数:
    - file_path: Excel文件路径（通过查询参数传递）

    返回:
    - JSON格式的Excel数据，包含所有工作表的内容
    """
    try:
        # 从查询参数获取文件路径
        file_path = request.args.get('file_path')

        if not file_path:
            return jsonify({
                "error": "缺少必需参数: file_path",
                "timestamp": datetime.now().isoformat()
            }), 400

        # 验证文件路径
        if not os.path.exists(file_path):
            return jsonify({
                "error": f"文件不存在: {file_path}",
                "timestamp": datetime.now().isoformat()
            }), 404

        # 使用ExcelFile对象获取所有工作表信息
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # 构建返回结果
        result = {
            "status": "success",
            "file_info": {
                "file_name": os.path.basename(file_path),
                "file_path": file_path,
                "sheet_count": len(sheet_names),
                "sheet_names": sheet_names,
                "read_time": datetime.now().isoformat()
            },
            "data": {}
        }

        # 读取所有工作表
        for sheet_name in sheet_names:
            try:
                # 读取当前工作表
                df = pd.read_excel(file_path, sheet_name=sheet_name)

                # 处理当前工作表数据
                sheet_data = {
                    "row_count": len(df),
                    "column_count": len(df.columns),
                    "column_names": df.columns.tolist(),
                    "data": df.where(pd.notnull(df), None).to_dict(orient='records')
                }

                # 添加到结果中
                result["data"][sheet_name] = sheet_data

            except Exception as e:
                # 如果某个工作表读取失败，记录错误信息
                result["data"][sheet_name] = {
                    "error": f"读取工作表失败: {str(e)}",
                    "row_count": 0,
                    "column_count": 0
                }

        return jsonify(result)

    except Exception as e:
        return jsonify({
            "error": f"读取Excel文件失败: {str(e)}",
            "file_path": file_path if 'file_path' in locals() else None,
            "timestamp": datetime.now().isoformat()
        }), 500


@app.route('/run-aspen-simulation', methods=['POST'])
def run_aspen_simulation():
    # 获取请求数据
    config = request.json
    if not config:
        return jsonify({"error": "请求体为空"}), 400
    
    # 创建唯一的时间戳
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 保存配置文件
    config_file_path = CONFIG_DIR / f"config_{timestamp}.json"
    try:
        with open(config_file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"配置文件已保存到: {config_file_path}")
    except Exception as e:
        print(f"保存配置文件时出错: {e}")
        return jsonify({"error": f"无法保存配置文件: {e}"}), 500

    # 初始化模拟管理器
    aspen_manager = AspenSimulationManager()
    try:
    # 尝试写入配置到ASPEN模拟文件
        # 创建新模拟 - 使用配置的模板路径
        template_path = str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE.exists() else None
        aspen_manager.create_new_simulation(template_path)

        # 创建唯一的结果输出文件名
        output_file_path = OUTPUT_DIR / f"output_{timestamp}.bkp"

        # 加载JSON配置
        loaded_config = aspen_manager.load_json_config(config)

        # 将配置写入Aspen
        aspen_manager.write_config_to_aspen(loaded_config)

    except Exception as e:
        # 获取详细的错误信息，包括具体是哪一步配置写入失败
        error_detail = f"配置写入失败: {str(e)}\n错误位置: {traceback.format_exc()}"
        print(f"n错误位置: {traceback.format_exc()}")
        error_message = analyze_aspen_error(error_detail)
        # 保存模拟文件
        aspen_manager.save_simulation(str(output_file_path))
        return jsonify({
            "success": False,
            "aspen_file_path": str(output_file_path),
            "config_file_path": str(config_file_path),
            "error_type": "模拟配置写入失败",
            "error_message": f"{error_message}: {str(e)}"
        }), 201

    try:
        # 运行模拟文件
        aspen_manager.run_simulation()

        # 获取ASPEN控制面板消息
        current_messages_str = aspen_manager.get_control_panel_messages()

        # 保存模拟文件
        aspen_manager.save_simulation(str(output_file_path))

        # 结果文件路径
        result_file_path = RESULT_DIR / f"aspen_result_export_{timestamp}.xlsx"

        if "No Errors" in current_messages_str:
            try:
                # 获取模拟文件运行结果
                result_absolute_path = aspen_manager.get_all_simulation_results(str(result_file_path), loaded_config)
            except Exception as e:
                print(f"保存结果文件错误: {str(e)}")

            # 返回生成的文件路径
            return jsonify({
                "success": True,
                "aspen_file_path": str(output_file_path),
                "config_file_path": str(config_file_path),
                "result_file_path": result_absolute_path,
                "message": "Aspen模拟已成功运行并保存"
            })
        elif "**  ERROR" in current_messages_str or "*** SEVERE ERROR" in current_messages_str:
            return jsonify({
                "success": False,
                "aspen_file_path": str(output_file_path),
                "config_file_path": str(config_file_path),
                "error_type": "模拟运行过程发生错误",
                "error_message": current_messages_str
            }), 201
        else:
            # 没有严重错误 -> 视为成功（可能有 WARNING 但不影响）
            return jsonify({
                "success": True,
                "aspen_file_path": str(output_file_path),
                "config_file_path": str(config_file_path),
                "result_file_path": result_absolute_path,
                "message": "模拟已完成（无严重错误）"
            })
    except Exception as e:
        # 获取ASPEN控制面板消息
        current_messages_str = aspen_manager.get_control_panel_messages()
        return jsonify({
            "success": False,
            "error_message": f"{str(e)}:{current_messages_str}",
            "error_type": "模拟运行过程失败"
        }), 201
    finally:
        # 确保关闭模拟
        if aspen_manager:
            try:
                aspen_manager.close_simulation()
            except:
                pass


@app.route('/download', methods=['GET'])
def download_file():
    """
    下载文件接口

    请求参数:
    - file_path: 文件路径（通过查询参数传递）

    返回:
    - 文件内容
    """
    try:
        from flask import send_file
        from pathlib import Path

        # 从查询参数获取文件路径
        file_path = request.args.get('file_path')

        if not file_path:
            return jsonify({
                "error": "缺少必需参数: file_path",
                "timestamp": datetime.now().isoformat()
            }), 400

        # 转换为Path对象
        file_path_obj = Path(file_path)

        # 验证文件是否存在
        if not file_path_obj.exists():
            return jsonify({
                "error": f"文件不存在: {file_path}",
                "timestamp": datetime.now().isoformat()
            }), 404

        # 验证是否为文件（不是目录）
        if not file_path_obj.is_file():
            return jsonify({
                "error": f"路径不是文件: {file_path}",
                "timestamp": datetime.now().isoformat()
            }), 400

        # 扩展允许的文件类型
        allowed_extensions = {'.bkp', '.apw', '.json', '.xlsx', '.xls', '.txt', '.log', '.out', '.csv'}
        if file_path_obj.suffix.lower() not in allowed_extensions:
            return jsonify({
                "error": f"不允许下载此类型文件: {file_path_obj.suffix}",
                "allowed_types": list(allowed_extensions),
                "timestamp": datetime.now().isoformat()
            }), 403

        # 根据文件扩展名设置Content-Type
        content_types = {
            '.bkp': 'application/octet-stream',
            '.apw': 'application/octet-stream',
            '.json': 'application/json',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xls': 'application/vnd.ms-excel',
            '.txt': 'text/plain',
            '.log': 'text/plain',
            '.out': 'text/plain',
            '.csv': 'text/csv'
        }

        mimetype = content_types.get(file_path_obj.suffix.lower(), 'application/octet-stream')

        # 使用send_file发送文件
        return send_file(
            str(file_path_obj),
            mimetype=mimetype,
            as_attachment=True,
            download_name=file_path_obj.name
        )

    except Exception as e:
        return jsonify({
            "error": f"下载文件失败: {str(e)}",
            "file_path": file_path if 'file_path' in locals() else None,
            "timestamp": datetime.now().isoformat()
        }), 500

@app.get("/health")
def health_check():
    return {
        "status": "ok",
        "timestamp": datetime.now().isoformat(),
        "aspen_available": True
    }

@app.route('/api/schema', methods=['GET'])
def get_schema():
    """
    获取 Schema 文件
    
    查询参数:
    - types: 设备类型，多个用逗号分隔，如 'base' 或 'Mixer,Heater' 或 'all'
    - format: 返回格式，'json' 或 'list'（默认 'json'）
    
    示例:
    - GET /api/schema?types=base
    - GET /api/schema?types=Mixer,Heater,Pump
    - GET /api/schema?types=all
    - GET /api/schema?types=all&format=list
    """
    try:
        # 获取查询参数
        types_param = request.args.get('types', 'base')
        format_param = request.args.get('format', 'json')
        
        # 解析设备类型
        if types_param.lower() == 'all':
            # 返回所有 schema 文件列表
            schema_files = list(SCHEMA_DIR.glob('*.json'))
            requested_types = [f.stem for f in schema_files]
        else:
            # 解析逗号分隔的类型
            requested_types = [t.strip() for t in types_param.split(',') if t.strip()]
        
        # 如果只是请求列表
        if format_param == 'list':
            available_schemas = list(SCHEMA_DIR.glob('*.json'))
            schema_list = []
            for schema_file in available_schemas:
                file_name = schema_file.name
                schema_type = schema_file.stem
                
                # 提取设备类型名称
                if schema_type == 'base_schema':
                    display_name = 'base'
                    description = '基础配置 Schema'
                elif schema_type.startswith('blocks_'):
                    block_type = schema_type.replace('blocks_', '').replace('_data', '')
                    display_name = block_type
                    description = f'{block_type} 设备配置 Schema'
                else:
                    display_name = schema_type
                    description = f'{schema_type} Schema'
                
                schema_list.append({
                    'type': display_name,
                    'file': file_name,
                    'description': description
                })
            
            return jsonify({
                'success': True,
                'count': len(schema_list),
                'schemas': sorted(schema_list, key=lambda x: x['type'])
            })
        
        # 读取并返回 schema 内容
        result = {}
        not_found = []
        
        for schema_type in requested_types:
            # 标准化类型名称
            if schema_type.lower() == 'base':
                schema_file = SCHEMA_DIR / 'base_schema.json'
                key = 'base'
            else:
                # 尝试多种文件名格式
                possible_files = [
                    SCHEMA_DIR / f'blocks_{schema_type}_data.json',
                    SCHEMA_DIR / f'{schema_type}.json',
                    SCHEMA_DIR / f'blocks_{schema_type}.json',
                ]
                
                schema_file = None
                for pf in possible_files:
                    if pf.exists():
                        schema_file = pf
                        break
                
                key = schema_type
            
            # 读取文件
            if schema_file and schema_file.exists():
                try:
                    with open(schema_file, 'r', encoding='utf-8') as f:
                        schema_content = json.load(f)
                    result[key] = schema_content
                except Exception as e:
                    result[key] = {
                        'error': f'读取文件失败: {str(e)}',
                        'file': str(schema_file)
                    }
            else:
                not_found.append(schema_type)
        
        # 构建响应
        response_data = {
            'success': len(not_found) == 0,
            'requested': requested_types,
            'found': len(result),
            'schemas': result
        }
        
        if not_found:
            response_data['not_found'] = not_found
            response_data['message'] = f'部分 Schema 未找到: {", ".join(not_found)}'
        
        return jsonify(response_data)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'message': '获取 Schema 失败'
        }), 500

@app.route('/api/schema/list', methods=['GET'])
def list_schemas():
    """
    列出所有可用的 Schema 文件
    
    返回所有可用的 Schema 类型列表
    """
    try:
        schema_files = list(SCHEMA_DIR.glob('*.json'))
        schema_list = []
        
        for schema_file in schema_files:
            file_name = schema_file.name
            schema_type = schema_file.stem
            file_size = schema_file.stat().st_size
            
            # 提取设备类型名称
            if schema_type == 'base_schema':
                display_name = 'base'
                category = 'base'
                description = '基础配置 Schema'
            elif schema_type.startswith('blocks_'):
                block_type = schema_type.replace('blocks_', '').replace('_data', '')
                display_name = block_type
                category = 'block'
                description = f'{block_type} 设备配置 Schema'
            else:
                display_name = schema_type
                category = 'other'
                description = f'{schema_type} Schema'
            
            schema_list.append({
                'type': display_name,
                'file': file_name,
                'category': category,
                'description': description,
                'size': file_size
            })
        
        return jsonify({
            'success': True,
            'count': len(schema_list),
            'schemas': sorted(schema_list, key=lambda x: (x['category'], x['type']))
        })
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'message': '列出 Schema 失败'
        }), 500

if __name__ == "__main__":
    cert_path, key_path, created, hosts = ensure_server_certificate(
        SSL_CERT_FILE,
        SSL_KEY_FILE,
        base_dir=Path(__file__).parent,
    )
    SSL_CERT_PATH = cert_path
    SSL_KEY_PATH = key_path

    if CONFIG_AVAILABLE:
        print_config()
        issues = validate_config()
        if issues:
            print("Warning: configuration issues were found, but service startup will continue.")

    import ssl

    ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    ssl_context.load_cert_chain(str(SSL_CERT_PATH), str(SSL_KEY_PATH))
    print("Starting Aspen simulation service in HTTPS mode")
    print(f"  Certificate: {SSL_CERT_PATH}")
    print(f"  Private key: {SSL_KEY_PATH}")
    if created:
        print("  A reusable self-signed certificate was created for this host.")
        print(f"  SAN: {', '.join(hosts)}")

    print(f"  Listen: {HOST}:{PORT}")
    print(f"  Debug: {DEBUG}")
    print("=" * 60)

    app.run(
        host=HOST,
        port=PORT,
        debug=DEBUG,
        use_reloader=False,
        ssl_context=ssl_context
    )
