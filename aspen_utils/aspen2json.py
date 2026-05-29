import win32com.client
import pythoncom
import json
import os
from typing import Dict, List, Any, Optional

class AspenToJSONConverter:
    def __init__(self, aspen_file_path):
        """初始化 Aspen Plus 连接"""
        self.aspen = None
        self.aspen_file_path = aspen_file_path
        self.data = {}

    def connect_to_aspen(self):
        """连接到 Aspen Plus 实例"""
        try:
            self.aspen = win32com.client.Dispatch("Apwn.Document")
            if os.path.exists(self.aspen_file_path):
                self.aspen.InitFromArchive2(os.path.abspath(self.aspen_file_path))
                print(f"成功加载 Aspen Plus 文件: {self.aspen_file_path}")
                return True
            else:
                print("文件不存在")
                return False
        except Exception as e:
            print(f"连接 Aspen Plus 失败: {e}")
            return False

    def disconnect(self):
        """断开与 Aspen Plus 的连接"""
        if self.aspen:
            self.aspen.Close()
            pythoncom.CoUninitialize()
            print("已断开与 Aspen Plus 的连接")

    def safe_get_node_value(self, node_path: str, default: Any = None) -> Any:
        """安全获取节点值，避免节点不存在时抛出异常"""
        try:
            node = self.aspen.Tree.FindNode(node_path)
            if node:
                return node.Value
            else:
                return default
        except Exception as e:
            print(f"获取节点 {node_path} 值时出错: {e}")
            return default

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

    def get_block_type(self, node_path, HAP_RECORDTYPE):
        node = self.aspen.Tree.FindNode(node_path)
        return node.AttributeValue(HAP_RECORDTYPE)

    def add_if_not_empty(self, data_dict, key, value, unit_key=None, unit_value=None, basis_key=None, basis_value=None):
        """如果值不为空，则将其添加到字典中"""
        if value is not None and value != "":
            data_dict[key] = value
            if unit_key and unit_value is not None and unit_value != "":
                data_dict[unit_key] = unit_value
            elif basis_key and basis_value is not None and basis_value != "":
                data_dict[basis_key] = basis_value

    def extract_metadata(self):
        """提取元数据"""
        try:
            description = self.safe_get_node_value(r"\Data\Results Summary\Run-Status\Output\DESCRIPTION", "Unknown")

            self.data["metadata"] = {
                "description": description
            }
            print("元数据提取完成")
        except Exception as e:
            print(f"提取元数据时出错: {e}")
            self.data["metadata"] = {"description": "Unknown"}
    def extract_setup(self):
        """提取设置数据"""
        try:
            setup_data = {}
            setup_data["sim_options"] = {}
            # 1. 提取设置-计算选项配置
            ENERGY_BAL_VALUE = self.safe_get_node_value(f"\Data\Setup\Sim-Options\Input\ENERGY_BAL") #设置-计算选项-执行热量平衡计算
            self.add_if_not_empty(setup_data["sim_options"], "energy_bal_value", ENERGY_BAL_VALUE)
            self.data["setup"] = setup_data
            print(f"设置数据提取完成")
        except Exception as e:
            print(f"提取设置数据时出错: {e}")
    def extract_components(self):
        """提取组分数据"""
        try:
            components = []

            # 1. 首先从 CID 目录获取所有子节点的值
            cid_nodes = self.get_child_nodes(r"\Data\Components\Comp-Lists\GLOBAL\Input\CID")
            cid_values = []

            for cid_node in cid_nodes:
                cid_value = self.safe_get_node_value(fr"\Data\Components\Comp-Lists\GLOBAL\Input\CID\{cid_node}")
                cid_values.append(cid_value)

            print(f"从 CID 目录获取到 {len(cid_values)} 个组分 ID")

            # 2. 使用 CID 值作为索引，从其他目录获取对应的值
            for i, cid in enumerate(cid_values, 1):
                # 获取组分名称
                name = self.safe_get_node_value(fr"\Data\Components\Specifications\Input\ANAME\{cid}", f"Component_{i}")

                # 获取 CAS 号
                casn = self.safe_get_node_value(fr"\Data\Components\Specifications\Input\CASN\{cid}", "")

                # 获取数据库名称
                dbname = self.safe_get_node_value(fr"\Data\Components\Specifications\Input\DBNAME\{cid}", "")

                # 数据库不存在的自定义组分不抽取
                if dbname is not None:
                    components.append({
                        "cid": cid,
                        "name": name,
                        "cas_number": casn,
                        "database_name": dbname
                    })
            self.data["components"] = components
            print(f"组分数据提取完成，共 {len(components)} 个组分")
        except Exception as e:
            print(f"提取组分数据时出错: {e}")
            self.data["components"] = []
    def extract_property_methods(self):
        """提取物性方法"""
        try:
            property_methods = []
            # 获取所有物性方法
            prop_methods_node = self.aspen.Tree.FindNode(r"\Data\Properties\Property Methods")
            # 获取基准方法
            basis_method = self.safe_get_node_value(
                fr"\Data\Properties\Specifications\Input\GBASEOPSET", "")

            if prop_methods_node and prop_methods_node.Elements.Count > 0:
                for method in prop_methods_node.Elements:
                    method_name = method.Name
                    if basis_method == method_name:
                        property_methods.append({
                            "method_name": method_name,
                            "is_basis_method": True
                        })
                    else:
                        property_methods.append({
                            "method_name": method_name,
                            "is_basis_method": False
                        })
            self.data["property_methods"] = property_methods
            print(f"物性方法提取完成，共 {len(property_methods)} 个方法")

        except Exception as e:
            print(f"提取物性方法时出错: {e}")
            self.data["property_methods"] = {}
    def extract_henry_components(self):
        """提取Henry组分"""
        try:
            henry_components = {}

            # 获取Henry组分集的子目录
            henry_sets = self.get_child_nodes(r"\Data\Components\Henry-Comps")

            for henry_set in henry_sets:
                # 获取当前Henry组分集的CID节点
                cid_nodes = self.get_child_nodes(fr"\Data\Components\Henry-Comps\{henry_set}\Input\CID")

                components_in_set = []
                for cid_node in cid_nodes:
                    # 获取CID节点的值（化学式）
                    formula = self.safe_get_node_value(
                        fr"\Data\Components\Henry-Comps\{henry_set}\Input\CID\{cid_node}")

                    if formula:
                        components_in_set.append({
                            "node": cid_node,
                            "formula": formula
                        })

                henry_components[henry_set] = {
                    "components": components_in_set
                }

            self.data["henry_components"] = henry_components
            print(f"Henry组分提取完成，共 {len(henry_components)} 个Henry组分集")

        except Exception as e:
            print(f"提取Henry组分时出错: {e}")
            self.data["henry_components"] = {}
    # def extract_custom_component_parameters(self):
    #     """提取自定义组分参数"""
    #     try:
    #         custom_params = {}
    #
    #         # 检查USRDEF目录是否存在
    #         usrdef_path = r"\Data\Properties\Parameters\Pure Components\USRDEF"
    #         usrdef_node = self.aspen.Tree.FindNode(usrdef_path)
    #
    #         if not usrdef_node:
    #             print("USRDEF目录不存在，跳过自定义组分参数提取")
    #             self.data["custom_component_parameters"] = {}
    #             return
    #
    #         # 获取SETNO目录下的参数和单位
    #         setno_nodes = self.get_child_nodes(fr"{usrdef_path}\Input\SETNO")
    #
    #         setno_params = {}
    #         for node in setno_nodes:
    #             value = self.safe_get_node_value(
    #                 fr"{usrdef_path}\Input\SETNO\{node}")
    #             units = self.safe_get_node_units(
    #                 fr"{usrdef_path}\Input\SETNO\{node}")
    #
    #             setno_params[node] = {
    #                 "value": value,
    #                 "units": units
    #             }
    #
    #         # 获取VALUE目录下的组分名称和值
    #         value_nodes = self.get_child_nodes(fr"{usrdef_path}\Input\UVALUE")
    #
    #         component_values = {}
    #         for node in value_nodes:
    #             value = self.safe_get_node_value(
    #                 fr"{usrdef_path}\Input\UVALUE\{node}")
    #
    #             component_values[node] = value
    #
    #         custom_params = {
    #             "setno_parameters": setno_params,
    #             "component_values": component_values
    #         }
    #
    #         self.data["custom_component_parameters"] = custom_params
    #         print(f"自定义组分参数提取完成，SETNO参数: {len(setno_params)} 个, 组分值: {len(component_values)} 个")
    #
    #     except Exception as e:
    #         print(f"提取自定义组分参数时出错: {e}")
    #         self.data["custom_component_parameters"] = {}
    def extract_blocks(self):
        """提取单元操作及其类型"""
        try:
            blocks_node = self.aspen.Tree.FindNode(r"\Data\Blocks")
            if not blocks_node:
                print("未找到Blocks节点")
                self.data["blocks"] = []
                return

            blocks = []
            for block_name in self.get_child_nodes(r"\Data\Blocks"):
                # 获取单元操作类型
                block_node = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block_name}")
                block_type = block_node.AttributeValue(6)

                blocks.append({
                    "name": block_name,
                    "type": block_type
                })

            self.data["blocks"] = blocks
            print(f"单元操作提取完成，共 {len(blocks)} 个单元操作")

        except Exception as e:
            print(f"提取单元操作时出错: {e}")
            self.data["blocks"] = []
    def extract_streams(self):
        """提取物流"""
        try:
            streams = self.get_child_nodes(r"\Data\Streams")
            self.data["streams"] = streams
            print(f"物流提取完成，共 {len(streams)} 个物流")
        except Exception as e:
            print(f"提取物流时出错: {e}")
            self.data["streams"] = []
    def extract_stream_connections(self):
        """提取物流连接"""
        try:
            connections = {}
            streams = self.data.get("streams", [])

            for stream in streams:
                try:
                    conn_path = fr"\Data\Streams\{stream}\Connections"
                    conn_nodes = self.get_child_nodes(conn_path)

                    if conn_nodes:
                        connections[stream] = {}
                        for conn_node in conn_nodes:
                            node_path = fr"{conn_path}\{conn_node}"
                            value = self.safe_get_node_value(node_path)
                            connections[stream][conn_node] = value
                except Exception as e:
                    print(f"提取物流 {stream} 连接时出错: {e}")
                    continue

            self.data["stream_connections"] = connections
            print("物流连接提取完成")
        except Exception as e:
            print(f"提取物流连接时出错: {e}")
            self.data["stream_connections"] = {}
    def extract_block_connections(self):
        """提取物流连接"""
        try:
            connections = {}
            blocks = self.data.get("blocks", [])
            for block in blocks:
                try:
                    conn_path = fr"\Data\Blocks\{block['name']}\Connections"
                    conn_nodes = self.get_child_nodes(conn_path)
                    if conn_nodes:
                        connections[block['name']] = {}
                        for conn_node in conn_nodes:
                            node_path = fr"{conn_path}\{conn_node}"
                            value = self.safe_get_node_value(node_path)
                            connections[block['name']][conn_node] = value
                except Exception as e:
                    print(f"提取设备 {block['name']} 连接时出错: {e}")
                    continue

            self.data["block_connections"] = connections
            print("设备连接提取完成")
        except Exception as e:
            print(f"提取设备连接时出错: {e}")
            self.data["block_connections"] = {}
    def extract_streams_data(self):
        """提取Streams流股数据"""
        try:
            stream_data = {}
            Streams = self.data.get("streams", [])
            for stream in Streams:
                stream_data[stream] = {}
                MIXED_SPEC = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\MIXED_SPEC\MIXED")
                stream_data[stream]["MIXED_SPEC"] = MIXED_SPEC
                PRES_VALUE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\PRES\MIXED")
                PRES_UNITS = self.safe_get_node_units(fr"\Data\Streams\{stream}\Input\PRES\MIXED")
                TEMP_VALUE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\TEMP\MIXED")
                TEMP_UNITS = self.safe_get_node_units(fr"\Data\Streams\{stream}\Input\TEMP\MIXED")
                VFRAC_VALUE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\VFRAC\MIXED")
                if MIXED_SPEC == "TP":
                    stream_data[stream]["pressure"] = {
                        "PRES_VALUE": PRES_VALUE,
                        "PRES_UNITS": PRES_UNITS
                    }
                    stream_data[stream]["temperature"] = {
                        "TEMP_VALUE": TEMP_VALUE,
                        "TEMP_UNITS": TEMP_UNITS
                    }
                elif MIXED_SPEC == "TV":
                    stream_data[stream]["temperature"] = {
                        "TEMP_VALUE": TEMP_VALUE,
                        "TEMP_UNITS": TEMP_UNITS
                    }
                    stream_data[stream]["vfrac"] = {
                        "VFRAC_VALUE": VFRAC_VALUE
                    }
                elif MIXED_SPEC == "PV":
                    stream_data[stream]["pressure"] = {
                        "PRES_VALUE": PRES_VALUE,
                        "PRES_UNITS": PRES_UNITS
                    }
                    stream_data[stream]["vfrac"] = {
                        "VFRAC_VALUE": VFRAC_VALUE
                    }
                # 提取流量数据
                FLOWBASE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\FLOWBASE")  # 规定-总流量-基准
                TOTFLOW_VALUE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\TOTFLOW") # 规定-总流量-值
                TOTFLOW_UNIT = self.safe_get_node_units(fr"\Data\Streams\{stream}\Input\TOTFLOW") # 规定-总流量-单位
                BASIS = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\BASIS")  # 规定-组成-基准
                flow_nodes = self.get_child_nodes(fr"\Data\Streams\{stream}\Input\FLOW\MIXED")   # 规定-组成
                flow_values = {}
                self.add_if_not_empty(flow_values, "FLOWBASE", FLOWBASE)
                self.add_if_not_empty(flow_values, "TOTFLOW_VALUE", TOTFLOW_VALUE, "TOTFLOW_UNITS", TOTFLOW_UNIT)
                self.add_if_not_empty(flow_values, "BASIS", BASIS)
                # 提取所有组分的name
                components = self.data.get("components", [])
                component_cids = [comp['cid'] for comp in components]
                for node in flow_nodes:
                    if node in component_cids: # 只提取components中的组分，自定义组分的配置不要提取
                        FLOW_VALUE = self.safe_get_node_value(fr"\Data\Streams\{stream}\Input\FLOW\MIXED\{node}")
                        FLOW_UNITS = self.safe_get_node_units(fr"\Data\Streams\{stream}\Input\FLOW\MIXED")
                        if FLOW_VALUE is not None and FLOW_VALUE != "":
                            flow_values[node] = {
                                "FLOW_VALUE": FLOW_VALUE,
                                "FLOW_UNITS": FLOW_UNITS,
                                "FLOW_BASIS": BASIS
                            }
                stream_data[f"{stream}"]["flow"] = flow_values
            self.data["stream_data"] = stream_data
            print("streams物流数据提取完成")
        except Exception as e:
            print(f"提取streams物流数据时出错: {e}")
            self.data["stream_data"] = {}
    def extract_convergence_data(self):
        """提取convergence数据"""
        try:
            convergence_data = {}
            # 收敛-收敛选项
            convergence_data["conv_options"] = {}
            #CONV_NODES = self.get_child_nodes(fr"\Data\Convergence\Convergence")  # 收敛节点
            #CONV_OPT_NODES = self.get_child_nodes(fr"\Data\Convergence\Conv-Options\Input\TEAR_METHOD")  # 收敛-选项
            # 默认值 - 撕裂收敛
            TOL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TOL")
            TRACE_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TRACE")
            TRACEOPT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TRACEOPT")
            COMPS_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\COMPS")
            STATE_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\STATE")
            FLASH_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\FLASH")
            UPDATE_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\UPDATE")
            VARITERHIST_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\VARITERHIST")
            self.add_if_not_empty(convergence_data["conv_options"], "tol", TOL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "trace", TRACE_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "traceopt", TRACEOPT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "comps", COMPS_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "state", STATE_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "flash", FLASH_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "update", UPDATE_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "variterhist", VARITERHIST_VALUE)
            # 默认方法
            TEAR_METHOD_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TEAR_METHOD")  # 收敛-选项-默认方法
            SPEC_METHOD_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SPEC_METHOD")
            MSPEC_METHOD_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\MSPEC_METHOD")
            COMB_METHOD_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\COMB_METHOD")
            OPT_METHOD_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\OPT_METHOD")
            self.add_if_not_empty(convergence_data["conv_options"], "tear_method", TEAR_METHOD_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "spec_method", SPEC_METHOD_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "mspec_method", MSPEC_METHOD_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "comb_method", COMB_METHOD_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "opt_method", OPT_METHOD_VALUE)
            # 顺序确定
            SPEC_LOOP_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SPEC_LOOP")
            USER_LOOP_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\USER_LOOP")
            TEAR_WEIGHT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TEAR_WEIGHT")
            LOOP_WEIGHT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\LOOP_WEIGHT")
            AFFECT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\AFFECT")
            CHECKSEQ_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\CHECKSEQ")
            TEAR_VAR_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\TEAR_VAR")
            self.add_if_not_empty(convergence_data["conv_options"], "spec_loop", SPEC_LOOP_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "user_loop", USER_LOOP_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "tear_weight", TEAR_WEIGHT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "loop_weight", LOOP_WEIGHT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "affect", AFFECT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "checkseq", CHECKSEQ_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "tear_var", TEAR_VAR_VALUE)
            # 方法 - Wegstein
            WEG_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\WEG_MAXIT") # 收敛-选项-迭代次数
            WEG_WAIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\WEG_WAIT")
            ACCELERATE_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\ACCELERATE")
            NACCELERATE_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NACCELERATE")
            WEG_QMIN_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\WEG_QMIN")
            WEG_QMAX_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\WEG_QMAX")
            self.add_if_not_empty(convergence_data["conv_options"], "weg_maxit", WEG_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "weg_wait", WEG_WAIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "accelerate", ACCELERATE_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "naccelerate", NACCELERATE_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "weg_qmin", WEG_QMIN_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "weg_qmax", WEG_QMAX_VALUE)
            # 方法 - 直接
            DIR_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\DIR_MAXIT")
            self.add_if_not_empty(convergence_data["conv_options"], "dir_maxit", DIR_MAXIT_VALUE)
            # 方法 - 正割
            SEC_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SEC_MAXIT")
            STEP_SIZ_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\STEP_SIZ")
            SEC_XTOL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SEC_XTOL")
            XFINAL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\XFINAL")
            BRACKET_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\BRACKET")
            STOP_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\STOP")
            self.add_if_not_empty(convergence_data["conv_options"], "sec_maxit", SEC_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "step_siz", STEP_SIZ_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sec_xtol", SEC_XTOL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "xfinal", XFINAL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "bracket", BRACKET_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "stop", STOP_VALUE)
            # 方法 - Broyden
            BR_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\BR_MAXIT")
            BR_XTOL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\BR_XTOL")
            BR_WAIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\BR_WAIT")
            self.add_if_not_empty(convergence_data["conv_options"], "br_maxit", BR_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "br_xtol", BR_XTOL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "br_wait", BR_WAIT_VALUE)
            # 方法 - Newton
            NEW_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NEW_MAXIT")
            NEW_MAXPASS_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NEW_MAXPASS")
            NEW_WAIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NEW_WAIT")
            NEW_XTOL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NEW_XTOL")
            OPT_N_JAC_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\OPT_N_JAC")
            RED_FACTOR_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\RED_FACTOR")
            REINIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\REINIT")
            self.add_if_not_empty(convergence_data["conv_options"], "new_maxit", NEW_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "new_maxpass", NEW_MAXPASS_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "new_wait", NEW_WAIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "new_xtol", NEW_XTOL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "opt_n_jac", OPT_N_JAC_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "red_factor", RED_FACTOR_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "reinit", REINIT_VALUE)
            # 方法 - SQP
            SQP_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_MAXIT")
            SQP_MAXPASS_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_MAXPASS")
            CONST_ITER_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\CONST_ITER")
            MAXLSPASS_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\MAXLSPASS")
            NLIMIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NLIMIT")
            SQP_TOL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_TOL")
            SQP_WAIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_WAIT")
            SQP_QMIN_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_QMIN")
            SQP_QMAX_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\SQP_QMAX")
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_maxit", SQP_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_maxpass", SQP_MAXPASS_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "const_iter", CONST_ITER_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "maxlspass", MAXLSPASS_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "nlimit", NLIMIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_tol", SQP_TOL_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_wait", SQP_WAIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_qmin", SQP_QMIN_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "sqp_qmax", SQP_QMAX_VALUE)
            # 方法 - BOBYQA
            BOBY_MAXIT_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\BOBY_MAXIT")
            NCONDITIONS_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\NCONDITIONS")
            INIT_REGION_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\INIT_REGION")
            FINAL_REGION_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\FINAL_REGION")
            INITPREF_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\INITPREF")
            PREFGROWI_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\PREFGROWI")
            PREFGROWF_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\PREFGROWF")
            EQPENTYP_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\EQPENTYP")
            INEQPENTYP_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\INEQPENTYP")
            PENSCL_VALUE = self.safe_get_node_value(fr"\Data\Convergence\Conv-Options\Input\PENSCL")
            self.add_if_not_empty(convergence_data["conv_options"], "boby_maxit", BOBY_MAXIT_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "nconditions", NCONDITIONS_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "init_region", INIT_REGION_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "final_region", FINAL_REGION_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "initpref", INITPREF_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "prefgrowi", PREFGROWI_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "prefgrowf", PREFGROWF_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "eqpentyp", EQPENTYP_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "ineqpentyp", INEQPENTYP_VALUE)
            self.add_if_not_empty(convergence_data["conv_options"], "penscl", PENSCL_VALUE)
            convergence_data["tear_data"] = []
            TEAR_NODES = self.get_child_nodes(fr"\Data\Convergence\Tear\Input\TOL")  # 收敛-撕裂-规定
            for tear_stream in TEAR_NODES:
                tear_stream_value = self.safe_get_node_value(fr"\Data\Convergence\Tear\Input\TOL\{tear_stream}")  # 收敛-撕裂-撕裂流股
                convergence_data["tear_data"].append({
                "tear_stream_name": tear_stream,
                "tear_stream_tol": tear_stream_value
            })
            # # 收敛-收敛
            # convergence_data["conv_data"] = []
            # CONV_NODES = self.get_child_nodes(fr"\Data\Convergence\Convergence")  # 收敛节点
            # for conv in CONV_NODES:
            #     conv_type = self.get_block_type(fr"\Data\Convergence\Convergence\{conv}", 6)  # 收敛类型
            #     tear_stream = []
            #     COMPS_NODES = self.get_child_nodes(fr"\Data\Convergence\Convergence\{conv}\Input\COMPS")  # 收敛-流股
            #     for comp in COMPS_NODES:
            #         STATE = self.safe_get_node_value(
            #             fr"\Data\Convergence\Convergence\{conv}\Input\STATE\{comp}")  # 收敛-状态变量
            #         TOL = self.safe_get_node_value(fr"\Data\Convergence\Convergence\{conv}\Input\TOL\{comp}")  # 收敛-允许误差
            #         tear_stream.append({
            #             "stream_id": comp,
            #             "STATE": STATE,
            #             "TOL": TOL
            #         })
            #     convergence_data["conv_data"].append({
            #         "conv_name": conv,
            #         "conv_type": conv_type,
            #         "tear_stream": tear_stream
            #     })
            # #收敛-序列
            # seq_data = []
            # SEQ_NODES = self.get_child_nodes(fr"\Data\Convergence\Sequence")  # 收敛-序列
            # for seq in SEQ_NODES:
            #     sep_type = self.get_block_type(fr"\Data\Convergence\Sequence\{seq}", 6)  # 序列类型
            #     calc_seq = []
            #     BLOCK_ID_NODES = self.get_child_nodes(fr"\Data\Convergence\Sequence\{seq}\Input\BLOCK_ID")  # 序列-计算顺序-模块
            #     for index, block_id_node in enumerate(BLOCK_ID_NODES):
            #         block_id = self.safe_get_node_value(fr"\Data\Convergence\Sequence\{seq}\Input\BLOCK_ID\{block_id_node}")
            #         block_type = self.safe_get_node_value(fr"\Data\Convergence\Sequence\{seq}\Input\BLOCK_TYPE\{block_id_node}")  # # 序列-计算顺序-模块类型
            #         calc_seq.append({
            #             "seq": index,
            #             "block_id": block_id,
            #             "block_type": block_type
            #         })
            #     seq_data.append({
            #         "sep_name": seq,
            #         "sep_type": sep_type,
            #         "calc_seq": calc_seq
            #     })
            # convergence_data["seq_data"] = seq_data
            self.data["convergence"] = convergence_data
            print(f"提取convergence数据完成")
        except Exception as e:
            print(f"提取convergence数据时出错: {e}")
    def extract_reactions_data(self):
        """提取reactions数据"""
        try:
            reactions_data = {}
            Reactions_NODES = self.get_child_nodes(fr"\Data\Reactions\Reactions")  # 反应
            for Reaction in Reactions_NODES:
                reactions_data[Reaction] = {}
                Reaction_TYPE = self.get_block_type(fr"\Data\Reactions\Reactions\{Reaction}", 6)  # 反应类型
                reactions_data[Reaction]["type"] = Reaction_TYPE
                COEF_NODES = self.get_child_nodes(fr"\Data\Reactions\Reactions\{Reaction}\Input\COEF")  # 反应-化学计量-反应物
                # reactions_data[Reaction]["COEF_DATA"] = {}
                reactions_data[Reaction]["REAC_DATA"] = []
                for REAC_ID in COEF_NODES:
                    reac_data = {}
                    REACTYPE = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\REACTYPE\{REAC_ID}")
                    reac_data["REAC_ID"] = REAC_ID
                    reac_data["REACTYPE"] = REACTYPE
                    reac_data["COEF_DATA"] = {}
                    COEF_SUBNODE = self.aspen.Tree.FindNode(fr"\Data\Reactions\Reactions\{Reaction}\Input\COEF\{REAC_ID}")  # 反应-化学计量-反应物
                    COEF_SUBNODES = self.get_child_nodes(fr"\Data\Reactions\Reactions\{Reaction}\Input\COEF\{REAC_ID}")  # 反应-化学计量-反应物
                    UNIQUE_COEF_SUBNODES = list(dict.fromkeys(COEF_SUBNODES))  # 将得到的二维列表去重
                    # 提取所有组分的name
                    components = self.data.get("components", [])
                    component_cids = [comp['cid'] for comp in components]
                    for i, COEF_MIXED_NODE in enumerate(UNIQUE_COEF_SUBNODES):
                        if COEF_MIXED_NODE[:-6] in component_cids: # 暂不提取自定义组分
                            COEF_MIXED_VALUE = COEF_SUBNODE.Elements(0, i).Value
                            reac_data["COEF_DATA"][COEF_MIXED_NODE[:-6]] = COEF_MIXED_VALUE
                    reac_data["COEF1_DATA"] = {}
                    COEF1_SUBNODE = self.aspen.Tree.FindNode(
                        fr"\Data\Reactions\Reactions\{Reaction}\Input\COEF1\{REAC_ID}")  # 反应-化学计量-反应物
                    COEF1_SUBNODES = self.get_child_nodes(
                        fr"\Data\Reactions\Reactions\{Reaction}\Input\COEF1\{REAC_ID}")  # 反应-化学计量-反应物
                    UNIQUE_COEF1_SUBNODES = list(dict.fromkeys(COEF1_SUBNODES))  # 将得到的二维列表去重
                    for i, COEF1_MIXED_NODE in enumerate(UNIQUE_COEF1_SUBNODES):
                        if COEF1_MIXED_NODE[:-6] in component_cids:  # 暂不提取自定义组分
                            COEF1_MIXED_VALUE = COEF1_SUBNODE.Elements(0, i).Value
                            reac_data["COEF1_DATA"][COEF1_MIXED_NODE[:-6]] = COEF1_MIXED_VALUE
                    # 动力学配置
                    PHASE = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\PHASE\{REAC_ID}")  # 动力学-反应相-类型
                    R_D_RBASIS = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\R_D_RBASIS\{REAC_ID}")  # 动力学-速率基准
                    PRE_EXP = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\PRE_EXP\{REAC_ID}")  # 动力学-反应相-K
                    T_EXP = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\T_EXP\{REAC_ID}")  # 动力学-反应相-n
                    ACT_ENERGY_VALUE = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\ACT_ENERGY\{REAC_ID}")  # 动力学-反应相-E
                    ACT_ENERGY_UNITS = self.safe_get_node_units(fr"\Data\Reactions\Reactions\{Reaction}\Input\ACT_ENERGY\{REAC_ID}")  # 动力学-反应相-E
                    T_REF_VALUE = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\T_REF\{REAC_ID}")  # 动力学-反应相-To
                    T_REF_UNITS = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\T_REF\{REAC_ID}")  # 动力学-反应相-To
                    R_D_CBASIS = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\R_D_CBASIS\{REAC_ID}")  # 动力学-反应相-基准
                    OPT_KINETIC = self.safe_get_node_value(fr"\Data\Reactions\Reactions\{Reaction}\Input\OPT_KINETIC")  # 动力学-使用内置幂定律BUILT/用户自定义SUBROUTINE
                    self.add_if_not_empty(reac_data, "PHASE", PHASE)
                    self.add_if_not_empty(reac_data, "R_D_RBASIS", R_D_RBASIS)
                    self.add_if_not_empty(reac_data, "PRE_EXP", PRE_EXP)
                    self.add_if_not_empty(reac_data, "T_EXP", T_EXP)
                    self.add_if_not_empty(reac_data, "ACT_ENERGY_VALUE", ACT_ENERGY_VALUE, "ACT_ENERGY_UNITS", ACT_ENERGY_UNITS)
                    self.add_if_not_empty(reac_data, "T_REF", T_REF_VALUE, "T_REF_UNITS", T_REF_UNITS)
                    self.add_if_not_empty(reac_data, "R_D_CBASIS", R_D_CBASIS)
                    self.add_if_not_empty(reac_data, "OPT_KINETIC", OPT_KINETIC)
                    reactions_data[Reaction]["REAC_DATA"].append(reac_data)
            self.data["reactions"] = reactions_data
            print(f"提取reactions数据完成")
        except Exception as e:
            print(f"提取reactions数据时出错: {e}")
    def extract_design_specs_data(self):
        """提取设计规定(Design-Spec)数据"""
        try:
            design_specs_data = {}
            # 获取所有设计规定节点
            DS_NODES = self.get_child_nodes(r"\Data\Flowsheeting Options\Design-Spec")

            for design_spec in DS_NODES:
                design_specs_data[design_spec] = {}
                base_path = fr"\Data\Flowsheeting Options\Design-Spec\{design_spec}\Input"
                # 1. 提取定义配置
                # 提取样本变量(FVN_*系列)
                design_specs_data[design_spec]["sampled_variables"] = []
                # 检查样本变量定义
                fvn_variable_path = fr"{base_path}\FVN_VARIABLE"
                # 尝试获取样本变量数组
                try:
                    fvn_variable_nodes = self.get_child_nodes(fvn_variable_path)
                    for fvn_variable in fvn_variable_nodes:
                        sampled_var = {}
                        sampled_var["variable_name"] = fvn_variable
                        # 提取其他FVN_*参数
                        fvn_params = [
                            ("OPT_CATEG", "opt_categ"),
                            ("FVN_STREAM", "stream"),
                            ("FVN_VARIABLE", "variable"),
                            ("FVN_COMPONEN", "component"),
                            ("FVN_SUBS", "substream"),
                            ("FVN_VARTYPE", "variable_type"),
                            ("FVN_PHYS_QTY", "physical_quantity"),
                            ("FVN_UOM", "units"),
                            ("FVN_BLOCK", "block"),
                            ("FVN_EO_NAME", "eo_name"),
                            ("FVN_ID1", "id1"),
                            ("FVN_ID2", "id2"),
                            ("FVN_ID3", "id3"),
                            ("FVN_DESCRIPT", "description"),
                            ("FVN_SENTENCE", "sentence"),
                            ("FVN_PARAMNO", "parameter_number"),
                            ("FVN_ATTRIB", "attribute"),
                            ("FVN_ELEM", "element"),
                            ("FVN_PROPSET", "property_set"),
                            ("FVN_INIT_VAL", "initial_value")
                        ]
                        for fvn_path, key in fvn_params:
                            try:
                                node = self.aspen.Tree.FindNode(fr"{base_path}\{fvn_path}")
                                if node is not None:
                                    subnode = node.Elements(f"{fvn_variable}")
                                    if subnode is not None:
                                        value = subnode.Value
                                        if value is not None:
                                            sampled_var[key] = value
                            except:
                                pass
                        # 只添加有内容的采样变量
                        if sampled_var:
                            design_specs_data[design_spec]["sampled_variables"].append(sampled_var)
                except Exception as e:
                    print(f"提取样本变量时出错: {e}")

                # 2. 提取规定配置
                design_specs_data[design_spec]["objective_function"] = {}
                # 提取表达式1
                expr1 = self.safe_get_node_value(fr"{base_path}\EXPR1")
                self.add_if_not_empty(design_specs_data[design_spec]["objective_function"],
                                      "EXPR1", expr1)
                # 提取容差
                tol = self.safe_get_node_value(fr"{base_path}\TOL")
                self.add_if_not_empty(design_specs_data[design_spec]["objective_function"],
                                      "TOL", tol)

                # 提取表达式2
                expr2 = self.safe_get_node_value(fr"{base_path}\EXPR2")
                self.add_if_not_empty(design_specs_data[design_spec]["objective_function"],
                                      "EXPR2", expr2)

                # 3. 提取操纵变量(VARY_*系列)
                design_specs_data[design_spec]["manipulated_variables"] = []

                # 检查操纵变量定义
                vary_variable_path = fr"{base_path}\VARYVARIABLE"

                # 尝试获取操纵变量值
                try:
                    vary_variable_node = self.aspen.Tree.FindNode(vary_variable_path)

                    # 如果VARYVARIABLE是单个值（不是数组）
                    if vary_variable_node is not None:
                        # 检查是否有值
                        try:
                            vary_value = vary_variable_node.Value
                            if vary_value is not None:
                                manipulated_var = {}
                                manipulated_var["variable_name"] = vary_value

                                # 提取其他VARY_*参数（单值版本）
                                vary_params = [
                                    ("VARYBLOCK", "block"),
                                    ("VARYPHYS_QTY", "physical_quantity"),
                                    ("VARY_VARTYPE", "variable_type"),
                                    ("VARYUOM", "units"),
                                    ("VARYSENTENCE", "sentence"),
                                    ("VARYSTREAM", "stream"),
                                    ("VARYCOMPONEN", "component"),
                                    ("VARYPARAMNO", "parameter_number"),
                                    ("VARYINIT_VAL", "initial_value"),
                                    ("VARYID1", "id1"),
                                    ("VARYID2", "id2"),
                                    ("VARYID3", "id3"),
                                    ("VARYDESCRIPT", "description"),
                                    ("VARYELEM", "element"),
                                    ("VARYEO_NAME", "eo_name"),
                                    ("VARYATTRIB", "attribute"),
                                    ("VARYSUBS", "substream"),
                                    ("VARYPROPSET", "property_set")
                                ]

                                for vary_path, key in vary_params:
                                    try:
                                        value = self.safe_get_node_value(fr"{base_path}\{vary_path}")
                                        if value is not None:
                                            manipulated_var[key] = value
                                    except:
                                        pass

                                # 提取VARYLINE1-4（如果有）
                                for line_num in range(1, 5):
                                    line_key = f"VARYLINE{line_num}"
                                    line_value = self.safe_get_node_value(fr"{base_path}\{line_key}")
                                    if line_value is not None:
                                        manipulated_var[f"line{line_num}"] = line_value

                                design_specs_data[design_spec]["manipulated_variables"].append(manipulated_var)
                        except:
                            # 可能是数组形式
                            mpbp_node = vary_variable_node.Elements("MPBP")
                            if mpbp_node is not None:
                                element_count = mpbp_node.Count

                                # 为每个操纵变量提取信息
                                for i in range(element_count):
                                    manipulated_var = {}

                                    # 提取变量名
                                    try:
                                        var_name = mpbp_node.Elements(0, i).Value
                                        manipulated_var["variable_name"] = var_name
                                    except:
                                        pass

                                    # 提取其他VARY_*参数（数组版本）
                                    vary_params = [
                                        ("VARYBLOCK", "block"),
                                        ("VARYPHYS_QTY", "physical_quantity"),
                                        ("VARY_VARTYPE", "variable_type"),
                                        ("VARYUOM", "units"),
                                        ("VARYSTREAM", "stream"),
                                        ("VARYCOMPONEN", "component"),
                                        ("VARYPARAMNO", "parameter_number"),
                                        ("VARYINIT_VAL", "initial_value"),
                                        ("VARYID1", "id1"),
                                        ("VARYID2", "id2"),
                                        ("VARYID3", "id3"),
                                        ("VARYDESCRIPT", "description"),
                                        ("VARYELEM", "element"),
                                        ("VARYEO_NAME", "eo_name"),
                                        ("VARYATTRIB", "attribute"),
                                        ("VARYSUBS", "substream"),
                                        ("VARYPROPSET", "property_set"),
                                        ("VARYSENTENCE", "sentence")
                                    ]

                                    for vary_path, key in vary_params:
                                        try:
                                            node = self.aspen.Tree.FindNode(fr"{base_path}\{vary_path}")
                                            if node is not None:
                                                mpbp_subnode = node.Elements("MPBP")
                                                if mpbp_subnode is not None and i < mpbp_subnode.Count:
                                                    value = mpbp_subnode.Elements(0, i).Value
                                                    if value is not None:
                                                        manipulated_var[key] = value
                                        except:
                                            pass

                                    # 只添加有内容的操纵变量
                                    if manipulated_var:
                                        design_specs_data[design_spec]["manipulated_variables"].append(manipulated_var)
                except Exception as e:
                    print(f"提取操纵变量时出错: {e}")

                # 提取边界和步长设置
                design_specs_data[design_spec]["bounds"] = {}

                # 提取全局边界
                lower = self.safe_get_node_value(fr"{base_path}\LOWER")
                self.add_if_not_empty(design_specs_data[design_spec]["bounds"],
                                      "LOWER", lower)

                upper = self.safe_get_node_value(fr"{base_path}\UPPER")
                self.add_if_not_empty(design_specs_data[design_spec]["bounds"],
                                      "UPPER", upper)

                # 提取步长设置
                step_size = self.safe_get_node_value(fr"{base_path}\STEP_SIZE")
                self.add_if_not_empty(design_specs_data[design_spec]["bounds"],
                                          "STEP_SIZE", step_size)

                max_step_size = self.safe_get_node_value(fr"{base_path}\MAX_STEP_SIZ")
                self.add_if_not_empty(design_specs_data[design_spec]["bounds"],
                                      "MAX_STEP_SIZ", max_step_size)

                # 提取阈值
                threshold = self.safe_get_node_value(fr"{base_path}\THRESHOLD")
                self.add_if_not_empty(design_specs_data[design_spec]["bounds"],
                                      "THRESHOLD", threshold)

            # 将提取的数据保存到类数据中
            self.data["design_specs"] = design_specs_data

            # 打印提取结果统计
            total_specs = len(design_specs_data)
            print(f"提取设计规定数据完成，共找到 {total_specs} 个设计规定")

            for spec_name, spec_data in design_specs_data.items():
                sampled_count = len(spec_data.get("sampled_variables", []))
                manipulated_count = len(spec_data.get("manipulated_variables", []))
                print(f"  {spec_name}: {sampled_count}个采样变量, {manipulated_count}个操纵变量")

            return design_specs_data

        except Exception as e:
            print(f"提取设计规定数据时出错: {e}")
            import traceback
            traceback.print_exc()
            return None
    def extract_block_Mixer_data(self):
        """提取block-Mixer模块数据"""
        try:
            blocks_Mixer_data = {}
            blocks_Mixer = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Mixer":
                    blocks_Mixer.append({
                        "name": block['name'],
                        "type": "Mixer"
                    })
            # 规定提取
            for block in blocks_Mixer:
                blocks_Mixer_data[block['name']] = {}
                try:
                    # Mixer-抽取规定
                    blocks_Mixer_data[block['name']]["SPEC_DATA"] = {}
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 闪蒸选项-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    T_EST_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\T_EST")  # 闪蒸选项-温度估值
                    T_EST_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\T_EST")
                    MAXIT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MAXIT")  # 闪蒸选项-最大迭代次数
                    TOL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TOL")  # 闪蒸选项-容许误差
                    self.add_if_not_empty(blocks_Mixer_data[block['name']]["SPEC_DATA"], "PRES_VALUE", PRES_VALUE, "PRES_UNITS", PRES_UNITS)
                    self.add_if_not_empty(blocks_Mixer_data[block['name']]["SPEC_DATA"], "T_EST_VALUE", T_EST_VALUE, "T_EST_UNITS", T_EST_UNITS)
                    self.add_if_not_empty(blocks_Mixer_data[block['name']]["SPEC_DATA"], "MAXIT_VALUE", MAXIT)
                    self.add_if_not_empty(blocks_Mixer_data[block['name']]["SPEC_DATA"], "TOL_VALUE", TOL)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}规定数据时出错: {e}")
            print(f"提取blocks模块Mixer所有数据完成")
            self.data["blocks_Mixer_data"] = blocks_Mixer_data
        except Exception as e:
            print(f"提取blocks模块blocks_Mixer_data数据时出错: {e}")
    def extract_block_Valve_data(self):
        """提取block-Valve模块数据"""
        try:
            blocks_Valve_data = {}
            blocks_Valve = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Valve":
                    blocks_Valve.append({
                        "name": block['name'],
                        "type": "Valve"
                    })
            # 规定提取
            for block in blocks_Valve:
                blocks_Valve_data[block['name']] = {}
                try:
                    blocks_Valve_data[block['name']]["JOB_DATA"] = {}
                    MODE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MODE")  # 作业-计算类型
                    blocks_Valve_data[block['name']]["JOB_DATA"]["MODE"] = MODE
                    if MODE == "ADIAB-FLASH":  # 当前只抽取指定出口压力下绝热闪蒸，可自行添加
                        P_OUT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\P_OUT")  # 作业-压力规范-出口压力
                        P_OUT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\P_OUT")  # 作业-压力规范-出口压力
                        NPHASE = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\NPHASE")  # 作业-闪蒸选项-有效相态
                        FLASH_MAXIT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLASH_MAXIT")  # 作业-闪蒸选项-最大迭代次数
                        FLASH_TOL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLASH_TOL")  # 作业-闪蒸选项-容许误差
                        self.add_if_not_empty(blocks_Valve_data[block['name']]["JOB_DATA"], "P_OUT_VALUE", P_OUT_VALUE, "P_OUT_UNITS", P_OUT_UNITS)
                        self.add_if_not_empty(blocks_Valve_data[block['name']]["JOB_DATA"], "NPHASE", NPHASE)
                        self.add_if_not_empty(blocks_Valve_data[block['name']]["JOB_DATA"], "FLASH_MAXIT", FLASH_MAXIT)
                        self.add_if_not_empty(blocks_Valve_data[block['name']]["JOB_DATA"], "FLASH_TOL", FLASH_TOL)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Valve所有数据完成")
            self.data["blocks_Valve_data"] = blocks_Valve_data
        except Exception as e:
            print(f"提取blocks模块blocks_Valve_data数据时出错: {e}")
    def extract_block_Compr_data(self):
        """提取block-Compr模块数据"""
        try:
            blocks_Compr_data = {}
            blocks_Compr = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Compr":
                    blocks_Compr.append({
                        "name": block['name'],
                        "type": "Compr"
                    })
            # 规定提取
            for block in blocks_Compr:
                blocks_Compr_data[block['name']] = {}
                try:
                    # Compr-抽取规定、公用工程
                    MODEL_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MODEL_TYPE")  # 规定-模型
                    TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TYPE")  # 规定-类型
                    OPT_SPEC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_SPEC")  # 规定-出口规范
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-排放压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    UTILITY_ID = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\UTILITY_ID")  # 公用工程(放规定一起)
                    blocks_Compr_data[block['name']]["SPEC_DATA"] = {
                        "MODEL_TYPE": MODEL_TYPE,
                        "TYPE": TYPE,
                    }
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Compr_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_Compr_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if OPT_SPEC is not None and OPT_SPEC != "":
                        blocks_Compr_data[block['name']]["SPEC_DATA"]["OPT_SPEC"] = OPT_SPEC
                    if UTILITY_ID is not None and UTILITY_ID != "":
                        blocks_Compr_data[block['name']]["SPEC_DATA"]["UTILITY_ID"] = UTILITY_ID
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Compr所有数据完成")
            self.data["blocks_Compr_data"] = blocks_Compr_data
        except Exception as e:
            print(f"提取blocks模块blocks_Compr_data数据时出错: {e}")
    def extract_block_Heater_data(self):
        """提取block-Heater模块数据"""
        try:
            blocks_Heater_data = {}
            blocks_Heater = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Heater":
                    blocks_Heater.append({
                        "name": block['name'],
                        "type": "Heater"
                    })
            # 规定提取
            for block in blocks_Heater:
                blocks_Heater_data[block['name']] = {}
                try:
                    SPEC_OPT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")  # 规定-温度
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    DELT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DELT")  # 规定-温度变化
                    DELT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DELT")
                    DEGSUP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DEGSUP")  # 规定-过热度
                    DEGSUP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DEGSUP")
                    DEGSUB_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DEGSUB")  # 规定-过冷度
                    DEGSUB_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DEGSUB")
                    VFRAC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VFRAC")  # 规定-汽相分率
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")  # 规定-负载
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    # UTILITY_ID = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\UTILITY_ID")  # 公用工程
                    blocks_Heater_data[block['name']]["SPEC_DATA"] = {
                        "SPEC_OPT": SPEC_OPT
                    }
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    if DELT_VALUE is not None and DELT_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DELT_VALUE"] = DELT_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DELT_UNITS"] = DELT_UNITS
                    if DEGSUP_VALUE is not None and DEGSUP_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DEGSUP_VALUE"] = DEGSUP_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DEGSUP_UNITS"] = DEGSUP_UNITS
                    if DEGSUB_VALUE is not None and DEGSUB_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DEGSUB_VALUE"] = DEGSUB_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DEGSUB_UNITS"] = DEGSUB_UNITS
                    if VFRAC_VALUE is not None and VFRAC_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["VFRAC_VALUE"] = VFRAC_VALUE
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        blocks_Heater_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Heater所有数据完成")
            self.data["blocks_Heater_data"] = blocks_Heater_data
        except Exception as e:
            print(f"提取blocks模块blocks_Heater_data数据时出错: {e}")
    def extract_block_Pump_data(self):
        """提取block-Pump模块数据"""
        try:
            blocks_Pump_data = {}
            blocks_Pump = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Pump":
                    blocks_Pump.append({
                        "name": block['name'],
                        "type": "Pump"
                    })
            # 规定提取
            for block in blocks_Pump:
                blocks_Pump_data[block['name']] = {}
                try:
                    PUMP_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PUMP_TYPE")  # 规定-模型
                    OPT_SPEC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_SPEC")  # 规定-出口规范
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-排放压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    UTILITY_ID = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\UTILITY_ID")  # 公用工程
                    blocks_Pump_data[block['name']]["SPEC_DATA"] = {
                        "PUMP_TYPE": PUMP_TYPE
                    }
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Pump_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_Pump_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if OPT_SPEC is not None and OPT_SPEC != "":
                        blocks_Pump_data[block['name']]["SPEC_DATA"]["OPT_SPEC"] = OPT_SPEC
                    if UTILITY_ID is not None and UTILITY_ID != "":
                        blocks_Pump_data[block['name']]["SPEC_DATA"]["UTILITY_ID"] = UTILITY_ID
                    # blocks_Pump_data[block['name']]["SPEC_DATA"] = {
                    #     "PUMP_TYPE": PUMP_TYPE,
                    #     "PRES_VALUE": PRES_VALUE,
                    #     "PRES_UNITS": PRES_UNITS,
                    #     "UTILITY_ID": UTILITY_ID
                    # }
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Pump所有数据完成")
            self.data["blocks_Pump_data"] = blocks_Pump_data
        except Exception as e:
            print(f"提取blocks模块blocks_Pump_data数据时出错: {e}")
    def extract_block_RStoic_data(self):
        """提取block-RStoic模块数据"""
        try:
            blocks_RStoic_data = {}
            blocks_RStoic = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "RStoic":
                    blocks_RStoic.append({
                        "name": block['name'],
                        "type": "RStoic"
                    })
            # 规定提取
            for block in blocks_RStoic:
                blocks_RStoic_data[block['name']] = {}
                try:
                    # 规定提取
                    SPEC_OPT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")  # 规定-温度
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    DELT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DELT")  # 规定-温度变化
                    DELT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DELT")
                    VFRAC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VFRAC")  # 规定-汽相分率
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")  # 规定-负载
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PHASE")  # 规定-有效相态
                    UTILITY_ID = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\UTILITY_ID")  # 公用工程
                    blocks_RStoic_data[block['name']]["SPEC_DATA"] = {
                        "SPEC_OPT": SPEC_OPT
                    }
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    if DELT_VALUE is not None and DELT_VALUE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["DELT_VALUE"] = DELT_VALUE
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["DELT_UNITS"] = DELT_UNITS
                    if VFRAC_VALUE is not None and VFRAC_VALUE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["VFRAC_VALUE"] = VFRAC_VALUE
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                    if PHASE is not None and PHASE != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["PHASE"] = PHASE
                    if UTILITY_ID is not None and UTILITY_ID != "":
                        blocks_RStoic_data[block['name']]["SPEC_DATA"]["UTILITY_ID"] = UTILITY_ID
                    # 反应提取
                    SERIES = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SERIES")  # 反应-反应连续发生
                    blocks_RStoic_data[block['name']]["REAC_DATA"] = {
                        "SERIES": SERIES
                    }
                    blocks_RStoic_data[block['name']]["REAC_DATA"]["REAC"] = []
                    KEY_SSID_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\KEY_SSID")  # 反应-反应编号
                    for SSID in KEY_SSID_NODES:
                        CONV = self.safe_get_node_value(
                            fr"\Data\Blocks\{block['name']}\Input\CONV\{SSID}")  # 反应-转化率
                        KEY_CID = self.safe_get_node_value(
                            fr"\Data\Blocks\{block['name']}\Input\KEY_CID\{SSID}")  # 反应-组分转化率
                        OPT_EXT_CONV = self.safe_get_node_value(
                            fr"\Data\Blocks\{block['name']}\Input\OPT_EXT_CONV\{SSID}")  # 反应-规范类型
                        EXTENT = self.safe_get_node_value(
                            fr"\Data\Blocks\{block['name']}\Input\EXTENT\{SSID}")  # 反应-摩尔反应进度
                        COEF_DATA = {}
                        COEF_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block['name']}\Input\COEF\{SSID}")  # 反应-化学计量-反应物
                        COEF_MIXED_NODE = self.get_child_nodes(
                            fr"\Data\Blocks\{block['name']}\Input\COEF\{SSID}")  # 反应-化学计量-反应物
                        UNIQUE_COEF_MIXED_NODES = list(dict.fromkeys(COEF_MIXED_NODE)) # 将得到的二维列表去重
                        for i, MIXED_NODE in enumerate(UNIQUE_COEF_MIXED_NODES):
                            COEF_MIXED_VALUE = COEF_NODE.Elements(0, i).Value
                            COEF_DATA[MIXED_NODE[:-6]] = COEF_MIXED_VALUE #最后六位 MIXED无需保留
                        # blocks_RStoic_data[block['name']]["REAC_DATA"][SSID]["COEF1_DATA"] = {}
                        COEF1_DATA = {}
                        COEF1_NODE = self.aspen.Tree.FindNode(fr"\Data\Blocks\{block['name']}\Input\COEF1\{SSID}")  # 反应-化学计量-反应物
                        COEF1_MIXED_NODE = self.get_child_nodes(
                            fr"\Data\Blocks\{block['name']}\Input\COEF1\{SSID}")  # 反应-化学计量-反应物
                        UNIQUE_COEF1_MIXED_NODES = list(dict.fromkeys(COEF1_MIXED_NODE)) # 将得到的二维列表去重
                        for i, MIXED_NODE in enumerate(UNIQUE_COEF1_MIXED_NODES):
                            COEF1_MIXED_VALUE = COEF1_NODE.Elements(0, i).Value
                            # blocks_RStoic_data[block['name']]["REAC_DATA"][SSID]["COEF1_DATA"][MIXED_NODE] = COEF1_MIXED_VALUE
                            COEF1_DATA[MIXED_NODE[:-6]] = COEF1_MIXED_VALUE #最后六位 MIXED无需保留
                        blocks_RStoic_data[block['name']]["REAC_DATA"]["REAC"].append({
                            "KEY_SSID": SSID,
                            "CONV": CONV,
                            "KEY_CID": KEY_CID,
                            "OPT_EXT_CONV": OPT_EXT_CONV,
                            "EXTENT": EXTENT,
                            "COEF_DATA": COEF_DATA,
                            "COEF1_DATA": COEF1_DATA
                        })
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块RStoic所有数据完成")
            self.data["blocks_RStoic_data"] = blocks_RStoic_data
        except Exception as e:
            print(f"提取blocks模块blocks_RStoic_dat数据时出错: {e}")
    def extract_block_RPlug_data(self):
        """提取block-RPlug模块数据"""
        try:
            blocks_RPlug_data = {}
            blocks_RPlug = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "RPlug":
                    blocks_RPlug.append({
                        "name": block['name'],
                        "type": "RPlug"
                    })
            # 规定提取
            for block in blocks_RPlug:
                blocks_RPlug_data[block['name']] = {}
                try:
                    TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TYPE")  # 规定-反应器类型
                    OPT_TSPEC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_TSPEC")  # 规定-操作条件
                    blocks_RPlug_data[block['name']]["SPEC_DATA"] = {
                        "TYPE": TYPE,
                        "OPT_TSPEC": OPT_TSPEC
                    }
                    if OPT_TSPEC == "CONST-TEMP":
                        REAC_TEMP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REAC_TEMP")  # 规定-反应器类型-操作条件-指定反应器温度
                        blocks_RPlug_data[block['name']]["SPEC_DATA"]["REAC_TEMP"] = REAC_TEMP
                    if OPT_TSPEC == "TEMP-PROF":
                        SPEC_TEMP_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\SPEC_TEMP")  # 规定-反应器类型-操作条件-温度分布-温度
                        SPEC_TEMP_DATA = {}
                        for SPEC_TEMP in SPEC_TEMP_NODES:
                            SPEC_TEMP_DATA[SPEC_TEMP] = {}
                            SPEC_TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_TEMP\{SPEC_TEMP}")
                            SPEC_TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SPEC_TEMP\{SPEC_TEMP}")
                            LOC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LOC\{SPEC_TEMP}")  # 规定-反应器类型-操作条件-温度分布-位置
                            if SPEC_TEMP_VALUE is not None and SPEC_TEMP_VALUE != "":
                                SPEC_TEMP_DATA[SPEC_TEMP]["SPEC_TEMP_VALUE"] = SPEC_TEMP_VALUE
                                SPEC_TEMP_DATA[SPEC_TEMP]["SPEC_TEMP_UNITS"] = SPEC_TEMP_UNITS
                            if LOC_VALUE is not None and LOC_VALUE != "":
                                SPEC_TEMP_DATA[SPEC_TEMP]["LOC_VALUE"] = LOC_VALUE
                        # 更新 SPEC_DATA 而不是完全替换，保留 TYPE 和 OPT_TSPEC
                        blocks_RPlug_data[block['name']]["SPEC_DATA"].update(SPEC_TEMP_DATA)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}规定数据时出错: {e}")
                    continue
                try:
                    # 配置提取
                    blocks_RPlug_data[block['name']]["CONFIG_DATA"] = {}
                    CHK_NTUBE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CHK_NTUBE")  # 配置-多管反应器
                    NTUBE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NTUBE")  # 配置-多管反应器-管数
                    LENGTH = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LENGTH")  # 配置-反应器维度-长度
                    DIAM = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DIAM")  # 配置-反应器维度-直径
                    PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PHASE")  # 配置-有效相-工艺流股
                    blocks_RPlug_data[block['name']]["CONFIG_DATA"]["PHASE"] = PHASE
                    if CHK_NTUBE is not None and CHK_NTUBE != "":
                        blocks_RPlug_data[block['name']]["CONFIG_DATA"]["CHK_NTUBE"] = CHK_NTUBE
                    if NTUBE is not None and NTUBE != "":
                        blocks_RPlug_data[block['name']]["CONFIG_DATA"]["NTUBE"] = NTUBE
                    if LENGTH is not None and LENGTH != "":
                        blocks_RPlug_data[block['name']]["CONFIG_DATA"]["LENGTH"] = LENGTH
                    if DIAM is not None and DIAM != "":
                        blocks_RPlug_data[block['name']]["CONFIG_DATA"]["DIAM"] = DIAM
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}配置数据时出错: {e}")
                    continue
                try:
                    #反应提取
                    REACSYS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REACSYS")  # 反应-反应体系
                    RXN_ID_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\RXN_ID")  # 反应-所选反应集
                    RXN_ID_DATA = {}
                    for RXN_ID in RXN_ID_NODES:
                        RXN_ID_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RXN_ID\{RXN_ID}")
                        RXN_ID_DATA[RXN_ID] = RXN_ID_VALUE
                    blocks_RPlug_data[block['name']]["REAC_DATA"] = {
                        "REACSYS": REACSYS,
                        "RXN_ID": RXN_ID_DATA
                    }
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}反应配置时出错: {e}")
                    continue
                try:
                    # 压力提取
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 压力-进口压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 压力-进口压力
                    OPT_PDROP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_PDROP ")  # 压力-通过反应器的压降
                    PDROP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PDROP ")  # 压力-压降-工艺流股
                    PDROP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PDROP ")  # 压力-压降-工艺流股
                    ROUGHNESS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\ROUGHNESS ")  # 压力-摩擦关联式-粗糙度
                    ROUGHNESS_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\ROUGHNESS ")  # 压力-摩擦关联式-粗糙度
                    DP_FCOR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DP_FCOR")  # 压力-摩擦关联式-压降关联式
                    DP_MULT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DP_MULT")  # 压力-摩擦关联式-压降比例因子
                    blocks_RPlug_data[block['name']]["PRES_DATA"] = {
                        "OPT_PDROP": OPT_PDROP
                    }
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if PDROP_VALUE is not None and PDROP_VALUE != "":
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["PDROP_VALUE"] = PDROP_VALUE
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["PDROP_UNITS"] = PDROP_UNITS
                    if ROUGHNESS_VALUE is not None and ROUGHNESS_VALUE != "":
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["ROUGHNESS_VALUE"] = ROUGHNESS_VALUE
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["ROUGHNESS_UNITS"] = ROUGHNESS_UNITS
                    if DP_FCOR is not None and DP_FCOR != "":
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["DP_FCOR"] = DP_FCOR
                    if DP_MULT is not None and DP_MULT != "":
                        blocks_RPlug_data[block['name']]["PRES_DATA"]["DP_MULT"] = DP_MULT
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}压力数据时出错: {e}")
                    continue
                try:
                    #催化剂
                    blocks_RPlug_data[block['name']]["CAT_DATA"] = {}
                    CAT_PRESENT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CAT_PRESENT")  # 催化剂-反应器内的催化剂
                    IGN_CAT_VOL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\IGN_CAT_VOL")  # 催化剂-忽略催化器体积
                    BED_VOIDAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BED_VOIDAGE")  # 催化剂-规定-床空隙率
                    CAT_RHO_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CAT_RHO")  # 催化剂-规定-颗粒密度
                    CAT_RHO_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CAT_RHO")  # 催化剂-规定-颗粒密度
                    CATWT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CATWT")  # 催化剂-规定-催化剂装填
                    CATWT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CATWT")  # 催化剂-规定-催化剂装填
                    if CAT_PRESENT is not None and CAT_PRESENT != "":
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["CAT_PRESENT"] = CAT_PRESENT
                    if IGN_CAT_VOL is not None and IGN_CAT_VOL != "":
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["IGN_CAT_VOL"] = IGN_CAT_VOL
                    if BED_VOIDAGE is not None and BED_VOIDAGE != "":
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["BED_VOIDAGE"] = BED_VOIDAGE
                    if CAT_RHO_VALUE is not None and CAT_RHO_VALUE != "":
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["CAT_RHO_VALUE"] = CAT_RHO_VALUE
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["CAT_RHO_UNITS"] = CAT_RHO_UNITS
                    if CATWT_VALUE is not None and CATWT_VALUE != "":
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["CATWT_VALUE"] = CATWT_VALUE
                        blocks_RPlug_data[block['name']]["CAT_DATA"]["CATWT_UNITS"] = CATWT_UNITS
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}催化剂数据时出错: {e}")
                    continue
            print(f"提取blocks模块RPlug所有数据完成")
            self.data["blocks_RPlug_data"] = blocks_RPlug_data
        except Exception as e:
            print(f"提取blocks模块blocks_RPlug_data数据时出错: {e}")
    def extract_block_Flash2_data(self):
        """提取block-Flash2模块数据"""
        try:
            blocks_Flash2_data = {}
            blocks_Flash2 = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Flash2":
                    blocks_Flash2.append({
                        "name": block['name'],
                        "type": "Flash2"
                    })
            # 规定提取
            for block in blocks_Flash2:
                blocks_Flash2_data[block['name']] = {}
                try:
                    SPEC_OPT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")  # 规定-温度
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    DELT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DELT")  # 规定-温度变化
                    DELT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DELT")
                    VFRAC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VFRAC")  # 规定-汽相分率
                    VFRAC_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\VFRAC")
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")  # 规定-负载
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    UTILITY_ID = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\UTILITY_ID")  # 公用工程
                    blocks_Flash2_data[block['name']]["SPEC_DATA"] = {
                        "SPEC_OPT": SPEC_OPT
                    }
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    if DELT_VALUE is not None and DELT_VALUE != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["DELT_VALUE"] = DELT_VALUE
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["DELT_UNITS"] = DELT_UNITS
                    if VFRAC_VALUE is not None and VFRAC_VALUE != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["VFRAC_VALUE"] = VFRAC_VALUE
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["VFRAC_UNITS"] = VFRAC_UNITS
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                    if UTILITY_ID is not None and UTILITY_ID != "":
                        blocks_Flash2_data[block['name']]["SPEC_DATA"]["UTILITY_ID"] = UTILITY_ID
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Flash2所有数据完成")
            self.data["blocks_Flash2_data"] = blocks_Flash2_data
        except Exception as e:
            print(f"提取blocks模块blocks_Flash3_data数据时出错: {e}")
    def extract_block_Flash3_data(self):
        """提取block-Flash3模块数据"""
        try:
            blocks_Flash3_data = {}
            blocks_Flash3 = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Flash3":
                    blocks_Flash3.append({
                        "name": block['name'],
                        "type": "Flash3"
                    })
            # 规定提取
            for block in blocks_Flash3:
                blocks_Flash3_data[block['name']] = {}
                try:
                    SPEC_OPT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_OPT")  # 规定-闪蒸计算类型
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")  # 规定-温度
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")  # 规定-负载
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    VFRAC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VFRAC")  # 规定-汽相分率
                    L2_COMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L2_COMP")

                    blocks_Flash3_data[block['name']]["SPEC_DATA"] = {
                        "SPEC_OPT": SPEC_OPT
                    }
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_Flash3_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        if TEMP_UNITS is not None and TEMP_UNITS != "":
                            blocks_Flash3_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Flash3_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        if PRES_UNITS is not None and PRES_UNITS != "":
                            blocks_Flash3_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_Flash3_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        if DUTY_UNITS is not None and DUTY_UNITS != "":
                            blocks_Flash3_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                    if VFRAC_VALUE is not None and VFRAC_VALUE != "":
                        blocks_Flash3_data[block['name']]["SPEC_DATA"]["VFRAC_VALUE"] = VFRAC_VALUE
                    if L2_COMP_VALUE is not None and L2_COMP_VALUE != "":
                        blocks_Flash3_data[block['name']]["SPEC_DATA"]["L2_COMP"] = L2_COMP_VALUE
                    
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Flash3所有数据完成")
            self.data["blocks_Flash3_data"] = blocks_Flash3_data
        except Exception as e:
            print(f"提取blocks模块Flash3数据时出错: {e}")
            self.data["blocks_Flash3_data"] = {}
    def extract_block_Decanter_data(self):
        """提取Decanter模块数据"""
        try:
            blocks_Decanter_data = {}
            blocks_Decanter = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Decanter":
                    blocks_Decanter.append({
                        "name": block['name'],
                        "type": "Decanter"
                    })
            # 规定提取
            for block in blocks_Decanter:
                blocks_Decanter_data[block['name']] = {}
                try:
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")  # 规定-倾析器规范-温度
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")  # 规定-倾析器规范-压力
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")  # 规定-倾析器规范-负荷
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    L2_CUTOFF = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L2_CUTOFF") # 规定-第二液相的组分摩尔分率
                    L2_COMPS_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\L2_COMPS") # 规定-第二液相的关键组分
                    blocks_Decanter_data[block['name']]["SPEC_DATA"] = {}
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                    if L2_CUTOFF is not None and L2_CUTOFF != "":
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["L2_CUTOFF"] = L2_CUTOFF
                    blocks_Decanter_data[block['name']]["SPEC_DATA"]["L2_COMPS"] = []
                    for L2_COMPS in L2_COMPS_NODES:
                        L2_COMPS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L2_COMPS\{L2_COMPS}")
                        blocks_Decanter_data[block['name']]["SPEC_DATA"]["L2_COMPS"].append(L2_COMPS_VALUE)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Decanter所有数据完成")
            self.data["blocks_Decanter_data"] = blocks_Decanter_data
        except Exception as e:
            print(f"提取blocks模块blocks_Decanter_data数据时出错: {e}")
    def extract_block_Sep_data(self):
        """提取block-Sep模块数据"""
        try:
            blocks_Sep_data = {}
            blocks_Sep = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Sep":
                    blocks_Sep.append({
                        "name": block['name'],
                        "type": "Sep"
                    })
            # 规定提取
            for block in blocks_Sep:
                blocks_Sep_data[block['name']] = {}
                try:
                    blocks_Sep_data[block['name']]["SPEC_DATA"] = {}
                    FLOW_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS")
                    for FLOW in FLOW_NODES:
                        blocks_Sep_data[block['name']]["SPEC_DATA"][FLOW] = []
                        # 提取所有组分ID
                        components = self.data.get("components", [])
                        component_cids = [comp['cid'] for comp in components]
                        COMP_ID_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS\{FLOW}\MIXED")
                        for COMP_ID in COMP_ID_NODES:
                            if COMP_ID in component_cids:  # 自定义组分的配置不要提取
                                FLOW_COMP_DATA = {"COMP_ID": COMP_ID}
                                FLOWBASIS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS\{FLOW}\MIXED\{COMP_ID}") # 规定-出口流股条件-基准
                                FRACS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FRACS\{FLOW}\MIXED\{COMP_ID}")  # 规定-出口流股条件-规定-分流分率
                                FLOWS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLOWS\{FLOW}\MIXED\{COMP_ID}")  # 规定-出口流股条件-规定-流量
                                FLOWS_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\FRACS\{FLOW}\MIXED\{COMP_ID}")  # 规定-出口流股条件-规定-流量
                                FLOWS_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\FRACS\{FLOW}\MIXED\{COMP_ID}", 13)  # 规定-出口流股条件-规定-流量
                                self.add_if_not_empty(FLOW_COMP_DATA, "FLOWBASIS", FLOWBASIS)
                                self.add_if_not_empty(FLOW_COMP_DATA, "FRACS", FRACS)
                                self.add_if_not_empty(FLOW_COMP_DATA, "FLOWS_VALUE", FLOWS_VALUE, "FLOWS_UNITS", FLOWS_UNITS, "FLOWS_BASIS", FLOWS_BASIS)
                                blocks_Sep_data[block['name']]["SPEC_DATA"][FLOW].append(FLOW_COMP_DATA)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Sep所有数据完成")
            self.data["blocks_Sep_data"] = blocks_Sep_data
        except Exception as e:
            print(f"提取blocks模块blocks_Sep_data数据时出错: {e}")
    def extract_block_Sep2_data(self):
        """提取block-Sep2模块数据"""
        try:
            blocks_Sep2_data = {}
            blocks_Sep2 = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Sep2":
                    blocks_Sep2.append({
                        "name": block['name'],
                        "type": "Sep2"
                    })
            # 规定提取
            for block in blocks_Sep2:
                blocks_Sep2_data[block['name']] = {}
                try:
                    blocks_Sep2_data[block['name']]["SPEC_DATA"] = {}
                    FLOW_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS\MIXED") #出口流股
                    for FLOW in FLOW_NODES:
                        blocks_Sep2_data[block['name']]["SPEC_DATA"][FLOW] = []
                        # 提取所有组分ID
                        components = self.data.get("components", [])
                        component_cids = [comp['cid'] for comp in components]
                        COMP_ID_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS\MIXED\{FLOW}")
                        for COMP_ID in COMP_ID_NODES:
                            if COMP_ID in component_cids:  # 自定义组分的配置不要提取
                                FLOW_COMP_DATA = {"COMP_ID": COMP_ID}
                                FLOWBASIS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLOWBASIS\MIXED\{FLOW}\{COMP_ID}") # 规定-出口流股条件-基准
                                FRACS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FRACS\MIXED\{FLOW}\{COMP_ID}")  # 规定-出口流股条件-规定-分流分率
                                FLOWS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FLOWS\MIXED\{FLOW}\{COMP_ID}")  # 规定-出口流股条件-规定-流量
                                FLOWS_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\FRACS\MIXED\{FLOW}\{COMP_ID}")  # 规定-出口流股条件-规定-流量
                                FLOWS_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\FRACS\MIXED\{FLOW}\{COMP_ID}", 13)  # 规定-出口流股条件-规定-流量
                                self.add_if_not_empty(FLOW_COMP_DATA, "FLOWBASIS", FLOWBASIS)
                                self.add_if_not_empty(FLOW_COMP_DATA, "FRACS", FRACS)
                                self.add_if_not_empty(FLOW_COMP_DATA, "FLOWS_VALUE", FLOWS_VALUE, "FLOWS_UNITS", FLOWS_UNITS, "FLOWS_BASIS", FLOWS_BASIS)
                                blocks_Sep2_data[block['name']]["SPEC_DATA"][FLOW].append(FLOW_COMP_DATA)
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Sep2所有数据完成")
            self.data["blocks_Sep2_data"] = blocks_Sep2_data
        except Exception as e:
            print(f"提取blocks模块blocks_Sep2_data数据时出错: {e}")
    def extract_block_RadFrac_data(self):
        """提取block-RadFrac模块数据"""
        try:
            blocks_RadFrac_data = {}
            blocks_RadFrac = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "RadFrac":
                    blocks_RadFrac.append({
                        "name": block['name'],
                        "type": "RadFrac"
                    })
            # 规定提取
            for block in blocks_RadFrac:
                blocks_RadFrac_data[block['name']] = {}
                try:
                    #配置抽取
                    CALC_MODE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CALC_MODE")  # 配置-计算类型
                    NSTAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSTAGE")  # 配置-塔板数
                    CONDENSER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CONDENSER") #配置-冷凝器
                    REBOILER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REBOILER") #配置-再沸器
                    NO_PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NO_PHASE") #配置-有效相态
                    BLKOPFREWAT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BLKOPFREWAT") #配置-有效相态
                    CONV_METH = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CONV_METH") #配置-收敛
                    BASIS_RR_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_RR") #配置-操作规范-回流比
                    BASIS_RR_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_RR", 13) #配置-操作规范-回流比
                    BASIS_L1_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_L1") #配置-操作规范-回流速率
                    BASIS_L1_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_L1") #配置-操作规范-回流速率
                    BASIS_L1_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_L1", 13) #配置-操作规范-回流速率
                    BASIS_D_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_D") #配置-操作规范-馏出物流率
                    BASIS_D_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_D") #配置-操作规范-馏出物流率
                    BASIS_D_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_D", 13) #配置-操作规范-馏出物流率
                    BASIS_B_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_B") #配置-操作规范-塔底物流率
                    BASIS_B_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_B") #配置-操作规范-塔底物流率
                    BASIS_B_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_B", 13) #配置-操作规范-塔底物流率
                    BASIS_VN_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_VN") #配置-操作规范-再沸蒸汽流速
                    BASIS_VN_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_VN") #配置-操作规范-再沸蒸汽流速
                    BASIS_VN_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_VN", 13) #配置-操作规范-再沸蒸汽流速
                    BASIS_BR_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_BR") #配置-操作规范-再沸比
                    BASIS_BR_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_L1", 13) #配置-操作规范-再沸比
                    Q1_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\Q1") #配置-操作规范-冷凝器负荷
                    Q1_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\Q1") #配置-操作规范-冷凝器负荷
                    QN_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\QN") #配置-操作规范-再沸器负荷
                    QN_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\QN") #配置-操作规范-再沸器负荷
                    DF_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\D:F") #配置-操作规范-馏出物进料比
                    DF_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\D:F", 13) #配置-操作规范-馏出物进料比
                    BF_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\B:F") #配置-操作规范-馏出物进料比
                    BF_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\B:F", 13) #配置-操作规范-馏出物进料比
                    # RW = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RW") #配置-自由水回流比  暂不需要
                    blocks_RadFrac_data[block['name']]['CONFIG_DATA'] = {
                        "CALC_MODE": CALC_MODE
                    }
                    # 配置-设置选项
                    if NSTAGE is not None and NSTAGE != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["NSTAGE"] = NSTAGE
                    if CONDENSER is not None and CONDENSER != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["CONDENSER"] = CONDENSER
                    if REBOILER is not None and REBOILER != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["REBOILER"] = REBOILER
                    if CONV_METH is not None and CONV_METH != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["CONV_METH"] = CONV_METH
                    if NO_PHASE is not None and NO_PHASE != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["NO_PHASE"] = NO_PHASE
                    if BLKOPFREWAT is not None and BLKOPFREWAT != "":
                        blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["BLKOPFREWAT"] = BLKOPFREWAT
                    # 配置-操作规范
                    blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"] = []
                    if BASIS_RR_VALUE is not None and BASIS_RR_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_RR_VALUE": BASIS_RR_VALUE,
                            "BASIS_RR_BASIS": BASIS_RR_BASIS
                        })
                    if BASIS_L1_VALUE is not None and BASIS_L1_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_L1_VALUE": BASIS_L1_VALUE,
                            "BASIS_L1_UNITS": BASIS_L1_UNITS,
                            "BASIS_L1_BASIS": BASIS_L1_BASIS
                        })
                    if BASIS_D_VALUE is not None and BASIS_D_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_D_VALUE": BASIS_D_VALUE,
                            "BASIS_D_UNITS": BASIS_D_UNITS,
                            "BASIS_D_BASIS": BASIS_D_BASIS
                        })
                    if BASIS_B_VALUE is not None and BASIS_B_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_B_VALUE": BASIS_B_VALUE,
                            "BASIS_B_UNITS": BASIS_B_UNITS,
                            "BASIS_B_BASIS": BASIS_B_BASIS
                        })
                    if BASIS_VN_VALUE is not None and BASIS_VN_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_VN_VALUE": BASIS_VN_VALUE,
                            "BASIS_VN_UNITS": BASIS_VN_UNITS,
                            "BASIS_VN_BASIS": BASIS_VN_BASIS
                        })
                    if BASIS_BR_VALUE is not None and BASIS_BR_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BASIS_BR_VALUE": BASIS_BR_VALUE,
                            "BASIS_BR_BASIS": BASIS_BR_BASIS
                        })
                    if Q1_VALUE is not None and Q1_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "Q1_VALUE": Q1_VALUE,
                            "Q1_UNITS": Q1_UNITS
                        })
                    if QN_VALUE is not None and QN_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "QN_VALUE": QN_VALUE,
                            "QN_UNITS": QN_UNITS
                        })
                    if DF_VALUE is not None and DF_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "DF_VALUE": DF_VALUE,
                            "DF_BASIS": DF_BASIS
                        })
                    if BF_VALUE is not None and BF_VALUE != "":
                        blocks_RadFrac_data[block['name']]['CONFIG_DATA']["OP_SPEC"].append({
                            "BF_VALUE": BF_VALUE,
                            "BF_BASIS": BF_BASIS
                        })
                    # if RW is not None and RW != "" and RW != 0:
                    #     blocks_RadFrac_data[block['name']]["CONFIG_DATA"]["RW"] = RW
                    #流股抽取
                    FEED_STAGE_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FEED_STAGE") #流股-进料流股
                    FEED_STAGE_DATA = []
                    for FEED_STAGE in FEED_STAGE_NODES:
                        FEED_STAGE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FEED_STAGE\{FEED_STAGE}")
                        FEED_CONVEN = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FEED_CONVEN\{FEED_STAGE}")  #流股-进料流股-常规
                        FEED_STAGE_DATA.append({
                            "FEED_STAGE": FEED_STAGE,
                            "FEED_STAGE_VALUE": FEED_STAGE_VALUE,
                            "FEED_CONVEN": FEED_CONVEN
                        })
                    blocks_RadFrac_data[block['name']]['FEED_STAGE_DATA'] = FEED_STAGE_DATA
                    PROD_STAGE_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\PROD_STAGE") #流股-产品流股
                    PROD_STAGE_DATA = []
                    for PROD_STAGE in PROD_STAGE_NODES:
                        PROD_STAGE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_STAGE\{PROD_STAGE}") #流股-产品流股-塔板
                        PROD_PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_PHASE\{PROD_STAGE}")  #流股-产品流股-相态
                        PROD_STAGE_DATA.append({
                            "PROD_STAGE": PROD_STAGE,
                            "PROD_STAGE_VALUE": PROD_STAGE_VALUE,
                            "PROD_PHASE": PROD_PHASE
                        })
                    blocks_RadFrac_data[block['name']]['PROD_STAGE_DATA'] = PROD_STAGE_DATA
                    #压力抽取
                    VIEW_PRES = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VIEW_PRES")  # 压力-查看
                    blocks_RadFrac_data[block['name']]['PRES_DATA'] = {}
                    blocks_RadFrac_data[block['name']]['PRES_DATA']["VIEW_PRES"] = VIEW_PRES
                    if VIEW_PRES == "TOP/BOTTOM": #压力-查看-塔顶塔底
                        VIEW_PRES_DATA = []
                        PRES1_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES1")  # 压力-查看-塔板1压力
                        PRES1_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES1")  # 压力-查看-塔板1压力
                        PRES2_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES2")  # 压力-查看-塔板2压力
                        PRES2_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES2")  # 压力-查看-塔板2压力
                        OPT_PRES_TOP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_PRES_TOP")  # 压力-查看-塔板2压力-选项
                        DP_COND_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DP_COND")  # 压力-查看-塔板2压力-冷凝器压降
                        DP_COND_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DP_COND")  # 压力-查看-塔板2压力-冷凝器压降
                        OPT_PRES = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_PRES")  # 压力-查看-塔其余部分压降-选项
                        DP_STAGE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DP_STAGE")  # 压力-查看-塔其余部分压降-塔板压降
                        DP_STAGE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DP_STAGE")  # 压力-查看-塔其余部分压降-塔板压降
                        DP_COL_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DP_COL")  # 压力-查看-塔其余部分压降-塔压降
                        DP_COL_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DP_COL")  # 压力-查看-塔其余部分压降-塔压降
                        VIEW_PRES_DATA.append({
                            "PRES1_VALUE": PRES1_VALUE,
                            "PRES1_UNITS": PRES1_UNITS,
                            "OPT_PRES_TOP": OPT_PRES_TOP,
                            "PRES2_VALUE": PRES2_VALUE,
                            "PRES2_UNITS": PRES2_UNITS,
                            "DP_COND_VALUE": DP_COND_VALUE,
                            "DP_COND_UNITS": DP_COND_UNITS,
                            "OPT_PRES": OPT_PRES,
                            "DP_STAGE_VALUE": DP_STAGE_VALUE,
                            "DP_STAGE_UNITS": DP_STAGE_UNITS,
                            "DP_COL_VALUE": DP_COL_VALUE,
                            "DP_COL_UNITS": DP_COL_UNITS
                        })
                        blocks_RadFrac_data[block['name']]['PRES_DATA']["STAGE_PRES"] = VIEW_PRES_DATA
                    if VIEW_PRES == "PROFILE": #压力-查看-压力分布
                        STAGE_PRES_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES")  # 压力-查看-压力分布
                        STAGE_PRES_DATA = []
                        for PRES_STAGE in STAGE_PRES_NODES:
                            STAGE_PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES\{PRES_STAGE}")
                            STAGE_PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES\{PRES_STAGE}")
                            STAGE_PRES_DATA.append({
                                "PRES_STAGE": PRES_STAGE,
                                "PRES_VALUE": STAGE_PRES_VALUE,
                                "PRES_UNITS": STAGE_PRES_UNITS
                            })
                        blocks_RadFrac_data[block['name']]['PRES_DATA']["STAGE_PRES"] = STAGE_PRES_DATA
                    #if view_pres == "PDROP":  # 压力-查看-塔段压降  暂未实现

                    # 冷凝器抽取
                    if CONDENSER != "NONE":
                        CONDENSER_DATA = {}
                        OPT_COND_SPC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_COND_SPC")  # 冷凝器-冷凝器规范
                        T1_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\T1")  # 冷凝器-冷凝器规范-温度
                        T1_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\T1")  # 冷凝器-冷凝器规范-温度
                        BASIS_RDV_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_RDV")  # 冷凝器-冷凝器规范-馏出物汽相分率
                        BASIS_RDV_BASIS = self.get_block_type(fr"\Data\Blocks\{block['name']}\Input\BASIS_RDV", 13)  # 冷凝器-冷凝器规范-馏出物汽相分率
                        SC_TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SC_TEMP")  # 冷凝器-冷凝器规范-过冷规范-过冷温度
                        SC_TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SC_TEMP")  # 冷凝器-冷凝器规范-过冷规范-过冷温度
                        SC_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SC_OPTION")  # 冷凝器-冷凝器规范
                        self.add_if_not_empty(CONDENSER_DATA, "OPT_COND_SPC", OPT_COND_SPC)
                        self.add_if_not_empty(CONDENSER_DATA, "T1_VALUE", T1_VALUE,"T1_UNITS", T1_UNITS)
                        self.add_if_not_empty(CONDENSER_DATA, "BASIS_RDV_VALUE", BASIS_RDV_VALUE, None, None, "BASIS_RDV_BASIS", BASIS_RDV_BASIS)
                        self.add_if_not_empty(CONDENSER_DATA, "SC_TEMP_VALUE", SC_TEMP_VALUE, "SC_TEMP_UNITS", SC_TEMP_UNITS)
                        self.add_if_not_empty(CONDENSER_DATA, "SC_OPTION", SC_OPTION)
                        blocks_RadFrac_data[block['name']]['CONDENSER_DATA'] = CONDENSER_DATA
                    # 规定-设计规范抽取
                    blocks_RadFrac_data[block['name']]['DESIGN_SPEC_DATA'] = {}
                    design_spec_node = self.get_child_nodes(
                        fr"\Data\Blocks\{block['name']}\Subobjects\Design Specs")
                    base_node = fr"\Data\Blocks\{block['name']}\Subobjects\Design Specs"
                    design_spec_data = []
                    for design_spec_id in design_spec_node:
                        SPEC_VALUE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\VALUE\{design_spec_id}")
                        # 提取 SPEC_VALUE 的单位
                        SPEC_VALUE_UNITS = self.safe_get_node_units(
                            fr"{base_node}\{design_spec_id}\Input\VALUE\{design_spec_id}")
                        SPEC_TYPE_VALUE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_TYPE\{design_spec_id}")
                        OPT_SPC_STR_VALUE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_STR\{design_spec_id}")
                        comp_data = []
                        COMPS_NODES = self.get_child_nodes(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_COMPS\{design_spec_id}")
                        for comp_id in COMPS_NODES:
                            comp_value = self.safe_get_node_value(
                                fr"{base_node}\{design_spec_id}\Input\SPEC_COMPS\{design_spec_id}\{comp_id}")
                            comp_data.append(comp_value)
                        spec_streams_data = []
                        SPEC_STREAMS_NODES = self.get_child_nodes(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_STREAMS\{design_spec_id}")
                        for spec_stream_id in SPEC_STREAMS_NODES:
                            spec_stream_value = self.safe_get_node_value(
                                fr"{base_node}\{design_spec_id}\Input\SPEC_STREAMS\{design_spec_id}\{spec_stream_id}")
                            spec_streams_data.append(spec_stream_value)
                        
                        # 提取新增的设计规范单个值参数
                        OPT_SPC_RAT = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_RAT\{design_spec_id}")
                        BASE_PHASE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\BASE_PHASE\{design_spec_id}")
                        BASE_STAGE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\BASE_STAGE\{design_spec_id}")
                        OPT_SPC_PRP1 = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_PRP1\{design_spec_id}")
                        PROPERTY = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\PROPERTY\{design_spec_id}")
                        OPT_SPC_PRP2 = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\OPT_SPC_PRP2\{design_spec_id}")
                        BASE_PROPERT = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\BASE_PROPERT\{design_spec_id}")
                        
                        # 提取缺失的设计规范参数
                        SPEC_DESCRIP = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_DESCRIP\{design_spec_id}")
                        SPEC_STAGE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_STAGE\{design_spec_id}")
                        SPEC_PHASE = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\SPEC_PHASE\{design_spec_id}")
                        SP_DEC_STRM = self.safe_get_node_value(
                            fr"{base_node}\{design_spec_id}\Input\SP_DEC_STRM\{design_spec_id}")
                        design_spec_data.append({
                            "SPEC_ID": design_spec_id,
                            "SPEC_VALUE": SPEC_VALUE,
                            "SPEC_VALUE_UNITS": SPEC_VALUE_UNITS,
                            "SPEC_TYPE_VALUE": SPEC_TYPE_VALUE,
                            "OPT_SPC_STR_VALUE": OPT_SPC_STR_VALUE,
                            "SPEC_DESCRIP": SPEC_DESCRIP,
                            "SPEC_STAGE": SPEC_STAGE,
                            "SPEC_PHASE": SPEC_PHASE,
                            "SP_DEC_STRM": SP_DEC_STRM,
                            "COMP_DATA": comp_data,
                            "SPEC_STREAMS": spec_streams_data,
                            "OPT_SPC_RAT": OPT_SPC_RAT,
                            "BASE_PHASE": BASE_PHASE,
                            "BASE_STAGE": BASE_STAGE,
                            "OPT_SPC_PRP1": OPT_SPC_PRP1,
                            "PROPERTY": PROPERTY,
                            "OPT_SPC_PRP2": OPT_SPC_PRP2,
                            "BASE_PROPERT": BASE_PROPERT
                        })
                        blocks_RadFrac_data[block['name']]['DESIGN_SPEC_DATA'] = design_spec_data
                    # 规定-变化抽取
                    blocks_RadFrac_data[block['name']]['VARY_DATA'] = {}
                    vary_node = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Subobjects\Vary")
                    base_node = fr"\Data\Blocks\{block['name']}\Subobjects\Vary"
                    vary_data = []
                    for vary_id in vary_node:
                        VAR_VALUE = self.safe_get_node_value(fr"{base_node}\{vary_id}\Input\VALUE\{vary_id}")
                        VARTYPE_VALUE = self.safe_get_node_value(fr"{base_node}\{vary_id}\Input\VARTYPE\{vary_id}")
                        LB_VALUE = self.safe_get_node_value(fr"{base_node}\{vary_id}\Input\LB\{vary_id}")
                        UB_VALUE = self.safe_get_node_value(fr"{base_node}\{vary_id}\Input\UB\{vary_id}")
                        STEP_VALUE = self.safe_get_node_value(fr"{base_node}\{vary_id}\Input\STEP\{vary_id}")
                        # 提取有单位的参数的单位信息
                        VARY_VALUE_UNITS = self.safe_get_node_units(fr"{base_node}\{vary_id}\Input\VALUE\{vary_id}")
                        LB_UNITS = self.safe_get_node_units(fr"{base_node}\{vary_id}\Input\LB\{vary_id}")
                        UB_UNITS = self.safe_get_node_units(fr"{base_node}\{vary_id}\Input\UB\{vary_id}")
                        # 提取缺失的变化参数
                        VARY_DESCRIP = self.safe_get_node_value(
                            fr"{base_node}\{vary_id}\Input\VARY_DESCRIP\{vary_id}")
                        comp_data = []
                        COMPS_NODES = self.get_child_nodes(fr"{base_node}\{vary_id}\Input\VARY_COMPS\{vary_id}")
                        for comp_id in COMPS_NODES:
                            comp_value = self.safe_get_node_value(
                                fr"{base_node}\{vary_id}\Input\Vary_COMPS\{vary_id}\{comp_id}")
                            comp_data.append(comp_value)
                        
                        # 提取新增的变化单个值参数
                        VARY_STAGE = self.safe_get_node_value(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE\{vary_id}")
                        VARY_STREAM = self.safe_get_node_value(
                            fr"{base_node}\{vary_id}\Input\VARY_STREAM\{vary_id}")
                        VARY_STAGE1 = self.safe_get_node_value(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE1\{vary_id}")
                        VARY_STAGE2 = self.safe_get_node_value(
                            fr"{base_node}\{vary_id}\Input\VARY_STAGE2\{vary_id}")
                        
                        vary_data.append({
                            "VARY_ID": vary_id,
                            "VARY_VALUE": VAR_VALUE,
                            "VARY_VALUE_UNITS": VARY_VALUE_UNITS,
                            "VARTYPE_VALUE": VARTYPE_VALUE,
                            "LB_VALUE": LB_VALUE,
                            "LB_UNITS": LB_UNITS,
                            "UB_VALUE": UB_VALUE,
                            "UB_UNITS": UB_UNITS,
                            "STEP_VALUE": STEP_VALUE,
                            "VARY_DESCRIP": VARY_DESCRIP,
                            "COMP_DATA": comp_data,
                            "VARY_STAGE": VARY_STAGE,
                            "VARY_STREAM": VARY_STREAM,
                            "VARY_STAGE1": VARY_STAGE1,
                            "VARY_STAGE2": VARY_STAGE2
                        })
                        blocks_RadFrac_data[block['name']]['VARY_DATA'] = vary_data

                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块RadFrac所有数据完成")
            self.data["blocks_RadFrac_data"] = blocks_RadFrac_data
        except Exception as e:
            print(f"提取blocks模块blocks_RadFrac_data数据时出错: {e}")
    def extract_block_DSTWU_data(self):
        """提取block-DSTWU模块数据"""
        try:
            blocks_DSTWU_data = {}
            blocks_DSTWU = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "DSTWU":
                    blocks_DSTWU.append({
                        "name": block['name'],
                        "type": "DSTWU"
                    })
            # 规定提取
            for block in blocks_DSTWU:
                blocks_DSTWU_data[block['name']] = {}
                try:
                    SPEC_DATA = {}
                    # 塔规范参数
                    OPT_NTRR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_NTRR")  # 塔规范-选择RR或NSTAGE
                    self.add_if_not_empty(SPEC_DATA, "OPT_NTRR", OPT_NTRR)
                    RR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RR")  # 塔规范-回流比
                    self.add_if_not_empty(SPEC_DATA, "RR", RR)
                    NSTAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSTAGE")  # 塔规范-塔板数
                    self.add_if_not_empty(SPEC_DATA, "NSTAGE", NSTAGE)
                    PTOP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PTOP")  # 压力-塔顶压力
                    PTOP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PTOP")  # 压力-塔顶压力
                    self.add_if_not_empty(SPEC_DATA, "PTOP", PTOP,"PTOP_UNITS", PTOP_UNITS)
                    PBOT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PBOT")  # 压力-塔底压力
                    PBOT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PBOT")  # 压力-塔底压力
                    self.add_if_not_empty(SPEC_DATA, "PBOT", PBOT, "PBOT_UNITS", PBOT_UNITS)
                    OPT_RDV = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_RDV")  # 冷凝器规范-选择LIQUID/VAPOR/VAPLIQ
                    self.add_if_not_empty(SPEC_DATA, "OPT_RDV", OPT_RDV)
                    RDV = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RDV")  # 冷凝器规范-汽相分率
                    self.add_if_not_empty(SPEC_DATA, "RDV", RDV)
                    LIGHTKEY = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LIGHTKEY")  # 关键组分-轻关键组分
                    self.add_if_not_empty(SPEC_DATA, "LIGHTKEY", LIGHTKEY)
                    RECOVL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RECOVL")  # 关键组分-轻关键组分回收率
                    self.add_if_not_empty(SPEC_DATA,"RECOVL", RECOVL)
                    RECOVH = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RECOVH")  # 关键组分-重关键组分回收率
                    self.add_if_not_empty(SPEC_DATA,"RECOVH", RECOVH)
                    HEAVYKEY = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HEAVYKEY")  # 关键组分-重关键组分
                    self.add_if_not_empty(SPEC_DATA,"HEAVYKEY", HEAVYKEY)
                    blocks_DSTWU_data[block['name']]["SPEC_DATA"]= SPEC_DATA
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块DSTWU所有数据完成")
            self.data["blocks_DSTWU_data"] = blocks_DSTWU_data
        except Exception as e:
            print(f"提取blocks模块blocks_DSTWU_data数据时出错: {e}")
    def extract_block_Distl_data(self):
        """提取block-Distl模块数据"""
        try:
            blocks_Distl_data = {}
            blocks_Distl = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Distl":
                    blocks_Distl.append({
                        "name": block['name'],
                        "type": "Distl"
                    })
            # 规定提取
            for block in blocks_Distl:
                blocks_Distl_data[block['name']] = {}
                try:
                    blocks_Distl_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 塔板数和进料位置（无单位）
                    NSTAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSTAGE")  # 塔板数
                    FEED_LOC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FEED_LOC")  # 进料塔板数
                    RR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RR")  # 回流比
                    D_F = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\D_F")  # 馏出物与进料摩尔比
                    COND_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\COND_TYPE")  # 冷凝器类型
                    
                    if NSTAGE is not None and NSTAGE != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["NSTAGE"] = NSTAGE
                    if FEED_LOC is not None and FEED_LOC != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["FEED_LOC"] = FEED_LOC
                    if RR is not None and RR != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["RR"] = RR
                    if D_F is not None and D_F != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["D_F"] = D_F
                    if COND_TYPE is not None and COND_TYPE != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["COND_TYPE"] = COND_TYPE
                    
                    # 压力（带单位，单位：10 = kPa）
                    PTOP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PTOP")  # 冷凝器压力
                    PTOP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PTOP")
                    PBOT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PBOT")  # 再沸器压力
                    PBOT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PBOT")
                    
                    if PTOP is not None and PTOP != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["PTOP"] = PTOP
                        if PTOP_UNITS is not None and PTOP_UNITS != "":
                            blocks_Distl_data[block['name']]["SPEC_DATA"]["PTOP_UNITS"] = PTOP_UNITS
                    if PBOT is not None and PBOT != "":
                        blocks_Distl_data[block['name']]["SPEC_DATA"]["PBOT"] = PBOT
                        if PBOT_UNITS is not None and PBOT_UNITS != "":
                            blocks_Distl_data[block['name']]["SPEC_DATA"]["PBOT_UNITS"] = PBOT_UNITS
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Distl所有数据完成")
            self.data["blocks_Distl_data"] = blocks_Distl_data
        except Exception as e:
            print(f"提取blocks模块Distl数据时出错: {e}")
            self.data["blocks_Distl_data"] = {}
    def extract_block_Dupl_data(self):
        """提取block-Dupl模块数据"""
        try:
            blocks_Dupl_data = {}
            blocks_Dupl = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Dupl":
                    blocks_Dupl.append({
                        "name": block['name'],
                        "type": "Dupl"
                    })
            # 规定提取
            for block in blocks_Dupl:
                blocks_Dupl_data[block['name']] = {}
                try:
                    blocks_Dupl_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 提取参数
                    OPSETNAME = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPSETNAME")  # 物性方法集名称（字符串）
                    CHEMISTRY = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CHEMISTRY")  # 化学计算（字符串）
                    TRUE_COMPS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TRUE_COMPS")  # 真实组分（字符串）
                    FRWATEROPSET = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FRWATEROPSET")  # 自由水物性方法集（字符串）
                    SOLU_WATER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SOLU_WATER")  # 可溶性水（整数）
                    print("***********************:",SOLU_WATER)
                    HENRY_COMPS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HENRY_COMPS")  # Henry组分（字符串）
                    
                    if OPSETNAME is not None and OPSETNAME != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["OPSETNAME"] = OPSETNAME
                    if CHEMISTRY is not None and CHEMISTRY != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["CHEMISTRY"] = CHEMISTRY
                    if TRUE_COMPS is not None and TRUE_COMPS != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["TRUE_COMPS"] = TRUE_COMPS
                    if FRWATEROPSET is not None and FRWATEROPSET != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["FRWATEROPSET"] = FRWATEROPSET
                    if SOLU_WATER is not None and SOLU_WATER != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["SOLU_WATER"] = SOLU_WATER
                    if HENRY_COMPS is not None and HENRY_COMPS != "":
                        blocks_Dupl_data[block['name']]["SPEC_DATA"]["HENRY_COMPS"] = HENRY_COMPS
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Dupl所有数据完成")
            self.data["blocks_Dupl_data"] = blocks_Dupl_data
        except Exception as e:
            print(f"提取blocks模块Dupl数据时出错: {e}")
            self.data["blocks_Dupl_data"] = {}
    def extract_block_Extract_data(self):
        """提取block-Extract模块数据"""
        try:
            blocks_Extract_data = {}
            blocks_Extract = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "Extract":
                    blocks_Extract.append({
                        "name": block['name'],
                        "type": "Extract"
                    })
            # 规定提取
            for block in blocks_Extract:
                blocks_Extract_data[block['name']] = {}
                try:
                    blocks_Extract_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 1. 塔设定
                    NSTAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSTAGE")  # 塔板数
                    OPT_THERMAL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_THERMAL")  # 热力学选项
                    
                    if NSTAGE is not None and NSTAGE != "":
                        blocks_Extract_data[block['name']]["SPEC_DATA"]["NSTAGE"] = NSTAGE
                    if OPT_THERMAL is not None and OPT_THERMAL != "":
                        blocks_Extract_data[block['name']]["SPEC_DATA"]["OPT_THERMAL"] = OPT_THERMAL
                    
                    # 根据 OPT_THERMAL 的值提取不同的参数
                    if OPT_THERMAL == "TEMP":
                        # 提取 TSPEC_TEMP（动态塔板节点）
                        TSPEC_TEMP_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\TSPEC_TEMP")
                        TSPEC_TEMP_DATA = {}
                        for stage_num in TSPEC_TEMP_NODES:
                            TSPEC_TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TSPEC_TEMP\{stage_num}")
                            TSPEC_TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TSPEC_TEMP\{stage_num}")
                            if TSPEC_TEMP_VALUE is not None and TSPEC_TEMP_VALUE != "":
                                TSPEC_TEMP_DATA[stage_num] = {
                                    "TSPEC_TEMP_VALUE": TSPEC_TEMP_VALUE
                                }
                                if TSPEC_TEMP_UNITS is not None and TSPEC_TEMP_UNITS != "":
                                    TSPEC_TEMP_DATA[stage_num]["TSPEC_TEMP_UNITS"] = TSPEC_TEMP_UNITS
                        if TSPEC_TEMP_DATA:
                            blocks_Extract_data[block['name']]["SPEC_DATA"]["TSPEC_TEMP"] = TSPEC_TEMP_DATA
                    
                    elif OPT_THERMAL == "DUTY":
                        # 提取 HEATER_DUTY（动态塔板节点）
                        HEATER_DUTY_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\HEATER_DUTY")
                        HEATER_DUTY_DATA = {}
                        for stage_num in HEATER_DUTY_NODES:
                            HEATER_DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HEATER_DUTY\{stage_num}")
                            HEATER_DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\HEATER_DUTY\{stage_num}")
                            if HEATER_DUTY_VALUE is not None and HEATER_DUTY_VALUE != "":
                                HEATER_DUTY_DATA[stage_num] = {
                                    "HEATER_DUTY_VALUE": HEATER_DUTY_VALUE
                                }
                                if HEATER_DUTY_UNITS is not None and HEATER_DUTY_UNITS != "":
                                    HEATER_DUTY_DATA[stage_num]["HEATER_DUTY_UNITS"] = HEATER_DUTY_UNITS
                        if HEATER_DUTY_DATA:
                            blocks_Extract_data[block['name']]["SPEC_DATA"]["HEATER_DUTY"] = HEATER_DUTY_DATA
                    
                    # 2. 关键组分
                    # 提取 COMP1_LIST（保留索引信息，支持不连续索引）
                    COMP1_LIST_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\COMP1_LIST")
                    COMP1_LIST = {}
                    for comp1_index in COMP1_LIST_NODES:
                        COMP1_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\COMP1_LIST\{comp1_index}")
                        if COMP1_VALUE is not None and COMP1_VALUE != "":
                            COMP1_LIST[comp1_index] = COMP1_VALUE
                    if COMP1_LIST:
                        blocks_Extract_data[block['name']]["SPEC_DATA"]["COMP1_LIST"] = COMP1_LIST
                    
                    # 提取 COMP2_LIST（保留索引信息，支持不连续索引）
                    COMP2_LIST_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\COMP2_LIST")
                    COMP2_LIST = {}
                    for comp2_index in COMP2_LIST_NODES:
                        COMP2_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\COMP2_LIST\{comp2_index}")
                        if COMP2_VALUE is not None and COMP2_VALUE != "":
                            COMP2_LIST[comp2_index] = COMP2_VALUE
                    if COMP2_LIST:
                        blocks_Extract_data[block['name']]["SPEC_DATA"]["COMP2_LIST"] = COMP2_LIST
                    
                    # 3. 压力
                    # 提取 STAGE_PRES（动态塔板节点）
                    STAGE_PRES_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES")
                    STAGE_PRES_DATA = {}
                    for stage_num in STAGE_PRES_NODES:
                        STAGE_PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES\{stage_num}")
                        STAGE_PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\STAGE_PRES\{stage_num}")
                        if STAGE_PRES_VALUE is not None and STAGE_PRES_VALUE != "":
                            STAGE_PRES_DATA[stage_num] = {
                                "STAGE_PRES_VALUE": STAGE_PRES_VALUE
                            }
                            if STAGE_PRES_UNITS is not None and STAGE_PRES_UNITS != "":
                                STAGE_PRES_DATA[stage_num]["STAGE_PRES_UNITS"] = STAGE_PRES_UNITS
                    if STAGE_PRES_DATA:
                        blocks_Extract_data[block['name']]["SPEC_DATA"]["STAGE_PRES"] = STAGE_PRES_DATA
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块Extract所有数据完成")
            self.data["blocks_Extract_data"] = blocks_Extract_data
        except Exception as e:
            print(f"提取blocks模块Extract数据时出错: {e}")
            self.data["blocks_Extract_data"] = {}
    def extract_block_FSplit_data(self):
        """提取block-FSplit模块数据"""
        try:
            blocks_FSplit_data = {}
            blocks_FSplit = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "FSplit":
                    blocks_FSplit.append({
                        "name": block['name'],
                        "type": "FSplit"
                    })
            # 规定提取
            for block in blocks_FSplit:
                blocks_FSplit_data[block['name']] = {}
                try:
                    blocks_FSplit_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 1. COMPS (无单位，只有值)
                    COMPS_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\COMPS")
                    COMPS_DATA = {}
                    for comp_subnode in COMPS_NODES:
                        MIXED_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\COMPS\{comp_subnode}\MIXED")
                        if MIXED_NODES:
                            COMPS_DATA[comp_subnode] = {}
                            COMPS_DATA[comp_subnode]["MIXED"] = {}
                            for leaf_node in MIXED_NODES:
                                COMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\COMPS\{comp_subnode}\MIXED\{leaf_node}")
                                if COMP_VALUE is not None and COMP_VALUE != "":
                                    COMPS_DATA[comp_subnode]["MIXED"][leaf_node] = COMP_VALUE
                            if not COMPS_DATA[comp_subnode]["MIXED"]:
                                del COMPS_DATA[comp_subnode]["MIXED"]
                                del COMPS_DATA[comp_subnode]
                    if COMPS_DATA:
                        blocks_FSplit_data[block['name']]["SPEC_DATA"]["COMPS"] = COMPS_DATA
                    
                    # 2. 参数列表：单位: 0 表示无单位，单位: 3 表示需要单位
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
                        PARAM_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\{param_name}")
                        PARAM_DATA = {}
                        for subnode in PARAM_NODES:
                            PARAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\{param_name}\{subnode}")
                            if PARAM_VALUE is not None and PARAM_VALUE != "":
                                PARAM_DATA[subnode] = {value_key: PARAM_VALUE}
                                if has_units:
                                    PARAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\{param_name}\{subnode}")
                                    if PARAM_UNITS is not None and PARAM_UNITS != "":
                                        PARAM_DATA[subnode][units_key] = PARAM_UNITS
                        if PARAM_DATA:
                            blocks_FSplit_data[block['name']]["SPEC_DATA"][param_name] = PARAM_DATA
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块FSplit所有数据完成")
            self.data["blocks_FSplit_data"] = blocks_FSplit_data
        except Exception as e:
            print(f"提取blocks模块FSplit数据时出错: {e}")
            self.data["blocks_FSplit_data"] = {}
    def extract_block_HeatX_data(self):
        """提取block-HeatX模块数据"""
        try:
            blocks_HeatX_data = {}
            blocks_HeatX = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "HeatX":
                    blocks_HeatX.append({
                        "name": block['name'],
                        "type": "HeatX"
                    })
            # 规定提取
            for block in blocks_HeatX:
                blocks_HeatX_data[block['name']] = {}
                try:
                    blocks_HeatX_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 按照指定顺序提取参数
                    # 1. MODE (无单位)
                    MODE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MODE")
                    if MODE is not None and MODE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MODE"] = MODE
                    
                    # 2. HSHELL_TUBE (无单位)
                    HSHELL_TUBE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HSHELL_TUBE")
                    if HSHELL_TUBE is not None and HSHELL_TUBE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HSHELL_TUBE"] = HSHELL_TUBE
                    
                    # 3. TYPE (无单位)
                    TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TYPE")
                    if TYPE is not None and TYPE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TYPE"] = TYPE
                    
                    # 4. PROGRAM_MODE (无单位)
                    PROGRAM_MODE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROGRAM_MODE")
                    if PROGRAM_MODE is not None and PROGRAM_MODE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["PROGRAM_MODE"] = PROGRAM_MODE
                    
                    # 5. SPEC (无单位)
                    SPEC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC")
                    if SPEC is not None and SPEC != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SPEC"] = SPEC
                    
                    # 6. VALUE (有单位)
                    VALUE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VALUE")
                    VALUE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\VALUE")
                    if VALUE_VALUE is not None and VALUE_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["VALUE_VALUE"] = VALUE_VALUE
                    if VALUE_UNITS is not None and VALUE_UNITS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["VALUE_UNITS"] = VALUE_UNITS
                    
                    # 7. AREA (有单位)
                    AREA_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\AREA")
                    AREA_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\AREA")
                    if AREA_VALUE is not None and AREA_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["AREA_VALUE"] = AREA_VALUE
                        if AREA_UNITS is not None and AREA_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["AREA_UNITS"] = AREA_UNITS
                    
                    # 8. UA (有单位)
                    UA_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\UA")
                    UA_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\UA")
                    if UA_VALUE is not None and UA_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["UA_VALUE"] = UA_VALUE
                        if UA_UNITS is not None and UA_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["UA_UNITS"] = UA_UNITS
                    
                    # 9. MIN_TAPP (有单位)
                    MIN_TAPP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MIN_TAPP")
                    MIN_TAPP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\MIN_TAPP")
                    if MIN_TAPP_VALUE is not None and MIN_TAPP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MIN_TAPP_VALUE"] = MIN_TAPP_VALUE
                        if MIN_TAPP_UNITS is not None and MIN_TAPP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["MIN_TAPP_UNITS"] = MIN_TAPP_UNITS
                    
                    # 10. FT_MIN (无单位)
                    FT_MIN = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FT_MIN")
                    if FT_MIN is not None and FT_MIN != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["FT_MIN"] = FT_MIN
                    
                    # 11. F_OPTION (无单位)
                    F_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\F_OPTION")
                    if F_OPTION is not None and F_OPTION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["F_OPTION"] = F_OPTION
                    
                    # 12. LMTD_CORRECT (无单位)
                    LMTD_CORRECT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LMTD_CORRECT")
                    if LMTD_CORRECT is not None and LMTD_CORRECT != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["LMTD_CORRECT"] = LMTD_CORRECT
                    
                    # 13. SIDE_VAR (无单位)
                    SIDE_VAR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SIDE_VAR")
                    if SIDE_VAR is not None and SIDE_VAR != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SIDE_VAR"] = SIDE_VAR
                    
                    # 14. CDP_OPTION (无单位)
                    CDP_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CDP_OPTION")
                    if CDP_OPTION is not None and CDP_OPTION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CDP_OPTION"] = CDP_OPTION
                    
                    # 15. PRES_COLD (有单位)
                    PRES_COLD_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES_COLD")
                    PRES_COLD_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES_COLD")
                    if PRES_COLD_VALUE is not None and PRES_COLD_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["PRES_COLD_VALUE"] = PRES_COLD_VALUE
                        if PRES_COLD_UNITS is not None and PRES_COLD_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["PRES_COLD_UNITS"] = PRES_COLD_UNITS
                    
                    # 16. CMAX_DP (无单位)
                    CMAX_DP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CMAX_DP")
                    if CMAX_DP is not None and CMAX_DP != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CMAX_DP"] = CMAX_DP
                    
                    # 17. CDP_SCALE (无单位)
                    CDP_SCALE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CDP_SCALE")
                    if CDP_SCALE is not None and CDP_SCALE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CDP_SCALE"] = CDP_SCALE
                    
                    # 18. TUBE_DP_FCOR (无单位)
                    TUBE_DP_FCOR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_DP_FCOR")
                    if TUBE_DP_FCOR is not None and TUBE_DP_FCOR != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_DP_FCOR"] = TUBE_DP_FCOR
                    
                    # 19. TUBE_DP_HCOR (无单位)
                    TUBE_DP_HCOR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_DP_HCOR")
                    if TUBE_DP_HCOR is not None and TUBE_DP_HCOR != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_DP_HCOR"] = TUBE_DP_HCOR
                    
                    # 20. TUBE_DP_PROF (无单位)
                    TUBE_DP_PROF = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_DP_PROF")
                    if TUBE_DP_PROF is not None and TUBE_DP_PROF != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_DP_PROF"] = TUBE_DP_PROF
                    
                    # 21. P_UPDATE (无单位)
                    P_UPDATE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\P_UPDATE")
                    if P_UPDATE is not None and P_UPDATE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["P_UPDATE"] = P_UPDATE
                    
                    # 22. U_OPTION (无单位)
                    U_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U_OPTION")
                    if U_OPTION is not None and U_OPTION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_OPTION"] = U_OPTION
                    
                    # 23. U (有单位)
                    U_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U")
                    U_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\U")
                    if U_VALUE is not None and U_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_VALUE"] = U_VALUE
                        if U_UNITS is not None and U_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_UNITS"] = U_UNITS
                    
                    # 24. B_B (有单位)
                    B_B_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\B_B")
                    B_B_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\B_B")
                    if B_B_VALUE is not None and B_B_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_B_VALUE"] = B_B_VALUE
                        if B_B_UNITS is not None and B_B_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_B_UNITS"] = B_B_UNITS
                    
                    # 25. B_L (有单位)
                    B_L_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\B_L")
                    B_L_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\B_L")
                    if B_L_VALUE is not None and B_L_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_L_VALUE"] = B_L_VALUE
                        if B_L_UNITS is not None and B_L_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_L_UNITS"] = B_L_UNITS
                    
                    # 26. B_V (有单位)
                    B_V_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\B_V")
                    B_V_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\B_V")
                    if B_V_VALUE is not None and B_V_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_V_VALUE"] = B_V_VALUE
                        if B_V_UNITS is not None and B_V_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["B_V_UNITS"] = B_V_UNITS
                    
                    # 27. L_B (有单位)
                    L_B_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L_B")
                    L_B_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\L_B")
                    if L_B_VALUE is not None and L_B_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_B_VALUE"] = L_B_VALUE
                        if L_B_UNITS is not None and L_B_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_B_UNITS"] = L_B_UNITS
                    
                    # 28. L_L (有单位)
                    L_L_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L_L")
                    L_L_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\L_L")
                    if L_L_VALUE is not None and L_L_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_L_VALUE"] = L_L_VALUE
                        if L_L_UNITS is not None and L_L_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_L_UNITS"] = L_L_UNITS
                    
                    # 29. L_V (有单位)
                    L_V_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\L_V")
                    L_V_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\L_V")
                    if L_V_VALUE is not None and L_V_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_V_VALUE"] = L_V_VALUE
                        if L_V_UNITS is not None and L_V_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["L_V_UNITS"] = L_V_UNITS
                    
                    # 30. V_B (有单位)
                    V_B_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\V_B")
                    V_B_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\V_B")
                    if V_B_VALUE is not None and V_B_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_B_VALUE"] = V_B_VALUE
                        if V_B_UNITS is not None and V_B_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_B_UNITS"] = V_B_UNITS
                    
                    # 31. V_L (有单位)
                    V_L_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\V_L")
                    V_L_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\V_L")
                    if V_L_VALUE is not None and V_L_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_L_VALUE"] = V_L_VALUE
                        if V_L_UNITS is not None and V_L_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_L_UNITS"] = V_L_UNITS
                    
                    # 32. V_V (有单位)
                    V_V_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\V_V")
                    V_V_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\V_V")
                    if V_V_VALUE is not None and V_V_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_V_VALUE"] = V_V_VALUE
                        if V_V_UNITS is not None and V_V_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["V_V_UNITS"] = V_V_UNITS
                    
                    # 33. U_REF_SIDE (无单位)
                    U_REF_SIDE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U_REF_SIDE")
                    if U_REF_SIDE is not None and U_REF_SIDE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_REF_SIDE"] = U_REF_SIDE
                    
                    # 34. UFLOW_BASIS (无单位)
                    UFLOW_BASIS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\UFLOW_BASIS")
                    if UFLOW_BASIS is not None and UFLOW_BASIS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["UFLOW_BASIS"] = UFLOW_BASIS
                    
                    # 35. BASIS_UFLOW (有单位)
                    BASIS_UFLOW_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_UFLOW")
                    BASIS_UFLOW_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_UFLOW")
                    if BASIS_UFLOW_VALUE is not None and BASIS_UFLOW_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["BASIS_UFLOW_VALUE"] = BASIS_UFLOW_VALUE
                        if BASIS_UFLOW_UNITS is not None and BASIS_UFLOW_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["BASIS_UFLOW_UNITS"] = BASIS_UFLOW_UNITS
                    
                    # 36. U_REF_VALUE (有单位)
                    U_REF_VALUE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U_REF_VALUE")
                    U_REF_VALUE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\U_REF_VALUE")
                    if U_REF_VALUE_VALUE is not None and U_REF_VALUE_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_REF_VALUE_VALUE"] = U_REF_VALUE_VALUE
                        if U_REF_VALUE_UNITS is not None and U_REF_VALUE_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_REF_VALUE_UNITS"] = U_REF_VALUE_UNITS
                    
                    # 37. U_EXPONENT (无单位)
                    U_EXPONENT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U_EXPONENT")
                    if U_EXPONENT is not None and U_EXPONENT != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_EXPONENT"] = U_EXPONENT
                    
                    # 38. U_SCALE (无单位)
                    U_SCALE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\U_SCALE")
                    if U_SCALE is not None and U_SCALE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["U_SCALE"] = U_SCALE
                    
                    # 39. CH_OPTION (无单位)
                    CH_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_OPTION")
                    if CH_OPTION is not None and CH_OPTION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_OPTION"] = CH_OPTION
                    
                    # 40. CH (有单位)
                    CH_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH")
                    CH_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CH")
                    if CH_VALUE is not None and CH_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_VALUE"] = CH_VALUE
                        if CH_UNITS is not None and CH_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_UNITS"] = CH_UNITS
                    
                    # 41. CH_B (有单位)
                    CH_B_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_B")
                    CH_B_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CH_B")
                    if CH_B_VALUE is not None and CH_B_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_B_VALUE"] = CH_B_VALUE
                        if CH_B_UNITS is not None and CH_B_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_B_UNITS"] = CH_B_UNITS
                    
                    # 42. CH_L (有单位)
                    CH_L_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_L")
                    CH_L_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CH_L")
                    if CH_L_VALUE is not None and CH_L_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_L_VALUE"] = CH_L_VALUE
                        if CH_L_UNITS is not None and CH_L_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_L_UNITS"] = CH_L_UNITS
                    
                    # 43. CH_V (有单位)
                    CH_V_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_V")
                    CH_V_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CH_V")
                    if CH_V_VALUE is not None and CH_V_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_V_VALUE"] = CH_V_VALUE
                        if CH_V_UNITS is not None and CH_V_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_V_UNITS"] = CH_V_UNITS
                    
                    # 44. CHFLOW_BASIS (无单位)
                    CHFLOW_BASIS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CHFLOW_BASIS")
                    if CHFLOW_BASIS is not None and CHFLOW_BASIS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CHFLOW_BASIS"] = CHFLOW_BASIS
                    
                    # 45. CH_EXPONENT (无单位)
                    CH_EXPONENT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_EXPONENT")
                    if CH_EXPONENT is not None and CH_EXPONENT != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_EXPONENT"] = CH_EXPONENT
                    
                    # 46. BASIS_CHFLOW (有单位)
                    BASIS_CHFLOW_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BASIS_CHFLOW")
                    BASIS_CHFLOW_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\BASIS_CHFLOW")
                    if BASIS_CHFLOW_VALUE is not None and BASIS_CHFLOW_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["BASIS_CHFLOW_VALUE"] = BASIS_CHFLOW_VALUE
                        if BASIS_CHFLOW_UNITS is not None and BASIS_CHFLOW_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["BASIS_CHFLOW_UNITS"] = BASIS_CHFLOW_UNITS
                    
                    # 47. CH_REF_VALUE (有单位)
                    CH_REF_VALUE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CH_REF_VALUE")
                    CH_REF_VALUE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CH_REF_VALUE")
                    if CH_REF_VALUE_VALUE is not None and CH_REF_VALUE_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_REF_VALUE_VALUE"] = CH_REF_VALUE_VALUE
                        if CH_REF_VALUE_UNITS is not None and CH_REF_VALUE_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["CH_REF_VALUE_UNITS"] = CH_REF_VALUE_UNITS
                    
                    # 48. TEMA_TYPE (无单位)
                    TEMA_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMA_TYPE")
                    if TEMA_TYPE is not None and TEMA_TYPE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TEMA_TYPE"] = TEMA_TYPE
                    
                    # 49. TUBE_NPASS (无单位)
                    TUBE_NPASS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_NPASS")
                    if TUBE_NPASS is not None and TUBE_NPASS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_NPASS"] = TUBE_NPASS
                    
                    # 50. ORIENTATION (无单位)
                    ORIENTATION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\ORIENTATION")
                    if ORIENTATION is not None and ORIENTATION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["ORIENTATION"] = ORIENTATION
                    
                    # 51. NSEAL_STRIP (无单位)
                    NSEAL_STRIP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSEAL_STRIP")
                    if NSEAL_STRIP is not None and NSEAL_STRIP != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["NSEAL_STRIP"] = NSEAL_STRIP
                    
                    # 52. TUBE_FLOW (无单位)
                    TUBE_FLOW = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_FLOW")
                    if TUBE_FLOW is not None and TUBE_FLOW != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_FLOW"] = TUBE_FLOW
                    
                    # 53. SHELL_BND_SP (有单位)
                    SHELL_BND_SP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SHELL_BND_SP")
                    SHELL_BND_SP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SHELL_BND_SP")
                    if SHELL_BND_SP_VALUE is not None and SHELL_BND_SP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_BND_SP_VALUE"] = SHELL_BND_SP_VALUE
                        if SHELL_BND_SP_UNITS is not None and SHELL_BND_SP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_BND_SP_UNITS"] = SHELL_BND_SP_UNITS
                    
                    # 54. SHELL_DIAM (有单位)
                    SHELL_DIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SHELL_DIAM")
                    SHELL_DIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SHELL_DIAM")
                    if SHELL_DIAM_VALUE is not None and SHELL_DIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_DIAM_VALUE"] = SHELL_DIAM_VALUE
                        if SHELL_DIAM_UNITS is not None and SHELL_DIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_DIAM_UNITS"] = SHELL_DIAM_UNITS
                    
                    # 55. SHELL_NPAR (无单位)
                    SHELL_NPAR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SHELL_NPAR")
                    if SHELL_NPAR is not None and SHELL_NPAR != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_NPAR"] = SHELL_NPAR
                    
                    # 56. SHELL_NSER (无单位)
                    SHELL_NSER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SHELL_NSER")
                    if SHELL_NSER is not None and SHELL_NSER != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_NSER"] = SHELL_NSER
                    
                    # 57. TUBE_TYPE (无单位)
                    TUBE_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_TYPE")
                    if TUBE_TYPE is not None and TUBE_TYPE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_TYPE"] = TUBE_TYPE
                    
                    # 58. TOTAL_NUMBER (无单位)
                    TOTAL_NUMBER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TOTAL_NUMBER")
                    if TOTAL_NUMBER is not None and TOTAL_NUMBER != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TOTAL_NUMBER"] = TOTAL_NUMBER
                    
                    # 59. PATTERN (无单位)
                    PATTERN = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PATTERN")
                    if PATTERN is not None and PATTERN != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["PATTERN"] = PATTERN
                    
                    # 60. MATERIAL (无单位)
                    MATERIAL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MATERIAL")
                    if MATERIAL is not None and MATERIAL != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MATERIAL"] = MATERIAL
                    
                    # 61. LENGTH (有单位)
                    LENGTH_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LENGTH")
                    LENGTH_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\LENGTH")
                    if LENGTH_VALUE is not None and LENGTH_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["LENGTH_VALUE"] = LENGTH_VALUE
                        if LENGTH_UNITS is not None and LENGTH_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["LENGTH_UNITS"] = LENGTH_UNITS
                    
                    # 62. PITCH (有单位)
                    PITCH_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PITCH")
                    PITCH_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PITCH")
                    if PITCH_VALUE is not None and PITCH_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["PITCH_VALUE"] = PITCH_VALUE
                        if PITCH_UNITS is not None and PITCH_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["PITCH_UNITS"] = PITCH_UNITS
                    
                    # 63. TCOND (有单位)
                    TCOND_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TCOND")
                    TCOND_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TCOND")
                    if TCOND_VALUE is not None and TCOND_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TCOND_VALUE"] = TCOND_VALUE
                        if TCOND_UNITS is not None and TCOND_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["TCOND_UNITS"] = TCOND_UNITS
                    
                    # 64. OUTSIDE_DIAM (有单位)
                    OUTSIDE_DIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OUTSIDE_DIAM")
                    OUTSIDE_DIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\OUTSIDE_DIAM")
                    if OUTSIDE_DIAM_VALUE is not None and OUTSIDE_DIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["OUTSIDE_DIAM_VALUE"] = OUTSIDE_DIAM_VALUE
                        if OUTSIDE_DIAM_UNITS is not None and OUTSIDE_DIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["OUTSIDE_DIAM_UNITS"] = OUTSIDE_DIAM_UNITS
                    
                    # 65. WALL_THICK (有单位)
                    WALL_THICK_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\WALL_THICK")
                    WALL_THICK_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\WALL_THICK")
                    if WALL_THICK_VALUE is not None and WALL_THICK_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["WALL_THICK_VALUE"] = WALL_THICK_VALUE
                        if WALL_THICK_UNITS is not None and WALL_THICK_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["WALL_THICK_UNITS"] = WALL_THICK_UNITS
                    
                    # 66. OPT_FHEIGHT (无单位)
                    OPT_FHEIGHT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_FHEIGHT")
                    if OPT_FHEIGHT is not None and OPT_FHEIGHT != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["OPT_FHEIGHT"] = OPT_FHEIGHT
                    
                    # 67. HEIGHT (有单位)
                    HEIGHT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HEIGHT")
                    HEIGHT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\HEIGHT")
                    if HEIGHT_VALUE is not None and HEIGHT_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HEIGHT_VALUE"] = HEIGHT_VALUE
                        if HEIGHT_UNITS is not None and HEIGHT_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["HEIGHT_UNITS"] = HEIGHT_UNITS
                    
                    # 68. ROOT_DIAM (有单位)
                    ROOT_DIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\ROOT_DIAM")
                    ROOT_DIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\ROOT_DIAM")
                    if ROOT_DIAM_VALUE is not None and ROOT_DIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROOT_DIAM_VALUE"] = ROOT_DIAM_VALUE
                        if ROOT_DIAM_UNITS is not None and ROOT_DIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROOT_DIAM_UNITS"] = ROOT_DIAM_UNITS
                    
                    # 69. OPT_FSPACING (无单位)
                    OPT_FSPACING = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_FSPACING")
                    if OPT_FSPACING is not None and OPT_FSPACING != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["OPT_FSPACING"] = OPT_FSPACING
                    
                    # 70. NPER_LENGTH (有单位)
                    NPER_LENGTH_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NPER_LENGTH")
                    NPER_LENGTH_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\NPER_LENGTH")
                    if NPER_LENGTH_VALUE is not None and NPER_LENGTH_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["NPER_LENGTH_VALUE"] = NPER_LENGTH_VALUE
                        if NPER_LENGTH_UNITS is not None and NPER_LENGTH_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["NPER_LENGTH_UNITS"] = NPER_LENGTH_UNITS
                    
                    # 71. THICKNESS (有单位)
                    THICKNESS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\THICKNESS")
                    THICKNESS_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\THICKNESS")
                    if THICKNESS_VALUE is not None and THICKNESS_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["THICKNESS_VALUE"] = THICKNESS_VALUE
                        if THICKNESS_UNITS is not None and THICKNESS_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["THICKNESS_UNITS"] = THICKNESS_UNITS
                    
                    # 72. AREA_RATIO (无单位)
                    AREA_RATIO = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\AREA_RATIO")
                    if AREA_RATIO is not None and AREA_RATIO != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["AREA_RATIO"] = AREA_RATIO
                    
                    # 73. EFFICIENCY (无单位)
                    EFFICIENCY = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\EFFICIENCY")
                    if EFFICIENCY is not None and EFFICIENCY != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["EFFICIENCY"] = EFFICIENCY
                    
                    # 74. BAFFLE_TYPE (无单位)
                    BAFFLE_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BAFFLE_TYPE")
                    if BAFFLE_TYPE is not None and BAFFLE_TYPE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["BAFFLE_TYPE"] = BAFFLE_TYPE
                    
                    # 75. NSEG_BAFFLE (无单位) - 只添加一次
                    NSEG_BAFFLE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSEG_BAFFLE")
                    if NSEG_BAFFLE is not None and NSEG_BAFFLE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["NSEG_BAFFLE"] = NSEG_BAFFLE
                    
                    # 76. RING_INDIAM (有单位)
                    RING_INDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RING_INDIAM")
                    RING_INDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\RING_INDIAM")
                    if RING_INDIAM_VALUE is not None and RING_INDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["RING_INDIAM_VALUE"] = RING_INDIAM_VALUE
                        if RING_INDIAM_UNITS is not None and RING_INDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["RING_INDIAM_UNITS"] = RING_INDIAM_UNITS
                    
                    # 77. RING_OUTDIAM (有单位)
                    RING_OUTDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RING_OUTDIAM")
                    RING_OUTDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\RING_OUTDIAM")
                    if RING_OUTDIAM_VALUE is not None and RING_OUTDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["RING_OUTDIAM_VALUE"] = RING_OUTDIAM_VALUE
                        if RING_OUTDIAM_UNITS is not None and RING_OUTDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["RING_OUTDIAM_UNITS"] = RING_OUTDIAM_UNITS
                    
                    # 78. ROD_DIAM (有单位)
                    ROD_DIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\ROD_DIAM")
                    ROD_DIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\ROD_DIAM")
                    if ROD_DIAM_VALUE is not None and ROD_DIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROD_DIAM_VALUE"] = ROD_DIAM_VALUE
                        if ROD_DIAM_UNITS is not None and ROD_DIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROD_DIAM_UNITS"] = ROD_DIAM_UNITS
                    
                    # 79. ROD_LENGTH (有单位)
                    ROD_LENGTH_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\ROD_LENGTH")
                    ROD_LENGTH_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\ROD_LENGTH")
                    if ROD_LENGTH_VALUE is not None and ROD_LENGTH_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROD_LENGTH_VALUE"] = ROD_LENGTH_VALUE
                        if ROD_LENGTH_UNITS is not None and ROD_LENGTH_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["ROD_LENGTH_UNITS"] = ROD_LENGTH_UNITS
                    
                    # 80. BAFFLE_CUT (无单位)
                    BAFFLE_CUT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\BAFFLE_CUT")
                    if BAFFLE_CUT is not None and BAFFLE_CUT != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["BAFFLE_CUT"] = BAFFLE_CUT
                    
                    # 81. IN_BFL_SP (有单位)
                    IN_BFL_SP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\IN_BFL_SP")
                    IN_BFL_SP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\IN_BFL_SP")
                    if IN_BFL_SP_VALUE is not None and IN_BFL_SP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["IN_BFL_SP_VALUE"] = IN_BFL_SP_VALUE
                        if IN_BFL_SP_UNITS is not None and IN_BFL_SP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["IN_BFL_SP_UNITS"] = IN_BFL_SP_UNITS
                    
                    # 82. SHELL_BFL_SP (有单位)
                    SHELL_BFL_SP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SHELL_BFL_SP")
                    SHELL_BFL_SP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SHELL_BFL_SP")
                    if SHELL_BFL_SP_VALUE is not None and SHELL_BFL_SP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_BFL_SP_VALUE"] = SHELL_BFL_SP_VALUE
                        if SHELL_BFL_SP_UNITS is not None and SHELL_BFL_SP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SHELL_BFL_SP_UNITS"] = SHELL_BFL_SP_UNITS
                    
                    # 83. SMID_BFL_SP (有单位)
                    SMID_BFL_SP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SMID_BFL_SP")
                    SMID_BFL_SP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SMID_BFL_SP")
                    if SMID_BFL_SP_VALUE is not None and SMID_BFL_SP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SMID_BFL_SP_VALUE"] = SMID_BFL_SP_VALUE
                        if SMID_BFL_SP_UNITS is not None and SMID_BFL_SP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SMID_BFL_SP_UNITS"] = SMID_BFL_SP_UNITS
                    
                    # 84. TUBES_IN_WIN (无单位)
                    TUBES_IN_WIN = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBES_IN_WIN")
                    if TUBES_IN_WIN is not None and TUBES_IN_WIN != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBES_IN_WIN"] = TUBES_IN_WIN
                    
                    # 85. TUBE_BFL_SP (有单位)
                    TUBE_BFL_SP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TUBE_BFL_SP")
                    TUBE_BFL_SP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TUBE_BFL_SP")
                    if TUBE_BFL_SP_VALUE is not None and TUBE_BFL_SP_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_BFL_SP_VALUE"] = TUBE_BFL_SP_VALUE
                        if TUBE_BFL_SP_UNITS is not None and TUBE_BFL_SP_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["TUBE_BFL_SP_UNITS"] = TUBE_BFL_SP_UNITS
                    
                    # 86. SNOZ_INDIAM (有单位)
                    SNOZ_INDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SNOZ_INDIAM")
                    SNOZ_INDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SNOZ_INDIAM")
                    if SNOZ_INDIAM_VALUE is not None and SNOZ_INDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SNOZ_INDIAM_VALUE"] = SNOZ_INDIAM_VALUE
                        if SNOZ_INDIAM_UNITS is not None and SNOZ_INDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SNOZ_INDIAM_UNITS"] = SNOZ_INDIAM_UNITS
                    
                    # 87. SNOZ_OUTDIAM (有单位)
                    SNOZ_OUTDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SNOZ_OUTDIAM")
                    SNOZ_OUTDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SNOZ_OUTDIAM")
                    if SNOZ_OUTDIAM_VALUE is not None and SNOZ_OUTDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SNOZ_OUTDIAM_VALUE"] = SNOZ_OUTDIAM_VALUE
                        if SNOZ_OUTDIAM_UNITS is not None and SNOZ_OUTDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["SNOZ_OUTDIAM_UNITS"] = SNOZ_OUTDIAM_UNITS
                    
                    # 88. TNOZ_INDIAM (有单位)
                    TNOZ_INDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TNOZ_INDIAM")
                    TNOZ_INDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TNOZ_INDIAM")
                    if TNOZ_INDIAM_VALUE is not None and TNOZ_INDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TNOZ_INDIAM_VALUE"] = TNOZ_INDIAM_VALUE
                        if TNOZ_INDIAM_UNITS is not None and TNOZ_INDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["TNOZ_INDIAM_UNITS"] = TNOZ_INDIAM_UNITS
                    
                    # 89. TNOZ_OUTDIAM (有单位)
                    TNOZ_OUTDIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TNOZ_OUTDIAM")
                    TNOZ_OUTDIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TNOZ_OUTDIAM")
                    if TNOZ_OUTDIAM_VALUE is not None and TNOZ_OUTDIAM_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["TNOZ_OUTDIAM_VALUE"] = TNOZ_OUTDIAM_VALUE
                        if TNOZ_OUTDIAM_UNITS is not None and TNOZ_OUTDIAM_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["TNOZ_OUTDIAM_UNITS"] = TNOZ_OUTDIAM_UNITS
                    
                    # 其他不在列表中的参数（放在最后）
                    # NUM_SHELLS (无单位)
                    NUM_SHELLS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NUM_SHELLS")
                    if NUM_SHELLS is not None and NUM_SHELLS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["NUM_SHELLS"] = NUM_SHELLS
                    
                    # SPECUN (无单位)
                    SPECUN = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPECUN")
                    if SPECUN is not None and SPECUN != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SPECUN"] = SPECUN
                    
                    # PRES_HOT (有单位)
                    PRES_HOT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES_HOT")
                    PRES_HOT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES_HOT")
                    if PRES_HOT_VALUE is not None and PRES_HOT_VALUE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["PRES_HOT_VALUE"] = PRES_HOT_VALUE
                        if PRES_HOT_UNITS is not None and PRES_HOT_UNITS != "":
                            blocks_HeatX_data[block['name']]["SPEC_DATA"]["PRES_HOT_UNITS"] = PRES_HOT_UNITS
                    
                    # SCUT_INTVLS (无单位)
                    SCUT_INTVLS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SCUT_INTVLS")
                    if SCUT_INTVLS is not None and SCUT_INTVLS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["SCUT_INTVLS"] = SCUT_INTVLS
                    
                    # MIN_FLS_PTS (无单位)
                    MIN_FLS_PTS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MIN_FLS_PTS")
                    if MIN_FLS_PTS is not None and MIN_FLS_PTS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MIN_FLS_PTS"] = MIN_FLS_PTS
                    
                    # MAX_NSHELLS (无单位)
                    MAX_NSHELLS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MAX_NSHELLS")
                    if MAX_NSHELLS is not None and MAX_NSHELLS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MAX_NSHELLS"] = MAX_NSHELLS
                    
                    # MIN_HRC_PTS (无单位)
                    MIN_HRC_PTS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MIN_HRC_PTS")
                    if MIN_HRC_PTS is not None and MIN_HRC_PTS != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["MIN_HRC_PTS"] = MIN_HRC_PTS
                    
                    # HDP_OPTION (无单位)
                    HDP_OPTION = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HDP_OPTION")
                    if HDP_OPTION is not None and HDP_OPTION != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HDP_OPTION"] = HDP_OPTION
                    
                    # HDP_SCALE (无单位)
                    HDP_SCALE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HDP_SCALE")
                    if HDP_SCALE is not None and HDP_SCALE != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HDP_SCALE"] = HDP_SCALE
                    
                    # HMAX_DP (无单位)
                    HMAX_DP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HMAX_DP")
                    if HMAX_DP is not None and HMAX_DP != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HMAX_DP"] = HMAX_DP
                    
                    # CDPPARM (无单位)
                    CDPPARM = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CDPPARM")
                    if CDPPARM is not None and CDPPARM != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CDPPARM"] = CDPPARM
                    
                    # HDPPARM (无单位)
                    HDPPARM = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HDPPARM")
                    if HDPPARM is not None and HDPPARM != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HDPPARM"] = HDPPARM
                    
                    # HDPPARMOP (无单位)
                    HDPPARMOP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HDPPARMOP")
                    if HDPPARMOP is not None and HDPPARMOP != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["HDPPARMOP"] = HDPPARMOP
                    
                    # CDPPARMOP (无单位)
                    CDPPARMOP = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CDPPARMOP")
                    if CDPPARMOP is not None and CDPPARMOP != "":
                        blocks_HeatX_data[block['name']]["SPEC_DATA"]["CDPPARMOP"] = CDPPARMOP
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块HeatX所有数据完成")
            self.data["blocks_HeatX_data"] = blocks_HeatX_data
        except Exception as e:
            print(f"提取blocks模块HeatX数据时出错: {e}")
            self.data["blocks_HeatX_data"] = {}
    def extract_block_MCompr_data(self):
        """提取block-MCompr模块数据"""
        try:
            blocks_MCompr_data = {}
            blocks_MCompr = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "MCompr":
                    blocks_MCompr.append({
                        "name": block['name'],
                        "type": "MCompr"
                    })
            # 规定提取
            for block in blocks_MCompr:
                blocks_MCompr_data[block['name']] = {}
                try:
                    blocks_MCompr_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 按照指定顺序提取参数
                    # 1. NSTAGE (无单位)
                    NSTAGE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NSTAGE")
                    if NSTAGE is not None and NSTAGE != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["NSTAGE"] = NSTAGE
                    
                    # 2. PROD_STAGE (节点本身有值，子节点也有值，两者值相同)
                    PROD_STAGE_NODE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_STAGE")  # 节点本身的值
                    PROD_STAGE_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\PROD_STAGE")
                    PROD_STAGE_DATA = []
                    for PROD_STAGE in PROD_STAGE_NODES:
                        # 子节点的值（动态流股名称，如MCOMPRO）
                        PROD_STREAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_STAGE\{PROD_STAGE}")
                        if PROD_STREAM_VALUE is not None and PROD_STREAM_VALUE != "":
                            PROD_STAGE_DATA.append({
                                "PROD_STAGE": PROD_STAGE,  # 动态流股名称
                                "PROD_STAGE_VALUE": PROD_STAGE_NODE_VALUE,  # 节点本身的值
                                "PROD_STREAM_VALUE": PROD_STREAM_VALUE  # 子节点的值
                            })
                    if PROD_STAGE_DATA:
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["PROD_STAGE"] = PROD_STAGE_DATA
                    
                    # 3. TYPE (无单位)
                    TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TYPE")
                    if TYPE is not None and TYPE != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["TYPE"] = TYPE
                    
                    # 4. OPT_SPEC (无单位)
                    OPT_SPEC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_SPEC")
                    if OPT_SPEC is not None and OPT_SPEC != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_SPEC"] = OPT_SPEC
                    
                    # 5. PRES (有单位，单位: 10)
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        if PRES_UNITS is not None and PRES_UNITS != "":
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    
                    # 6. TYPE_STG (无单位)
                    TYPE_STG = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TYPE_STG")
                    if TYPE_STG is not None and TYPE_STG != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["TYPE_STG"] = TYPE_STG
                    
                    # 7. CALC_SPEED (无单位)
                    CALC_SPEED = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CALC_SPEED")
                    if CALC_SPEED is not None and CALC_SPEED != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["CALC_SPEED"] = CALC_SPEED
                    
                    # 8. GPSA_BASIS (无单位)
                    GPSA_BASIS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\GPSA_BASIS")
                    if GPSA_BASIS is not None and GPSA_BASIS != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["GPSA_BASIS"] = GPSA_BASIS
                    
                    # 9. CPR_METHOD (无单位)
                    CPR_METHOD = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CPR_METHOD")
                    if CPR_METHOD is not None and CPR_METHOD != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["CPR_METHOD"] = CPR_METHOD
                    
                    # 10. FEED_STAGE (节点本身有值，子节点也有值，两者值相同)
                    FEED_STAGE_NODE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FEED_STAGE")  # 节点本身的值
                    FEED_STAGE_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\FEED_STAGE")
                    FEED_STAGE_DATA = []
                    for FEED_STAGE in FEED_STAGE_NODES:
                        # 子节点的值（动态流股名称，如MCOMPRI）
                        FEED_STREAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\FEED_STAGE\{FEED_STAGE}")
                        if FEED_STREAM_VALUE is not None and FEED_STREAM_VALUE != "":
                            FEED_STAGE_DATA.append({
                                "FEED_STAGE": FEED_STAGE,  # 动态流股名称
                                "FEED_STAGE_VALUE": FEED_STAGE_NODE_VALUE,  # 节点本身的值
                                "FEED_STREAM_VALUE": FEED_STREAM_VALUE  # 子节点的值
                            })
                    if FEED_STAGE_DATA:
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["FEED_STAGE"] = FEED_STAGE_DATA
                    
                    # 11. GLOBAL (节点本身有值，子节点也有值，两者值相同)
                    GLOBAL_NODE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\GLOBAL")  # 节点本身的值
                    GLOBAL_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\GLOBAL")
                    GLOBAL_DATA = {}
                    for GLOBAL in GLOBAL_NODES:
                        # 子节点的值（动态流股名称，如MCOMPRO）
                        PROD_STREAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\GLOBAL\{GLOBAL}")
                        if PROD_STREAM_VALUE is not None and PROD_STREAM_VALUE != "":
                            GLOBAL_DATA[GLOBAL] = {
                                "GLOBAL_VALUE": GLOBAL_NODE_VALUE,  # 节点本身的值
                                "PROD_STREAM_VALUE": PROD_STREAM_VALUE  # 子节点的值
                            }
                    if GLOBAL_DATA:
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["GLOBAL"] = GLOBAL_DATA
                    
                    # 12. PROD_PHASE (节点本身有值，子节点也有值，两者值相同)
                    PROD_PHASE_NODE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_PHASE")  # 节点本身的值
                    PROD_PHASE_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\PROD_PHASE")
                    PROD_PHASE_DATA = []
                    for PROD_PHASE in PROD_PHASE_NODES:
                        # 子节点的值（动态流股名称，如MCOMPRO）
                        PROD_STREAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PROD_PHASE\{PROD_PHASE}")
                        if PROD_STREAM_VALUE is not None and PROD_STREAM_VALUE != "":
                            PROD_PHASE_DATA.append({
                                "PROD_PHASE": PROD_PHASE,  # 动态流股名称
                                "PROD_PHASE_VALUE": PROD_PHASE_NODE_VALUE,  # 节点本身的值
                                "PROD_STREAM_VALUE": PROD_STREAM_VALUE  # 子节点的值
                            })
                    if PROD_PHASE_DATA:
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["PROD_PHASE"] = PROD_PHASE_DATA
                    
                    # 13. TEMP (有单位，单位: 4)
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_MCompr_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        if TEMP_UNITS is not None and TEMP_UNITS != "":
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    
                    # 14-32. 按顺序提取带stage_num的参数（动态提取所有stage_num值）
                    # 先获取 CLFR 节点下的所有子节点（这些就是 stage_num）
                    CLFR_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\CLFR")
                    stage_num_list = sorted(CLFR_NODES, key=lambda x: int(x) if x.isdigit() else 0)  # 排序确保顺序一致
                    
                    # 如果没有找到 CLFR 节点，尝试从其他参数中提取 stage_num
                    if not stage_num_list:
                        # 尝试从 CL_TEMP 或其他参数中提取
                        CL_TEMP_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\CL_TEMP")
                        if CL_TEMP_NODES:
                            stage_num_list = sorted(CL_TEMP_NODES, key=lambda x: int(x) if x.isdigit() else 0)
                    
                    # 对每个 stage_num 提取所有参数
                    for stage_num in stage_num_list:
                        # 14. CLFR\{stage_num} (无单位)
                        CLFR_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CLFR\{stage_num}")
                        if CLFR_VALUE is not None and CLFR_VALUE != "":
                            if "CLFR" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["CLFR"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["CLFR"][stage_num] = CLFR_VALUE
                        
                        # 14. CL_TEMP\{stage_num} (有单位，单位: 4)
                        CL_TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CL_TEMP\{stage_num}")
                        CL_TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\CL_TEMP\{stage_num}")
                        if CL_TEMP_VALUE is not None and CL_TEMP_VALUE != "":
                            if "CL_TEMP" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["CL_TEMP"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["CL_TEMP"][stage_num] = {
                                "CL_TEMP_VALUE": CL_TEMP_VALUE
                            }
                            if CL_TEMP_UNITS is not None and CL_TEMP_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["CL_TEMP"][stage_num]["CL_TEMP_UNITS"] = CL_TEMP_UNITS
                        
                        # 15. COOLER_UTL\{stage_num} (无单位)
                        COOLER_UTL_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\COOLER_UTL\{stage_num}")
                        if COOLER_UTL_VALUE is not None and COOLER_UTL_VALUE != "":
                            if "COOLER_UTL" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["COOLER_UTL"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["COOLER_UTL"][stage_num] = COOLER_UTL_VALUE
                        
                        # 16. C_S_PRES\{stage_num} (有单位，单位: 10)
                        C_S_PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\C_S_PRES\{stage_num}")
                        C_S_PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\C_S_PRES\{stage_num}")
                        if C_S_PRES_VALUE is not None and C_S_PRES_VALUE != "":
                            if "C_S_PRES" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["C_S_PRES"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["C_S_PRES"][stage_num] = {
                                "C_S_PRES_VALUE": C_S_PRES_VALUE
                            }
                            if C_S_PRES_UNITS is not None and C_S_PRES_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["C_S_PRES"][stage_num]["C_S_PRES_UNITS"] = C_S_PRES_UNITS
                        
                        # 17. DELP\{stage_num} (有单位，单位: 10)
                        DELP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DELP\{stage_num}")
                        DELP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DELP\{stage_num}")
                        if DELP_VALUE is not None and DELP_VALUE != "":
                            if "DELP" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["DELP"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["DELP"][stage_num] = {
                                "DELP_VALUE": DELP_VALUE
                            }
                            if DELP_UNITS is not None and DELP_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["DELP"][stage_num]["DELP_UNITS"] = DELP_UNITS
                        
                        # 18. DUTY\{stage_num} (有单位，单位: 18)
                        DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY\{stage_num}")
                        DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY\{stage_num}")
                        if DUTY_VALUE is not None and DUTY_VALUE != "":
                            if "DUTY" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["DUTY"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["DUTY"][stage_num] = {
                                "DUTY_VALUE": DUTY_VALUE
                            }
                            if DUTY_UNITS is not None and DUTY_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["DUTY"][stage_num]["DUTY_UNITS"] = DUTY_UNITS
                        
                        # 19. MEFF\{stage_num} (无单位)
                        MEFF_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\MEFF\{stage_num}")
                        if MEFF_VALUE is not None and MEFF_VALUE != "":
                            if "MEFF" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["MEFF"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["MEFF"][stage_num] = MEFF_VALUE
                        
                        # 20. OPT_CLFR\{stage_num} (无单位)
                        OPT_CLFR_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_CLFR\{stage_num}")
                        if OPT_CLFR_VALUE is not None and OPT_CLFR_VALUE != "":
                            if "OPT_CLFR" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CLFR"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CLFR"][stage_num] = OPT_CLFR_VALUE
                        
                        # 21. OPT_CLSPEC\{stage_num} (无单位)
                        OPT_CLSPEC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_CLSPEC\{stage_num}")
                        if OPT_CLSPEC_VALUE is not None and OPT_CLSPEC_VALUE != "":
                            if "OPT_CLSPEC" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CLSPEC"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CLSPEC"][stage_num] = OPT_CLSPEC_VALUE
                        
                        # 22. OPT_CSPEC\{stage_num} (无单位)
                        OPT_CSPEC_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_CSPEC\{stage_num}")
                        if OPT_CSPEC_VALUE is not None and OPT_CSPEC_VALUE != "":
                            if "OPT_CSPEC" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CSPEC"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_CSPEC"][stage_num] = OPT_CSPEC_VALUE
                        
                        # 23. OPT_TEMP\{stage_num} (无单位)
                        OPT_TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_TEMP\{stage_num}")
                        if OPT_TEMP_VALUE is not None and OPT_TEMP_VALUE != "":
                            if "OPT_TEMP" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_TEMP"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["OPT_TEMP"][stage_num] = OPT_TEMP_VALUE
                        
                        # 24. PDROP\{stage_num} (有单位，单位: 10)
                        PDROP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PDROP\{stage_num}")
                        PDROP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PDROP\{stage_num}")
                        if PDROP_VALUE is not None and PDROP_VALUE != "":
                            if "PDROP" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["PDROP"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["PDROP"][stage_num] = {
                                "PDROP_VALUE": PDROP_VALUE
                            }
                            if PDROP_UNITS is not None and PDROP_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["PDROP"][stage_num]["PDROP_UNITS"] = PDROP_UNITS
                        
                        # 25. PEFF\{stage_num} (无单位)
                        PEFF_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PEFF\{stage_num}")
                        if PEFF_VALUE is not None and PEFF_VALUE != "":
                            if "PEFF" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["PEFF"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["PEFF"][stage_num] = PEFF_VALUE
                        
                        # 26. POWER\{stage_num} (有单位，单位: 3)
                        POWER_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\POWER\{stage_num}")
                        POWER_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\POWER\{stage_num}")
                        if POWER_VALUE is not None and POWER_VALUE != "":
                            if "POWER" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["POWER"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["POWER"][stage_num] = {
                                "POWER_VALUE": POWER_VALUE
                            }
                            if POWER_UNITS is not None and POWER_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["POWER"][stage_num]["POWER_UNITS"] = POWER_UNITS
                        
                        # 27. PRATIO\{stage_num} (无单位)
                        PRATIO_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRATIO\{stage_num}")
                        if PRATIO_VALUE is not None and PRATIO_VALUE != "":
                            if "PRATIO" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["PRATIO"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["PRATIO"][stage_num] = PRATIO_VALUE
                        
                        # 28. SEFF\{stage_num} (无单位)
                        SEFF_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SEFF\{stage_num}")
                        if SEFF_VALUE is not None and SEFF_VALUE != "":
                            if "SEFF" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["SEFF"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["SEFF"][stage_num] = SEFF_VALUE
                        
                        # 29. SPECS_UTL\{stage_num} (无单位)
                        SPECS_UTL_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPECS_UTL\{stage_num}")
                        if SPECS_UTL_VALUE is not None and SPECS_UTL_VALUE != "":
                            if "SPECS_UTL" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["SPECS_UTL"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["SPECS_UTL"][stage_num] = SPECS_UTL_VALUE
                        
                        # 31. TEMP\{stage_num} (有单位，单位: 4)
                        TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP\{stage_num}")
                        TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP\{stage_num}")
                        if TEMP_VALUE is not None and TEMP_VALUE != "":
                            if "TEMP" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["TEMP"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["TEMP"][stage_num] = {
                                "TEMP_VALUE": TEMP_VALUE
                            }
                            if TEMP_UNITS is not None and TEMP_UNITS != "":
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["TEMP"][stage_num]["TEMP_UNITS"] = TEMP_UNITS
                        
                        # 32. TRATIO\{stage_num} (无单位)
                        TRATIO_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TRATIO\{stage_num}")
                        if TRATIO_VALUE is not None and TRATIO_VALUE != "":
                            if "TRATIO" not in blocks_MCompr_data[block['name']]["SPEC_DATA"]:
                                blocks_MCompr_data[block['name']]["SPEC_DATA"]["TRATIO"] = {}
                            blocks_MCompr_data[block['name']]["SPEC_DATA"]["TRATIO"][stage_num] = TRATIO_VALUE
                    
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块MCompr所有数据完成")
            self.data["blocks_MCompr_data"] = blocks_MCompr_data
        except Exception as e:
            print(f"提取blocks模块MCompr数据时出错: {e}")
            self.data["blocks_MCompr_data"] = {}
    def extract_block_RCSTR_data(self):
        """提取block-RCSTR模块数据"""
        try:
            blocks_RCSTR_data = {}
            blocks_RCSTR = []
            blocks = self.data.get("blocks", [])
            for block in blocks:
                if block['type'] == "RCSTR":
                    blocks_RCSTR.append({
                        "name": block['name'],
                        "type": "RCSTR"
                    })
            # 规定提取
            for block in blocks_RCSTR:
                blocks_RCSTR_data[block['name']] = {}
                try:
                    blocks_RCSTR_data[block['name']]["SPEC_DATA"] = {}
                    
                    # 按照指定顺序提取参数
                    # 1. HTRANMODE (无单位)
                    HTRANMODE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\HTRANMODE")
                    if HTRANMODE is not None and HTRANMODE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["HTRANMODE"] = HTRANMODE
                    
                    # 2. PRES (有单位)
                    PRES_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    PRES_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PRES")
                    if PRES_VALUE is not None and PRES_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["PRES_VALUE"] = PRES_VALUE
                        if PRES_UNITS is not None and PRES_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["PRES_UNITS"] = PRES_UNITS
                    
                    # 3. SPEC_OPT (无单位)
                    SPEC_OPT = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_OPT")
                    if SPEC_OPT is not None and SPEC_OPT != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SPEC_OPT"] = SPEC_OPT
                    
                    # 4. NPHASE (无单位)
                    NPHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\NPHASE")
                    if NPHASE is not None and NPHASE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["NPHASE"] = NPHASE
                    
                    # 5. TEMP (有单位)
                    TEMP_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    TEMP_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\TEMP")
                    if TEMP_VALUE is not None and TEMP_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["TEMP_VALUE"] = TEMP_VALUE
                        if TEMP_UNITS is not None and TEMP_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["TEMP_UNITS"] = TEMP_UNITS
                    
                    # 6. DUTY (有单位)
                    DUTY_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    DUTY_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\DUTY")
                    if DUTY_VALUE is not None and DUTY_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["DUTY_VALUE"] = DUTY_VALUE
                        if DUTY_UNITS is not None and DUTY_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["DUTY_UNITS"] = DUTY_UNITS
                    
                    # 7. VFRAC (无单位)
                    VFRAC = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VFRAC")
                    if VFRAC is not None and VFRAC != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["VFRAC"] = VFRAC
                    
                    # 8. SPEC_TYPE (无单位)
                    SPEC_TYPE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_TYPE")
                    if SPEC_TYPE is not None and SPEC_TYPE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SPEC_TYPE"] = SPEC_TYPE
                    
                    # 9. SPEC_PHASE (无单位)
                    SPEC_PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SPEC_PHASE")
                    if SPEC_PHASE is not None and SPEC_PHASE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SPEC_PHASE"] = SPEC_PHASE
                    
                    # 10. REACT_VOL (有单位)
                    REACT_VOL_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REACT_VOL")
                    REACT_VOL_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\REACT_VOL")
                    if REACT_VOL_VALUE is not None and REACT_VOL_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["REACT_VOL_VALUE"] = REACT_VOL_VALUE
                        if REACT_VOL_UNITS is not None and REACT_VOL_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["REACT_VOL_UNITS"] = REACT_VOL_UNITS
                    
                    # 11. REACT_VOL_FR (无单位)
                    REACT_VOL_FR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REACT_VOL_FR")
                    if REACT_VOL_FR is not None and REACT_VOL_FR != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["REACT_VOL_FR"] = REACT_VOL_FR
                    
                    # 12. PH_RES_TIME (有单位)
                    PH_RES_TIME_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PH_RES_TIME")
                    PH_RES_TIME_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\PH_RES_TIME")
                    if PH_RES_TIME_VALUE is not None and PH_RES_TIME_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["PH_RES_TIME_VALUE"] = PH_RES_TIME_VALUE
                        if PH_RES_TIME_UNITS is not None and PH_RES_TIME_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["PH_RES_TIME_UNITS"] = PH_RES_TIME_UNITS
                    
                    # 13. PHASE (无单位)
                    PHASE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\PHASE")
                    if PHASE is not None and PHASE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["PHASE"] = PHASE
                    
                    # 14. VOL (有单位)
                    VOL_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\VOL")
                    VOL_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\VOL")
                    if VOL_VALUE is not None and VOL_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["VOL_VALUE"] = VOL_VALUE
                        if VOL_UNITS is not None and VOL_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["VOL_UNITS"] = VOL_UNITS
                    
                    # 15. RES_TIME (有单位)
                    RES_TIME_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RES_TIME")
                    RES_TIME_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\RES_TIME")
                    if RES_TIME_VALUE is not None and RES_TIME_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["RES_TIME_VALUE"] = RES_TIME_VALUE
                        if RES_TIME_UNITS is not None and RES_TIME_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["RES_TIME_UNITS"] = RES_TIME_UNITS
                    
                    # 16. CHK_MASSTR (无单位)
                    CHK_MASSTR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CHK_MASSTR")
                    if CHK_MASSTR is not None and CHK_MASSTR != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["CHK_MASSTR"] = CHK_MASSTR
                    
                    # 17. REACSYS (无单位)
                    REACSYS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\REACSYS")
                    if REACSYS is not None and REACSYS != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["REACSYS"] = REACSYS
                    
                    # 18. RXN_ID (动态节点列表，无单位)
                    try:
                        RXN_ID_NODES = self.get_child_nodes(fr"\Data\Blocks\{block['name']}\Input\RXN_ID")
                        RXN_ID_DATA = {}
                        for RXN_ID in RXN_ID_NODES:
                            RXN_ID_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\RXN_ID\{RXN_ID}")
                            if RXN_ID_VALUE is not None and RXN_ID_VALUE != "":
                                RXN_ID_DATA[RXN_ID] = RXN_ID_VALUE
                        if RXN_ID_DATA:
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["RXN_ID"] = RXN_ID_DATA
                    except Exception as e:
                        print(f"提取blocks模块{block['type']}_{block['name']}RXN_ID数据时出错: {e}")
                    
                    # 19. SUBBYPASS (有单位)
                    SUBBYPASS_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SUBBYPASS")
                    SUBBYPASS_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SUBBYPASS")
                    if SUBBYPASS_VALUE is not None and SUBBYPASS_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUBBYPASS_VALUE"] = SUBBYPASS_VALUE
                        if SUBBYPASS_UNITS is not None and SUBBYPASS_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUBBYPASS_UNITS"] = SUBBYPASS_UNITS
                    
                    # 20. CRYSTSYS (无单位)
                    CRYSTSYS = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CRYSTSYS")
                    if CRYSTSYS is not None and CRYSTSYS != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["CRYSTSYS"] = CRYSTSYS
                    
                    # 21. LOWER (有单位)
                    LOWER_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\LOWER")
                    LOWER_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\LOWER")
                    if LOWER_VALUE is not None and LOWER_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["LOWER_VALUE"] = LOWER_VALUE
                        if LOWER_UNITS is not None and LOWER_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["LOWER_UNITS"] = LOWER_UNITS
                    
                    # 22. SUB_RRSBN (有单位)
                    SUB_RRSBN_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SUB_RRSBN")
                    SUB_RRSBN_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SUB_RRSBN")
                    if SUB_RRSBN_VALUE is not None and SUB_RRSBN_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUB_RRSBN_VALUE"] = SUB_RRSBN_VALUE
                        if SUB_RRSBN_UNITS is not None and SUB_RRSBN_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUB_RRSBN_UNITS"] = SUB_RRSBN_UNITS
                    
                    # 23. SUB_STDDEV (有单位)
                    SUB_STDDEV_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\SUB_STDDEV")
                    SUB_STDDEV_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\SUB_STDDEV")
                    if SUB_STDDEV_VALUE is not None and SUB_STDDEV_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUB_STDDEV_VALUE"] = SUB_STDDEV_VALUE
                        if SUB_STDDEV_UNITS is not None and SUB_STDDEV_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["SUB_STDDEV_UNITS"] = SUB_STDDEV_UNITS
                    
                    # 24. S_OPT (有单位)
                    S_OPT_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\S_OPT")
                    S_OPT_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\S_OPT")
                    if S_OPT_VALUE is not None and S_OPT_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["S_OPT_VALUE"] = S_OPT_VALUE
                        if S_OPT_UNITS is not None and S_OPT_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["S_OPT_UNITS"] = S_OPT_UNITS
                    
                    # 25. USER_SLOWER (有单位)
                    USER_SLOWER_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\USER_SLOWER")
                    USER_SLOWER_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\USER_SLOWER")
                    if USER_SLOWER_VALUE is not None and USER_SLOWER_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["USER_SLOWER_VALUE"] = USER_SLOWER_VALUE
                        if USER_SLOWER_UNITS is not None and USER_SLOWER_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["USER_SLOWER_UNITS"] = USER_SLOWER_UNITS
                    
                    # 26. USER_SVALUE (有单位)
                    USER_SVALUE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\USER_SVALUE")
                    USER_SVALUE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\USER_SVALUE")
                    if USER_SVALUE_VALUE is not None and USER_SVALUE_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["USER_SVALUE_VALUE"] = USER_SVALUE_VALUE
                        if USER_SVALUE_UNITS is not None and USER_SVALUE_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["USER_SVALUE_UNITS"] = USER_SVALUE_UNITS
                    
                    # 27. AGITATOR (无单位)
                    AGITATOR = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\AGITATOR")
                    if AGITATOR is not None and AGITATOR != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["AGITATOR"] = AGITATOR
                    
                    # 28. AGITRATE (有单位)
                    AGITRATE_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\AGITRATE")
                    AGITRATE_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\AGITRATE")
                    if AGITRATE_VALUE is not None and AGITRATE_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["AGITRATE_VALUE"] = AGITRATE_VALUE
                        if AGITRATE_UNITS is not None and AGITRATE_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["AGITRATE_UNITS"] = AGITRATE_UNITS
                    
                    # 29. IMPELLR_DIAM (有单位)
                    IMPELLR_DIAM_VALUE = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\IMPELLR_DIAM")
                    IMPELLR_DIAM_UNITS = self.safe_get_node_units(fr"\Data\Blocks\{block['name']}\Input\IMPELLR_DIAM")
                    if IMPELLR_DIAM_VALUE is not None and IMPELLR_DIAM_VALUE != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["IMPELLR_DIAM_VALUE"] = IMPELLR_DIAM_VALUE
                        if IMPELLR_DIAM_UNITS is not None and IMPELLR_DIAM_UNITS != "":
                            blocks_RCSTR_data[block['name']]["SPEC_DATA"]["IMPELLR_DIAM_UNITS"] = IMPELLR_DIAM_UNITS
                    
                    # 30. POWERNUMBER (无单位)
                    POWERNUMBER = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\POWERNUMBER")
                    if POWERNUMBER is not None and POWERNUMBER != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["POWERNUMBER"] = POWERNUMBER
                    
                    # 31. OPT_PSD (无单位)
                    OPT_PSD = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_PSD")
                    if OPT_PSD is not None and OPT_PSD != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["OPT_PSD"] = OPT_PSD
                    
                    # 32. CONST_METHOD (无单位)
                    CONST_METHOD = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\CONST_METHOD")
                    if CONST_METHOD is not None and CONST_METHOD != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["CONST_METHOD"] = CONST_METHOD
                    
                    # 33. OPT_SUBPSD (无单位)
                    OPT_SUBPSD = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_SUBPSD")
                    if OPT_SUBPSD is not None and OPT_SUBPSD != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["OPT_SUBPSD"] = OPT_SUBPSD
                    
                    # 29. OPT_OVERALL (无单位)
                    OPT_OVERALL = self.safe_get_node_value(fr"\Data\Blocks\{block['name']}\Input\OPT_OVERALL")
                    if OPT_OVERALL is not None and OPT_OVERALL != "":
                        blocks_RCSTR_data[block['name']]["SPEC_DATA"]["OPT_OVERALL"] = OPT_OVERALL
                        
                except Exception as e:
                    print(f"提取blocks模块{block['type']}_{block['name']}数据时出错: {e}")
                    continue
            print(f"提取blocks模块RCSTR所有数据完成")
            self.data["blocks_RCSTR_data"] = blocks_RCSTR_data
        except Exception as e:
            print(f"提取blocks模块RCSTR数据时出错: {e}")
            self.data["blocks_RCSTR_data"] = {}

    def extract_all_data(self):
        """提取所有数据"""
        print("开始提取 Aspen Plus 数据...")
        self.extract_setup()
        self.extract_components()
        self.extract_property_methods()
        # self.extract_henry_components()  # 新增：提取Henry组分
        self.extract_blocks()
        self.extract_streams()
        self.extract_block_connections()
        self.extract_streams_data()
        self.extract_reactions_data()
        self.extract_convergence_data()
        self.extract_design_specs_data()
        self.extract_block_Mixer_data()
        self.extract_block_Valve_data()
        self.extract_block_Compr_data()
        self.extract_block_Heater_data()
        self.extract_block_Pump_data()
        self.extract_block_RStoic_data()
        self.extract_block_RPlug_data()
        self.extract_block_Flash2_data()
        self.extract_block_Flash3_data()
        self.extract_block_Decanter_data()
        self.extract_block_Sep_data()
        self.extract_block_Sep2_data()
        self.extract_block_RadFrac_data()
        self.extract_block_DSTWU_data()
        self.extract_block_Distl_data()
        self.extract_block_Dupl_data()
        self.extract_block_Extract_data()
        self.extract_block_FSplit_data()
        self.extract_block_HeatX_data()
        self.extract_block_MCompr_data()
        self.extract_block_RCSTR_data()
        print("所有数据提取完成")

    def save_to_json(self, output_path: str):
        """将提取的数据保存为 JSON 文件"""
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
            print(f"数据已保存到: {output_path}")
            return True
        except Exception as e:
            print(f"保存 JSON 文件时出错: {e}")
            return False


# 使用示例
if __name__ == "__main__":
    converter = AspenToJSONConverter(r"E:\DICP\Build_Aspenplus_process\《化工过程模拟实训—Aspen+Plus教程》_第2版_勘误_模拟源文件-20180606\《化工过程模拟实训—Aspen Plus教程》_第2版_例题和习题模拟源文件_20180606\11-第11章_工艺流程模拟\习题\Solution11.10.bkp")

    if converter.connect_to_aspen():
        try:
            # 提取所有数据
            converter.extract_all_data()

            # 保存为 JSON 文件
            output_json_path = r"E:\DICP\Build_Aspenplus_process\json\Solution11.10.json"
            converter.save_to_json(output_json_path)

        except Exception as e:
            print(f"处理过程中出错: {e}")
        finally:
            # 断开连接
            converter.disconnect()
    else:
        print("无法连接到 Aspen Plus 文件")