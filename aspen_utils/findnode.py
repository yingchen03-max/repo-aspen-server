import os
from datetime import datetime
import win32com.client
import pythoncom
import json
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
            # print(f"获取节点 {node_path} 值时出错: {e}")
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
            # print(f"获取节点 {node_path} 单位时出错: {e}")
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
            # print(f"获取 {parent_path} 子节点时出错: {e}")
            return []

    def export_aspen_node_structure(self, base_path, output_file=None, max_depth=None):
        """
        递归获取ASPEN节点路径下的所有子节点结构并输出到文件

        参数:
        - base_path: 要遍历的ASPEN节点路径，例如 r'\Data\Flowsheeting Options\Design-Spec'
        - output_file: 输出文件路径，如果为None则自动生成文件名
        - max_depth: 最大递归深度，None表示无限制

        返回:
        - 包含所有节点路径的列表
        - 同时保存到本地文件
        """
        try:
            # 生成默认输出文件名
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_path = base_path.replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                output_file = f"aspen_nodes_{safe_path}_{timestamp}.txt"

            # 用于存储所有发现的节点路径
            all_nodes = []

            def traverse_node(current_path, current_depth=0):
                """递归遍历节点的辅助函数"""
                # 检查深度限制
                if max_depth is not None and current_depth > max_depth:
                    return

                try:
                    # 获取当前节点的子节点
                    child_nodes = self.get_child_nodes(current_path)

                    for child in child_nodes:
                        # 构建完整子节点路径
                        if current_path.endswith('\\'):
                            child_path = current_path + child
                        else:
                            child_path = current_path + '\\' + child

                        # 添加到结果列表
                        all_nodes.append(child_path)

                        # 递归遍历子节点
                        traverse_node(child_path, current_depth + 1)

                except Exception as e:
                    # 记录错误但不中断遍历
                    error_msg = f"访问节点 {current_path} 时出错: {e}"
                    all_nodes.append(f"# ERROR: {error_msg}")
                    print(f"警告: {error_msg}")

            print(f"开始遍历ASPEN节点: {base_path}")
            print(f"最大深度: {'无限制' if max_depth is None else max_depth}")

            # 开始遍历
            traverse_node(base_path)

            # 保存到文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"# ASPEN节点结构导出报告\n")
                f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 根路径: {base_path}\n")
                f.write(f"# 总节点数: {len(all_nodes)}\n")
                f.write(f"# ==========================================\n\n")

                for i, node_path in enumerate(all_nodes, 1):
                    # 计算节点深度（通过反斜杠数量）
                    depth = node_path.count('\\') - base_path.count('\\')
                    indent = "  " * depth

                    # 如果是错误信息，特殊标记
                    if node_path.startswith("# ERROR:"):
                        f.write(f"{indent}{node_path}\n")
                    else:
                        f.write(f"{i:4d}. {indent}{node_path}\n")

            print(f"节点遍历完成!")
            print(f"发现 {len(all_nodes)} 个节点")
            print(f"已保存到: {os.path.abspath(output_file)}")

            return all_nodes, output_file

        except Exception as e:
            print(f"遍历ASPEN节点时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], None

    def export_aspen_node_structure_with_values(self, base_path, output_file=None, max_depth=3):
        """
        获取节点结构并包含关键节点的值

        参数:
        - base_path: 要遍历的ASPEN节点路径
        - output_file: 输出文件路径
        - max_depth: 最大递归深度，默认3层

        返回:
        - 包含节点路径和值的字典
        """
        try:
            # 生成默认输出文件名
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_path = base_path.replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                output_file = f"aspen_nodes_with_values_{safe_path}_{timestamp}.txt"

            # 用于存储节点信息
            nodes_info = []

            def traverse_node_with_values(current_path, current_depth=0):
                """递归遍历节点并获取值的辅助函数"""
                # 检查深度限制
                if max_depth is not None and current_depth > max_depth:
                    return

                try:
                    # 获取当前节点的子节点
                    child_nodes = self.get_child_nodes(current_path)

                    for child in child_nodes:
                        # 构建完整子节点路径
                        if current_path.endswith('\\'):
                            child_path = current_path + child
                        else:
                            child_path = current_path + '\\' + child

                        # 尝试获取节点值
                        node_value = None
                        node_units = None

                        try:
                            # 先尝试获取值
                            node_value = self.safe_get_node_value(child_path)

                            # 如果获取到值，再尝试获取单位
                            if node_value is not None:
                                try:
                                    node_units = self.safe_get_node_units(child_path)
                                except:
                                    pass
                        except:
                            pass

                        # 创建节点信息字典
                        node_info = {
                            'path': child_path,
                            'depth': current_depth + 1,
                            'has_value': node_value is not None,
                            'value': node_value,
                            'units': node_units
                        }

                        nodes_info.append(node_info)

                        # 如果节点有值，通常不需要进一步遍历其子节点（根据Aspen结构）
                        # 但为了完整性，我们可以继续遍历
                        traverse_node_with_values(child_path, current_depth + 1)

                except Exception as e:
                    # 记录错误
                    error_info = {
                        'path': current_path,
                        'depth': current_depth,
                        'error': str(e),
                        'has_value': False
                    }
                    nodes_info.append(error_info)
                    print(f"警告: 访问节点 {current_path} 时出错: {e}")

            print(f"开始遍历ASPEN节点并获取值: {base_path}")

            # 开始遍历
            traverse_node_with_values(base_path)

            # 保存到文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"# ASPEN节点结构及值导出报告\n")
                f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 根路径: {base_path}\n")
                f.write(f"# 最大深度: {max_depth}\n")
                f.write(f"# 总节点数: {len(nodes_info)}\n")
                f.write(f"# ==========================================\n\n")

                for i, node_info in enumerate(nodes_info, 1):
                    # 根据深度缩进
                    indent = "  " * node_info.get('depth', 0)

                    # 处理错误信息
                    if 'error' in node_info:
                        f.write(f"{i:4d}. {indent}# ERROR at {node_info['path']}: {node_info['error']}\n")
                        continue

                    # 正常节点信息
                    path_display = node_info['path']

                    if node_info['has_value']:
                        value_str = f" = {node_info['value']}"
                        if node_info['units']:
                            value_str += f" [{node_info['units']}]"
                        f.write(f"{i:4d}. {indent}{path_display}{value_str}\n")
                    else:
                        f.write(f"{i:4d}. {indent}{path_display}\n")

            print(f"节点遍历完成!")
            print(f"发现 {len(nodes_info)} 个节点")
            print(f"已保存到: {os.path.abspath(output_file)}")

            return nodes_info, output_file

        except Exception as e:
            print(f"遍历ASPEN节点时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], None

    def export_aspen_node_structure_with_values(self, base_path, output_file=None, max_depth=3, only_with_value=False):
        """
        获取节点结构并包含关键节点的值

        参数:
        - base_path: 要遍历的ASPEN节点路径
        - output_file: 输出文件路径
        - max_depth: 最大递归深度，默认3层
        - only_with_value: 是否只输出包含值的节点，True则只输出有值的节点，False输出所有节点

        返回:
        - 包含节点路径和值的字典
        """
        try:
            # 生成默认输出文件名
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_path = base_path.replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                suffix = "_with_values_only" if only_with_value else "_with_values"
                output_file = f"aspen_nodes{suffix}_{safe_path}_{timestamp}.txt"

            # 用于存储节点信息
            nodes_info = []

            def traverse_node_with_values(current_path, current_depth=0):
                """递归遍历节点并获取值的辅助函数"""
                # 检查深度限制
                if max_depth is not None and current_depth > max_depth:
                    return

                try:
                    # 获取当前节点的子节点
                    child_nodes = self.get_child_nodes(current_path)

                    for child in child_nodes:
                        # 构建完整子节点路径
                        if current_path.endswith('\\'):
                            child_path = current_path + child
                        else:
                            child_path = current_path + '\\' + child

                        # 尝试获取节点值
                        node_value = None
                        node_units = None

                        try:
                            # 先尝试获取值
                            node_value = self.safe_get_node_value(child_path)

                            # 如果获取到值，再尝试获取单位
                            if node_value is not None:
                                try:
                                    node_units = self.safe_get_node_units(child_path)
                                except:
                                    pass
                        except:
                            pass

                        # 判断是否需要记录此节点
                        # 如果只记录有值的节点且节点无值，则跳过
                        if only_with_value and node_value is None:
                            # 如果设置了only_with_value=True，即使当前节点无值，我们仍需要检查其子节点
                            # 因为子节点可能有值
                            pass
                        else:
                            # 创建节点信息字典
                            node_info = {
                                'path': child_path,
                                'depth': current_depth + 1,
                                'has_value': node_value is not None,
                                'value': node_value,
                                'units': node_units
                            }

                            nodes_info.append(node_info)

                        # 继续遍历子节点（无论当前节点是否有值）
                        traverse_node_with_values(child_path, current_depth + 1)

                except Exception as e:
                    # 记录错误
                    error_info = {
                        'path': current_path,
                        'depth': current_depth,
                        'error': str(e),
                        'has_value': False
                    }

                    # 如果只记录有值的节点，错误节点不记录
                    if not only_with_value:
                        nodes_info.append(error_info)

                    print(f"警告: 访问节点 {current_path} 时出错: {e}")

            print(f"开始遍历ASPEN节点并获取值: {base_path}")
            print(f"只输出有值节点: {only_with_value}")

            # 开始遍历
            traverse_node_with_values(base_path)

            # 如果只输出有值的节点，过滤一下列表（可能包含之前添加的无值节点）
            if only_with_value:
                nodes_info = [node for node in nodes_info if node.get('has_value', False) and 'error' not in node]

            # 保存到文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"# ASPEN节点结构及值导出报告\n")
                f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 根路径: {base_path}\n")
                f.write(f"# 最大深度: {max_depth}\n")
                f.write(f"# 只输出有值节点: {only_with_value}\n")
                f.write(f"# 总节点数: {len(nodes_info)}\n")
                f.write(f"# ==========================================\n\n")

                for i, node_info in enumerate(nodes_info, 1):
                    # 根据深度缩进
                    indent = "  " * node_info.get('depth', 0)

                    # 处理错误信息
                    if 'error' in node_info:
                        if not only_with_value:  # 只有非only_with_value模式才输出错误
                            f.write(f"{i:4d}. {indent}# ERROR at {node_info['path']}: {node_info['error']}\n")
                        continue

                    # 正常节点信息
                    path_display = node_info['path']

                    if node_info['has_value']:
                        value_str = f" = {node_info['value']}"
                        if node_info['units']:
                            value_str += f" [{node_info['units']}]"
                        f.write(f"{i:4d}. {indent}{path_display}{value_str}\n")
                    else:
                        # 如果没有值且不是only_with_value模式，输出路径
                        if not only_with_value:
                            f.write(f"{i:4d}. {indent}{path_display}\n")

            print(f"节点遍历完成!")
            print(f"发现 {len(nodes_info)} 个节点")
            print(f"已保存到: {os.path.abspath(output_file)}")

            return nodes_info, output_file

        except Exception as e:
            print(f"遍历ASPEN节点时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], None

    def find_nodes_by_pattern(self, base_path, pattern, output_file=None, case_sensitive=False):
        """
        在指定节点下搜索包含特定模式的所有节点

        参数:
        - base_path: 搜索的根路径
        - pattern: 要搜索的模式（字符串）
        - output_file: 输出文件路径
        - case_sensitive: 是否区分大小写

        返回:
        - 匹配的节点路径列表
        """
        try:
            # 获取所有节点
            all_nodes, _ = self.export_aspen_node_structure(base_path, None, None)

            # 过滤匹配的节点
            if case_sensitive:
                matched_nodes = [node for node in all_nodes if pattern in node and not node.startswith("# ERROR:")]
            else:
                pattern_lower = pattern.lower()
                matched_nodes = [node for node in all_nodes if
                                 pattern_lower in node.lower() and not node.startswith("# ERROR:")]

            # 生成默认输出文件名
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_pattern = pattern.replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                output_file = f"aspen_search_{safe_pattern}_{timestamp}.txt"

            # 保存到文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"# ASPEN节点搜索报告\n")
                f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 搜索根路径: {base_path}\n")
                f.write(f"# 搜索模式: '{pattern}'\n")
                f.write(f"# 匹配节点数: {len(matched_nodes)}\n")
                f.write(f"# ==========================================\n\n")

                for i, node_path in enumerate(matched_nodes, 1):
                    f.write(f"{i:4d}. {node_path}\n")

            print(f"搜索完成!")
            print(f"在 '{base_path}' 下搜索模式 '{pattern}'")
            print(f"找到 {len(matched_nodes)} 个匹配项")
            print(f"已保存到: {os.path.abspath(output_file)}")

            return matched_nodes, output_file

        except Exception as e:
            print(f"搜索节点时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], None

    def find_nodes_by_value(self, base_path, target_value, value_type="exact",
                            case_sensitive=False, max_depth=None, output_file=None):
        """
        根据节点的值搜索节点路径

        参数:
        - base_path: 搜索的根路径
        - target_value: 要搜索的值（可以是字符串、数字等）
        - value_type: 匹配类型，可选 "exact"（精确匹配）、"contains"（包含）、"startswith"（开头匹配）、"endswith"（结尾匹配）
        - case_sensitive: 是否区分大小写（仅对字符串有效）
        - max_depth: 最大递归深度，None表示无限制
        - output_file: 输出文件路径，如果为None则自动生成

        返回:
        - 包含匹配节点路径和值的列表
        """
        try:
            # 生成默认输出文件名
            if output_file is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                safe_base = base_path.replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                safe_value = str(target_value)[:50].replace('\\', '_').replace('/', '_').replace(':', '').strip('_')
                output_file = f"aspen_value_search_{safe_base}_{safe_value}_{timestamp}.txt"

            # 用于存储匹配结果
            matched_nodes = []

            def traverse_and_search(current_path, current_depth=0):
                """递归遍历节点并搜索值的辅助函数"""
                # 检查深度限制
                if max_depth is not None and current_depth > max_depth:
                    return

                try:
                    # 首先尝试获取当前节点的值
                    node_value = self.safe_get_node_value(current_path)

                    # 如果获取到值，检查是否匹配
                    if node_value is not None:
                        is_match = False

                        # 根据匹配类型检查
                        if value_type == "exact":
                            # 精确匹配
                            if isinstance(target_value, str) and isinstance(node_value, str):
                                if case_sensitive:
                                    is_match = (node_value == target_value)
                                else:
                                    is_match = (node_value.lower() == target_value.lower())
                            else:
                                # 对于非字符串类型，尝试直接比较
                                try:
                                    is_match = (node_value == target_value)
                                except:
                                    # 如果比较失败，转换为字符串比较
                                    is_match = (str(node_value) == str(target_value))

                        elif value_type == "contains":
                            # 包含匹配
                            node_str = str(node_value)
                            target_str = str(target_value)

                            if case_sensitive:
                                is_match = (target_str in node_str)
                            else:
                                is_match = (target_str.lower() in node_str.lower())

                        elif value_type == "startswith":
                            # 开头匹配
                            node_str = str(node_value)
                            target_str = str(target_value)

                            if case_sensitive:
                                is_match = node_str.startswith(target_str)
                            else:
                                is_match = node_str.lower().startswith(target_str.lower())

                        elif value_type == "endswith":
                            # 结尾匹配
                            node_str = str(node_value)
                            target_str = str(target_value)

                            if case_sensitive:
                                is_match = node_str.endswith(target_str)
                            else:
                                is_match = node_str.lower().endswith(target_str.lower())

                        elif value_type == "numeric_range":
                            # 数值范围匹配（当target_value是元组(min, max)时）
                            try:
                                if isinstance(target_value, (tuple, list)) and len(target_value) == 2:
                                    min_val, max_val = target_value
                                    # 尝试将节点值转换为数值
                                    node_num = float(node_value)
                                    is_match = (min_val <= node_num <= max_val)
                            except (ValueError, TypeError):
                                is_match = False

                        # 如果匹配，添加到结果列表
                        if is_match:
                            # 获取单位（如果存在）
                            units = None
                            try:
                                units = self.safe_get_node_units(current_path)
                            except:
                                pass

                            matched_nodes.append({
                                "path": current_path,
                                "value": node_value,
                                "units": units,
                                "depth": current_depth
                            })

                    # 然后获取子节点并递归遍历
                    child_nodes = self.get_child_nodes(current_path)
                    for child in child_nodes:
                        # 构建子节点路径
                        if current_path.endswith('\\'):
                            child_path = current_path + child
                        else:
                            child_path = current_path + '\\' + child

                        # 递归遍历
                        traverse_and_search(child_path, current_depth + 1)

                except Exception as e:
                    # 记录错误但不中断遍历
                    error_msg = f"访问节点 {current_path} 时出错: {e}"
                    print(f"警告: {error_msg}")

            print(f"开始搜索节点值: {target_value}")
            print(f"根路径: {base_path}")
            print(f"匹配类型: {value_type}")
            print(f"区分大小写: {case_sensitive}")

            # 开始遍历搜索
            traverse_and_search(base_path)

            # 保存结果到文件
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(f"# ASPEN节点值搜索报告\n")
                f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# 根路径: {base_path}\n")
                f.write(f"# 搜索值: {target_value}\n")
                f.write(f"# 匹配类型: {value_type}\n")
                f.write(f"# 区分大小写: {case_sensitive}\n")
                f.write(f"# 最大深度: {'无限制' if max_depth is None else max_depth}\n")
                f.write(f"# 匹配节点数: {len(matched_nodes)}\n")
                f.write(f"# ==========================================\n\n")

                for i, node_info in enumerate(matched_nodes, 1):
                    indent = "  " * node_info["depth"]
                    path_display = node_info["path"]

                    # 构建值显示字符串
                    value_str = f" = {node_info['value']}"
                    if node_info['units']:
                        value_str += f" [{node_info['units']}]"

                    f.write(f"{i:4d}. {indent}{path_display}{value_str}\n")

                    print(f"找到匹配路径：{indent}{path_display}{value_str}\n")

            print(f"搜索完成!")
            print(f"找到 {len(matched_nodes)} 个匹配项")
            print(f"已保存到: {output_file}")

            return matched_nodes, output_file

        except Exception as e:
            print(f"搜索节点值时出错: {e}")
            import traceback
            traceback.print_exc()
            return [], None

# 使用方法
if __name__ == "__main__":
    converter = AspenToJSONConverter(r"D:\aspen\orgfile\Example9.1-DesignSpec.bkp")
    design_spec_path = r'\Data\Flowsheeting Options\Design-Spec\DS-1\Input'
    if converter.connect_to_aspen():
        # # 示例1: 搜索节点下的子节点结构
        # nodes, file1 = converter.export_aspen_node_structure(design_spec_path)
        # 示例2：搜索节点下的子节点结构及对应value
        nodes_with_values, file2 = converter.export_aspen_node_structure_with_values(
            design_spec_path,
            max_depth=3,
            only_with_value=True
        )
        # # 示例3: 搜索包含特定值的节点
        # tol_nodes, file3 = converter.find_nodes_by_value(
        #     design_spec_path,
        #     "150",
        #     value_type="exact",
        #     max_depth=3
        # )

