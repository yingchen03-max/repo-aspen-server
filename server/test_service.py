"""
Aspen Agent 服务测试脚本
用于测试服务是否正常运行
"""
import requests
import json
import sys
from pathlib import Path

# 配置
API_URL = "https://localhost:6000"
VERIFY_SSL = False  # 开发环境跳过 SSL 验证

def test_health():
    """测试健康检查端点"""
    print("=" * 60)
    print("测试 1: 健康检查")
    print("=" * 60)
    
    try:
        response = requests.get(
            f"{API_URL}/health",
            verify=VERIFY_SSL,
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            print(f"✓ 服务正常运行")
            print(f"  状态: {data.get('status')}")
            print(f"  时间: {data.get('timestamp')}")
            print(f"  Aspen 可用: {data.get('aspen_available')}")
            return True
        else:
            print(f"✗ 健康检查失败: HTTP {response.status_code}")
            return False
            
    except requests.exceptions.ConnectionError:
        print(f"✗ 无法连接到服务: {API_URL}")
        print(f"  请确保服务已启动")
        return False
    except Exception as e:
        print(f"✗ 测试失败: {e}")
        return False

def test_schema_list():
    """测试 Schema 列表端点"""
    print("\n" + "=" * 60)
    print("测试 2: Schema 列表")
    print("=" * 60)
    
    try:
        response = requests.get(
            f"{API_URL}/api/schema/list",
            verify=VERIFY_SSL,
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            if data.get('success'):
                print(f"✓ 获取 Schema 列表成功")
                print(f"  可用 Schema 数量: {data.get('count')}")
                
                # 显示前5个
                schemas = data.get('schemas', [])
                print(f"  示例 Schema:")
                for schema in schemas[:5]:
                    print(f"    - {schema['type']}: {schema['description']}")
                if len(schemas) > 5:
                    print(f"    ... 还有 {len(schemas) - 5} 个")
                return True
            else:
                print(f"✗ 获取失败: {data.get('message')}")
                return False
        else:
            print(f"✗ 请求失败: HTTP {response.status_code}")
            return False
            
    except Exception as e:
        print(f"✗ 测试失败: {e}")
        return False

def test_schema_get():
    """测试获取 Schema 端点"""
    print("\n" + "=" * 60)
    print("测试 3: 获取 Schema")
    print("=" * 60)
    
    test_cases = [
        ('base', '基础 Schema'),
        ('Mixer,Heater', '多个设备 Schema'),
    ]
    
    all_passed = True
    for types, description in test_cases:
        try:
            print(f"\n  测试: {description} (types={types})")
            response = requests.get(
                f"{API_URL}/api/schema",
                params={'types': types},
                verify=VERIFY_SSL,
                timeout=5
            )
            
            if response.status_code == 200:
                data = response.json()
                if data.get('success'):
                    print(f"    ✓ 成功获取 {data.get('found')} 个 Schema")
                    for schema_type in data.get('schemas', {}).keys():
                        print(f"      - {schema_type}")
                else:
                    print(f"    ✗ 获取失败: {data.get('message')}")
                    all_passed = False
            else:
                print(f"    ✗ 请求失败: HTTP {response.status_code}")
                all_passed = False
                
        except Exception as e:
            print(f"    ✗ 测试失败: {e}")
            all_passed = False
    
    return all_passed

def test_simulation(config_file=None):
    """测试模拟端点"""
    print("\n" + "=" * 60)
    print("测试 4: 运行模拟")
    print("=" * 60)
    
    if config_file is None:
        print("⚠ 未提供配置文件，跳过模拟测试")
        print("  使用方法: python test_service.py <config.json>")
        return None
    
    config_path = Path(config_file)
    if not config_path.exists():
        print(f"✗ 配置文件不存在: {config_file}")
        return False
    
    try:
        # 读取配置
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        print(f"✓ 已读取配置文件: {config_file}")
        
        # 发送请求
        print(f"  正在提交模拟请求...")
        response = requests.post(
            f"{API_URL}/run-aspen-simulation",
            json=config,
            verify=VERIFY_SSL,
            timeout=300  # 5分钟超时
        )
        
        result = response.json()
        
        if result.get('success'):
            print(f"✓ 模拟成功完成")
            print(f"  Aspen 文件: {result.get('aspen_file_path')}")
            print(f"  配置文件: {result.get('config_file_path')}")
            print(f"  结果文件: {result.get('result_file_path')}")
            return True
        else:
            print(f"✗ 模拟失败")
            print(f"  错误类型: {result.get('error_type')}")
            print(f"  错误信息: {result.get('error_message', '')[:200]}...")
            if result.get('aspen_file_path'):
                print(f"  Aspen 文件: {result.get('aspen_file_path')}")
            return False
            
    except requests.exceptions.Timeout:
        print(f"✗ 请求超时（超过5分钟）")
        return False
    except Exception as e:
        print(f"✗ 测试失败: {e}")
        return False

def main():
    """主函数"""
    print("\n" + "=" * 60)
    print("Aspen Agent 服务测试")
    print("=" * 60)
    print(f"API URL: {API_URL}")
    print(f"SSL 验证: {'启用' if VERIFY_SSL else '禁用'}")
    print()
    
    # 禁用 SSL 警告
    if not VERIFY_SSL:
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    
    # 测试健康检查
    health_ok = test_health()
    
    if not health_ok:
        print("\n" + "=" * 60)
        print("测试失败: 服务未运行或无法访问")
        print("=" * 60)
        sys.exit(1)
    
    # 测试 Schema API
    schema_list_ok = test_schema_list()
    schema_get_ok = test_schema_get()
    
    # 测试模拟（如果提供了配置文件）
    config_file = sys.argv[1] if len(sys.argv) > 1 else None
    sim_ok = test_simulation(config_file) if config_file else None
    
    # 总结
    print("\n" + "=" * 60)
    print("测试总结")
    print("=" * 60)
    print(f"健康检查: {'✓ 通过' if health_ok else '✗ 失败'}")
    print(f"Schema 列表: {'✓ 通过' if schema_list_ok else '✗ 失败'}")
    print(f"Schema 获取: {'✓ 通过' if schema_get_ok else '✗ 失败'}")
    print(f"模拟测试: {'✓ 通过' if sim_ok else '✗ 失败' if sim_ok is not None else '⚠ 跳过'}")
    print("=" * 60)
    
    if not (health_ok and schema_list_ok and schema_get_ok):
        sys.exit(1)
    if sim_ok is False:
        sys.exit(1)

if __name__ == "__main__":
    main()
