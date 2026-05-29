"""
Aspen Agent 配置文件
用于管理所有路径和环境配置
"""
import os
from pathlib import Path
from dotenv import load_dotenv
from ssl_utils import ensure_server_certificate, resolve_ssl_paths

# 加载环境变量
load_dotenv()

# ==================== 基础配置 ====================
# 服务器配置
HOST = os.getenv("ASPEN_HOST", "0.0.0.0")
PORT = int(os.getenv("ASPEN_SIMULATOR_PORT", "6666"))
DEBUG = os.getenv("ASPEN_DEBUG", "True").lower() == "true"

# ==================== 路径配置 ====================
# 基础目录 - 可以通过环境变量覆盖
BASE_DIR = Path(os.getenv("ASPEN_BASE_DIR", "D:/aspen"))
BASE_DIR.mkdir(parents=True, exist_ok=True)

# 模板文件目录
TEMPLATE_DIR = BASE_DIR / "orgfile"
TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_TEMPLATE = TEMPLATE_DIR / "test.bkp"

# 输出文件目录
OUTPUT_DIR = BASE_DIR / "bkpfile"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 结果文件目录
RESULT_DIR = BASE_DIR / "resultfile"
RESULT_DIR.mkdir(parents=True, exist_ok=True)

# 配置文件保存目录
CONFIG_DIR = BASE_DIR / "configfile"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

# ==================== SSL 配置 ====================
# SSL 证书路径
SSL_CERT_FILE = os.getenv("SSL_CERT_FILE", "ssl/cert.pem")
SSL_KEY_FILE = os.getenv("SSL_KEY_FILE", "ssl/key.pem")
SSL_CERT_PATH, SSL_KEY_PATH = resolve_ssl_paths(SSL_CERT_FILE, SSL_KEY_FILE, Path(__file__).parent)

# ==================== Aspen Plus 配置 ====================
# Aspen Plus 可执行文件路径（可选）
ASPEN_EXECUTABLE = os.getenv("ASPEN_EXECUTABLE", "")

# Schema 文件目录
SCHEMA_DIR = Path(__file__).parent.parent / "schema"

# ==================== 日志配置 ====================
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")
LOG_FILE = BASE_DIR / "logs" / "aspen_agent.log"
LOG_FILE.parent.mkdir(parents=True, exist_ok=True)

# ==================== 配置验证 ====================
def validate_config():
    """验证配置是否正确"""
    issues = []
    
    # 检查必要的目录
    required_dirs = {
        "BASE_DIR": BASE_DIR,
        "TEMPLATE_DIR": TEMPLATE_DIR,
        "OUTPUT_DIR": OUTPUT_DIR,
        "RESULT_DIR": RESULT_DIR,
        "CONFIG_DIR": CONFIG_DIR,
    }
    
    for name, path in required_dirs.items():
        if not path.exists():
            issues.append(f"{name} 不存在: {path}")
    
    if not SSL_CERT_PATH.exists():
        issues.append(f"SSL 证书文件不存在: {SSL_CERT_PATH}")
    if not SSL_KEY_PATH.exists():
        issues.append(f"SSL 密钥文件不存在: {SSL_KEY_PATH}")
    
    return issues

def print_config():
    """打印当前配置"""
    print("=" * 60)
    print("Aspen Agent 配置")
    print("=" * 60)
    print(f"服务器地址: {HOST}:{PORT}")
    print(f"调试模式: {DEBUG}")
    print(f"基础目录: {BASE_DIR}")
    print(f"模板目录: {TEMPLATE_DIR}")
    print(f"输出目录: {OUTPUT_DIR}")
    print(f"结果目录: {RESULT_DIR}")
    print(f"配置目录: {CONFIG_DIR}")
    print(f"SSL 证书: {SSL_CERT_PATH}")
    print(f"SSL 密钥: {SSL_KEY_PATH}")
    print("=" * 60)
    
    # 验证配置
    issues = validate_config()
    if issues:
        print("\n⚠️  配置警告:")
        for issue in issues:
            print(f"  - {issue}")
        print()

if __name__ == "__main__":
    cert_path, key_path, created, hosts = ensure_server_certificate(
        SSL_CERT_FILE,
        SSL_KEY_FILE,
        base_dir=Path(__file__).parent,
    )
    if created:
        print(f"已创建 SSL 证书: {cert_path}")
        print(f"已创建 SSL 密钥: {key_path}")
        print(f"证书 SAN: {', '.join(hosts)}")
    print_config()
