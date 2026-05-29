# Quick Start

## 1. 安装依赖

```bat
cd server
pip install -r requirements.txt
python -m pywin32_postinstall -install
```

## 2. 准备配置

```bat
copy .env.example .env
notepad .env
```

最小配置：

```env
ASPEN_BASE_DIR=D:/aspen
ASPEN_SIMULATOR_PORT=6000
SSL_CERT_FILE=ssl/cert.pem
SSL_KEY_FILE=ssl/key.pem
```

可选扩展：

```env
ASPEN_SSL_HOSTS=localhost,127.0.0.1,my-hostname,10.10.10.20
```

## 3. 启动服务

推荐：

```bat
cd server
start.bat
```

或手工执行：

```bat
cd server
python bootstrap_ssl.py
python aspenagent.py
```

启动后应看到类似输出：

```text
Starting Aspen simulation service in HTTPS mode
  Certificate: D:\code\dicp\aspen_python\server\ssl\cert.pem
  Private key: D:\code\dicp\aspen_python\server\ssl\key.pem
  Listen: 0.0.0.0:6000
```

## 4. 验证服务

```bash
curl -k https://localhost:6000/health
```

或：

```bat
cd server
python test_service.py
```

## 5. 发起模拟请求

```bash
curl -k -X POST https://localhost:6000/run-aspen-simulation ^
  -H "Content-Type: application/json" ^
  -d @config.json
```

## 注意

- 服务只支持 HTTPS
- 不再支持 HTTP 启动
- 如果证书文件不存在，会自动生成固定自签名证书
- 生产环境建议替换为正式证书
