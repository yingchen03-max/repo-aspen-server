# Server README

## 服务概览

`server/` 目录提供 Aspen Plus 的 Flask 服务封装，核心能力包括：

- Aspen 模拟执行
- 健康检查
- Schema 查询
- HTTPS 服务启动

## 关键文件

- `aspenagent.py`：服务主入口
- `config.py`：配置读取与校验
- `bootstrap_ssl.py`：启动前证书准备脚本
- `ssl_utils.py`：固定 SSL 证书生成与路径解析
- `start.bat`：Windows 启动脚本
- `test_service.py`：接口联通测试脚本

## 配置

复制并编辑配置文件：

```bat
copy .env.example .env
notepad .env
```

常用配置项：

```env
ASPEN_HOST=0.0.0.0
ASPEN_SIMULATOR_PORT=6000
ASPEN_BASE_DIR=D:/aspen
SSL_CERT_FILE=ssl/cert.pem
SSL_KEY_FILE=ssl/key.pem
```

如需让自动生成证书覆盖指定主机名或 IP：

```env
ASPEN_SSL_HOSTS=localhost,127.0.0.1,my-hostname,10.10.10.20
```

## HTTPS 机制

当前服务只支持 HTTPS。

启动流程如下：

1. 解析 `SSL_CERT_FILE` 和 `SSL_KEY_FILE`
2. 如果文件不存在，则生成固定自签名证书
3. 使用 `ssl.SSLContext` 加载证书并启动服务
4. 不再允许回退到 HTTP
5. 不再使用 Flask `adhoc` 临时证书

默认生成位置：

- `server/ssl/cert.pem`
- `server/ssl/key.pem`

## 启动

推荐方式：

```bat
start.bat
```

该脚本会：

- 检查 Python
- 检查依赖
- 校验配置
- 生成或复用 SSL 证书
- 启动 HTTPS 服务

手工启动：

```bat
python bootstrap_ssl.py
python aspenagent.py
```

## 测试

健康检查：

```bash
curl -k https://localhost:6000/health
```

Schema 列表：

```bash
curl -k https://localhost:6000/api/schema/list
```

查询指定 Schema：

```bash
curl -k "https://localhost:6000/api/schema?types=base"
curl -k "https://localhost:6000/api/schema?types=Mixer,Heater,Pump"
```

执行模拟：

```bash
curl -k -X POST https://localhost:6000/run-aspen-simulation ^
  -H "Content-Type: application/json" ^
  -d @config.json
```

Python 测试脚本：

```bat
python test_service.py
```

## 常见问题

### SSL 连接失败

优先检查：

- `server/ssl/` 下证书文件是否存在
- 当前访问地址是否包含在证书 SAN 中
- 客户端是否对自签名证书执行了严格校验

如果是跨主机访问，建议在 `.env` 中配置：

```env
ASPEN_SSL_HOSTS=服务域名,服务IP,localhost,127.0.0.1
```

### 客户端校验证书失败

开发或测试环境：

- `curl` 使用 `-k`
- Python `requests` 使用 `verify=False`

生产环境：

- 使用受信任 CA 签发的证书
- 将 `SSL_CERT_FILE`、`SSL_KEY_FILE` 指向正式证书路径

### Aspen COM 连接失败

检查：

- Aspen Plus 是否已安装
- 当前账号是否有 Aspen 权限
- `pywin32` 是否已安装并执行过 `pywin32_postinstall`
