# Aspen Python

Aspen Plus 自动化服务项目，提供 Aspen 模拟执行、Schema 查询和结果导出能力。

当前版本已经改为：

- 只允许 HTTPS 启动
- 启动前自动检查并生成固定 SSL 证书
- 不再支持 HTTP 启动
- 不再使用 Flask `adhoc` 临时证书

## 目录结构

```text
aspen_python/
|-- aspen_utils/            Aspen 辅助脚本
|-- schema/                 JSON Schema 定义
|-- server/                 Flask 服务
|   |-- aspenagent.py       服务入口
|   |-- config.py           配置文件
|   |-- bootstrap_ssl.py    启动前证书准备
|   |-- ssl_utils.py        SSL 证书生成与校验
|   |-- start.bat           Windows 启动脚本
|   `-- ssl/                运行时生成的证书目录
|-- QUICKSTART.md           快速启动
`-- README.md
```

## 运行要求

- Windows
- Python 3.8+
- 已安装 Aspen Plus
- 当前用户具备 Aspen COM 调用权限

## HTTPS 说明

服务启动时会优先使用 `server/.env` 中配置的证书路径：

```env
SSL_CERT_FILE=ssl/cert.pem
SSL_KEY_FILE=ssl/key.pem
```

如果文件不存在，启动脚本会自动创建一套固定自签名证书，而不是使用临时证书。生成后的默认位置：

- `server/ssl/cert.pem`
- `server/ssl/key.pem`

生成证书时会自动写入以下 SAN：

- `localhost`
- `127.0.0.1`
- `::1`
- 当前主机名
- 当前主机可发现的本机 IP

如果跨主机访问时仍出现主机名校验问题，可以在 `.env` 中增加：

```env
ASPEN_SSL_HOSTS=localhost,127.0.0.1,my-hostname,10.10.10.20
```

## 启动方式

推荐使用：

```bat
cd server
start.bat
```

启动脚本会执行以下步骤：

1. 检查 Python 和依赖
2. 校验配置
3. 创建或复用固定 SSL 证书
4. 以 HTTPS 模式启动 Flask 服务

也可以直接运行：

```bat
cd server
python bootstrap_ssl.py
python aspenagent.py
```

## 常用接口

健康检查：

```bash
curl -k https://localhost:6000/health
```

执行模拟：

```bash
curl -k -X POST https://localhost:6000/run-aspen-simulation ^
  -H "Content-Type: application/json" ^
  -d @config.json
```

查询 Schema：

```bash
curl -k "https://localhost:6000/api/schema?types=base"
curl -k "https://localhost:6000/api/schema?types=Mixer"
curl -k "https://localhost:6000/api/schema?types=all&format=list"
```

## 常见问题

### 1. 拷贝到其他主机后 SSL 连接失败

常见原因：

- 旧版本服务使用了临时证书
- 新主机没有预置证书文件
- 证书 SAN 不包含当前访问使用的主机名或 IP

现在的处理方式：

- 启动前自动生成固定证书
- 后续重启复用同一套证书
- 可通过 `ASPEN_SSL_HOSTS` 扩展 SAN

### 2. 证书是自签名，客户端校验失败

开发或内网测试场景可以：

- `curl` 使用 `-k`
- Python `requests` 使用 `verify=False`

生产环境建议替换为受信任证书，并将 `SSL_CERT_FILE`、`SSL_KEY_FILE` 指向正式证书文件。

### 3. 服务无法连接 Aspen Plus

检查：

- Aspen Plus 是否已安装
- 是否能手工打开 Aspen Plus
- 是否以具备权限的 Windows 用户运行服务
- `pywin32` 是否安装完整

## 更多文档

- [快速启动](./QUICKSTART.md)
- [服务端说明](./server/README.md)
