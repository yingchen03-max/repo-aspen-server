# Schema API Examples

服务当前只支持 HTTPS，下面示例统一使用：

- `https://localhost:6000`
- 本地测试时可使用 `curl -k`
- Python `requests` 可使用 `verify=False`

## 列出全部 Schema

```bash
curl -k https://localhost:6000/api/schema/list
```

```python
import requests

response = requests.get(
    "https://localhost:6000/api/schema/list",
    verify=False,
    timeout=10,
)
data = response.json()
print(data["count"])
```

## 获取基础 Schema

```bash
curl -k "https://localhost:6000/api/schema?types=base"
```

```python
import requests

response = requests.get(
    "https://localhost:6000/api/schema",
    params={"types": "base"},
    verify=False,
    timeout=10,
)
base_schema = response.json()["schemas"]["base"]
```

## 获取多个设备 Schema

```bash
curl -k "https://localhost:6000/api/schema?types=Mixer,Heater,Pump"
```

```python
import requests

response = requests.get(
    "https://localhost:6000/api/schema",
    params={"types": "Mixer,Heater,Pump"},
    verify=False,
    timeout=10,
)
schemas = response.json()["schemas"]
```

## 只返回 Schema 列表

```bash
curl -k "https://localhost:6000/api/schema?types=all&format=list"
```

## 提交模拟前先拉取 Schema

```python
import requests

session = requests.Session()
session.verify = False

schema_resp = session.get(
    "https://localhost:6000/api/schema",
    params={"types": "Mixer,Heater"},
    timeout=10,
)
schemas = schema_resp.json()["schemas"]

config = {
    "setup": {
        "TITLE": "demo",
        "PROP_SET": "NRTL-RK",
    },
    "components": ["WATER", "ETHANOL"],
    "blocks": [
        {"name": "MIX1", "type": "Mixer"},
        {"name": "HEAT1", "type": "Heater"},
    ],
}

run_resp = session.post(
    "https://localhost:6000/run-aspen-simulation",
    json=config,
    timeout=120,
)
print(run_resp.json())
```

## JavaScript 示例

```javascript
const axios = require("axios");
const https = require("https");

const client = axios.create({
  baseURL: "https://localhost:6000",
  httpsAgent: new https.Agent({ rejectUnauthorized: false }),
  timeout: 10000,
});

async function main() {
  const response = await client.get("/api/schema", {
    params: { types: "Mixer,Heater,Pump" },
  });
  console.log(Object.keys(response.data.schemas));
}

main().catch((error) => {
  console.error(error.message);
});
```

## 保存返回结果

```bash
curl -k "https://localhost:6000/api/schema?types=Mixer" -o mixer_schema.json
curl -k "https://localhost:6000/api/schema?types=all&format=list" -o schema_list.json
```

## jq 处理示例

```bash
curl -ks "https://localhost:6000/api/schema/list" | jq -r ".schemas[].type"
curl -ks "https://localhost:6000/api/schema/list" | jq ".schemas[] | select(.category == \"block\") | .type"
```

## 封装校验器

```python
import requests
import jsonschema


class ConfigValidator:
    def __init__(self, api_url="https://localhost:6000", verify_ssl=False):
        self.api_url = api_url
        self.verify_ssl = verify_ssl
        self.schemas = {}

    def load_schema(self, schema_type):
        if schema_type not in self.schemas:
            response = requests.get(
                f"{self.api_url}/api/schema",
                params={"types": schema_type},
                verify=self.verify_ssl,
                timeout=10,
            )
            self.schemas[schema_type] = response.json()["schemas"][schema_type]
        return self.schemas[schema_type]

    def validate(self, config, schema_type):
        schema = self.load_schema(schema_type)
        jsonschema.validate(config, schema)


validator = ConfigValidator()
validator.validate({}, "Mixer")
```

## 生产环境建议

- 使用正式 CA 证书替换自签名证书
- 将 `SSL_CERT_FILE` 和 `SSL_KEY_FILE` 指向正式证书
- 客户端启用证书校验，不要继续使用 `-k` 或 `verify=False`
