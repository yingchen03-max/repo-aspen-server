# Schema API 设计文档

## 设计理念

Schema API 设计遵循以下原则：

1. **灵活性**: 支持单个、多个或全部 Schema 查询
2. **RESTful**: 使用标准 HTTP 方法和状态码
3. **易用性**: 简单的查询参数，清晰的响应结构
4. **性能**: 支持按需加载，避免不必要的数据传输
5. **容错性**: 部分失败不影响其他结果返回

## API 端点设计

### 端点 1: GET /api/schema/list

**用途**: 快速浏览所有可用的 Schema

**优点**:
- 轻量级，只返回元数据
- 适合用于下拉菜单、自动完成等场景
- 包含文件大小信息，便于客户端决策

**响应结构**:
```json
{
  "success": true,
  "count": 23,
  "schemas": [
    {
      "type": "设备类型名称",
      "file": "文件名",
      "category": "分类（base/block/other）",
      "description": "描述",
      "size": 文件大小（字节）
    }
  ]
}
```

### 端点 2: GET /api/schema

**用途**: 获取具体的 Schema 内容

**查询参数**:
- `types`: 设备类型（必需）
  - 单个: `types=Mixer`
  - 多个: `types=Mixer,Heater,Pump`
  - 全部: `types=all`
  - 基础: `types=base`
  
- `format`: 返回格式（可选）
  - `json` (默认): 返回完整 Schema 内容
  - `list`: 仅返回列表（等同于 /api/schema/list）

**优点**:
- 支持批量获取，减少请求次数
- 灵活的类型匹配（自动处理文件名变体）
- 部分失败不影响其他结果

**响应结构**:
```json
{
  "success": true/false,
  "requested": ["请求的类型列表"],
  "found": 找到的数量,
  "not_found": ["未找到的类型"],  // 仅在有未找到时出现
  "message": "错误或警告信息",      // 仅在有问题时出现
  "schemas": {
    "类型名": Schema内容对象
  }
}
```

## 设计决策

### 1. 为什么使用查询参数而不是路径参数？

**选择**: `GET /api/schema?types=Mixer,Heater`  
**而不是**: `GET /api/schema/Mixer,Heater`

**原因**:
- 支持多个类型更自然
- 查询参数更灵活，易于扩展
- 符合 RESTful 最佳实践（集合资源的过滤）

### 2. 为什么支持 format 参数？

**原因**:
- 有时只需要列表（如下拉菜单）
- 有时需要完整内容（如配置验证）
- 避免创建过多端点
- 提供统一的查询接口

### 3. 为什么部分失败仍返回 200？

**原因**:
- 部分成功仍有价值
- 客户端可以根据 `success` 字段判断
- `not_found` 列表提供详细信息
- 符合实际使用场景（用户可能输错类型名）

### 4. 文件名匹配策略

支持多种文件名格式：
```python
# 对于类型 "Mixer"，尝试匹配：
- blocks_Mixer_data.json  # 标准格式
- Mixer.json              # 简化格式
- blocks_Mixer.json       # 变体格式
```

**原因**:
- 提高容错性
- 支持不同的命名约定
- 用户体验更好

## 使用场景

### 场景 1: 前端动态表单

```
用户选择设备类型 → 调用 /api/schema?types=<type>
→ 根据 Schema 生成表单 → 用户填写 → 提交模拟
```

### 场景 2: 配置验证工具

```
用户上传配置 → 解析配置中的设备类型
→ 调用 /api/schema?types=<types> → 验证配置
→ 显示验证结果
```

### 场景 3: 文档生成

```
调用 /api/schema/list → 获取所有类型
→ 逐个调用 /api/schema?types=<type>
→ 生成 Markdown/HTML 文档
```

### 场景 4: IDE 插件

```
用户编辑配置文件 → IDE 检测设备类型
→ 调用 /api/schema?types=<type>
→ 提供自动完成和验证
```

## 扩展建议

### 未来可能的扩展

1. **Schema 版本控制**:
   ```
   GET /api/schema?types=Mixer&version=v1.0
   ```

2. **Schema 搜索**:
   ```
   GET /api/schema/search?q=temperature
   ```

3. **Schema 差异对比**:
   ```
   GET /api/schema/diff?type=Mixer&v1=1.0&v2=2.0
   ```

4. **Schema 验证端点**:
   ```
   POST /api/schema/validate
   Body: { "type": "Mixer", "config": {...} }
   ```

5. **Schema 合并**:
   ```
   GET /api/schema/merged?types=Mixer,Heater
   返回合并后的完整 Schema
   ```

## 性能考虑

1. **缓存**: 考虑添加内存缓存，Schema 文件不常变化
2. **压缩**: 对大型响应启用 gzip 压缩
3. **分页**: 如果 Schema 数量很大，考虑分页
4. **CDN**: 生产环境可以将 Schema 文件放到 CDN

## 安全考虑

1. **路径遍历**: 已通过限制在 SCHEMA_DIR 目录内防止
2. **DOS 攻击**: 考虑添加速率限制
3. **文件大小**: 限制单次请求的 Schema 数量
4. **输入验证**: 验证 types 参数格式

---

这个设计平衡了灵活性、性能和易用性，适合各种使用场景。
