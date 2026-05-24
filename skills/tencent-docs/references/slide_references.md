# 幻灯片（Slide / PPT）参考文档

本文件包含腾讯文档 MCP 幻灯片相关工具的使用指南和注意事项。

---

## 概述

幻灯片通过 `create_slide` 工具创建，AI 自动根据用户描述和参考资料生成 PPT 内容。该接口为**异步接口**，需配合 `slide_progress` 工具轮询进度。

**多轮对话编辑**：支持通过 `session_id` 对已生成的 PPT 进行多轮编辑，无需重新创建。

**推荐方式**：使用 `generate_slide.js` 脚本自动完成创建/编辑和进度轮询的完整流程，避免 agent 手动轮询的不确定性。

---

## 工具列表

| 工具名称 | 功能说明 |
|---------|---------|
| create_slide | 创建或编辑幻灯片（AI 自动生成内容，异步接口，支持多轮对话） |
| slide_progress | 查询幻灯片生成进度 |

---

## 工具详细说明

### 1. create_slide

#### 功能说明
根据用户描述和参考资料，由 AI 自动生成或编辑幻灯片内容。支持两种模式：
- **首次创建**：不传 `session_id`，发起新的 PPT 生成任务
- **多轮编辑**：传入之前返回的 `session_id`，对已有 PPT 进行修改

#### 调用示例

**示例1：根据主题生成 PPT（首次创建）**
```json
{
  "description": "生成一份主题为'2024年度销售总结'的PPT，要求包含业绩回顾、亮点项目、问题分析和来年规划四个章节"
}
```

**示例2：根据参考材料生成 PPT（首次创建）**
```json
{
  "reference_context": "第一季度销售额达到1200万，同比增长25%。主要增长来自华南区域，新客户占比40%。存在问题：北方市场渗透率不足，客单价偏低。",
  "description": "根据材料生成PPT，要求风格简洁专业，重点突出数据亮点"
}
```

**示例3：多轮编辑已有 PPT**
```json
{
  "session_id": "session_1234567890",
  "description": "把第3页的标题改成'2024年Q1销售回顾'，并在最后增加一页'下一步计划'"
}
```

#### 参数说明
- `description` (string, 必填): 用户对 PPT 的要求描述。首次创建样例：【生成一份主题为xxx的PPT，要求xxxx】；多轮编辑样例：【把第3页标题改成xxx】【增加一页关于xxx的内容】
- `reference_context` (string, 可选): 生成 PPT 的参考资料，必须是 UTF-8 文本格式。**仅当用户明确指定需要根据某段内容/材料生成PPT时才传此参数，不要自由发挥填充内容**
- `session_id` (string, 可选): 会话ID，用于多轮对话编辑PPT。首次创建时不传，后续编辑时传入上一次 create_slide 返回的 session_id

#### 返回值说明
```json
{
  "session_id": "session_1234567890",
  "error": "",
  "trace_id": "trace_1234567890"
}
```

> ⚠️ **注意**：`create_slide` 为异步接口，返回 `session_id` 后需轮询进度。**推荐使用 `generate_slide.js` 脚本**自动处理完整流程。

### 2. slide_progress

#### 功能说明
查询幻灯片生成进度，与 `create_slide` 配合使用。通常由 `generate_slide.js` 脚本自动调用，无需手动轮询。

#### 状态说明
- `in_progress`：进行中，继续轮询
- `completed`：已完成，幻灯片已生成，从响应中获取 `file_url`
- `failed`：失败，停止轮询
- `canceled`：已取消，停止轮询
- `not_found`：未找到（`session_id` 不正确或已过期），停止轮询

#### 调用示例
```json
{
  "session_id": "session_1234567890"
}
```

#### 参数说明
- `session_id` (string, 必填): `create_slide` 返回的异步任务 session_id

#### 返回值说明
```json
{
  "status": "completed",
  "file_url": "https://docs.qq.com/slide/DV2h5cWJ0R1lQb0lH",
  "error": "",
  "trace_id": "trace_1234567890"
}
```

---

## 典型工作流

### ✅ 推荐做法：使用 generate_slide.js 脚本

使用 JS 脚本自动完成创建/编辑和进度轮询的完整流程，避免 agent 手动轮询的不确定性。

#### 首次创建幻灯片
```bash
node generate_slide.js --description "生成一份关于AI技术发展趋势的PPT"
```

#### 根据参考材料创建幻灯片
```bash
node generate_slide.js --description "根据材料生成PPT，重点突出数据" --reference_context "Q1销售额1200万，同比增长25%..."
```

#### 多轮编辑已有幻灯片
```bash
node generate_slide.js --description "把第3页标题改成'新标题'，增加一页总结" --session_id "session_1234567890"
```

#### 脚本输出格式

**成功时：**
```
SLIDE_COMPLETED
SESSION_ID:<session_id>
FILE_URL:<file_url>
```

**失败时：**
```
SLIDE_FAILED
ERROR:<error_message>
```

### Agent 执行指引

#### 标准工作流

1. **解析用户意图**：判断是首次创建还是多轮编辑
   - 首次创建：用户首次要求生成PPT，无 session_id
   - 多轮编辑：用户要求修改已有PPT，有 session_id
2. **执行脚本**：使用 `generate_slide.js` 脚本完成任务
3. **解析输出**：从脚本输出中提取 `SESSION_ID` 和 `FILE_URL`
4. **反馈用户**：
   - 成功：返回 `file_url` 链接，并提示用户如需编辑可继续对话
   - 失败：返回错误信息

#### 多轮编辑场景示例

```
用户: "帮我生成一份关于AI的PPT"
Agent: 执行 node generate_slide.js --description "生成一份关于AI的PPT"
Agent: ✅ 生成完成！链接: https://docs.qq.com/slide/xxx
       session_id: session_abc123（如需编辑可继续对话）

用户: "把第2页标题改成'AI的应用场景'"
Agent: 执行 node generate_slide.js --description "把第2页标题改成'AI的应用场景'" --session_id "session_abc123"
Agent: ✅ 编辑完成！链接: https://docs.qq.com/slide/xxx

用户: "在最后增加一页总结"
Agent: 执行 node generate_slide.js --description "在最后增加一页总结" --session_id "session_abc123"
Agent: ✅ 编辑完成！链接: https://docs.qq.com/slide/xxx
```

#### ❌ 避免的做法

```bash
# ❌ 错误1：手动调用 create_slide + slide_progress 循环轮询
# agent 手动轮询存在不确定性，可能因超时或中断导致失败

# ❌ 错误2：多轮编辑时不传 session_id
# 会重新创建一份新PPT，而非编辑已有的

# ❌ 错误3：静默等待不反馈进度
# 脚本会自动输出进度信息，agent 应转达给用户
```

---

## 注意事项

- `create_slide` 为异步接口，推荐使用 `generate_slide.js` 脚本自动处理
- 脚本轮询间隔：每 20 秒一次
- 最长等待时间：20 分钟
- `reference_context` 仅在用户明确指定需要根据某段内容/材料生成 PPT 时才传
- 多轮编辑时必须传入 `session_id`，否则会创建新的 PPT
- `generate_slide.js` 脚本为跨平台 JS 脚本，需要 Node.js >= 14 运行环境

### 文件上传和图片处理指导

**文件内容解析**：当用户上传文件时，agent 应首先使用自身能力读取文件内容并解析为文本：
- 对于文本文件（.txt, .md, .docx, .pdf 等）：直接提取文本内容
- 对于表格文件（.xlsx, .csv 等）：提取数据并转换为描述性文本
- 对于代码文件：提取关键代码片段和注释

**图片内容解析**：当用户上传图片时，agent 应使用 OCR 或图片识别能力提取文字内容：
- 使用 OCR 技术识别图片中的文字
- 对于图表、流程图等，描述图片的主要内容和结构
- 将识别结果转换为清晰的文本描述

**处理流程**：
1. 读取用户上传的文件/图片内容
2. 使用 agent 自身能力解析为结构化文本
3. 将解析后的文本作为 `reference_context` 参数传递给 slide skill
4. 在 `description` 中说明基于解析内容生成 PPT 的要求

**示例**：
```bash
# 用户上传了销售数据表格，agent 解析后生成 PPT
node generate_slide.js --description "根据销售数据分析生成PPT，包含趋势分析和关键指标" --reference_context "2024年Q1销售额1200万，同比增长25%，主要增长来自华南区域..."

# 用户上传了产品架构图，agent 解析后生成 PPT  
node generate_slide.js --description "根据产品架构图生成技术介绍PPT" --reference_context "系统采用微服务架构，包含用户服务、订单服务、支付服务等模块..."
```
