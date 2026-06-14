# Excel 公式生成器

> 不用背函数，描述需求就能拿到能直接粘贴的公式。支持 Excel / WPS / Google Sheets 全平台，基础功能完全免费，AI 增强按需使用。

[![HTML](https://img.shields.io/badge/HTML-单文件-blue.svg)]()
[![Platform](https://img.shields.io/badge/Platform-Excel%20%7C%20WPS%20%7C%20Google%20Sheets-green.svg)]()
[![AI 支持](https://img.shields.io/badge/AI-BYOK-orange.svg)]()

---

## 🚀 快速使用

### 方式一：直接下载使用

下载以下文件放在同一目录，用浏览器打开 `index.html` 即可：

```
index.html
```

### 方式二：在线体验

- 主要地址：https://jotarou.com/tools/excel/
- 备用地址：https://jotaroustar.github.io/excel-formula-generator

---

## ✨ 功能特性

### 🟢 免费基础模式（无需 API Key）

**场景化速查（5 大类 · 交互式填空）**

放弃传统"按函数名分类"，改为按业务场景分类，降低非技术用户的认知门槛：

| 场景 | 涵盖需求 |
|------|---------|
| 🔎 数据查找匹配 | XLOOKUP / VLOOKUP / INDEX+MATCH 按条件找对应值 |
| 🧮 多条件统计求和 | COUNTIFS / SUMIFS / COUNTIF / SUMIF 多维度汇总 |
| ✂️ 文本拆分与提取 | LEFT / RIGHT / MID / TRIM / TEXTSPLIT / TEXTJOIN |
| 📅 日期与工期计算 | DATEDIF / NETWORKDAYS / TODAY 天数、月数、工作日 |
| 🏆 去重与排名 | RANK / UNIQUE / LARGE / COUNTIF 去重、排序、标记 |

**工作流：** 选场景 → 填写表单 → 一键生成可直接粘贴的公式

**版本自动降级（防呆设计）**

切换目标软件版本后，生成逻辑自动调整：

| 场景 | Excel 365 / WPS 最新版 | Excel 2019 及以下 |
|------|----------------------|-----------------|
| 数据查找 | `=XLOOKUP(...)` | 自动改为 `=INDEX(MATCH(...))` |
| 文本拆分 | `=TEXTSPLIT(...)` | 自动改为 `=SUBSTITUTE(REPT(...))` 兼容写法 |

**公式解释器**

粘贴看不懂的祖传公式，工具自动：
- 提取并列出所有嵌套函数名
- 按大白话逐层解释各函数作用
- 说明多层嵌套的计算顺序

内置 40+ 函数词典，覆盖 VLOOKUP、XLOOKUP、SUMIFS、DATEDIF、IFERROR、TEXTJOIN 等高频函数。

**报错诊断手册**

点击错误代码，查看图文并茂的成因分析与修复方案：

| 错误代码 | 常见原因 |
|---------|---------|
| `#N/A` | 查找值不存在 / 格式不一致 / 区域未锁定 |
| `#REF!` | 引用列被删除 / 剪切粘贴错位 / 跨表引用失效 |
| `#VALUE!` | 文本参与数值运算 / 日期格式异常 / 数组大小不匹配 |
| `#NAME?` | 函数名拼错 / 文本漏引号 / 低版本用了新函数 |
| `#DIV/0!` | 除数为空 / 除数计算结果为零 |

---

### 🧠 AI 智能助手模式（BYOK · 需 API Key）

**描述需求 → 直接生成公式**

```
用户输入：统计 A 列中销售额大于 100 且 B 列为「已完成」的行数
AI 输出：=COUNTIFS(A:A,">100",B:B,"已完成")
```

前端在发送请求前自动封装环境上下文：

```
角色：Excel/WPS 公式专家
目标软件：[当前选择的平台和版本]
用户表头样本：[可选粘贴]
要求：只输出纯公式，不要 Markdown 代码块标记
```

支持 **`[+ 粘贴我的表头样本]`** 功能，让 AI 理解你的真实列名，生成更精准的公式。

**公式解释器（AI 深度版）**

粘贴复杂公式，AI 逐层拆解参数含义，比内置词典更准确处理自定义写法。

**报错修复（结构化诊断）**

填写结构化信息后由 AI 给出精准修复方案：

```
报错公式：=VLOOKUP(A2,Sheet2!A:C,4,0)
报错类型：#REF!
数据源现在有几列：3 列（删了一列）
补充说明：昨天删过 D 列
```

信息不足时 AI 会明确追问缺失信息，而非强行给出错误修复。

---

### 🎨 UX 细节

- **新手模式**：生成公式后附带参数中文对照卡片（纯净公式与说明物理隔离，直接复制不受影响）
- **深色 / 浅色模式**：跟随系统或手动切换，保护长期面对表格的打工人的眼睛
- **历史记录**：本地缓存最近 10 条生成的公式，刷新页面不丢失
- **Ctrl + Enter 快捷生成**：AI 模式下无需点按钮
- **一键复制**：点击代码块任意位置或右上角复制按钮，自动写入剪贴板
- **示例预览表格**：每个场景附带静态示例数据，展示公式预期效果
- **平滑滚动**：生成结果后自动滚动到结果区域

---

## 🛠 多平台支持

| 软件 | 版本 | 支持函数集 |
|------|------|---------|
| Microsoft Excel | 365 / 最新版 | 全部，含 XLOOKUP / FILTER / TEXTSPLIT |
| Microsoft Excel | 2019 及以下 | 自动降级为兼容函数，不生成新版专属函数 |
| WPS 表格 | 最新版 | 同 Excel 365，支持新函数 |
| Google Sheets | — | 支持 SPLIT / QUERY 等 Sheets 专属函数 |

---

## 🔌 多种使用模式

**【免费模式】** 无需任何 Key，打开即用，覆盖场景速查、公式解释、报错手册全部功能。

**【AI 模式（BYOK）】** 填入自己的 API Key，调用大模型实现自然语言转公式、报错修复。支持自定义 Base URL，可接入任意 OpenAI 兼容接口（包括 [jotarou.com API](https://jotarou.com)）。

---

## 🤖 支持的模型

| 模型 | 说明 |
|------|------|
| DeepSeek Chat | 速度快、成本低，日常推荐 |
| GPT-4o mini | OpenAI 轻量模型 |
| Claude 3.5 Haiku | 逻辑严谨，公式准确率高 |
| Gemini 2.0 Flash | Google 多语言能力强 |

支持自定义 Base URL，可接入任意 OpenAI 兼容接口。

---

## 🔒 隐私说明

- **公式库模式**：纯前端运行，所有计算在浏览器本地完成，零数据上传
- **AI 模式**：输入内容经 jotarou.com API 网关转发至大模型服务商，网关不持久化对话内容
- **API Key 存储**：Key 仅存储在你的浏览器本地（localStorage），不经过本工具服务器存储

---

## 🧱 技术栈

- 纯 HTML + CSS + JavaScript，单文件，无需构建工具
- 零外部依赖，无框架，无 CDN 引用，离线可用（AI 模式需网络）
- CSS 变量驱动深色 / 浅色主题，`color-mix()` 实现毛玻璃导航栏
- localStorage 实现历史记录与配置持久化
- 调用 OpenAI 兼容接口，前端封装 System Prompt 约束 AI 输出格式

---

## 📁 文件结构

```
index.html   # 全部代码（HTML + CSS + JS 单文件，53KB）
```

---

## 📋 使用场景覆盖

适用人群：

- **职场打工人**：日常表格处理、数据汇总、工期计算
- **电商运营**：销售数据统计、多条件筛选、排名
- **HR / 行政**：考勤天数计算、名单去重、信息提取
- **数据分析初学者**：学习函数逻辑，理解嵌套公式结构
- **接手祖传表格的人**：快速看懂前任留下的复杂公式

---

## 🔗 相关项目

| 项目 | 说明 |
|------|------|
| [快递单号提取器](https://jotarou.com/tools/express/) | 批量提取快递单号，支持截图 AI 识别 |
| [简历优化器](https://jotarou.com/tools/resume/) | AI 逐条分析建议 / JD 匹配度分析 |
| [日报周报生成器](https://jotarou.com/tools/report/) | 输入工作内容，一键生成专业报告 |
| [文案生成器](https://jotarou.com/tools/copywriter/) | 朋友圈/社交媒体文案，覆盖多平台多风格 |
| [AI 面试模拟器](https://jotarou.com/tools/interview/) | 全岗位面试练习，AI 实时点评打分 |
| [英文邮件生成器](https://jotarou.com/tools/email/) | 12 场景职场英文邮件，告别 Chinglish |
| [jotarou.com API](https://jotarou.com) | AI API 中转服务，支持 Claude / GPT / DeepSeek |

---
---

<div align="center">
  <p>由 <a href="https://jotarou.com">jotarou.com</a> 提供支持 · 开源免费</p>
  <p>@jotaroustar 用心制作 ❤️</p>
</div>
© 2026 jotarou.com · 代码版权所有，禁止直接复制用于商业产品
