# 国家行政机关公文格式处理工具（GB/T 9704-2012）

## 简介

本工具是专门用于处理党政机关公文格式的 Python 脚本集，严格按照《GB/T 9704-2012 党政机关公文格式》国家标准设计。

### 主要功能

- ✅ **格式诊断**：自动检测文档中的格式问题
- ✅ **标点修复**：智能修复中英文标点混用
- ✅ **公文排版**：一键应用标准公文格式
- ✅ **表格优化**：自动调整表格布局和对齐
- ✅ **页码生成**：自动生成标准页码（一字线，奇右偶左）
- ✅ **版记处理**：自动处理抄送、印发单位等版记格式

## 环境要求

- Python 3.7+
- python-docx

使用 `uv run --with python-docx` 可自动安装依赖。

## 快速开始

### 1. 格式诊断

分析文档存在的问题：

```bash
cd gbt-9704-2012-skills
uv run --with python-docx python3 scripts/analyzer.py your_document.docx
```

### 2. 修复标点

修复中英文标点混用问题：

```bash
uv run --with python-docx python3 scripts/punctuation.py input.docx output.docx
```

### 3. 应用公文格式

一键排版为标准的公文格式：

```bash
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx
```

## 详细使用说明

### 格式诊断（analyzer.py）

诊断报告包含以下内容：

- **标点问题**：英文括号、引号、冒号、句号等
- **序号问题**：序号格式不统一、层级跳跃
- **段落格式**：缺少首行缩进、行距不统一
- **字体问题**：字体种类过多、字号不统一
- **公文结构**：缺少标题、落款、日期等

**示例：**

```bash
uv run --with python-docx python3 scripts/analyzer.py test.docx
```

**输出：**

```
==================================================
              公文格式诊断报告
==================================================

【标点问题】共 8 处
  - 英文括号：第 2、5、8 段
  - 英文引号：第 3、7 段

【序号问题】共 2 处
  - 序号格式不统一：同时存在 "1." 和 "1、"

【段落格式问题】共 5 处
  - 缺少首行缩进：第 3、5、7、9、11 段

【字体问题】共 2 处
  - 字体种类过多：检测到 6 种字体

--------------------------------------------------
共发现 17 处格式问题

建议：
  - 运行 punctuation.py 修复标点问题
  - 运行 formatter.py 统一公文格式
```

### 标点修复（punctuation.py）

**基本用法：**

```bash
# 智能模式（推荐）
uv run --with python-docx python3 scripts/punctuation.py input.docx output.docx

# 强制全部转为中文标点
uv run --with python-docx python3 scripts/punctuation.py input.docx output.docx --mode chinese

# 只修复特定类型
uv run --with python-docx python3 scripts/punctuation.py input.docx output.docx --fix brackets,quotes
```

**支持的修复类型：**

- `brackets` - 括号
- `quotes` - 引号
- `colon` - 冒号
- `comma` - 逗号
- `period` - 句号
- `semicolon` - 分号
- `question` - 问号
- `exclamation` - 叹号
- `ellipsis` - 省略号
- `dash` - 破折号

### 公文排版（formatter.py）

**基本用法：**

```bash
# 使用标准公文格式
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx

# 使用自定义配置
uv run --with python-docx python3 scripts/formatter.py input.docx output.docx --preset custom
```

**自动识别的公文要素：**

1. **标题**：关于 XXX 的通知/报告/请示/函
2. **主送机关**：XXX：
3. **正文**：各层级标题和段落
4. **附件说明**：附件：XXX
5. **发文机关署名**：单位名称
6. **成文日期**：2024 年 1 月 15 日
7. **版记**：抄送、印发单位

## 公文格式标准

### 页面设置

```
纸张：A4（210mm × 297mm）
页边距：上 37mm，下 35mm，左 28mm，右 26mm
版心：156mm × 225mm（不含页码）
```

### 字体字号

| 要素 | 字体 | 字号 | 说明 |
|------|------|------|------|
| 公文标题 | 方正小标宋简体 | 二号（22pt） | 居中 |
| 主送机关 | 仿宋_GB2312 | 三号（16pt） | 顶格 |
| 一级标题 | 黑体 | 三号（16pt） | "一、" |
| 二级标题 | 楷体_GB2312 | 三号（16pt） | "（一）" |
| 三级标题 | 仿宋_GB2312 | 三号（16pt） | "1." |
| 四级标题 | 仿宋_GB2312 | 三号（16pt） | "（1）" |
| 正文 | 仿宋_GB2312 | 三号（16pt） | 首行缩进 2 字符 |
| 落款 | 仿宋_GB2312 | 三号（16pt） | 右对齐 |
| 页码 | 宋体 | 四号（14pt） | 一字线，奇右偶左 |

### 段落格式

```
正文：首行缩进 2 字符，行距固定值 28 磅
标题：首行缩进 2 字符，行距固定值 28 磅
主送机关：顶格，行距固定值 28 磅
落款：右对齐，行距固定值 28 磅
```

### 页码格式

```
位置：距版心下边缘约 7mm
格式：— 1 —（一字线 + 空格 + 页码 + 空格 + 一字线）
对齐：奇数页居右，偶数页居左
字体：四号宋体
```

## 推荐工作流程

```bash
# 步骤 1：诊断
uv run --with python-docx python3 scripts/analyzer.py messy.docx

# 步骤 2：修复标点
uv run --with python-docx python3 scripts/punctuation.py messy.docx temp.docx

# 步骤 3：应用公文格式
uv run --with python-docx python3 scripts/formatter.py temp.docx clean.docx

# 步骤 4：再次诊断确认
uv run --with python-docx python3 scripts/analyzer.py clean.docx
```

## 自定义配置

如果需要调整公文格式参数，可以编辑 `presets/custom.json` 文件。

**配置项说明：**

- `page` - 页面边距设置
- `title` - 主标题格式
- `recipient` - 主送机关格式
- `heading1-4` - 一至四级标题格式
- `body` - 正文格式
- `signature` - 落款格式
- `date` - 日期格式
- `attachment` - 附件格式
- `table` - 表格格式
- `page_number` - 页码设置

每个配置项包含：

- `font_cn` - 中文字体
- `font_en` - 英文字体
- `size` - 字号（pt）
- `bold` - 是否加粗
- `align` - 对齐方式（left/center/right/justify）
- `indent` - 首行缩进（pt）
- `line_spacing` - 行距（磅）

## 常见问题

### Q: 为什么输出的文档字体显示不正确？

A: 需要系统安装对应的中文字体：
- 方正小标宋简体
- 仿宋_GB2312
- 楷体_GB2312
- 黑体

Windows 系统一般自带这些字体，Mac/Linux 系统需要手动安装。

### Q: 可以处理带表格的公文吗？

A: 可以。会自动识别并优化表格格式，包括：
- 边框（0.5 磅细实线）
- 表头（居中，可加粗）
- 内容对齐（数字右对齐，短文本居中）
- 列宽（根据内容自动调整）

### Q: 成文日期能自动转换格式吗？

A: 可以识别多种日期格式，但不会自动转换格式。建议在输入文档中就使用标准的阿拉伯数字日期格式："2024 年 1 月 15 日"。

### Q: 如何处理版记（抄送、印发）？

A: 会自动识别"抄送："、"印发："等关键字，按标准格式排版（左右各空 1 字，三号仿宋）。

### Q: 支持旧版 .doc 格式吗？

A: 不支持。需要先用 Word 将 .doc 文件另存为 .docx 格式。

## 文件结构

```
gbt-9704-2012-skills/
├── SKILL.md                 # 使用说明
├── README.md                # 详细介绍
├── scripts/
│   ├── analyzer.py          # 格式诊断
│   ├── punctuation.py       # 标点修复
│   └── formatter.py         # 公文排版
└── presets/
    └── custom.json          # 自定义配置
```

## 版本信息

- 版本：1.0.0
- 标准依据：GB/T 9704-2012 党政机关公文格式
- 创建日期：2026-04-08

## 项目来源与致谢

本工具是在 [document-format-skills](https://github.com/KaguraNanaga/document-format-skills) 项目基础上，针对国家行政机关公文格式（GB/T 9704-2012）进行的专项开发和优化。

### 主要改进

- ✅ **公文标准化**：严格按照 GB/T 9704-2012 标准重新设计格式预设
- ✅ **公文要素识别**：增强了对公文特有要素的自动识别（标题、主送、落款等）
- ✅ **页码规范化**：实现标准公文页码格式（一字线，奇右偶左）
- ✅ **格式诊断增强**：增加公文结构完整性检查
- ✅ **表格优化**：针对公文表格特点优化布局算法

感谢原项目作者的优秀工作，本工具继承了原项目的核心功能，并针对党政机关公文场景进行了深度定制。

## 许可证

本工具遵循MIT开源协议。
