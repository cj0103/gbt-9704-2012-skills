# 快速使用指南

## 一分钟快速上手

### 场景 1：我有一份杂乱的文档，需要快速排版

```bash
# 进入目录
cd gbt-9704-2012-skills

# 一键排版（推荐）
uv run --with python-docx python3 scripts/formatter.py 你的文档.docx 输出文档.docx
```

### 场景 2：文档有很多英文标点，需要修复

```bash
# 修复标点
uv run --with python-docx python3 scripts/punctuation.py 你的文档.docx 修复后.docx

# 再应用公文格式
uv run --with python-docx python3 scripts/formatter.py 修复后.docx 最终文档.docx
```

### 场景 3：不知道文档有什么问题

```bash
# 先诊断
uv run --with python-docx python3 scripts/analyzer.py 你的文档.docx
```

## 完整工作流程

```bash
# 步骤 1：诊断问题
uv run --with python-docx python3 scripts/analyzer.py input.docx

# 步骤 2：修复标点
uv run --with python-docx python3 scripts/punctuation.py input.docx temp.docx

# 步骤 3：应用标准格式
uv run --with python-docx python3 scripts/formatter.py temp.docx output.docx

# 步骤 4：验证结果
uv run --with python-docx python3 scripts/analyzer.py output.docx
```

## 常用命令速查

### 诊断命令
```bash
# 基本诊断
python3 scripts/analyzer.py file.docx

# 输出 JSON 格式
python3 scripts/analyzer.py file.docx --json
```

### 标点修复命令
```bash
# 智能修复（默认）
python3 scripts/punctuation.py input.docx output.docx

# 强制转中文标点
python3 scripts/punctuation.py input.docx output.docx --mode chinese

# 只修复括号和引号
python3 scripts/punctuation.py input.docx output.docx --fix brackets,quotes
```

### 格式排版命令
```bash
# 使用标准格式
python3 scripts/formatter.py input.docx output.docx

# 使用自定义配置
python3 scripts/formatter.py input.docx output.docx --preset custom
```

## 公文格式要点

### 标题层级
- 一级标题：一、二、三、
- 二级标题：（一）（二）（三）
- 三级标题：1. 2. 3.
- 四级标题：（1）（2）（3）

### 字体要求
- 主标题：方正小标宋简体，二号
- 一级标题：黑体，三号
- 二级标题：楷体_GB2312，三号
- 正文：仿宋_GB2312，三号

### 段落格式
- 正文：首行缩进 2 字符，行距 28 磅
- 主送机关：顶格
- 落款：右对齐

## 常见问题快速解决

### Q: 没有安装 python-docx 怎么办？
A: 使用 `uv run --with python-docx` 会自动安装，无需手动安装。

### Q: 输出文档字体显示异常？
A: 需要安装中文字体：方正小标宋简体、仿宋_GB2312、楷体_GB2312、黑体。

### Q: 如何自定义格式？
A: 编辑 `presets/custom.json` 文件，调整字体、字号、行距等参数。

### Q: 表格处理不理想？
A: 表格会自动优化，如需调整可修改 `custom.json` 中的 `table` 配置。

## 获取帮助

```bash
# 查看命令帮助
python3 scripts/analyzer.py --help
python3 scripts/punctuation.py --help
python3 scripts/formatter.py --help
```

## 示例演示

运行完整示例：
```bash
python3 test_example.py
```

这会创建一个测试文档，并演示完整的诊断→修复→排版→验证流程。
