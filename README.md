# PPT2MD

将 PowerPoint (.pptx) 文件解析并转换为 Markdown 格式的工具。

## 功能特性

- **文本提取**：提取所有文本框、占位符中的文本，保留加粗、斜体、下划线等格式
- **标题映射**：自动根据占位符类型映射 Markdown 标题层级（`#` ~ `######`）
- **列表识别**：识别项目符号和缩进层级，转换为 Markdown 列表
- **图片提取**：提取所有图片到 `images/` 文件夹，在 Markdown 中记录位置和尺寸信息
- **表格转换**：将 PPT 表格转换为 Markdown 表格语法
- **公式转换**：将 OMML 数学公式转换为 LaTeX 语法（`$...$` / `$$...$$`）
- **图表处理**：提取图表数据为表格，支持导出图表为图片（Windows）
- **SmartArt**：提取 SmartArt 图形中的文本内容
- **演讲者备注**：提取每张幻灯片的备注内容
- **文档属性**：提取标题、作者等元数据写入 YAML frontmatter
- **批量转换**：支持一次转换目录下所有 .pptx 文件

## 安装

```bash
pip install -e .
```

## 使用

### 基本用法

```bash
ppt2md input.pptx
```

### 指定输出目录

```bash
ppt2md input.pptx -o ./output/
```

### 完整参数

```bash
ppt2md input.pptx \
    -o ./output/ \
    --include-notes \
    --skip-empty \
    --image-dpi 150 \
    --debug
```

### 批量转换

```bash
ppt2md ./pptx_folder/ -o ./output/
```

## 命令行参数

| 参数 | 说明 |
|------|------|
| `input` | 输入 .pptx 文件或目录 |
| `-o, --output` | 输出目录（默认当前目录） |
| `--output-file` | 指定输出 Markdown 文件名 |
| `--formula-as-image` | 将公式渲染为图片而非 LaTeX |
| `--include-notes` | 包含演讲者备注 |
| `--include-hidden` | 包含隐藏幻灯片 |
| `--skip-empty` | 跳过空白幻灯片 |
| `--keep-original-format` | 保留原始图片格式（不转换 PNG） |
| `--image-dpi` | 图片 DPI（默认 96） |
| `--no-frontmatter` | 不生成 YAML frontmatter |
| `--verbose` | 显示详细处理信息 |
| `--debug` | 调试模式 |

## 输出结构

```
output/
├── input.md              # Markdown 文件
└── images/               # 提取的图片
    ├── slide_01_img_01.png
    ├── slide_01_img_02.png
    ├── slide_02_chart_01.png
    └── ...
```

## Markdown 输出格式

```markdown
---
title: "演示文稿标题"
author: "作者"
source: "input.pptx"
---

# 演示文稿标题

---

## Slide 1: 标题

**加粗文本** 和 *斜体文本*

- 列表项 1
  - 嵌套列表项

<!-- position: x=10.28cm, y=6.57cm, width=28.81cm, height=10.28cm -->
![image](images/slide_01_img_01.png)

---

## Slide 2: 数据

| 列1 | 列2 | 列3 |
|-----|-----|-----|
| A   | B   | C   |

$$E = mc^2$$

<!-- Notes -->
### Speaker Notes

备注内容...
```

## 技术栈

- Python 3.10+
- python-pptx
- Pillow
- lxml

## 项目结构

```
ppt2md/
├── __init__.py
├── main.py              # CLI 入口
├── parser/
│   ├── presentation.py  # 演示文稿解析
│   ├── slide.py         # 幻灯片解析
│   ├── text.py          # 文本提取
│   ├── formula.py       # 公式转换
│   ├── image.py         # 图片提取
│   ├── table.py         # 表格转换
│   ├── chart.py         # 图表处理
│   ├── smartart.py      # SmartArt 解析
│   ├── media.py         # 多媒体提取
│   └── notes.py         # 备注提取
├── converter/
│   ├── md_writer.py     # Markdown 写入
│   ├── format_utils.py  # 格式转换
│   └── position_utils.py # 位置单位转换
└── utils/
    ├── file_utils.py    # 文件工具
    └── emu_utils.py     # EMU 单位转换
```

## 已知限制

- SmartArt 通过 XML 降级解析，复杂布局可能不完美
- 图表数据提取依赖 COM 自动化（仅 Windows）
- Markdown 无法完美还原 PPT 的精确排版
- 不支持动画和转场效果
- 仅支持 .pptx 格式，不支持旧版 .ppt

## License

MIT
