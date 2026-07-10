# PPTDecomp

> PowerPoint Decompiler — 语义级 PPT 反编译与无损重建系统。

输入一个 `.pptx` 文件，输出结构化 Markdown（含完整语义信息）。再拿这个 Markdown，可以**无损复原**出与原 PPT 一模一样的 `.pptx`。还可以将 PPT 重建为自包含的 HTML 页面。

**这不是格式转换器，是反编译器。** MD 只是中间表示，核心是语义级提取与无损重建。

---

## 核心能力

| 阶段 | 能力 | 说明 |
|------|------|------|
| **提取** | PPT → MD | 将 PPT 二进制结构反编译为人类可读的结构化 Markdown，包含文本、字体、颜色、位置、大小、公式、图片、表格、形状、渐变、组合等全部信息 |
| **重建** | MD → PPTX | 从 MD 中的语义信息无损重建原始 PPTX，支持公式、颜色、渐变、连接线、组合形状、旋转等 |
| **导出** | PPT → HTML | 将 PPT 重建为自包含的 HTML 页面，图片 base64 内嵌，公式 KaTeX 渲染，支持绝对定位布局 |

## 功能特性

- **文本提取**：提取所有文本框、占位符中的文本，保留加粗、斜体、下划线、删除线、上标/下标等格式
- **标题映射**：自动根据占位符类型映射 Markdown 标题层级（`#` ~ `######`）
- **列表识别**：识别项目符号和缩进层级，转换为 Markdown 列表
- **图片提取**：提取所有图片，在 Markdown 中记录位置和尺寸信息
- **表格转换**：将 PPT 表格转换为 Markdown 表格语法，支持合并单元格
- **公式转换**：将 OMML 数学公式转换为 LaTeX 语法（`$...$` / `$$...$$`），支持上下标、分数、积分、矩阵、修饰符等
- **图表处理**：提取图表数据为表格，支持导出图表为图片（Windows）
- **SmartArt**：提取 SmartArt 图形中的文本内容
- **形状还原**：矩形、椭圆、圆角矩形、箭头、连接线、流程图符号等，含填充色、渐变、边框
- **组合形状**：嵌套组的递归提取与重建
- **颜色系统**：主题色（accent1-6）解析为绝对 sRGB，渐变、透明度完整保留
- **演讲者备注**：提取每张幻灯片的备注内容
- **文档属性**：提取标题、作者等元数据写入 YAML frontmatter
- **批处理**：支持一次转换目录下所有 .pptx 文件
- **HTML 导出**：将 PPT 渲染为自包含 HTML，支持 KaTeX 公式、图片 base64 内嵌、精确还原布局

## 安装

```bash
git clone https://github.com/caiyilian/PPTDecomp.git
cd PPTDecomp
pip install -e .
```

### 可选依赖

```bash
# HTML → PNG 截图（用于闭环对比测试）
pip install playwright
python -m playwright install chromium

# PPT → EMF → PNG 导出（用于 roundtrip 像素级对比）
pip install pywin32 numpy
```

## 使用

### PPT → MD（反编译）

```bash
ppt2md input.pptx
ppt2md input.pptx -o ./output/
ppt2md input.pptx -o ./output/ --include-notes
```

### MD → PPTX（重建）

```bash
python -c "from ppt2md.converter.reverse import convert_md_to_pptx; convert_md_to_pptx('output.md', 'rebuilt.pptx')"
```

### PPT → HTML（导出）

```bash
python ppt2html.py input.pptx -o output.html
python ppt2html.py input.pptx -o output.html --screenshot  # 同时生成截图
```

### Roundtrip 闭环对比

```bash
python compare_roundtrip.py input.pptx -o ./cmp
# 生成左右对比图（原始 vs 重建），逐页验证
```

## 命令行参数

### `ppt2md`

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

### `ppt2html`

| 参数 | 说明 |
|------|------|
| `input` | 输入 .pptx 文件 |
| `-o, --output` | 输出 HTML 文件路径 |
| `--screenshot` | 同时用 Playwright 生成截图 PNG |

## 输出结构

```
output/
├── input.md              # Markdown 文件（含完整语义元数据）
├── input.html            # HTML 重建（可选）
├── input.png             # HTML 截图（可选）
└── images/               # 提取的图片
    ├── slide_01_img_01.png
    └── ...
```

## 技术栈

- **Python** 3.10+
- **python-pptx** — PPTX 读写
- **Pillow** — 图片处理
- **lxml** — XML 解析
- **pywin32** — Windows COM（PPT → EMF 导出，可选）
- **playwright** — HTML 截图（可选）

## 项目结构

```
PPTDecomp/
├── ppt2md/                    # 核心模块
│   ├── main.py                # CLI 入口
│   ├── parser/                # PPT 解析（反编译方向）
│   │   ├── presentation.py    # 演示文稿解析
│   │   ├── slide.py           # 幻灯片解析
│   │   ├── text.py            # 文本提取
│   │   ├── formula.py         # OMML 公式 → LaTeX
│   │   ├── image.py           # 图片提取
│   │   ├── table.py           # 表格转换
│   │   ├── chart.py           # 图表处理
│   │   ├── smartart.py        # SmartArt 解析
│   │   ├── media.py           # 多媒体提取
│   │   └── notes.py           # 备注提取
│   └── converter/             # MD → PPTX（重建方向）
│       ├── reverse.py         # MD→PPTX 重建入口
│       ├── build.py           # 构建 PPTX 形状
│       ├── metadata.py        # 元数据提取（fill、line、body 等）
│       ├── html.py            # PPT→HTML 导出
│       ├── format_utils.py    # 格式工具
│       ├── md_writer.py       # Markdown 写入
│       └── position_utils.py  # 位置单位转换
├── ppt2html.py                # HTML 导出 CLI
├── compare_roundtrip.py       # roundtrip 对比工具
├── test_all_features.md       # 全功能测试 MD
├── ori.pptx                   # 测试用 PPT
└── method12.pptx              # 测试用 PPT
```

## 为什么不是"格式转换器"

市面上普通的 PPT 转 MD 工具是**格式转换器**：读 PPT → 吐出 MD，信息有损，不可逆。

PPTDecomp 是**反编译器（Decompiler）**：

- 像反编译器把机器码还原成源代码一样，它把 PPT 的二进制结构还原成人类可读的语义信息
- 这种语义信息是完备的：Markdown 只是选中的"中间表示格式"，换成 JSON/YAML 也一样成立
- 核心不是"转成 MD"，而是**提取语义 + 无损重建**

## 已知限制

- SmartArt 通过 XML 降级解析，复杂布局可能不完美
- 图表数据提取依赖 COM 自动化（仅 Windows）
- 不支持动画和转场效果
- 仅支持 .pptx 格式，不支持旧版 .ppt

## License

MIT