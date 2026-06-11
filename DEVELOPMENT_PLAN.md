# PPTX 转 Markdown 分阶段开发方案

## 项目概述

将 PowerPoint (.pptx) 文件解析并转换为 Markdown 格式，提取所有可识别的内容（文本、公式、图片、表格、图表等），图片提取到独立文件夹，Markdown 中记录图片路径及位置/尺寸信息。

## 技术选型

| 组件 | 技术方案 |
|------|---------|
| 解析库 | `python-pptx` |
| 公式处理 | OMML XML → LaTeX |
| 图片处理 | `Pillow` (PIL) |
| 图表导出 | `win32com.client` (Windows COM) |
| 语言 | Python 3.10+ |

## 分阶段实施计划

---

### Phase 1：项目脚手架

**目标**：搭建项目骨架，可运行空命令。

**工作内容**：
- 创建目录结构：`ppt2md/`（源码）、`tests/`（测试）
- 编写 `pyproject.toml`（项目元数据、依赖声明、CLI 入口点）
- 创建 `ppt2md/__init__.py`、`ppt2md/main.py`（CLI 骨架，使用 `argparse`）
- 创建 `tests/__init__.py`、`tests/conftest.py`（pytest fixtures）
- 初始化 `.gitignore`、编写 `README.md`
- 创建 `scratch/` 目录（测试输出，已 gitignore）

**验证**：`pip install -e .` 成功，`ppt2md --help` 输出帮助信息。

---

### Phase 2：基础 PPTX 解析

**目标**：能打开 .pptx 文件并列出所有幻灯片编号。

**工作内容**：
- 实现 `ppt2md/parser/presentation.py`：打开 Presentation 对象
- 实现 `ppt2md/parser/slide.py`：遍历 slides，返回幻灯片列表
- 在 CLI 中接入：`ppt2md input.pptx` 打印幻灯片数量
- 编写测试：用简单 pptx 验证能正确读取幻灯片数

**验证**：对 `method.pptx` 输出 "Found N slides"。

---

### Phase 3：文本帧基础提取

**目标**：提取幻灯片中所有文本框的文本内容。

**工作内容**：
- 实现 `ppt2md/parser/text.py`：遍历 `slide.shapes`，检测 `has_text_frame`
- 提取 `paragraph.text` 和 `run.text`
- 处理空文本帧和空段落的跳过逻辑
- 初步输出：按幻灯片分组打印所有文本
- 编写测试：验证纯文本提取的正确性

**验证**：输出包含幻灯片中所有可见文本。

---

### Phase 4：文本格式转 Markdown 标记

**目标**：将文本格式（加粗、斜体、下划线等）转换为 Markdown 内联标记。

**工作内容**：
- 实现 `ppt2md/converter/format_utils.py`：检测 `run.font.bold`、`run.font.italic`、`run.font.underline`、`run.font.strike`
- 映射规则：`bold → **text**`、`italic → *text*`、`underline → <u>text</u>`、`strikethrough → ~~text~~`
- 处理同一段落内多 run 的混合格式拼接
- 编写测试：验证各种格式组合的 Markdown 输出

**验证**：加粗文本输出 `**加粗**`，混合格式正确拼接。

---

### Phase 5：占位符类型识别与标题层级映射

**目标**：根据占位符类型自动映射 Markdown 标题层级。

**工作内容**：
- 检测 `shape.placeholder_format` 的 `idx` 值，识别占位符类型：
  - `0` = 标题 → `#`
  - `1` = 副标题 / 中心标题
  - `2` = 日期
  - `3` = 幻灯片编号
  - `4` = 页脚
  - `5` = 页眉
  - `6` = 正文
  - `7` = 对象
  - `8` = 议程
  - `9` = 图表标题
  - `10` = 图片标题
  - `11` = 副标题
- 映射：标题占位符 → `## Slide N: {text}`，正文占位符 → 普通段落
- 跳过日期、页脚、页眉、编号等元信息占位符
- 编写测试：验证标题/正文/副标题的正确映射

**验证**：标题占位符输出为 `##`，正文为普通段落，页脚等被跳过。

---

### Phase 6：列表和项目符号检测

**目标**：识别 PPT 中的列表项并转换为 Markdown 列表。

**工作内容**：
- 检测 `paragraph.level`（缩进层级）和 `paragraph.alignment`
- 检测段落文本中的项目符号字符（•、▸、● 等）
- 映射规则：
  - level 0 → `- item`
  - level 1 → `  - item`（2空格缩进）
  - level 2 → `    - item`（4空格缩进）
- 检测有序列表（数字+点）并映射为 `1. item`
- 编写测试：验证多层嵌套列表的输出

**验证**：一级列表输出 `- `，二级列表输出 `  - `，嵌套正确。

---

### Phase 7：图片提取与保存

**目标**：从幻灯片中提取所有图片并保存到 `images/` 文件夹。

**工作内容**：
- 实现 `ppt2md/parser/image.py`：遍历 shapes，检测 `shape.shape_type == MSO_SHAPE_TYPE.PICTURE` 或 `shape.image`
- 提取 `shape.image.blob` 和 `shape.image.content_type`
- 根据 content_type 确定文件扩展名（png/jpg/gif/bmp/tiff/emf/wmf）
- 命名规则：`slide_{NN}_img_{MM}.{ext}`
- 创建 `output/images/` 目录并保存文件
- 编写测试：验证图片文件被正确提取和保存

**验证**：`images/` 文件夹中出现正确的图片文件。

---

### Phase 8：图片位置与尺寸信息记录

**目标**：在 Markdown 中以注释形式记录每张图片的位置和尺寸。

**工作内容**：
- 读取 `shape.left`、`shape.top`、`shape.width`、`shape.height`（EMU 单位）
- 实现 `ppt2md/converter/position_utils.py`：EMU → cm 和 px 转换
- 在 MD 中输出注释：
  ```
  <!-- position: x=10.28cm, y=6.57cm, width=28.81cm, height=10.28cm -->
  ![image](images/slide_01_img_01.png)
  ```
- 处理 `None` 值（某些形状可能没有明确的位置信息）
- 编写测试：验证位置信息的正确转换和格式

**验证**：MD 中图片前出现位置注释，单位正确。

---

### Phase 9：图片裁剪处理

**目标**：处理图片的裁剪偏移信息。

**工作内容**：
- 检测图片形状的裁剪属性：`crop_left`、`crop_right`、`crop_top`、`crop_bottom`
- 裁剪值为比例（0.0-1.0），记录原始图片尺寸和裁剪区域
- 在 MD 注释中追加裁剪信息：
  ```
  <!-- position: ..., crop: left=0.1, right=0.05, top=0.0, bottom=0.1 -->
  ```
- 可选：使用 Pillow 实际裁剪图片
- 编写测试：验证裁剪信息的正确记录

**验证**：有裁剪的图片正确记录裁剪比例。

---

### Phase 10：图片去重

**目标**：避免同一张图片被重复保存。

**工作内容**：
- 建立 `rId → saved_filename` 映射表（`image_map`）
- 每个 shape 的图片通过 `slide.part.related_parts` 的 rId 索引
- 已存在的 rId 复用已有文件名，不重复保存
- 在 MD 中引用同一个文件
- 编写测试：验证重复引用的图片只保存一次

**验证**：同一图片多处引用时，`images/` 中只有一个文件，MD 中多处引用同一路径。

---

### Phase 11：表格提取为 Markdown 表格

**目标**：将 PPT 表格转换为 Markdown 表格语法。

**工作内容**：
- 实现 `ppt2md/parser/table.py`：检测 `shape.has_table`
- 遍历 `table.rows` 和 `table.columns`，提取 `cell.text`
- 生成 Markdown 表格：
  ```
  | Header 1 | Header 2 |
  |----------|----------|
  | cell 1   | cell 2   |
  ```
- 处理单元格内多段落（用 `<br>` 或空格连接）
- 检测合并单元格（`cell.is_merge_target`），在注释中标记
- 编写测试：验证各种表格（含合并）的输出

**验证**：表格正确转为 Markdown 表格语法。

---

### Phase 12：演讲者备注提取

**目标**：提取每张幻灯片的演讲者备注。

**工作内容**：
- 检测 `slide.has_notes_slide`
- 提取 `slide.notes_slide.notes_text_frame.text`
- 在幻灯片内容末尾添加备注区块：
  ```
  <!-- Notes -->
  ### Speaker Notes

  备注内容...
  ```
- 处理空备注和格式化备注（段落、列表）
- 编写测试：验证备注的正确提取和格式

**验证**：有备注的幻灯片输出备注区块。

---

### Phase 13：形状位置排序

**目标**：按视觉阅读顺序排列幻灯片中的形状。

**工作内容**：
- 实现排序函数：按 `shape.top`（行）升序，相同行内按 `shape.left`（列）升序
- 处理 `None` 值（将 None 视为最大值排到最后）
- 定义分组策略：top 差值在阈值内视为同一行
- 排序后按顺序输出内容到 MD
- 编写测试：验证排序后的输出顺序

**验证**：内容按从上到下、从左到右的顺序输出。

---

### Phase 14：OMML 公式检测与基础 LaTeX 转换

**目标**：检测幻灯片中的数学公式并转换为基础 LaTeX。

**工作内容**：
- 实现 `ppt2md/parser/formula.py`：在 shape XML 中查找 `<m:oMath>` 和 `<m:oMathPara>` 节点
- 解析 OMML 命名空间 `http://schemas.openxmlformats.org/officeDocument/2006/math`
- 实现基础元素转换：
  - `m:r`（文本 run）→ 直接文本
  - `m:f`（分数）→ `\frac{a}{b}`
  - `m:rad`（根号）→ `\sqrt{x}`
  - `m:sSub`（下标）→ `x_{n}`
  - `m:sSup`（上标）→ `x^{n}`
  - `m:d`（括号）→ `\left( ... \right)`
- 在 MD 中输出 `$$...$$`（行间公式）
- 编写测试：验证基础公式的 LaTeX 输出

**验证**：简单公式正确转为 LaTeX。

---

### Phase 15：OMML 公式全面转换

**目标**：支持所有 OMML 数学元素的 LaTeX 转换。

**工作内容**：
- 扩展 `formula.py`，支持以下元素：
  - `m:nary`（求和/积分/乘积）→ `\sum_{i=0}^{n}`
  - `m:func`（函数）→ `\sin`, `\cos`, `\log`
  - `m:acc`（重音符号）→ `\hat{x}`, `\bar{x}`, `\vec{x}`
  - `m:bar`（横线）→ `\overline{x}`, `\underline{x}`
  - `m:limLow`/`m:limUpp`（极限）→ 下标/上标位置
  - `m:m`（矩阵）→ `\begin{matrix}...\end{matrix}`
  - `m:eqArr`（方程组）→ 多行公式对齐
  - `m:sPre`（前置上下标）
  - `m:groupChr`（群组字符）→ `\underbrace{...}`
  - `m:borderBox`/`m:box`
- 处理嵌套公式（公式内包含公式）
- 编写测试：验证各元素的转换正确性

**验证**：复杂嵌套公式正确转为 LaTeX。

---

### Phase 16：图表数据提取

**目标**：提取 PPT 图表的类型和数据。

**工作内容**：
- 实现 `ppt2md/parser/chart.py`：检测 `shape.has_chart`
- 读取 `chart.chart_type`（柱状图、饼图、折线图、散点图等）
- 读取 `chart.series` 获取数据系列
- 尝试读取 `chart.chart_data` 中的分类和值
- 输出格式：
  ```
  > **Chart**: 柱状图，3 个系列，5 个分类

  | Category | Series 1 | Series 2 | Series 3 |
  |----------|----------|----------|----------|
  | A        | 10       | 20       | 30       |
  ```
- 编写测试：验证图表数据提取

**验证**：图表类型和数据正确提取为表格。

---

### Phase 17：图表图片导出（COM 自动化）

**目标**：通过 Windows COM 自动化将图表导出为图片。

**工作内容**：
- 实现 `ppt2md/parser/chart.py` 的图片导出功能
- 使用 `win32com.client` 连接 PowerPoint 应用
- 通过 `shape.Export()` 方法将图表形状导出为 PNG
- 仅在 Windows 环境且 COM 可用时启用
- 回退策略：COM 不可用时跳过图片导出，仅保留数据表格
- 编写测试：在 Windows 环境下验证图表图片导出

**验证**：图表导出为 PNG 图片并在 MD 中引用。

---

### Phase 18：SmartArt 与组合形状递归解析

**目标**：提取 SmartArt 中的文本和组合形状中的所有子形状。

**工作内容**：
- 实现递归解析：检测 `shape.shape_type == MSO_SHAPE_TYPE.GROUP`
- 遍历 `shape.shapes` 子形状，递归调用解析器
- SmartArt 处理：SmartArt 在 .pptx 中存储为形状组，通过 XML 解析
  - 查找 `<dgm:relPoints>` 或 `<p:sp>` 节点
  - 提取所有文本节点，按层级组织为嵌套列表
- 处理自选图形（矩形、圆角矩形等）中的文本
- 编写测试：验证组合形状和 SmartArt 的文本提取

**验证**：组合形状中的所有文本被正确提取。

---

### Phase 19：多媒体与 OLE 对象处理

**目标**：提取音频、视频和嵌入的 OLE 对象。

**工作内容**：
- 检测 `shape.shape_type == MSO_SHAPE_TYPE.MEDIA`
- 提取音频/视频文件的 blob 和 content_type
- 保存到 `output/media/` 文件夹，命名规则：`slide_{NN}_media_{MM}.{ext}`
- OLE 对象处理：检测嵌入的 Excel/Word 对象
  - 尝试提取 OLE 流中的内容
  - 降级方案：在 MD 中标记为 `[Embedded Object: type]`
- 编写测试：验证多媒体文件提取

**验证**：音频/视频文件被正确保存，OLE 对象被标记。

---

### Phase 20：背景图片提取

**目标**：提取幻灯片的背景图片。

**工作内容**：
- 检测幻灯片背景类型：
  - 纯色填充 → 跳过
  - 渐变填充 → 跳过
  - 图片填充 → 提取背景图片
  - 母版背景 → 从 slide_layout.slide_master 获取
- 提取背景图片 blob 并保存
- 命名规则：`slide_{NN}_bg.{ext}`
- 在 MD 中添加背景图片引用（放在内容之前）
- 编写测试：验证背景图片提取

**验证**：有背景图片的幻灯片正确提取背景。

---

### Phase 21：文档属性 → YAML Frontmatter

**目标**：提取演示文稿的元数据并写入 YAML frontmatter。

**工作内容**：
- 读取 `presentation.core_properties`：
  - `title`、`author`、`subject`、`category`
  - `created`、`modified`
  - `keywords`、`comments`
  - `last_modified_by`、`revision`
- 在 MD 文件开头生成 YAML frontmatter：
  ```
  ---
  title: "演示文稿标题"
  author: "作者"
  created: "2024-01-01"
  modified: "2024-01-02"
  source: "原始文件.pptx"
  ---
  ```
- 处理空值和特殊字符（冒号、引号等）
- 编写测试：验证 frontmatter 的正确生成

**验证**：MD 文件开头包含正确的 YAML frontmatter。

---

### Phase 22：节检测与幻灯片编号

**目标**：检测演示文稿中的节（Section）并在 MD 中添加节标题。

**工作内容**：
- 遍历 `presentation.slides` 的节信息
- 使用 `itermediate_slides` 或 XML 解析检测节边界
- 在 MD 中添加节标题：
  ```
  # Section: 章节名称

  ---
  ```
- 为每张幻灯片添加编号注释：`<!-- Slide N -->`
- 编写测试：验证节标题的正确插入

**验证**：有节的演示文稿正确输出节标题。

---

### Phase 23：母版/布局占位符过滤与隐藏幻灯片处理

**目标**：过滤母版默认占位符文本，处理隐藏幻灯片。

**工作内容**：
- 检测占位符是否为默认文本（如"单击此处添加标题"）
  - 通过 `shape.placeholder_format.idx` 和 `shape.text` 判断
  - 常见默认文本列表匹配
- 默认占位符文本跳过或标记为 `[placeholder]`
- 隐藏幻灯片检测：
  - 检查 `slide.slide_id` 在 `presentation.slides` 中的状态
  - 默认跳过隐藏幻灯片，CLI 参数 `--include-hidden` 可包含
- 空幻灯片检测和跳过
- 编写测试：验证过滤逻辑

**验证**：默认占位符文本被跳过，隐藏幻灯片按选项处理。

---

### Phase 24：空幻灯片与空白形状过滤

**目标**：智能识别并跳过无内容的幻灯片和形状。

**工作内容**：
- 空幻灯片检测：遍历所有 shapes，判断是否有实际内容
  - 有文本、图片、表格、图表 → 有内容
  - 仅有线条、空白形状 → 空白
- 空形状过滤：跳过无文本的纯形状
- 空段落过滤：跳过无文本的空行
- CLI 参数 `--skip-empty` 控制是否跳过空幻灯片
- 在 MD 中可选输出 `<!-- Empty Slide -->` 占位
- 编写测试：验证空白过滤逻辑

**验证**：空幻灯片被正确识别和过滤。

---

### Phase 25：完整 CLI 选项集成

**目标**：实现所有命令行参数并整合到统一 CLI。

**工作内容**：
- 完善 `ppt2md/main.py` 的 argparse 配置：
  - `-o / --output`：输出目录
  - `--output-file`：指定输出文件名
  - `--formula-as-image`：公式渲染为图片
  - `--include-notes`：包含演讲者备注
  - `--include-hidden`：包含隐藏幻灯片
  - `--skip-empty`：跳过空幻灯片
  - `--keep-original-format`：保留原始图片格式
  - `--image-dpi`：图片 DPI
  - `--no-frontmatter`：不生成 frontmatter
  - `--debug`：调试模式
- 实现输出目录自动创建
- 实现进度显示（`--verbose`）
- 编写测试：验证各参数组合

**验证**：所有 CLI 参数正确生效。

---

### Phase 26：批量转换模式

**目标**：支持一次转换目录下所有 .pptx 文件。

**工作内容**：
- 检测输入路径是文件还是目录
- 目录模式下扫描所有 `.pptx` 文件
- 为每个文件创建独立的输出目录
- 汇总报告：成功/失败数量
- 并发处理选项（`--jobs N`）
- 编写测试：验证批量转换

**验证**：批量转换目录下所有 pptx 文件。

---

### Phase 27：错误处理与边界情况

**目标**：健壮的错误处理，优雅降级。

**工作内容**：
- 捕获并处理常见异常：
  - `PackageNotFoundError`：无效的 pptx 文件
  - `KeyError`：缺失的 XML 节点
  - `ValueError`：无效的数据格式
  - `PermissionError`：文件权限问题
  - `IOError`：磁盘空间不足
- 损坏的图片跳过并警告
- 损坏的公式降级为原始 XML 文本
- 未知形状类型跳过并记录
- 统一的错误日志格式
- 返回码：0=成功，1=部分失败，2=全部失败
- 编写测试：用各种异常输入验证错误处理

**验证**：异常文件不崩溃，输出警告信息，返回正确退出码。

---

### Phase 28：集成测试与最终清理

**目标**：用真实 PPT 文件进行端到端测试，清理代码。

**工作内容**：
- 准备测试 PPT 文件集（`scratch/test_files/`）：
  - 纯文本 PPT
  - 含图片 PPT
  - 含表格 PPT
  - 含公式 PPT
  - 含图表 PPT
  - 含 SmartArt PPT
  - 含组合形状 PPT
  - 含备注 PPT
  - 空白 PPT
  - 复杂混合内容 PPT
- 编写端到端测试脚本
- 代码审查：检查重复代码、命名规范、文档字符串
- 性能测试：大文件（100+ 页）的处理时间
- 更新 README：完整使用文档、API 文档、示例
- 最终版本号确定和标签

**验证**：所有测试通过，代码整洁，文档完整。
