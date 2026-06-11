# PPTX 转 Markdown 开发方案

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

---

## P1：项目脚手架

| Step | 动作 | 说明 |
|------|------|------|
| P1.1 | 编写代码 | 创建目录结构 `ppt2md/`、`tests/`、`scratch/`；编写 `pyproject.toml`（项目元数据、依赖 python-pptx/Pillow/lxml、CLI 入口 ppt2md）；创建 `ppt2md/__init__.py`、`ppt2md/main.py`（argparse 骨架：input 参数、-o 输出目录、--help）；创建 `tests/__init__.py`、`tests/conftest.py`（pytest fixtures） |
| P1.2 | 测试验证 | 运行 `pip install -e .` 确认安装成功；运行 `ppt2md --help` 确认输出帮助信息；在 `scratch/test_files/` 中用 python-pptx 生成一个简单测试 pptx；编写 `tests/test_phase1.py` 测试 CLI 基础功能；运行 `pytest tests/test_phase1.py` 确认通过 |
| P1.3 | Git 提交推送 | `git checkout -b feature/phase-1-scaffolding`；`git add` 所有新文件（不包括 scratch/）；`git commit -m "feat: Phase 1 - Project Scaffolding"`；`git push -u origin feature/phase-1-scaffolding` |
| P1.4 | 创建 Issue | `gh issue create --title "Phase 1: Project Scaffolding" --body "## Goal\nSet up project scaffolding with CLI entry point.\n\n## Changes\n- Project structure: ppt2md/, tests/\n- pyproject.toml with dependencies\n- CLI skeleton with argparse\n- Pytest configuration\n\n## Verification\n- pip install -e . succeeds\n- ppt2md --help works"` |
| P1.5 | 创建 PR | `gh pr create --title "feat: Phase 1 - Project Scaffolding" --body "Closes #N\n\n## Summary\nProject scaffolding with CLI entry point.\n\n## Changes\n- ppt2md/ package with CLI\n- pyproject.toml\n- tests/ with conftest"` |
| P1.6 | 合并 PR | `gh pr merge --merge` |
| P1.7 | 关闭 Issue | 检查 Issue 是否自动关闭（PR 中有 Closes #N），如未关闭则 `gh issue close N` |
| P1.8 | 清理分支 | `git checkout main`；`git pull`；`git branch -d feature/phase-1-scaffolding`；`git push origin --delete feature/phase-1-scaffolding` |

---

## P2：基础 PPTX 解析

| Step | 动作 | 说明 |
|------|------|------|
| P2.1 | 编写代码 | 创建 `ppt2md/parser/__init__.py`、`ppt2md/parser/presentation.py`：用 `pptx.Presentation(file_path)` 打开文件，返回 slide 数量和基本信息；在 `ppt2md/main.py` 中接入：打开文件后打印幻灯片数量 |
| P2.2 | 测试验证 | 在 `scratch/test_files/` 中用 python-pptx 生成含 3 张幻灯片的测试文件；编写 `tests/test_phase2.py`：测试能正确读取幻灯片数；运行 `pytest tests/test_phase2.py` 确认通过 |
| P2.3 | Git 提交推送 | 分支 `feature/phase-2-basic-parsing`；commit `feat: Phase 2 - Basic PPTX Parsing`；push |
| P2.4 | 创建 Issue | 标题 `Phase 2: Basic PPTX Parsing`（英文正文） |
| P2.5 | 创建 PR | 关联 Issue（`Closes #N`），英文描述 |
| P2.6 | 合并 PR | `gh pr merge --merge` |
| P2.7 | 关闭 Issue | 确认关闭 |
| P2.8 | 清理分支 | 删除临时分支，回到 main |

---

## P3：文本帧提取

| Step | 动作 | 说明 |
|------|------|------|
| P3.1 | 编写代码 | 创建 `ppt2md/parser/text.py`：遍历 slide.shapes，检测 has_text_frame，提取 paragraph.text 和 run.text，跳过空文本帧和空段落 |
| P3.2 | 测试验证 | 测试用 pptx 包含文本框和占位符；验证所有文本被正确提取；验证空文本被跳过 |
| P3.3 | Git 提交推送 | 分支 `feature/phase-3-text-extraction`；commit `feat: Phase 3 - Text Frame Extraction`；push |
| P3.4 | 创建 Issue | 标题 `Phase 3: Text Frame Extraction`（英文） |
| P3.5 | 创建 PR | 关联 Issue，英文描述 |
| P3.6 | 合并 PR | `gh pr merge --merge` |
| P3.7 | 关闭 Issue | 确认关闭 |
| P3.8 | 清理分支 | 删除临时分支，回到 main |

---

## P4：文本格式转 Markdown

| Step | 动作 | 说明 |
|------|------|------|
| P4.1 | 编写代码 | 创建 `ppt2md/converter/__init__.py`、`ppt2md/converter/format_utils.py`：检测 run.font.bold→`**text**`、italic→`*text*`、underline→`<u>text</u>`、strikethrough→`~~text~~`；处理同一段落内多 run 混合格式拼接 |
| P4.2 | 测试验证 | 测试加粗、斜体、下划线、删除线、混合格式；验证 Markdown 输出正确 |
| P4.3 | Git 提交推送 | 分支 `feature/phase-4-text-formatting`；commit `feat: Phase 4 - Text Formatting to Markdown`；push |
| P4.4 | 创建 Issue | 标题 `Phase 4: Text Formatting to Markdown`（英文） |
| P4.5 | 创建 PR | 关联 Issue，英文描述 |
| P4.6 | 合并 PR | `gh pr merge --merge` |
| P4.7 | 关闭 Issue | 确认关闭 |
| P4.8 | 清理分支 | 删除临时分支，回到 main |

---

## P5：占位符类型映射

| Step | 动作 | 说明 |
|------|------|------|
| P5.1 | 编写代码 | 检测 shape.placeholder_format.idx：0→标题(`##`)、6→正文(段落)、跳过日期/页脚/页眉/编号等元信息占位符 |
| P5.2 | 测试验证 | 测试标题占位符输出 `##`，正文为普通段落，页脚等被跳过 |
| P5.3 | Git 提交推送 | 分支 `feature/phase-5-placeholder-mapping`；commit `feat: Phase 5 - Placeholder Type Mapping`；push |
| P5.4 | 创建 Issue | 标题 `Phase 5: Placeholder Type Mapping`（英文） |
| P5.5 | 创建 PR | 关联 Issue，英文描述 |
| P5.6 | 合并 PR | `gh pr merge --merge` |
| P5.7 | 关闭 Issue | 确认关闭 |
| P5.8 | 清理分支 | 删除临时分支，回到 main |

---

## P6：列表和项目符号

| Step | 动作 | 说明 |
|------|------|------|
| P6.1 | 编写代码 | 检测 paragraph.level（缩进层级）：level 0→`- item`，level 1→`  - item`，level 2→`    - item`；检测有序列表（数字+点）→`1. item` |
| P6.2 | 测试验证 | 测试多层嵌套列表输出，验证缩进正确 |
| P6.3 | Git 提交推送 | 分支 `feature/phase-6-list-bullet`；commit `feat: Phase 6 - List and Bullet Detection`；push |
| P6.4 | 创建 Issue | 标题 `Phase 6: List and Bullet Detection`（英文） |
| P6.5 | 创建 PR | 关联 Issue，英文描述 |
| P6.6 | 合并 PR | `gh pr merge --merge` |
| P6.7 | 关闭 Issue | 确认关闭 |
| P6.8 | 清理分支 | 删除临时分支，回到 main |

---

## P7：图片提取

| Step | 动作 | 说明 |
|------|------|------|
| P7.1 | 编写代码 | 创建 `ppt2md/parser/image.py`：遍历 shapes，检测 shape.image，提取 shape.image.blob 和 content_type，按扩展名保存到 `output/images/`，命名 `slide_{NN}_img_{MM}.{ext}` |
| P7.2 | 测试验证 | 测试用 pptx 包含图片；验证图片文件正确保存到 images/ 目录；验证文件名格式正确 |
| P7.3 | Git 提交推送 | 分支 `feature/phase-7-image-extraction`；commit `feat: Phase 7 - Image Extraction`；push |
| P7.4 | 创建 Issue | 标题 `Phase 7: Image Extraction`（英文） |
| P7.5 | 创建 PR | 关联 Issue，英文描述 |
| P7.6 | 合并 PR | `gh pr merge --merge` |
| P7.7 | 关闭 Issue | 确认关闭 |
| P7.8 | 清理分支 | 删除临时分支，回到 main |

---

## P8：图片位置与尺寸

| Step | 动作 | 说明 |
|------|------|------|
| P8.1 | 编写代码 | 创建 `ppt2md/converter/position_utils.py`：EMU→cm（÷360000）和 px（÷914400×dpi）转换；在 MD 中输出 `<!-- position: x=10.28cm, y=6.57cm, width=28.81cm, height=10.28cm -->` 注释 |
| P8.2 | 测试验证 | 验证位置信息格式和单位正确 |
| P8.3 | Git 提交推送 | 分支 `feature/phase-8-image-position`；commit `feat: Phase 8 - Image Position and Size`；push |
| P8.4 | 创建 Issue | 标题 `Phase 8: Image Position and Size`（英文） |
| P8.5 | 创建 PR | 关联 Issue，英文描述 |
| P8.6 | 合并 PR | `gh pr merge --merge` |
| P8.7 | 关闭 Issue | 确认关闭 |
| P8.8 | 清理分支 | 删除临时分支，回到 main |

---

## P9：图片裁剪处理

| Step | 动作 | 说明 |
|------|------|------|
| P9.1 | 编写代码 | 检测 crop_left/right/top/bottom（比例 0.0-1.0），在 MD 注释中追加 `crop: left=0.1, right=0.05` |
| P9.2 | 测试验证 | 验证裁剪信息正确记录 |
| P9.3 | Git 提交推送 | 分支 `feature/phase-9-image-crop`；commit `feat: Phase 9 - Image Crop Handling`；push |
| P9.4 | 创建 Issue | 标题 `Phase 9: Image Crop Handling`（英文） |
| P9.5 | 创建 PR | 关联 Issue，英文描述 |
| P9.6 | 合并 PR | `gh pr merge --merge` |
| P9.7 | 关闭 Issue | 确认关闭 |
| P9.8 | 清理分支 | 删除临时分支，回到 main |

---

## P10：图片去重

| Step | 动作 | 说明 |
|------|------|------|
| P10.1 | 编写代码 | 建立 `rId → saved_filename` 映射表；已存在的 rId 复用文件名，不重复保存 |
| P10.2 | 测试验证 | 验证同一图片多处引用只保存一次 |
| P10.3 | Git 提交推送 | 分支 `feature/phase-10-image-dedup`；commit `feat: Phase 10 - Image Deduplication`；push |
| P10.4 | 创建 Issue | 标题 `Phase 10: Image Deduplication`（英文） |
| P10.5 | 创建 PR | 关联 Issue，英文描述 |
| P10.6 | 合并 PR | `gh pr merge --merge` |
| P10.7 | 关闭 Issue | 确认关闭 |
| P10.8 | 清理分支 | 删除临时分支，回到 main |

---

## P11：表格提取

| Step | 动作 | 说明 |
|------|------|------|
| P11.1 | 编写代码 | 创建 `ppt2md/parser/table.py`：检测 has_table，遍历 rows/columns 提取 cell.text，生成 Markdown 表格语法；处理单元格内多段落用 `<br>` 连接 |
| P11.2 | 测试验证 | 验证表格正确转为 Markdown 表格 |
| P11.3 | Git 提交推送 | 分支 `feature/phase-11-table-extraction`；commit `feat: Phase 11 - Table Extraction`；push |
| P11.4 | 创建 Issue | 标题 `Phase 11: Table Extraction`（英文） |
| P11.5 | 创建 PR | 关联 Issue，英文描述 |
| P11.6 | 合并 PR | `gh pr merge --merge` |
| P11.7 | 关闭 Issue | 确认关闭 |
| P11.8 | 清理分支 | 删除临时分支，回到 main |

---

## P12：演讲者备注

| Step | 动作 | 说明 |
|------|------|------|
| P12.1 | 编写代码 | 创建 `ppt2md/parser/notes.py`：检测 has_notes_slide，提取 notes_text_frame.text，在 MD 末尾输出 `<!-- Notes -->` + `### Speaker Notes` 区块 |
| P12.2 | 测试验证 | 验证备注正确提取和格式化 |
| P12.3 | Git 提交推送 | 分支 `feature/phase-12-speaker-notes`；commit `feat: Phase 12 - Speaker Notes`；push |
| P12.4 | 创建 Issue | 标题 `Phase 12: Speaker Notes`（英文） |
| P12.5 | 创建 PR | 关联 Issue，英文描述 |
| P12.6 | 合并 PR | `gh pr merge --merge` |
| P12.7 | 关闭 Issue | 确认关闭 |
| P12.8 | 清理分支 | 删除临时分支，回到 main |

---

## P13：形状位置排序

| Step | 动作 | 说明 |
|------|------|------|
| P13.1 | 编写代码 | 按 shape.top 升序、同行内 shape.left 升序排列；处理 None 值（视为最大值排最后）；top 差值在阈值内视为同一行 |
| P13.2 | 测试验证 | 验证排序后输出顺序符合视觉阅读顺序 |
| P13.3 | Git 提交推送 | 分支 `feature/phase-13-shape-sorting`；commit `feat: Phase 13 - Shape Position Sorting`；push |
| P13.4 | 创建 Issue | 标题 `Phase 13: Shape Position Sorting`（英文） |
| P13.5 | 创建 PR | 关联 Issue，英文描述 |
| P13.6 | 合并 PR | `gh pr merge --merge` |
| P13.7 | 关闭 Issue | 确认关闭 |
| P13.8 | 清理分支 | 删除临时分支，回到 main |

---

## P14：OMML 基础公式转换

| Step | 动作 | 说明 |
|------|------|------|
| P14.1 | 编写代码 | 创建 `ppt2md/parser/formula.py`：查找 `<m:oMath>` 和 `<m:oMathPara>` 节点；实现 m:r(文本)、m:f(分数→`\frac{a}{b}`)、m:rad(根号→`\sqrt{x}`)、m:sSub(下标→`x_{n}`)、m:sSup(上标→`x^{n}`)、m:d(括号→`\left( \right)`) → LaTeX |
| P14.2 | 测试验证 | 验证基础公式正确转为 LaTeX |
| P14.3 | Git 提交推送 | 分支 `feature/phase-14-omml-basic`；commit `feat: Phase 14 - OMML Basic LaTeX`；push |
| P14.4 | 创建 Issue | 标题 `Phase 14: OMML Basic LaTeX`（英文） |
| P14.5 | 创建 PR | 关联 Issue，英文描述 |
| P14.6 | 合并 PR | `gh pr merge --merge` |
| P14.7 | 关闭 Issue | 确认关闭 |
| P14.8 | 清理分支 | 删除临时分支，回到 main |

---

## P15：OMML 全面公式转换

| Step | 动作 | 说明 |
|------|------|------|
| P15.1 | 编写代码 | 扩展 formula.py：m:nary(求和/积分→`\sum_{i=0}^{n}`)、m:func(函数→`\sin`)、m:acc(重音→`\hat{x}`)、m:bar(横线→`\overline{x}`)、m:limLow/m:limUpp(极限)、m:m(矩阵→`\begin{matrix}`)、m:eqArr(方程组)、m:sPre(前置上下标)、m:groupChr(群组字符→`\underbrace`) |
| P15.2 | 测试验证 | 验证复杂嵌套公式正确转为 LaTeX |
| P15.3 | Git 提交推送 | 分支 `feature/phase-15-omml-comprehensive`；commit `feat: Phase 15 - OMML Comprehensive LaTeX`；push |
| P15.4 | 创建 Issue | 标题 `Phase 15: OMML Comprehensive LaTeX`（英文） |
| P15.5 | 创建 PR | 关联 Issue，英文描述 |
| P15.6 | 合并 PR | `gh pr merge --merge` |
| P15.7 | 关闭 Issue | 确认关闭 |
| P15.8 | 清理分支 | 删除临时分支，回到 main |

---

## P16：图表数据提取

| Step | 动作 | 说明 |
|------|------|------|
| P16.1 | 编写代码 | 创建 `ppt2md/parser/chart.py`：检测 has_chart，读取 chart_type、series、chart_data，输出 MD 表格 |
| P16.2 | 测试验证 | 验证图表类型和数据正确提取 |
| P16.3 | Git 提交推送 | 分支 `feature/phase-16-chart-data`；commit `feat: Phase 16 - Chart Data Extraction`；push |
| P16.4 | 创建 Issue | 标题 `Phase 16: Chart Data Extraction`（英文） |
| P16.5 | 创建 PR | 关联 Issue，英文描述 |
| P16.6 | 合并 PR | `gh pr merge --merge` |
| P16.7 | 关闭 Issue | 确认关闭 |
| P16.8 | 清理分支 | 删除临时分支，回到 main |

---

## P17：图表图片导出

| Step | 动作 | 说明 |
|------|------|------|
| P17.1 | 编写代码 | `win32com.client` 连接 PowerPoint，`shape.Export()` 导出图表为 PNG；仅 Windows 可用，COM 不可用时跳过 |
| P17.2 | 测试验证 | 验证图表图片正确导出（Windows 环境） |
| P17.3 | Git 提交推送 | 分支 `feature/phase-17-chart-image`；commit `feat: Phase 17 - Chart Image Export`；push |
| P17.4 | 创建 Issue | 标题 `Phase 17: Chart Image Export`（英文） |
| P17.5 | 创建 PR | 关联 Issue，英文描述 |
| P17.6 | 合并 PR | `gh pr merge --merge` |
| P17.7 | 关闭 Issue | 确认关闭 |
| P17.8 | 清理分支 | 删除临时分支，回到 main |

---

## P18：SmartArt 与组合形状

| Step | 动作 | 说明 |
|------|------|------|
| P18.1 | 编写代码 | 递归解析 GROUP 形状（shape.shape_type == MSO_SHAPE_TYPE.GROUP）；XML 解析 SmartArt 文本节点 |
| P18.2 | 测试验证 | 验证组合形状和 SmartArt 文本正确提取 |
| P18.3 | Git 提交推送 | 分支 `feature/phase-18-smartart`；commit `feat: Phase 18 - SmartArt and Group Shapes`；push |
| P18.4 | 创建 Issue | 标题 `Phase 18: SmartArt and Group Shapes`（英文） |
| P18.5 | 创建 PR | 关联 Issue，英文描述 |
| P18.6 | 合并 PR | `gh pr merge --merge` |
| P18.7 | 关闭 Issue | 确认关闭 |
| P18.8 | 清理分支 | 删除临时分支，回到 main |

---

## P19：多媒体与 OLE 对象

| Step | 动作 | 说明 |
|------|------|------|
| P19.1 | 编写代码 | 创建 `ppt2md/parser/media.py`：提取音视频 blob 到 `output/media/`；OLE 对象标记为 `[Embedded Object: type]` |
| P19.2 | 测试验证 | 验证多媒体文件正确保存 |
| P19.3 | Git 提交推送 | 分支 `feature/phase-19-media-ole`；commit `feat: Phase 19 - Media and OLE Objects`；push |
| P19.4 | 创建 Issue | 标题 `Phase 19: Media and OLE Objects`（英文） |
| P19.5 | 创建 PR | 关联 Issue，英文描述 |
| P19.6 | 合并 PR | `gh pr merge --merge` |
| P19.7 | 关闭 Issue | 确认关闭 |
| P19.8 | 清理分支 | 删除临时分支，回到 main |

---

## P20：背景图片

| Step | 动作 | 说明 |
|------|------|------|
| P20.1 | 编写代码 | 检测图片填充背景，提取 blob 保存为 `slide_{NN}_bg.{ext}`，在 MD 中放在内容之前引用 |
| P20.2 | 测试验证 | 验证背景图片正确提取 |
| P20.3 | Git 提交推送 | 分支 `feature/phase-20-background`；commit `feat: Phase 20 - Background Image`；push |
| P20.4 | 创建 Issue | 标题 `Phase 20: Background Image`（英文） |
| P20.5 | 创建 PR | 关联 Issue，英文描述 |
| P20.6 | 合并 PR | `gh pr merge --merge` |
| P20.7 | 关闭 Issue | 确认关闭 |
| P20.8 | 清理分支 | 删除临时分支，回到 main |

---

## P21：文档属性 Frontmatter

| Step | 动作 | 说明 |
|------|------|------|
| P21.1 | 编写代码 | 读取 presentation.core_properties（title/author/created/modified/source），生成 YAML frontmatter |
| P21.2 | 测试验证 | 验证 frontmatter 格式正确 |
| P21.3 | Git 提交推送 | 分支 `feature/phase-21-frontmatter`；commit `feat: Phase 21 - Document Properties Frontmatter`；push |
| P21.4 | 创建 Issue | 标题 `Phase 21: Document Properties Frontmatter`（英文） |
| P21.5 | 创建 PR | 关联 Issue，英文描述 |
| P21.6 | 合并 PR | `gh pr merge --merge` |
| P21.7 | 关闭 Issue | 确认关闭 |
| P21.8 | 清理分支 | 删除临时分支，回到 main |

---

## P22：节检测

| Step | 动作 | 说明 |
|------|------|------|
| P22.1 | 编写代码 | 检测 presentation.sections，在 MD 中插入 `# Section: ...` 标题 |
| P22.2 | 测试验证 | 验证节标题正确插入 |
| P22.3 | Git 提交推送 | 分支 `feature/phase-22-section-detection`；commit `feat: Phase 22 - Section Detection`；push |
| P22.4 | 创建 Issue | 标题 `Phase 22: Section Detection`（英文） |
| P22.5 | 创建 PR | 关联 Issue，英文描述 |
| P22.6 | 合并 PR | `gh pr merge --merge` |
| P22.7 | 关闭 Issue | 确认关闭 |
| P22.8 | 清理分支 | 删除临时分支，回到 main |

---

## P23：母版/布局占位符过滤

| Step | 动作 | 说明 |
|------|------|------|
| P23.1 | 编写代码 | 匹配默认占位符文本（"单击此处添加标题"等中英文），跳过或标记为 `[placeholder]` |
| P23.2 | 测试验证 | 验证默认占位符被正确过滤 |
| P23.3 | Git 提交推送 | 分支 `feature/phase-23-master-filtering`；commit `feat: Phase 23 - Master Layout Filtering`；push |
| P23.4 | 创建 Issue | 标题 `Phase 23: Master Layout Filtering`（英文） |
| P23.5 | 创建 PR | 关联 Issue，英文描述 |
| P23.6 | 合并 PR | `gh pr merge --merge` |
| P23.7 | 关闭 Issue | 确认关闭 |
| P23.8 | 清理分支 | 删除临时分支，回到 main |

---

## P24：空白幻灯片过滤

| Step | 动作 | 说明 |
|------|------|------|
| P24.1 | 编写代码 | 检测无内容幻灯片（无文本/图片/表格/图表）；CLI 参数 `--skip-empty` 控制 |
| P24.2 | 测试验证 | 验证空幻灯片被正确识别和过滤 |
| P24.3 | Git 提交推送 | 分支 `feature/phase-24-empty-filtering`；commit `feat: Phase 24 - Empty Slide Filtering`；push |
| P24.4 | 创建 Issue | 标题 `Phase 24: Empty Slide Filtering`（英文） |
| P24.5 | 创建 PR | 关联 Issue，英文描述 |
| P24.6 | 合并 PR | `gh pr merge --merge` |
| P24.7 | 关闭 Issue | 确认关闭 |
| P24.8 | 清理分支 | 删除临时分支，回到 main |

---

## P25：完整 CLI 集成

| Step | 动作 | 说明 |
|------|------|------|
| P25.1 | 编写代码 | 完善所有 CLI 参数：-o、--include-notes、--include-hidden、--skip-empty、--formula-as-image、--keep-original-format、--image-dpi、--no-frontmatter、--verbose、--debug |
| P25.2 | 测试验证 | 验证各参数组合正确生效 |
| P25.3 | Git 提交推送 | 分支 `feature/phase-25-full-cli`；commit `feat: Phase 25 - Full CLI Integration`；push |
| P25.4 | 创建 Issue | 标题 `Phase 25: Full CLI Integration`（英文） |
| P25.5 | 创建 PR | 关联 Issue，英文描述 |
| P25.6 | 合并 PR | `gh pr merge --merge` |
| P25.7 | 关闭 Issue | 确认关闭 |
| P25.8 | 清理分支 | 删除临时分支，回到 main |

---

## P26：批量转换

| Step | 动作 | 说明 |
|------|------|------|
| P26.1 | 编写代码 | 检测目录输入，扫描所有 .pptx，为每个创建独立输出目录，汇总报告成功/失败数量 |
| P26.2 | 测试验证 | 验证批量转换目录下所有 pptx 文件 |
| P26.3 | Git 提交推送 | 分支 `feature/phase-26-batch`；commit `feat: Phase 26 - Batch Conversion`；push |
| P26.4 | 创建 Issue | 标题 `Phase 26: Batch Conversion`（英文） |
| P26.5 | 创建 PR | 关联 Issue，英文描述 |
| P26.6 | 合并 PR | `gh pr merge --merge` |
| P26.7 | 关闭 Issue | 确认关闭 |
| P26.8 | 清理分支 | 删除临时分支，回到 main |

---

## P27：错误处理

| Step | 动作 | 说明 |
|------|------|------|
| P27.1 | 编写代码 | 捕获 PackageNotFoundError/KeyError/ValueError/PermissionError/IOError；损坏图片跳过并警告；损坏公式降级为原始 XML；未知形状跳过并记录；返回码 0=成功 1=部分失败 2=全部失败 |
| P27.2 | 测试验证 | 用异常输入验证错误处理，确认返回正确退出码 |
| P27.3 | Git 提交推送 | 分支 `feature/phase-27-error-handling`；commit `feat: Phase 27 - Error Handling`；push |
| P27.4 | 创建 Issue | 标题 `Phase 27: Error Handling`（英文） |
| P27.5 | 创建 PR | 关联 Issue，英文描述 |
| P27.6 | 合并 PR | `gh pr merge --merge` |
| P27.7 | 关闭 Issue | 确认关闭 |
| P27.8 | 清理分支 | 删除临时分支，回到 main |

---

## P28：集成测试与清理

| Step | 动作 | 说明 |
|------|------|------|
| P28.1 | 编写代码 | 在 `scratch/test_files/` 准备测试 PPT 文件集（纯文本、含图片、含表格、含公式、含图表、含 SmartArt、含组合形状、含备注、空白、复杂混合）；编写端到端测试；代码审查；更新 README 完整文档 |
| P28.2 | 测试验证 | 所有测试通过；性能测试（100+ 页 PPT） |
| P28.3 | Git 提交推送 | 分支 `feature/phase-28-integration`；commit `feat: Phase 28 - Integration Testing`；push |
| P28.4 | 创建 Issue | 标题 `Phase 28: Integration Testing`（英文） |
| P28.5 | 创建 PR | 关联 Issue，英文描述 |
| P28.6 | 合并 PR | `gh pr merge --merge` |
| P28.7 | 关闭 Issue | 确认关闭 |
| P28.8 | 清理分支 | 删除临时分支，回到 main |

---

## 输出格式规范

### Markdown 文件结构

```markdown
---
title: "演示文稿标题"
author: "作者"
source: "input.pptx"
---

# 演示文稿标题

---

## Slide 1: 标题

正文内容。

- 列表项 1
- 列表项 2

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

### 图片命名

```
images/
├── slide_01_img_01.png
├── slide_01_img_02.png
├── slide_02_chart_01.png
├── slide_01_bg.png
└── ...
```
