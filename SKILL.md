---
name: cn-patent-fasttrack
description: >-
  中国发明与实用新型专利零探索快车道技能（支持硬件/结构/智能穿戴类与软件/算法流程类双分支）。
  当用户希望从技术构想、草图、代码、论文或旧版草稿快速产出中国专利申请全套文件、做竞品专利定向规避、
  设计全排列组合与多组独立权利要求防线、绘制3D工业设计级黑白结构线稿或标准算法流程图、
  或一键核验标号并导出国知局五分件Word、合稿Word、矢量线稿PDF与完整资料包ZIP时使用。
---

# 中国专利零探索快车道（CN Patent Fast-Track）

## Overview

本技能将「竞品规避拆解 → 全排列组合/多独权防线设计 → 件号表先行锁号 → 3D立体结构线稿（或标准算法流程图）渲染与视觉自检 → 一键联检并导出国知局五分件 Word/PDF/ZIP」固化为一条零探索流水线。无需在多层子技能目录中盲搜，也无需每次临时手写数百行 SVG 转 PNG、标号核对或 Word 排版脚本。

## Dependencies

- **`patent-disclosure-skill`**（位于 `~/.codex/skills/patent-disclosure-skill` 及 `~/.gemini/config/skills/patent-disclosure-skill`）：本技能内置 CLI 自动直连复用其 `.venv` 环境（Playwright、Pillow、`python-docx`、PyYAML）以及 `audit_claims.py`、`render_invention_figures.py`、`md_to_docx.py`、`latex_delimiters.py`；仅在需要深度审查答复（OA）、专利地图或外观设计六视图检查时，才去读取 `patent-disclosure-skill/SKILL.md`。
- **`uv`**（`~/.local/bin/uv`）：统一通过 `uv run --no-project` 启动本技能 CLI（脚本内部会自动切换至已预装依赖的 `.venv` 解释器，零下载耗时）。

## Quick Start

```bash
export PATH="$HOME/.local/bin:$PATH"
SKILL_CLI="$HOME/.gemini/config/skills/cn-patent-fasttrack/scripts/patent_fasttrack.py"

# 1. 初始化案件目录（硬件选 hardware，软件算法选 software）
uv run --no-project "$SKILL_CLI" init \
  --case-dir "./智能穿戴传感专利_20261125" \
  --mode hardware \
  --title "智能穿戴传感组件及设备" \
  --output "./智能穿戴传感专利_20261125/.fasttrack_init.json"

# 2. 渲染纯黑白 300DPI 紧裁 PNG 与汇总矢量 PDF（软件类可加 --plan invention_figures.yaml）
uv run --no-project "$SKILL_CLI" render-figures \
  --case-dir "./智能穿戴传感专利_20261125" \
  --dpi 300 --pad-px 24 \
  --output "./智能穿戴传感专利_20261125/.fasttrack_render.json"

# 3. 标号三方对齐 + 权要形式审查 + 摘要字数 + 公式定界符联检
uv run --no-project "$SKILL_CLI" audit \
  --case-dir "./智能穿戴传感专利_20261125" \
  --output "./智能穿戴传感专利_20261125/.fasttrack_audit.json"

# 4. 一键生成国知局五分件 Word + 审阅合稿 Word + 完整申请资料包 ZIP
uv run --no-project "$SKILL_CLI" build-package \
  --case-dir "./智能穿戴传感专利_20261125" \
  --title "智能穿戴传感组件" \
  --output "./智能穿戴传感专利_20261125/.fasttrack_package.json"
```

## Workflow

遇到新专利撰写或改稿任务时，严格按以下 **5 步快车道** 执行：

### 1. 分支判定与竞品定向规避（先读 `references/claim_and_fto_playbook.md`）
- **判定案件分支**：
  - **硬件/结构/智能穿戴类（`hardware`）** → 同步阅读 `references/hardware_3d_svg_guide.md`；
  - **软件/算法/信号处理类（`software`）** → 同步阅读 `references/software_flow_guide.md`。
- **不阻塞输入原则**：只有当决定独权 1 核心边界的关键技术结构/步骤不明时，才向用户集中提问（最多 2–3 问）；发明人姓名、样机实测公差或实验数值缺失时，用工程示例范围写入说明书并记入《问题清单.md》，**严禁因缺交底目录卡死主流程**。
- **竞品规避检索**：若需检索竞品专利，可调用 `patent-disclosure-skill` 的 `cnipa_search.py` / `gp_search.py` 或直接使用内置 `search_web`；若遇 WAF 或 429 限流，立即改用 `search_web` 查公开文本并在《复核记录与规避分析.md》注明核验边界。

### 2. 初始化案件与锁死《件号登记表.yaml》
- 在当前工作区下运行 `init` 子命令创建 `{案件名称}_{YYYYMMDD}/`。
- **先锁号、后写文画图**：在写正文和画图前，先把全部图幅与附图标记（硬件 `100..199`，软件步骤 `S101..` 与模块名）登记进 `件号登记表.yaml`，作为全案唯一标号真值源。

### 3. 全排列组合权利要求与正文一次成稿
- **硬件类推荐三独权防线**：① 独权 1 卡住核心物理装配/耦合骨架，从权穷举全部分区拓扑、材料马库什并列、多模态单元排列组合；② 针对高价值复合子模组单设第二独权（防竞品更换外框底座绕开，备实审分案）；③ 整机及动态通道选通系统设第三独权（固化主控电路异位布置、通道物理隔离等结构避让特征）。
- **软件类推荐四独权防线**：方法独权 1 + 退化降级/自校准分支从权 + 一一对应的装置模块独权 + 电子设备独权 + 存储介质独权；所有数学公式严格用 `$...$` 或 `$$...$$` 包裹。
- 说明书摘要控制在 **200–280 字**（含标点严禁超 300 字），交代技术问题、核心结构/步骤手段与主要用途。

### 4. 附图矢量生成与强制视觉读图自检
- **硬件类**：在案件 `scripts/` 下编写参数化 Python 脚本生成 `可编辑线稿/图1_*.svg`～`图N_*.svg`，严格采用双曲倒角轮廓 + 侧裙厚度脊线、未受载/受载双态剖面、30° 等轴测爆炸图与微孔弹性体/剖面 `<pattern>` 图例，杜绝扁平药丸画法。
- **软件类**：编写 `invention_figures.yaml`，运行 `render-figures --plan invention_figures.yaml` 自动生成标准黑白流程图与虚线框模块图。
- **强制视觉质检门禁**：运行 `render-figures` 后，**必须调用 `read_file` 逐张查看生成的 `附图PNG/图*.png`**，肉眼确认：① 整机/底座左右严格水平对称无歪斜；② 引出线圆点准确落在零件墨线上、无交叉重叠；③ 无灰阶底色、无残留 `图1` 或中文解释字样。不合格先改图再往下走。

### 5. 一键联检（`audit`）与五分件打包（`build-package`）
- 运行 `audit` 并读取输出 JSON：确保 `missing_in_svgs == []`、`unregistered_in_svgs == []`、权要 `errors == 0`、摘要 ≤ 300 字、无裸括号 LaTeX 公式；若有报错一次性修完。
- 运行 `build-package` 一次性产出：国知局五分件 Word（`权利要求书.docx`、`说明书.docx`、`说明书摘要.docx`、`说明书附图.docx`、`摘要附图.docx`［纯图无图号］）、审阅合稿 Word、汇总线稿 PDF 与完整申请资料包 ZIP。
- 回复末尾统一附上可点击 `file://` 链接，并以 `## 交付后请确认` 摘要提示《问题清单.md》与《复核记录与规避分析.md》中的待核事项。

## Utility Scripts

所有子命令均通过 `uv run --no-project` 调用 `scripts/patent_fasttrack.py`，且必须显式传入 `--output <json>` 将详细报告写入文件（终端仅打印单行状态）：

1. **`init`**：初始化硬件或软件专利案件目录骨架与 YAML/Markdown 模板
   - 必选参数：`--case-dir <dir>`、`--mode {hardware,software}`、`--title <案件标题>`、`--output <json>`；可选 `--force`
2. **`render-figures`**：CNIPA 纯黑白静态扫描 + 无头浏览器紧裁导出 300 DPI PNG 与 A4 矢量汇总 PDF
   - 必选参数：`--case-dir <dir>`、`--output <json>`；可选 `--plan <invention_figures.yaml>`、`--dpi <int>`（默认 300）、`--pad-px <int>`（默认 24）、`--pdf-name <文件名.pdf>`
3. **`audit`**：标号/步骤三方一致性 + `audit_claims.py` 权要审查 + 摘要字数 + 公式定界符一站式联检
   - 必选参数：`--case-dir <dir>`、`--output <json>`（存在 `ERROR` 时退出码为 `1`）
4. **`build-package`**：构建国知局五分件 Word（含原生 OMML 公式渲染）、合稿审阅 Word 与完整资料包 ZIP
   - 必选参数：`--case-dir <dir>`、`--title <文件前缀>`、`--output <json>`

## Rate Limiting

- `patent_fasttrack.py` 的四个核心子命令均为纯本地文件与图形排版计算，无外部网络请求。
- 若调用外部检索脚本（如 `gp_search.py`），严格执行单次请求不循环重试，遇 HTTP `429` 或退出码 `4` 立即停止并切换 `search_web`。脚本内置 `FileLockRateLimiter`（基于 `fcntl` 跨进程锁与 `time.monotonic()`，默认间隔 ≥ 1.0 秒，5xx 指数退避，429 抛出 `RateLimitError`）。

## Common Mistakes

1. **在 SVG 附图内部写 `图1` 标题、中文解释句或使用半透明灰色（`opacity` / `#ccc`）**：违反国知局黑白附图规范且会导致 Word 附图页出现双层「图1」。子图解释统一写入说明书「附图说明」段，材质区分一律用纯黑白 `<pattern>` 剖面线或点阵。
2. **用扁平矩形/椭圆硬凑硬件结构图，或未做 `read_file` 视觉检查就交付**：导致爆炸图缺乏立体厚度、正视图外框左右歪斜或引出线悬空。每次改完 SVG 必须渲染 PNG 并用 `read_file` 肉眼质检。
3. **`摘要附图.docx` 误加 `图1`/`图2` 图号，或软件公式裸写圆括号 `(x_i)`**：国知局电子申请要求《摘要附图》分件内仅放单幅干净插图不带图号文字（`build-package` 已自动处理）；说明书公式必须用 `$x_i$` 包裹以生成原生 OMML 公式。
