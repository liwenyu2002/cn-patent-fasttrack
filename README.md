# cn-patent-fasttrack

中国发明与实用新型专利**零探索快车道技能**（支持**硬件/结构/智能穿戴类**与**软件/算法流程类**双分支），适用于 Antigravity、Codex 与 Claude Code。

在 [`patent-disclosure-skill`](https://github.com/handsomestWei/patent-disclosure-skill) 的校验与转换底层之上，固化了一套高省 Token 的端到端作业流与单入口 CLI，省去每次翻找多层子技能目录以及临时手写 SVG 转 PNG、标号三方核对和国知局五分件 Word 打包脚本的开销。

## 核心能力

- **双分支按需路由**：
  - **硬件/结构类（`hardware`）**：内置 3D 双曲倒角轮廓、未受载/受载双态物理剖面、30° 等轴测爆炸图、纯黑白微孔弹性体/剖面 `<pattern>` 图例及圆点引线函数，解决纯 2D 代码绘图扁平失真的问题。
  - **软件/算法类（`software`）**：内置标准 `invention_figures.yaml` 方法流程图与虚线框模块图模板，自动将 `$...$` / `$$...$$` 公式渲染为可编辑 Word 原生 OMML 公式（`<m:oMath>`）。
- **多组独权与竞品规避手册**：提供「核心骨架独权 + 拓扑/混排穷举从权 + 高价值子模组独立权利要求（防绕开/备分案）+ 整机系统独权」防线模板及全部技术特征避让核查表。
- **单文件固化 CLI（`scripts/patent_fasttrack.py`）**：
  - `init`：一键初始化硬件或软件专利目录骨架、`件号登记表.yaml` 与 Markdown 模板；
  - `render-figures`：国知局纯黑白零灰阶静态扫描 + 无头 Chromium 内容包围盒紧裁导出 300 DPI PNG 与 A4 矢量汇总 PDF；
  - `audit`：三方标号/步骤一致性（YAML ↔ SVG ↔ 正文）+ 权利要求形式审查 + 摘要 300 字上限 + 裸括号公式拦截；
  - `build-package`：一键导出国知局电子申请五分件 Word（`权利要求书.docx`、`说明书.docx`、`说明书摘要.docx`、`说明书附图.docx`、`摘要附图.docx`）、审阅合稿 Word 与完整资料包 ZIP。

## 目录结构

```text
cn-patent-fasttrack/
├── SKILL.md                              # 主指令（5步快车道 SOP、双分支路由与视觉质检门禁）
├── README.md                             # 安装与使用说明
├── scripts/
│   └── patent_fasttrack.py               # 统一 CLI（init / render-figures / audit / build-package）
└── references/
    ├── hardware_3d_svg_guide.md          # 3D 工业设计级黑白专利结构线稿规范与可复用代码库
    ├── software_flow_guide.md            # 算法流程图/模块框图 YAML 模板与 Word OMML 公式规范
    └── claim_and_fto_playbook.md         # 全排列组合从权矩阵、子模组分案独权与竞品规避手册
```

## 前置依赖与安装

1. 确保本机已安装 [`patent-disclosure-skill`](https://github.com/handsomestWei/patent-disclosure-skill)（已完成 `.venv` 与 Playwright Chromium 初始化）及 [`uv`](https://docs.astral.sh/uv/)。
2. 克隆至对应 Agent 的全局技能目录（可建立软链接跨端共用）：

```bash
# Antigravity 全局技能目录
git clone https://github.com/liwenyu2002/cn-patent-fasttrack.git ~/.gemini/config/skills/cn-patent-fasttrack

# 可选：同步软链至 Codex / Claude Code
ln -sfn ~/.gemini/config/skills/cn-patent-fasttrack ~/.codex/skills/cn-patent-fasttrack
ln -sfn ~/.gemini/config/skills/cn-patent-fasttrack ~/.agents/skills/cn-patent-fasttrack
```

> 若 `patent-disclosure-skill` 安装在自定义路径，可通过环境变量 `PATENT_DISCLOSURE_SKILL_DIR=/path/to/patent-disclosure-skill` 指定。
