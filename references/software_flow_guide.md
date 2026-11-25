# 软件/算法流程类专利附图与撰写手册

本手册用于指导**算法、信号处理、人工智能、控制系统及计算机软件类**中国发明专利的快速成稿，直接复用 `render_invention_figures.py` 生成符合国知局规范的纯黑白流程图与模块框图，并确保数学公式原生转换为可编辑 Word OMML 公式。

---

## 一、标准图幅规划与 `invention_figures.yaml` 规范

软件/算法类发明专利推荐配置 **3–4 幅标准附图**，无需手写底层 SVG 坐标，直接编写案件目录下的 `invention_figures.yaml`，再调用 `patent_fasttrack.py render-figures --plan invention_figures.yaml` 一键渲染：

1. **图 1（方法主流程图，常设 `abstract: true` 作摘要附图）**：`kind: flowchart`
   - 节点标签统一以 `S101`、`S102`… 开头（如 `S101 获取多通道原始生理电信号`），判断分支用 `shape: diamond`，左右分支通过 `col: -1` / `col: 1` 排布，回跳循环边用 `route: right` 或 `route: left`。
2. **图 2（子算法/模型推理细化流程图，可选）**：`kind: flowchart`
   - 展开独权 1 中最核心创新步骤的内部子步骤（`S201`–`S205`），支撑从属权利要求与创造性答辩。
3. **图 3（装置/系统虚拟模块框图）**：`kind: block_diagram`
   - 外部大框用虚线分组（`columns` / `extras`），内部实线矩形为功能模块（严禁画件号引出线，模块名直接写在框内，与装置独立权利要求的模块名称逐字一致）。

### 可直接复用的 `invention_figures.yaml` 模板

```yaml
figures:
  - fig: 1
    kind: flowchart
    abstract: true
    nodes:
      - {id: S101, label: "S101 获取多通道原始输入数据", row: 0, col: 0, shape: rect}
      - {id: S102, label: "S102 计算通道质量指标并判断是否达标", row: 1, col: 0, shape: diamond}
      - {id: S103, label: "S103 执行全通道特征融合与目标推理", row: 2, col: 0, shape: rect}
      - {id: S104, label: "S104 切换降级子集通道并补偿权重", row: 2, col: 1, shape: rect}
      - {id: S105, label: "S105 输出控制指令或状态估计结果", row: 3, col: 0, shape: rect}
    edges:
      - {from: S101, to: S102}
      - {from: S102, to: S103, label: "是"}
      - {from: S102, to: S104, label: "否"}
      - {from: S103, to: S105}
      - {from: S104, to: S105}

  - fig: 2
    kind: block_diagram
    columns:
      - id: g_front
        label: 数据预处理与质量评估单元
        nodes:
          - {id: m_acq, label: 多通道获取模块}
          - {id: m_qc, label: 信号质量评估模块}
      - id: g_core
        label: 自适应融合推理单元
        nodes:
          - {id: m_fuse, label: 全通道融合模块}
          - {id: m_deg, label: 降级权重补偿模块}
          - {id: m_out, label: 结果输出模块}
    extras:
      - id: g_cfg
        label: 在线阈值更新单元
        above: g_core
        nodes:
          - {id: m_cal, label: 基线自校准模块}
    edges:
      - {from: m_acq, to: m_qc}
      - {from: m_qc, to: m_fuse}
      - {from: m_qc, to: m_deg}
      - {from: m_cal, to: m_fuse}
      - {from: m_fuse, to: m_out}
      - {from: m_deg, to: m_out}
```

> **进阶混合图提示**：若算法专利包含「神经网络张量维度变换图」或「时序滑动窗口/波形对齐示意图」，前两幅用 `invention_figures.yaml` 生成至 `可编辑线稿/` 后，再单独补充一张纯黑白矢量 `图3_时序窗口对齐示意图.svg` 放入 `可编辑线稿/`，最后统一运行 `render-figures`（不带 `--plan`）即可将全部 SVG 合并导出 PNG 与汇总 PDF。

---

## 二、软件/算法专利撰写与客体合规要点（《专利审查指南》第二部分第九章）

1. **避免落入《专利法》第 25 条智力活动规则**：
   - 独权 1 的每个步骤处理对象必须是**具体的物理传感信号、图像帧、网络数据包、工程控制量或存储介质参数**，不得写成脱离技术领域的纯抽象数学运算。
2. **数学公式排版硬性规范（支持 Word 原生 OMML 公式）**：
   - 行内公式**必须**用单美元符包裹，例如 `$w_i = \exp(-\alpha \cdot e_i)$`；
   - 独立行公式**必须**单独成段并用双美元符 `$$...$$` 包裹；
   - **严禁**把含下标/反斜杠的 LaTeX 裸写在圆括号里（如 `(M_{\mathrm{total}})` 或 `(x_i)`），否则 `audit` 会报 `LATEX_DELIM` 错误以防 Word 导出乱码；
   - 每个公式下方紧跟「其中，`$w_i$` 表示第 `$i$` 通道的动态融合权重……」逐项释义全部符号物理量纲与含义。
3. **四组独立权利要求闭环**：
   - **权利要求 1（方法）** → **从属权利要求 2–7（子算法/异常分支/在线校准）** → **独立权利要求 8（与权 1 步骤一一对应的虚拟装置模块）** → **独立权利要求 9（包括处理器和存储器的电子设备）** → **独立权利要求 10（计算机可读存储介质）**。
   - 说明书「具体实施方式」末尾必须包含与独权 8、9、10 对应的装置实施例、电子设备硬件结构段及存储介质段，确保《专利法》第 26 条第 4 款说明书充分支持。
