# 硬件/结构类专利 3D 立体黑白线稿绘制手册

本手册沉淀中国发明与实用新型硬件结构附图的 **3D 参数化矢量生成规范**，解决纯 2D 代码拼图“扁平、死板、缺乏厚度感”以及 AI 生图“标号乱码、灰阶违规”的两难问题。

---

## 一、国知局（CNIPA）黑白结构线稿四条红线

1. **纯黑白零灰阶**：背景纯白（`#ffffff`），线条与填充点阵纯黑（`#000000`）。**严禁**使用彩色、渐变（`linearGradient`）、半透明（`opacity < 1`、`rgba`）或灰色块（`#888` 等）；不同材质一律通过 `<pattern>` 工程剖面线/微孔点阵区分。
2. **图内零解释文字、零图号标题**：结构图中只允许出现阿拉伯数字件号（如 `130`、`151、152`）、剖切符号（`A`—`A`）和子图号 `(a)`、`(b)`、`(c)`。**严禁**在 SVG 内写 `图1`、`工作侧`、`放大图` 等解释字样（`图1` 标题由排版器在图外下方统一生成，子图解释写入说明书「附图说明」段落）。
3. **引出线落点精准、外轮廓留白标号**：每条引出线末端必须准确落在对应零件的实体构造墨线或剖面内部，引出线之间严禁交叉，数字标号统一排布在图形外围留白区并保持水平对齐。
4. **对称性与规避特征零走样**：正视图（如整机对称外框、双侧安装座）必须围绕中心轴数学镜像生成，消除左右高低歪斜或搭接断缝；对竞品专利做了结构规避的部位（例如「无金属螺旋弹簧悬架，改用预压微孔缓冲层」「传感探头端腔体内不设主控 PCB 基板」「收发窗口间设独立遮挡隔筋」），严禁在图中画出被规避的禁忌构件。

---

## 二、为什么以及如何画出“工业设计级 3D 形象感”

避免直接用扁平 `<rect>` 和 `<ellipse>` 拼凑零件，遵循以下四项几何建模技法：

1. **双曲倒角轮廓 + 侧裙厚度脊线（3D Bevel & Side Skirt）**
   - 任何立体壳体、弧面探头、按键或穿戴模组，先定义外轮廓闭合三次贝塞尔路径 `SILHOUETTE_D`（线宽 `1.9–2.2 px`），再偏移一条内部转折脊线 `TURN_RIDGE_D`（线宽 `1.0–1.2 px`），立刻产生真实的双曲弧面与侧壁厚度感。
   - 对多分区边界使用 **Chaikin 圆角细分**（`chaikin_smooth`），消除生硬折角，形成真实模具圆角。
2. **双态物理对比剖面（Unloaded vs. Loaded State）**
   - 涉及弹性缓冲、卡扣锁止、浮动退让或密封压缩的机构，优先在一张图中上下并列 `(b)` 未受载预压态与 `(c)` 受载工作态（用虚线保留原基准轮廓 + 双点划线勾勒外部接触面），让审查员一眼看懂动态机械机理。
3. **30° 等轴测爆炸装配图（3D Isometric Exploded View）**
   - 爆炸图中的每一级零件均赋予前后两层偏移轮廓与侧边连线以呈现真实板厚、沉台止挡台阶、翻边止挡肩与通孔内壁，各层沿平行点划线装配箭头展开。
4. **四分之一立体剖切放大图（3D Quarter-Cutaway Detail）**
   - 对同轴嵌套或腔体复合件（如外环安装座 + 密封绝缘衬套 + 中心换能芯体），用局部 3D 扇形剖切同时展示顶面环绕拓扑与内部台阶压持关系。

---

## 三、可直接复用的标准工程剖面 `<defs>` 与引线函数

在编写案件专属 `scripts/gen_patent_svgs.py` 时，直接拷入以下经过实测验证的纯黑白材质图例与引线辅助函数：

```python
FONT = "Arial, Helvetica, sans-serif"

PATENT_DEFS = """  <defs>
    <!-- 刚性绝缘本体：42° 细剖面线 -->
    <pattern id="p-body" patternUnits="userSpaceOnUse" width="8.5" height="8.5" patternTransform="rotate(42)">
      <rect width="8.5" height="8.5" fill="white"/>
      <line x1="0" y1="0" x2="0" y2="8.5" stroke="black" stroke-width="0.82"/>
    </pattern>
    <!-- 背盖/压环：-42° 反向剖面线（区分相邻配合件） -->
    <pattern id="p-cover" patternUnits="userSpaceOnUse" width="9.5" height="9.5" patternTransform="rotate(-42)">
      <rect width="9.5" height="9.5" fill="white"/>
      <line x1="0" y1="0" x2="0" y2="9.5" stroke="black" stroke-width="0.82"/>
    </pattern>
    <!-- 金属/导电壳体：-54° 密剖面线 -->
    <pattern id="p-metal" patternUnits="userSpaceOnUse" width="6.0" height="6.0" patternTransform="rotate(-54)">
      <rect width="6.0" height="6.0" fill="white"/>
      <line x1="0" y1="0" x2="0" y2="6.0" stroke="black" stroke-width="0.92"/>
    </pattern>
    <!-- 遮光隔筋/密封件：45° 加密剖面线 -->
    <pattern id="p-septum" patternUnits="userSpaceOnUse" width="4.0" height="4.0" patternTransform="rotate(45)">
      <rect width="4.0" height="4.0" fill="white"/>
      <line x1="0" y1="0" x2="0" y2="4.0" stroke="black" stroke-width="1.05"/>
    </pattern>
    <!-- 硅胶/微孔发泡弹性体：纯黑白微孔点阵（切忌画生硬方格网或金属螺旋弹簧） -->
    <pattern id="p-elastomer" patternUnits="userSpaceOnUse" width="13" height="13">
      <rect width="13" height="13" fill="white"/>
      <circle cx="3.0" cy="3.2" r="1.2" fill="none" stroke="black" stroke-width="0.82"/>
      <circle cx="9.6" cy="3.5" r="0.8" fill="black"/>
      <circle cx="6.4" cy="7.8" r="1.3" fill="none" stroke="black" stroke-width="0.82"/>
      <circle cx="2.4" cy="10.6" r="0.75" fill="black"/>
      <circle cx="10.5" cy="10.2" r="1.15" fill="none" stroke="black" stroke-width="0.82"/>
    </pattern>
  </defs>"""


def chaikin_smooth(pts: list[tuple[float, float]], iterations: int = 3) -> list[tuple[float, float]]:
    """多边形圆角细分，生成平滑模具倒角边界。"""
    cur = list(pts)
    for _ in range(iterations):
        nxt = []
        n = len(cur)
        for i in range(n):
            p0, p1 = cur[i], cur[(i + 1) % n]
            nxt.append((0.75 * p0[0] + 0.25 * p1[0], 0.75 * p0[1] + 0.25 * p1[1]))
            nxt.append((0.25 * p0[0] + 0.75 * p1[0], 0.25 * p0[1] + 0.75 * p1[1]))
        cur = nxt
    return cur


def callout(
    tx: float, ty: float,
    kx: float, ky: float,
    px: float, py: float,
    label: str,
    anchor: str = "middle",
    fsize: int = 23,
) -> str:
    """带落点实心小圆点的折线引出标号：(px,py) 为零件墨线落点，(kx,ky) 为折点，(tx,ty) 为外围文字锚点。"""
    lx = tx + 6 if anchor == "start" else (tx - 6 if anchor == "end" else tx)
    ly = ty + 8 if anchor in ("start", "end") else (ty + 22 if ty > ky else ty - 8)
    return (
        f'<polyline points="{px:.1f},{py:.1f} {kx:.1f},{ky:.1f} {tx:.1f},{ty:.1f}" '
        f'fill="none" stroke="black" stroke-width="1.25" stroke-linejoin="round"/>'
        f'<circle cx="{px:.1f}" cy="{py:.1f}" r="2.6" fill="black"/>'
        f'<text x="{lx:.1f}" y="{ly:.1f}" font-family="{FONT}" font-size="{fsize}" '
        f'text-anchor="{anchor}" fill="black">{label}</text>'
    )
```

---

## 四、推荐标准图幅组合（硬件结构发明 5–6 幅）

- **图 1**：整机装配位置正视图 `(a)`（严格左右镜像对称）+ 核心模组工作面立体放大图 `(b)`
- **图 2（常作摘要附图）**：主视图 `(a)`（带 `A—A` 剖切线）+ 未受载预压剖面 `(b)` + 工作受载自适应变形剖面 `(c)`
- **图 3**：局部光路/气路/流体/电路密封横剖 `(a)` + 整机远置主控/多通道选通连接框图 `(b)`
- **图 4**：核心模组 3D 等轴测爆炸分解装配图（清晰展示各级沉台、止挡肩、缓冲层、封盖与装配流向）
- **图 5**：高价值核心子模组（如同轴复合探头、快拆锁紧接头）3D 剖切放大图及替代换能芯剖面（支撑第二独立权利要求及分案）
- **图 6**：多拓扑划分 `(a1–a4)` + 多模态排列组合矩阵 `(b1–b4)` + 替代柔性复合表层变型剖面 `(c)`（把所有能想到的变型一图锁死）
