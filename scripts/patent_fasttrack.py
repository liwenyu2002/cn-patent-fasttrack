#!/usr/bin/env python3
"""Unified zero-exploration CLI for cn-patent-fasttrack.

Subcommands:
  init            Initialize a hardware or software patent case directory and templates.
  render-figures  Lint SVGs for CNIPA pure-B&W compliance and export 300-DPI tight PNGs + vector PDF.
  audit           Verify 3-way numeral/step parity, claims formal rules, abstract length, and LaTeX math.
  build-package   Generate 5 CNIPA split Word docs (with OMML math), full review Word doc, and ZIP bundle.

Always invoke with `uv run --no-project` and pass `--output <json>` to keep stdout concise.
"""
from __future__ import annotations

import argparse
import fcntl
import glob
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any
from urllib import error as urllib_error
from urllib import parse as urllib_parse
from urllib import request as urllib_request

def _locate_base_skill_dir() -> Path:
    env_dir = os.environ.get("PATENT_DISCLOSURE_SKILL_DIR")
    candidates = [
        Path(env_dir).expanduser() if env_dir else None,
        Path.home() / ".codex/skills/patent-disclosure-skill",
        Path.home() / ".gemini/config/skills/patent-disclosure-skill",
        Path.home() / ".agents/skills/patent-disclosure-skill",
        Path.home() / ".claude/skills/patent-disclosure-skill",
    ]
    for cand in candidates:
        if cand and (cand / "skills/patent-application/tools").is_dir():
            return cand
    return Path.home() / ".codex/skills/patent-disclosure-skill"


BASE_SKILL_DIR = _locate_base_skill_dir()
VENV_PY = BASE_SKILL_DIR / ".venv/bin/python"
APP_TOOLS_DIR = BASE_SKILL_DIR / "skills/patent-application/tools"


def _ensure_venv() -> None:
    """Re-exec seamlessly into patent-disclosure-skill's prebuilt .venv if needed."""
    try:
        import docx  # noqa: F401
        import PIL  # noqa: F401
        import playwright  # noqa: F401
        import yaml  # noqa: F401
    except ImportError:
        if VENV_PY.is_file() and os.environ.get("_CN_PATENT_FT_REEXEC") != "1":
            env = dict(os.environ)
            env["_CN_PATENT_FT_REEXEC"] = "1"
            os.execve(str(VENV_PY), [str(VENV_PY), *sys.argv], env)
        print(
            f"Missing Python dependencies (docx/PIL/playwright/yaml) and {VENV_PY} not found.",
            file=sys.stderr,
        )
        sys.exit(1)


_ensure_venv()

import docx  # noqa: E402
import yaml  # noqa: E402
from docx.enum.table import WD_TABLE_ALIGNMENT  # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH  # noqa: E402
from docx.oxml import OxmlElement, parse_xml  # noqa: E402
from docx.oxml.ns import nsdecls, qn  # noqa: E402
from docx.shared import Mm, Pt, RGBColor  # noqa: E402
from PIL import Image, ImageChops  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

if str(APP_TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(APP_TOOLS_DIR))


class RateLimitError(Exception):
    """Raised when an external patent endpoint returns HTTP 429."""


class FileLockRateLimiter:
    """Cross-process file-lock rate limiter using time.monotonic()."""

    def __init__(self, requests_per_second: float = 1.0, lock_name: str = "cn_patent_fasttrack.lock") -> None:
        self.requests_per_second = max(0.05, requests_per_second)
        self.delay = 1.0 / self.requests_per_second
        self.lock_path = Path(tempfile.gettempdir()) / lock_name
        self._last_mono = 0.0

    def get_json(self, base_url: str, path: str, params: dict[str, Any] | None = None, retries: int = 3) -> Any:
        url = f"{base_url.rstrip('/')}/{path.lstrip('/')}"
        if params:
            url = f"{url}?{urllib_parse.urlencode(params)}"
        self.lock_path.touch(exist_ok=True)
        with self.lock_path.open("r+", encoding="utf-8") as lock_f:
            fcntl.flock(lock_f.fileno(), fcntl.LOCK_EX)
            try:
                for attempt in range(retries):
                    elapsed = time.monotonic() - self._last_mono
                    if elapsed < self.delay:
                        time.sleep(self.delay - elapsed)
                    req = urllib_request.Request(url, headers={"Accept": "application/json"})
                    try:
                        with urllib_request.urlopen(req, timeout=30) as resp:
                            self._last_mono = time.monotonic()
                            return json.loads(resp.read().decode("utf-8"))
                    except urllib_error.HTTPError as exc:
                        self._last_mono = time.monotonic()
                        if exc.code == 429:
                            raise RateLimitError(
                                f"HTTP 429 Too Many Requests from {url} "
                                f"(rate limit: {self.requests_per_second} req/s)."
                            ) from exc
                        if exc.code >= 500:
                            wait = 2**attempt
                            print(
                                f"Transient HTTP {exc.code} from {url}, retry {attempt + 1}/{retries} in {wait}s...",
                                file=sys.stderr,
                            )
                            if attempt == retries - 1:
                                raise RuntimeError(
                                    f"Server error {exc.code} from {url} after {retries} retries."
                                ) from exc
                            time.sleep(wait)
                            continue
                        try:
                            body = exc.read().decode("utf-8", errors="replace")[:1000]
                        except OSError:
                            body = str(exc.reason)
                        raise RuntimeError(f"HTTP {exc.code} from {url}: {body}") from exc
            finally:
                fcntl.flock(lock_f.fileno(), fcntl.LOCK_UN)


def write_json_output(data: dict[str, Any], output_path: Path, summary_msg: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{summary_msg} Report written to: {output_path}")


def _find_chromium_executable() -> str | None:
    candidates = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        *sorted(
            glob.glob(
                os.path.expanduser(
                    "~/Library/Caches/ms-playwright/chromium*/chrome-mac*/"
                    "Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing"
                )
            ),
            reverse=True,
        ),
    ]
    return next((c for c in candidates if os.path.exists(c)), None)


def _write_if_needed(path: Path, content: str, *, force: bool) -> str:
    if path.exists() and not force:
        return "skipped_existing"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return "written"


def cmd_init(args: argparse.Namespace) -> int:
    case_dir = Path(args.case_dir).expanduser().resolve()
    mode = args.mode
    title = args.title.strip()
    force = bool(args.force)

    (case_dir / "可编辑线稿").mkdir(parents=True, exist_ok=True)
    (case_dir / "附图PNG").mkdir(parents=True, exist_ok=True)
    (case_dir / "scripts").mkdir(parents=True, exist_ok=True)

    files_status: dict[str, str] = {}

    if mode == "hardware":
        reg_yaml = f"""mode: hardware
title: "{title}"
figures:
  - id: 图1
    title: 整机安装位置及核心组件结构示意图
    file: 可编辑线稿/图1_整机与核心结构.svg
  - id: 图2
    title: 主剖面及未受载与受载工作状态对比示意图
    file: 可编辑线稿/图2_主剖面与双态对比.svg
    is_abstract_figure: true
  - id: 图3
    title: 三维轴测爆炸分解示意图
    file: 可编辑线稿/图3_轴测爆炸分解.svg
  - id: 图4
    title: 多拓扑划分与排列组合变型示意图
    file: 可编辑线稿/图4_拓扑与排列组合.svg

parts:
  - {{id: "100", name: "主体框架"}}
  - {{id: "110", name: "基座托体"}}
  - {{id: "120", name: "功能嵌件"}}
  - {{id: "130", name: "弹性支撑层"}}
  - {{id: "140", name: "背盖"}}
"""
    else:
        reg_yaml = f"""mode: software
title: "{title}"
figures:
  - id: 图1
    title: 方法主流程示意图
    file: 可编辑线稿/图1_方法主流程.svg
    is_abstract_figure: true
  - id: 图2
    title: 装置功能模块框图
    file: 可编辑线稿/图2_装置模块框图.svg

steps:
  - {{id: "S101", name: "获取多通道原始输入信号"}}
  - {{id: "S102", name: "评估通道质量并判定是否满足融合条件"}}
  - {{id: "S103", name: "执行全通道自适应特征融合推理"}}
  - {{id: "S104", name: "切换降级子集通道并执行权重补偿"}}
  - {{id: "S105", name: "输出目标控制指令或状态估计结果"}}

parts: []
"""
        inv_figs_yaml = """figures:
  - fig: 1
    kind: flowchart
    abstract: true
    nodes:
      - {id: S101, label: "S101 获取多通道原始输入信号", row: 0, col: 0, shape: rect}
      - {id: S102, label: "S102 评估通道质量是否达标", row: 1, col: 0, shape: diamond}
      - {id: S103, label: "S103 全通道自适应特征融合", row: 2, col: 0, shape: rect}
      - {id: S104, label: "S104 切换降级通道并补偿权重", row: 2, col: 1, shape: rect}
      - {id: S105, label: "S105 输出状态估计或控制指令", row: 3, col: 0, shape: rect}
    edges:
      - {from: S101, to: S102}
      - {from: S102, to: S103, label: "是"}
      - {from: S102, to: S104, label: "否"}
      - {from: S103, to: S105}
      - {from: S104, to: S105}

  - fig: 2
    kind: block_diagram
    columns:
      - id: g_in
        label: 信号采集与质量评估单元
        nodes:
          - {id: m_acq, label: 多通道获取模块}
          - {id: m_qc, label: 通道质量评估模块}
      - id: g_proc
        label: 自适应融合与输出单元
        nodes:
          - {id: m_fuse, label: 全通道融合模块}
          - {id: m_deg, label: 降级权重补偿模块}
          - {id: m_out, label: 结果输出模块}
    edges:
      - {from: m_acq, to: m_qc}
      - {from: m_qc, to: m_fuse}
      - {from: m_qc, to: m_deg}
      - {from: m_fuse, to: m_out}
      - {from: m_deg, to: m_out}
"""
        files_status["invention_figures.yaml"] = _write_if_needed(
            case_dir / "invention_figures.yaml", inv_figs_yaml, force=force
        )

    files_status["件号登记表.yaml"] = _write_if_needed(case_dir / "件号登记表.yaml", reg_yaml, force=force)
    files_status["权利要求书.md"] = _write_if_needed(
        case_dir / "权利要求书.md", f"# 权利要求书\n\n1. 一种{title}，其特征在于，包括：\n", force=force
    )
    files_status["说明书.md"] = _write_if_needed(
        case_dir / "说明书.md",
        f"# {title}\n\n## 技术领域\n\n[0001] 本发明涉及……技术领域。\n\n## 背景技术\n\n[0002] ……\n\n"
        f"## 发明内容\n\n[0003] ……\n\n## 附图说明\n\n[0004] 图1示出……\n\n## 具体实施方式\n\n[0005] ……\n",
        force=force,
    )
    files_status["说明书摘要.md"] = _write_if_needed(
        case_dir / "说明书摘要.md", f"# 说明书摘要\n\n本发明公开了一种{title}……\n", force=force
    )
    files_status["复核记录与规避分析.md"] = _write_if_needed(
        case_dir / "复核记录与规避分析.md",
        f"# {title} · 竞品规避分析与定稿复核记录\n\n| 竞品专利 | 核心限定 | 本案区别与规避设计 | 待跟踪事项 |\n| --- | --- | --- | --- |\n",
        force=force,
    )
    files_status["问题清单.md"] = _write_if_needed(
        case_dir / "问题清单.md",
        "# 问题清单（供发明人与专利代理师会稿确认，不入正式申请文件）\n\n1. 申请人与发明人著录信息待补。\n",
        force=force,
    )

    report = {
        "ok": True,
        "command": "init",
        "mode": mode,
        "title": title,
        "case_dir": str(case_dir),
        "files": files_status,
    }
    write_json_output(report, Path(args.output).expanduser().resolve(), f"Success! Initialized {mode} case at {case_dir}.")
    return 0


def lint_svg_cnipa_bw(svg_path: Path) -> list[dict[str, str]]:
    """Static check for CNIPA pure black-and-white patent figure rules."""
    txt = svg_path.read_text(encoding="utf-8")
    issues: list[dict[str, str]] = []

    for m in re.finditer(r'\b(?:fill-|stroke-)?opacity\s*=\s*["\']([^"\']+)["\']', txt, re.I):
        try:
            if float(m.group(1)) < 0.999:
                issues.append({"level": "ERROR", "code": "SEMI_TRANSPARENT", "detail": m.group(0)})
        except ValueError:
            pass

    if re.search(r"<(?:linear|radial)Gradient\b", txt, re.I) or "rgba(" in txt.lower():
        issues.append({"level": "ERROR", "code": "GRADIENT_OR_RGBA", "detail": "Gradient or rgba() prohibited in CNIPA line art."})

    allowed_hex = {"#000", "#000000", "#fff", "#ffffff"}
    bad_hex = sorted({h.lower() for h in re.findall(r"#[0-9a-fA-F]{3,8}\b", txt) if h.lower() not in allowed_hex})
    if bad_hex:
        issues.append({"level": "ERROR", "code": "NON_BW_HEX_COLOR", "detail": f"Disallowed hex colors: {bad_hex}"})

    for m in re.finditer(r"<text\b[^>]*>(.*?)</text>", txt, re.S | re.I):
        inner = re.sub(r"<[^>]+>", "", m.group(1)).strip()
        if re.match(r"^图\s*\d+\s*$", inner):
            issues.append(
                {
                    "level": "ERROR",
                    "code": "IN_SVG_FIG_TITLE",
                    "detail": f"Remove '{inner}' inside SVG; figure captions belong outside the drawing.",
                }
            )
    return issues


def _tight_crop_png(png_path: Path, pad: int = 24) -> tuple[int, int]:
    im = Image.open(png_path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox:
        cropped = im.crop(
            (
                max(0, bbox[0] - pad),
                max(0, bbox[1] - pad),
                min(im.width, bbox[2] + pad),
                min(im.height, bbox[3] + pad),
            )
        )
        cropped.save(png_path)
        return cropped.size
    return im.size


def _ordered_svgs(case_dir: Path, reg: dict[str, Any]) -> list[Path]:
    svg_dir = case_dir / "可编辑线稿"
    ordered: list[Path] = []
    for item in reg.get("figures") or []:
        if isinstance(item, dict) and item.get("file"):
            p = case_dir / str(item["file"])
            if p.is_file():
                ordered.append(p)
    if ordered:
        return ordered
    return sorted(svg_dir.glob("图*.svg"))


def cmd_render_figures(args: argparse.Namespace) -> int:
    case_dir = Path(args.case_dir).expanduser().resolve()
    svg_dir = case_dir / "可编辑线稿"
    png_dir = case_dir / "附图PNG"
    svg_dir.mkdir(parents=True, exist_ok=True)
    png_dir.mkdir(parents=True, exist_ok=True)

    reg_path = case_dir / "件号登记表.yaml"
    reg = yaml.safe_load(reg_path.read_text(encoding="utf-8")) if reg_path.is_file() else {}
    reg = reg or {}

    plan_arg = Path(args.plan).expanduser().resolve() if args.plan else None
    if plan_arg is None and (case_dir / "invention_figures.yaml").is_file() and not list(svg_dir.glob("图*.svg")):
        plan_arg = case_dir / "invention_figures.yaml"

    plan_generated: list[str] = []
    if plan_arg and plan_arg.is_file():
        import render_invention_figures as rif

        plan_data = yaml.safe_load(plan_arg.read_text(encoding="utf-8")) or {}
        with tempfile.TemporaryDirectory() as tmp_d:
            tmp_out = Path(tmp_d)
            rif.render_plan(plan_data, tmp_out, plan_dir=plan_arg.parent)
            fig_map: dict[int, Path] = {}
            for idx, item in enumerate(reg.get("figures") or [], 1):
                if isinstance(item, dict) and item.get("file"):
                    m_num = re.search(r"(\d+)", str(item.get("id") or idx))
                    f_no = int(m_num.group(1)) if m_num else idx
                    fig_map[f_no] = case_dir / str(item["file"])
            for gen_svg in sorted(tmp_out.glob("图*.svg")):
                m_no = re.search(r"(\d+)", gen_svg.stem)
                f_no = int(m_no.group(1)) if m_no else 1
                dest_svg = fig_map.get(f_no, svg_dir / gen_svg.name)
                dest_svg.parent.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(gen_svg, dest_svg)
                plan_generated.append(str(dest_svg))

    svgs = _ordered_svgs(case_dir, reg)
    if not svgs:
        report = {"ok": False, "error": f"No SVG files found in {svg_dir}"}
        write_json_output(report, Path(args.output).expanduser().resolve(), "Failed: no SVGs to render.")
        return 1

    lint_results: dict[str, list[dict[str, str]]] = {}
    lint_errors = 0
    for svg_p in svgs:
        issues = lint_svg_cnipa_bw(svg_p)
        if issues:
            lint_results[svg_p.name] = issues
            lint_errors += sum(1 for x in issues if x["level"] == "ERROR")

    exe = _find_chromium_executable()
    scale = max(2.4, float(args.dpi) / 96.0)
    rendered_pngs: list[dict[str, Any]] = []
    pages_html: list[str] = []

    with sync_playwright() as p:
        launch_kwargs: dict[str, Any] = {"headless": True}
        if exe:
            launch_kwargs["executable_path"] = exe
        browser = p.chromium.launch(**launch_kwargs)
        page = browser.new_page(device_scale_factor=scale)

        for idx, svg_p in enumerate(svgs, 1):
            raw_svg = svg_p.read_text(encoding="utf-8")
            m_vb = re.search(
                r'viewBox\s*=\s*["\']\s*(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s*["\']',
                raw_svg,
            )
            if m_vb:
                vb_w = max(1.0, float(m_vb.group(3)))
                vb_h = max(1.0, float(m_vb.group(4)))
                vp_w = 1000
                vp_h = max(200, int(round(1000.0 * vb_h / vb_w)))
            else:
                vp_w, vp_h = 1000, 1250

            # Strip fixed mm width/height on root <svg> during rasterization so CSS 1000px width governs cleanly
            clean_svg = re.sub(
                r"(<svg\b[^>]*?)\s+width\s*=\s*[\"'][^\"']+[\"']",
                r"\1",
                raw_svg,
                count=1,
            )
            clean_svg = re.sub(
                r"(<svg\b[^>]*?)\s+height\s*=\s*[\"'][^\"']+[\"']",
                r"\1",
                clean_svg,
                count=1,
            )

            page.set_viewport_size({"width": vp_w + 80, "height": vp_h + 80})
            html_doc = (
                "<!doctype html><html><head><meta charset='utf-8'>"
                f"<style>html,body{{margin:0;padding:20px;background:#fff;}}"
                f"svg{{display:block;width:{vp_w}px;height:auto;}}</style>"
                f"</head><body>{clean_svg}</body></html>"
            )
            page.set_content(html_doc, wait_until="load")
            out_png = png_dir / f"{svg_p.stem}.png"
            page.screenshot(path=str(out_png), full_page=True)
            w_px, h_px = _tight_crop_png(out_png, pad=int(args.pad_px))
            rendered_pngs.append({"svg": svg_p.name, "png": str(out_png), "width_px": w_px, "height_px": h_px})
            pages_html.append(f'<div class="sheet"><div class="figbox">{clean_svg}</div><div class="cap">图{idx}</div></div>')

        pdf_name = args.pdf_name or f"{reg.get('title', case_dir.name)}_附图线稿汇总.pdf"
        pdf_path = case_dir / pdf_name
        pdf_html = (
            "<!doctype html><html><head><meta charset='utf-8'><style>"
            "@page { size: A4 portrait; margin: 20mm 18mm 18mm 20mm; }"
            "body { margin: 0; padding: 0; background: white; font-family: 'Songti SC', 'STSong', serif; }"
            ".sheet { page-break-after: always; display: flex; flex-direction: column; "
            "align-items: center; justify-content: center; height: 255mm; }"
            ".sheet:last-child { page-break-after: auto; }"
            ".figbox { display: flex; align-items: center; justify-content: center; max-height: 232mm; width: 100%; }"
            ".figbox svg { max-width: 165mm; max-height: 230mm; width: 100%; height: auto; }"
            ".cap { margin-top: 5mm; font-size: 13pt; color: #000; text-align: center; }"
            f"</style></head><body>{''.join(pages_html)}</body></html>"
        )
        page.set_content(pdf_html, wait_until="load")
        page.pdf(path=str(pdf_path), format="A4", print_background=True)
        browser.close()

    ok = lint_errors == 0
    report = {
        "ok": ok,
        "command": "render-figures",
        "case_dir": str(case_dir),
        "plan_generated_svgs": plan_generated,
        "rendered_pngs": rendered_pngs,
        "summary_pdf": str(pdf_path),
        "bw_lint_errors": lint_errors,
        "bw_lint_details": lint_results,
    }
    write_json_output(
        report,
        Path(args.output).expanduser().resolve(),
        f"{'Success!' if ok else 'Rendered with B&W lint errors!'} {len(rendered_pngs)} PNGs + PDF generated.",
    )
    return 0 if ok else 1


def cmd_audit(args: argparse.Namespace) -> int:
    import audit_claims
    from latex_delimiters import find_bare_paren_latex

    case_dir = Path(args.case_dir).expanduser().resolve()
    reg_path = case_dir / "件号登记表.yaml"
    claims_path = case_dir / "权利要求书.md"
    spec_path = case_dir / "说明书.md"
    abs_path = case_dir / "说明书摘要.md"

    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []

    if not reg_path.is_file():
        errors.append({"section": "register", "code": "MISSING_REGISTER", "message": "件号登记表.yaml 不存在"})
        reg: dict[str, Any] = {}
    else:
        reg = yaml.safe_load(reg_path.read_text(encoding="utf-8")) or {}

    svgs = _ordered_svgs(case_dir, reg)
    claims_txt = claims_path.read_text(encoding="utf-8") if claims_path.is_file() else ""
    spec_txt = spec_path.read_text(encoding="utf-8") if spec_path.is_file() else ""
    abs_txt = abs_path.read_text(encoding="utf-8") if abs_path.is_file() else ""
    combined_txt = f"{claims_txt}\n{spec_txt}"

    # 1. Numeral parity (for hardware or hybrid cases with parts)
    parts_list = reg.get("parts") or []
    reg_part_ids = {str(p["id"]).strip() for p in parts_list if isinstance(p, dict) and p.get("id")}
    svg_part_ids: set[str] = set()
    for svg_p in svgs:
        s_txt = svg_p.read_text(encoding="utf-8")
        for m in re.finditer(r"<text\b[^>]*>(.*?)</text>", s_txt, re.S | re.I):
            inner = re.sub(r"<[^>]+>", "", m.group(1)).strip()
            if re.match(r"^\d{2,4}[A-Za-z]?(?:[、,，/]\d{2,4}[A-Za-z]?)*$", inner):
                for tok in re.split(r"[、,，/]", inner):
                    svg_part_ids.add(tok.strip())
            else:
                for pid in reg_part_ids:
                    if re.search(rf"(?<![0-9A-Za-z]){re.escape(pid)}(?![0-9A-Za-z])", inner):
                        svg_part_ids.add(pid)

    missing_in_svgs = sorted(reg_part_ids - svg_part_ids)
    unregistered_in_svgs = sorted(svg_part_ids - reg_part_ids) if reg_part_ids else []
    missing_in_text = sorted(
        pid
        for pid in reg_part_ids
        if not re.search(rf"(?<![0-9A-Za-z]){re.escape(pid)}(?![0-9A-Za-z])", combined_txt)
    )

    if missing_in_svgs:
        errors.append({"section": "numerals", "code": "MISSING_IN_SVGS", "items": missing_in_svgs})
    if unregistered_in_svgs:
        errors.append({"section": "numerals", "code": "UNREGISTERED_IN_SVGS", "items": unregistered_in_svgs})
    if missing_in_text:
        errors.append({"section": "numerals", "code": "MISSING_IN_SPEC_OR_CLAIMS", "items": missing_in_text})

    # 1b. Software step parity (if steps registered)
    steps_list = reg.get("steps") or []
    reg_step_ids = {str(s["id"]).strip() for s in steps_list if isinstance(s, dict) and s.get("id")}
    all_svg_raw = "\n".join(p.read_text(encoding="utf-8") for p in svgs)
    missing_steps_svg = sorted(sid for sid in reg_step_ids if sid not in all_svg_raw)
    missing_steps_spec = sorted(sid for sid in reg_step_ids if sid not in spec_txt)
    if missing_steps_svg:
        errors.append({"section": "steps", "code": "STEPS_MISSING_IN_SVGS", "items": missing_steps_svg})
    if missing_steps_spec:
        errors.append({"section": "steps", "code": "STEPS_MISSING_IN_SPEC", "items": missing_steps_spec})

    # 2. Claims formal check
    claim_findings_raw = audit_claims.audit(claims_txt) if claims_txt else []
    c_err, c_warn = audit_claims.summarize(claim_findings_raw)
    for f in claim_findings_raw:
        entry = {"section": "claims", "claim": f.claim, "code": f.code, "message": f.message}
        if f.level == "ERROR":
            errors.append(entry)
        else:
            warnings.append(entry)
    if not claims_txt:
        errors.append({"section": "claims", "code": "MISSING_CLAIMS", "message": "权利要求书.md 为空或不存在"})

    # 3. Abstract length & abstract figure check
    abs_lines = [
        ln.strip()
        for ln in abs_txt.splitlines()
        if ln.strip() and not ln.strip().startswith("#") and not ln.strip().startswith("摘要附图")
    ]
    abs_clean = re.sub(r"\s+", "", "".join(abs_lines))
    abs_len = len(abs_clean)
    if abs_len == 0:
        errors.append({"section": "abstract", "code": "EMPTY_ABSTRACT", "message": "说明书摘要正文为空"})
    elif abs_len > 300:
        errors.append(
            {
                "section": "abstract",
                "code": "ABSTRACT_TOO_LONG",
                "message": f"说明书摘要共 {abs_len} 字（含标点），超过国知局 300 字上限",
            }
        )

    abs_fig_png = _resolve_abstract_png(case_dir, reg)
    if abs_fig_png is None or not abs_fig_png.is_file():
        warnings.append(
            {
                "section": "abstract",
                "code": "MISSING_ABSTRACT_PNG",
                "message": "尚未找到摘要附图对应的 PNG（请先运行 render-figures）",
            }
        )

    # 4. Bare-parenthesis LaTeX delimiter check
    for md_name, md_content in (("权利要求书.md", claims_txt), ("说明书.md", spec_txt)):
        hits = find_bare_paren_latex(md_content)
        if hits:
            errors.append(
                {
                    "section": "latex_delimiters",
                    "file": md_name,
                    "code": "BARE_PAREN_LATEX",
                    "count": len(hits),
                    "message": "发现裸圆括号公式，请改用 $...$ 包裹以生成 Word OMML 公式",
                }
            )

    ok = len(errors) == 0
    report = {
        "ok": ok,
        "command": "audit",
        "case_dir": str(case_dir),
        "summary": {
            "registered_parts": len(reg_part_ids),
            "registered_steps": len(reg_step_ids),
            "claim_errors": c_err,
            "claim_warnings": c_warn,
            "abstract_chars": abs_len,
            "total_errors": len(errors),
            "total_warnings": len(warnings),
        },
        "numerals": {
            "missing_in_svgs": missing_in_svgs,
            "unregistered_in_svgs": unregistered_in_svgs,
            "missing_in_text": missing_in_text,
        },
        "errors": errors,
        "warnings": warnings,
    }
    write_json_output(
        report,
        Path(args.output).expanduser().resolve(),
        f"Audit {'PASSED' if ok else 'FAILED'}: {len(errors)} errors, {len(warnings)} warnings.",
    )
    return 0 if ok else 1


def _resolve_abstract_png(case_dir: Path, reg: dict[str, Any]) -> Path | None:
    png_dir = case_dir / "附图PNG"
    figs = [f for f in (reg.get("figures") or []) if isinstance(f, dict)]
    chosen = next((f for f in figs if f.get("is_abstract_figure")), figs[0] if figs else None)
    if chosen and chosen.get("file"):
        stem = Path(str(chosen["file"])).stem
        cand = png_dir / f"{stem}.png"
        if cand.is_file():
            return cand
    pngs = sorted(png_dir.glob("图*.png"))
    return pngs[0] if pngs else None


BODY_FONT = "宋体"
HEAD_FONT = "黑体"


def _apply_rfonts(rPr: Any, east: str = BODY_FONT, ascii_f: str = "Times New Roman") -> None:
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), east)
    rFonts.set(qn("w:ascii"), ascii_f)
    rFonts.set(qn("w:hAnsi"), ascii_f)
    rFonts.set(qn("w:cs"), east)
    rFonts.set(qn("w:hint"), "eastAsia")
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "zh-CN")


def _setup_cnipa_doc() -> Any:
    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_width = Mm(210)
    sec.page_height = Mm(297)
    sec.top_margin = Mm(25)
    sec.bottom_margin = Mm(20)
    sec.left_margin = Mm(25)
    sec.right_margin = Mm(20)
    style = doc.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = Pt(10.5)
    style.font.color.rgb = RGBColor(0, 0, 0)
    _apply_rfonts(style._element.get_or_add_rPr(), east=BODY_FONT)
    style.paragraph_format.line_spacing = 1.45
    style.paragraph_format.space_after = Pt(4)
    return doc


def _add_styled_run(p: Any, text: str, *, bold: bool = False, size: float = 10.5, font_east: str = BODY_FONT) -> Any:
    r = p.add_run(text)
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = font_east
    r.font.color.rgb = RGBColor(0, 0, 0)
    _apply_rfonts(r._element.get_or_add_rPr(), east=font_east)
    return r


def _set_table_borders(tbl: Any) -> None:
    tblPr = tbl._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def _add_rich_para(doc: Any, text: str, *, indent_first: bool = True, space_after: float = 4, size: float = 10.5) -> Any:
    p = doc.add_paragraph()
    if indent_first:
        p.paragraph_format.first_line_indent = Pt(21)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.45
    for part in re.split(r"(\*\*.*?\*\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            _add_styled_run(p, part[2:-2], bold=True, size=size)
        else:
            _add_styled_run(p, part, bold=False, size=size)
    return p


def _add_heading(doc: Any, text: str, level: int = 1) -> Any:
    p = doc.add_paragraph()
    if level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        _add_styled_run(p, text, bold=True, size=15, font_east=HEAD_FONT)
    elif level == 1:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(6)
        _add_styled_run(p, text, bold=True, size=12, font_east=HEAD_FONT)
    else:
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        _add_styled_run(p, text, bold=True, size=11, font_east=HEAD_FONT)
    return p


def _append_md_to_doc(doc: Any, md_text: str, *, base_dir: Path, skip_h1: bool = False, indent_body: bool = True) -> None:
    """Render Markdown into doc; delegates math-bearing docs to md_to_docx OMML engine."""
    if "$" in md_text:
        from copy import deepcopy
        from md_to_docx import convert_md_to_docx

        cleaned_lines = [
            ln for ln in md_text.splitlines() if not (skip_h1 and ln.startswith("# "))
        ]
        sub_doc = convert_md_to_docx("\n".join(cleaned_lines), base_dir=base_dir, prefer_omml=True)
        for child in list(sub_doc.element.body):
            if child.tag.endswith("sectPr"):
                continue
            doc.element.body.insert(-1, deepcopy(child))
        return

    lines = md_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1
            continue
        if line.startswith("# "):
            if not skip_h1:
                _add_heading(doc, line[2:].strip(), level=0)
            i += 1
        elif line.startswith("## "):
            _add_heading(doc, line[3:].strip(), level=1)
            i += 1
        elif line.startswith("### "):
            _add_heading(doc, line[4:].strip(), level=2)
            i += 1
        elif line.strip().startswith("|"):
            tbl_lines: list[str] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                tbl_lines.append(lines[i].strip())
                i += 1
            rows: list[list[str]] = []
            for tl in tbl_lines:
                cells = [c.strip() for c in tl.strip("|").split("|")]
                if all(re.match(r"^[-:]+$", c) for c in cells):
                    continue
                rows.append(cells)
            if rows:
                tbl = doc.add_table(rows=len(rows), cols=len(rows[0]))
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                _set_table_borders(tbl)
                for r_i, row in enumerate(rows):
                    for c_i, val in enumerate(row):
                        if c_i >= len(rows[0]):
                            break
                        cell = tbl.cell(r_i, c_i)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)
                        p.paragraph_format.line_spacing = 1.25
                        _add_styled_run(p, val, bold=(r_i == 0), size=9.0)
        elif line.startswith("- "):
            _add_rich_para(doc, "• " + line[2:].strip(), indent_first=True, space_after=3)
            i += 1
        else:
            _add_rich_para(doc, line.strip(), indent_first=indent_body, space_after=4)
            i += 1


def _add_figure_page(
    doc: Any,
    png_path: Path,
    *,
    caption_text: str | None = None,
    max_w_mm: float = 156.0,
    max_h_mm: float = 218.0,
    page_break_before: bool = False,
) -> None:
    if page_break_before:
        doc.add_page_break()
    with Image.open(png_path) as im:
        w, h = im.size
    ratio = w / max(1, h)
    w_mm = max_w_mm
    h_mm = w_mm / ratio
    if h_mm > max_h_mm:
        h_mm = max_h_mm
        w_mm = h_mm * ratio
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.add_run().add_picture(str(png_path), width=Mm(w_mm))
    if caption_text:
        cp = doc.add_paragraph()
        cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cp.paragraph_format.space_before = Pt(4)
        cp.paragraph_format.space_after = Pt(0)
        _add_styled_run(cp, caption_text, bold=False, size=11)


def cmd_build_package(args: argparse.Namespace) -> int:
    case_dir = Path(args.case_dir).expanduser().resolve()
    prefix = args.title.strip()
    reg_path = case_dir / "件号登记表.yaml"
    reg = yaml.safe_load(reg_path.read_text(encoding="utf-8")) if reg_path.is_file() else {}
    reg = reg or {}

    claims_md = (case_dir / "权利要求书.md").read_text(encoding="utf-8")
    spec_md = (case_dir / "说明书.md").read_text(encoding="utf-8")
    abs_md = (case_dir / "说明书摘要.md").read_text(encoding="utf-8")

    spec_h1_match = re.search(r"^#\s+(.+)$", spec_md, re.M)
    full_patent_title = spec_h1_match.group(1).strip() if spec_h1_match else str(reg.get("title") or prefix)

    abs_body_lines = [
        ln.strip()
        for ln in abs_md.splitlines()
        if ln.strip() and not ln.strip().startswith("#") and not ln.strip().startswith("摘要附图")
    ]

    png_dir = case_dir / "附图PNG"
    ordered_pngs: list[Path] = []
    abs_fig_idx = 1
    for idx, item in enumerate(reg.get("figures") or [], 1):
        if isinstance(item, dict) and item.get("file"):
            cand = png_dir / f"{Path(str(item['file'])).stem}.png"
            if cand.is_file():
                ordered_pngs.append(cand)
            if item.get("is_abstract_figure"):
                abs_fig_idx = idx
    if not ordered_pngs:
        ordered_pngs = sorted(png_dir.glob("图*.png"))

    if not ordered_pngs:
        report = {"ok": False, "error": "附图PNG 目录中无可用 PNG，请先执行 render-figures"}
        write_json_output(report, Path(args.output).expanduser().resolve(), "Failed: missing rendered PNGs.")
        return 1

    abs_png = _resolve_abstract_png(case_dir, reg) or ordered_pngs[0]

    # 1. 权利要求书.docx
    d_claims = _setup_cnipa_doc()
    _append_md_to_doc(d_claims, claims_md, base_dir=case_dir, skip_h1=True, indent_body=True)
    p_claims = case_dir / "权利要求书.docx"
    d_claims.save(str(p_claims))

    # 2. 说明书.docx
    d_spec = _setup_cnipa_doc()
    _append_md_to_doc(d_spec, spec_md, base_dir=case_dir, skip_h1=False, indent_body=True)
    p_spec = case_dir / "说明书.docx"
    d_spec.save(str(p_spec))

    # 3. 说明书摘要.docx
    d_abs = _setup_cnipa_doc()
    _append_md_to_doc(d_abs, "\n\n".join(abs_body_lines), base_dir=case_dir, skip_h1=True, indent_body=True)
    p_abs = case_dir / "说明书摘要.docx"
    d_abs.save(str(p_abs))

    # 4. 说明书附图.docx
    d_figs = _setup_cnipa_doc()
    for idx, png_p in enumerate(ordered_pngs, 1):
        _add_figure_page(d_figs, png_p, caption_text=f"图{idx}", page_break_before=(idx > 1))
    p_figs = case_dir / "说明书附图.docx"
    d_figs.save(str(p_figs))

    # 5. 摘要附图.docx (CNIPA rule: image only, NO '图X' caption!)
    d_absfig = _setup_cnipa_doc()
    _add_figure_page(d_absfig, abs_png, caption_text=None, page_break_before=False)
    p_absfig = case_dir / "摘要附图.docx"
    d_absfig.save(str(p_absfig))

    # 6. 合稿审阅 Word
    d_all = _setup_cnipa_doc()
    _add_heading(d_all, full_patent_title, level=0)
    _add_heading(d_all, "说明书摘要", level=1)
    _append_md_to_doc(d_all, "\n\n".join(abs_body_lines), base_dir=case_dir, skip_h1=True, indent_body=True)
    _add_rich_para(d_all, f"摘要附图为图{abs_fig_idx}。", indent_first=True)
    d_all.add_page_break()
    _add_heading(d_all, "权利要求书", level=1)
    _append_md_to_doc(d_all, claims_md, base_dir=case_dir, skip_h1=True, indent_body=True)
    d_all.add_page_break()
    _add_heading(d_all, "说明书", level=1)
    _append_md_to_doc(d_all, spec_md, base_dir=case_dir, skip_h1=True, indent_body=True)
    for idx, png_p in enumerate(ordered_pngs, 1):
        _add_figure_page(d_all, png_p, caption_text=f"图{idx}", page_break_before=True)
    p_all = case_dir / f"{prefix}_发明专利申请全稿.docx"
    d_all.save(str(p_all))

    # 7. ZIP 完整资料包
    zip_path = case_dir / f"{prefix}_完整申请资料包.zip"
    svgs = _ordered_svgs(case_dir, reg)
    pdfs = sorted(case_dir.glob("*.pdf"))
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for doc_file in [p_claims, p_spec, p_abs, p_figs, p_absfig, p_all]:
            zf.write(doc_file, doc_file.name)
        for pdf_file in pdfs:
            zf.write(pdf_file, pdf_file.name)
        for aux_name in ("复核记录与规避分析.md", "问题清单.md", "件号登记表.yaml", "invention_figures.yaml"):
            aux_p = case_dir / aux_name
            if aux_p.is_file():
                zf.write(aux_p, aux_name)
        for svg_p in svgs:
            zf.write(svg_p, f"可编辑线稿/{svg_p.name}")
        for png_p in ordered_pngs:
            zf.write(png_p, f"附图PNG/{png_p.name}")

    written = [str(p_claims), str(p_spec), str(p_abs), str(p_figs), str(p_absfig), str(p_all), str(zip_path)]
    report = {
        "ok": True,
        "command": "build-package",
        "case_dir": str(case_dir),
        "patent_title": full_patent_title,
        "abstract_figure": abs_png.name,
        "figure_count": len(ordered_pngs),
        "deliverables": written,
    }
    write_json_output(
        report,
        Path(args.output).expanduser().resolve(),
        f"Success! Built 5 CNIPA split Word files + full draft + ZIP ({len(ordered_pngs)} figures).",
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="CN Patent Fast-Track CLI (init / render-figures / audit / build-package)")
    sub = parser.add_subparsers(dest="command", required=True)

    p_init = sub.add_parser("init", help="Initialize patent case directory and YAML/Markdown templates")
    p_init.add_argument("--case-dir", required=True, help="Target case directory")
    p_init.add_argument("--mode", required=True, choices=["hardware", "software"], help="Patent branch mode")
    p_init.add_argument("--title", required=True, help="Patent working title")
    p_init.add_argument("--force", action="store_true", help="Overwrite existing starter files")
    p_init.add_argument("--output", required=True, help="Output JSON report path")

    p_ren = sub.add_parser("render-figures", help="Lint SVGs and render tight 300-DPI B&W PNGs + A4 PDF")
    p_ren.add_argument("--case-dir", required=True, help="Target case directory")
    p_ren.add_argument("--plan", default=None, help="Optional invention_figures.yaml path for software diagrams")
    p_ren.add_argument("--dpi", type=int, default=300, help="Target rasterization DPI (default: 300)")
    p_ren.add_argument("--pad-px", type=int, default=24, help="Whitespace padding around tight-cropped PNG (default: 24)")
    p_ren.add_argument("--pdf-name", default=None, help="Optional summary PDF filename")
    p_ren.add_argument("--output", required=True, help="Output JSON report path")

    p_aud = sub.add_parser("audit", help="Run 3-way numeral/step parity, claims check, abstract length, and LaTeX check")
    p_aud.add_argument("--case-dir", required=True, help="Target case directory")
    p_aud.add_argument("--output", required=True, help="Output JSON report path")

    p_pkg = sub.add_parser("build-package", help="Build 5 CNIPA Word files, full review Word file, and ZIP bundle")
    p_pkg.add_argument("--case-dir", required=True, help="Target case directory")
    p_pkg.add_argument("--title", required=True, help="Filename prefix for full Word and ZIP deliverables")
    p_pkg.add_argument("--output", required=True, help="Output JSON report path")

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.command == "init":
        return cmd_init(args)
    if args.command == "render-figures":
        return cmd_render_figures(args)
    if args.command == "audit":
        return cmd_audit(args)
    if args.command == "build-package":
        return cmd_build_package(args)
    print(f"Unknown command: {args.command}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
