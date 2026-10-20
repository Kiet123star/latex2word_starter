"""Native Microsoft Word DrawingML Shape & Graph Compiler (TikZ / PGFPlots -> <wpg:wgp>).

Converts LaTeX TikZ geometry diagrams (\begin{tikzpicture}) and PGFPlots function graphs
(\begin{axis}) directly into editable Office Open XML WordprocessingGroup (<wpg:wgp>)
and WordprocessingShape (<wps:wsp>) vector primitives with native OMML math labels.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from docx.oxml import parse_xml

from .omml_math import latex_to_omml_xml_string

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

EMU_PER_CM = 360000
EMU_PER_PT = 12700

COLOR_PALETTE: dict[str, tuple[int, int, int]] = {
    "black": (15, 23, 42),
    "white": (255, 255, 255),
    "red": (220, 38, 38),
    "blue": (37, 99, 235),
    "green": (22, 163, 74),
    "teal": (13, 148, 136),
    "cyan": (6, 182, 212),
    "magenta": (219, 39, 119),
    "yellow": (234, 179, 8),
    "orange": (234, 88, 12),
    "purple": (147, 51, 234),
    "violet": (124, 58, 237),
    "brown": (146, 64, 14),
    "gray": (100, 116, 139),
    "grey": (100, 116, 139),
    "darkgray": (51, 65, 85),
    "lightgray": (203, 213, 225),
}


def resolve_tikz_color(color_spec: str, default_hex: str = "0F172A") -> str:
    """Resolve TikZ color expressions like 'blue', 'gray!20', 'cyan!10', 'blue!70!black' to 6-char hex."""
    if not color_spec:
        return default_hex
    spec = color_spec.strip().lower()
    if spec.startswith("#") and len(spec) == 7:
        return spec[1:].upper()
    parts = [p.strip() for p in spec.split("!") if p.strip()]
    if not parts:
        return default_hex
    base_name = parts[0]
    if base_name not in COLOR_PALETTE:
        return default_hex
    r1, g1, b1 = COLOR_PALETTE[base_name]
    if len(parts) == 1:
        return f"{r1:02X}{g1:02X}{b1:02X}"
    try:
        pct = float(parts[1]) / 100.0
    except ValueError:
        pct = 1.0
    pct = max(0.0, min(1.0, pct))
    c2_name = parts[2] if len(parts) >= 3 and parts[2] in COLOR_PALETTE else "white"
    r2, g2, b2 = COLOR_PALETTE[c2_name]
    r = int(round(r1 * pct + r2 * (1.0 - pct)))
    g = int(round(g1 * pct + g2 * (1.0 - pct)))
    b = int(round(b1 * pct + b2 * (1.0 - pct)))
    return f"{r:02X}{g:02X}{b:02X}"


@dataclass
class PrimLine:
    x1: float
    y1: float
    x2: float
    y2: float
    color: str = "0F172A"
    width_pt: float = 1.0
    dashed: bool = False
    dotted: bool = False
    double_line: bool = False
    arrow_start: bool = False
    arrow_end: bool = False


@dataclass
class PrimPath:
    points: list[tuple[float, float]]
    closed: bool = False
    smooth: bool = False
    stroke_color: str | None = "0F172A"
    fill_color: str | None = None
    width_pt: float = 1.1
    dashed: bool = False
    dotted: bool = False
    arrow_start: bool = False
    arrow_end: bool = False


@dataclass
class PrimCircle:
    cx: float
    cy: float
    radius_cm: float
    is_radius_pt: bool = False
    stroke_color: str | None = "0F172A"
    fill_color: str | None = None
    width_pt: float = 1.0
    dashed: bool = False


@dataclass
class PrimLabel:
    x: float
    y: float
    text: str
    anchor: str = "center"
    xshift_pt: float = 0.0
    yshift_pt: float = 0.0
    color: str = "0F172A"
    bold: bool = False
    font_size_pt: float = 10.5


@dataclass
class Scene2D:
    lines: list[PrimLine] = field(default_factory=list)
    paths: list[PrimPath] = field(default_factory=list)
    circles: list[PrimCircle] = field(default_factory=list)
    labels: list[PrimLabel] = field(default_factory=list)
    clip_rect: tuple[float, float, float, float] | None = None
    scale: float = 1.0


def _safe_eval_math(expr: str, x_val: float = 0.0, macros: dict[str, float] | None = None) -> float:
    """Evaluate a LaTeX/TikZ/PGFPlots math expression safely."""
    s = expr.strip()
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1].strip()
    if macros:
        for k in sorted(macros.keys(), key=len, reverse=True):
            s = s.replace(k, f"({macros[k]})")
    s = s.replace(r"\x", f"({x_val})")
    s = re.sub(r"\bx\b", f"({x_val})", s)
    s = s.replace("^", "**")
    s = s.replace(r"\pi", "math.pi").replace("pi", "math.pi")
    s = s.replace("deg(", "math.degrees(").replace("rad(", "math.radians(")
    s = re.sub(r"\bsin\s*\(", "_sin_deg(", s)
    s = re.sub(r"\bcos\s*\(", "_cos_deg(", s)
    s = re.sub(r"\btan\s*\(", "_tan_deg(", s)
    s = re.sub(r"\bsqrt\s*\(", "math.sqrt(", s)
    s = re.sub(r"\bln\s*\(", "math.log(", s)
    s = re.sub(r"\blog\s*\(", "math.log10(", s)
    s = re.sub(r"\bexp\s*\(", "math.exp(", s)
    s = re.sub(r"\babs\s*\(", "abs(", s)
    s = re.sub(r"(\d)\s*\(", r"\1*(", s)

    env = {
        "__builtins__": {},
        "math": math,
        "abs": abs,
        "min": min,
        "max": max,
        "_sin_deg": lambda v: math.sin(math.radians(v)),
        "_cos_deg": lambda v: math.cos(math.radians(v)),
        "_tan_deg": lambda v: math.tan(math.radians(v)),
    }
    return float(eval(s, env, {}))


def _parabola_bend_points(
    p1: tuple[float, float], vertex: tuple[float, float], p2: tuple[float, float], steps: int = 24
) -> list[tuple[float, float]]:
    """Sample exact points along TikZ `(p1) parabola bend (vertex) (p2)`."""
    x1, y1 = p1
    xv, yv = vertex
    x2, y2 = p2
    pts: list[tuple[float, float]] = []
    half = max(6, steps // 2)
    for i in range(half):
        t = i / float(half)
        x = x1 + (xv - x1) * t
        if abs(x1 - xv) > 1e-9:
            y = yv + (y1 - yv) * ((x - xv) / (x1 - xv)) ** 2
        else:
            y = y1 + (yv - y1) * t
        pts.append((x, y))
    for i in range(half + 1):
        t = i / float(half)
        x = xv + (x2 - xv) * t
        if abs(x2 - xv) > 1e-9:
            y = yv + (y2 - yv) * ((x - xv) / (x2 - xv)) ** 2
        else:
            y = yv + (y2 - yv) * t
        pts.append((x, y))
    return pts


def _parse_style_opts(opt_str: str) -> dict[str, any]:
    """Parse TikZ [...] style options into stroke, fill, width, dash, arrows, domain, anchor, shifts."""
    info: dict[str, any] = {
        "color": None,
        "fill": None,
        "width_pt": 1.0,
        "dashed": False,
        "dotted": False,
        "double": False,
        "arrow_start": False,
        "arrow_end": False,
        "smooth": False,
        "domain": None,
        "samples": 80,
        "anchor": "center",
        "pos": 0.5,
        "xshift_pt": 0.0,
        "yshift_pt": 0.0,
        "bold": False,
        "scale": 1.0,
    }
    if not opt_str:
        return info

    tokens: list[str] = []
    cur: list[str] = []
    depth = 0
    for ch in opt_str:
        if ch in "{(":
            depth += 1
        elif ch in "})":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            tokens.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        tokens.append("".join(cur).strip())

    for tok in tokens:
        if not tok:
            continue
        t_low = tok.lower().strip()
        if t_low in ("->", "-stealth", "-latex", "-to", ">=stealth", ">=latex"):
            if not t_low.startswith(">="):
                info["arrow_end"] = True
        elif t_low in ("<-", "stealth-", "latex-"):
            info["arrow_start"] = True
        elif t_low in ("<->", "stealth-stealth", "latex-latex"):
            info["arrow_start"] = True
            info["arrow_end"] = True
        elif t_low in ("dashed", "densely dashed", "loosely dashed"):
            info["dashed"] = True
        elif t_low in ("dotted", "densely dotted"):
            info["dotted"] = True
        elif t_low == "double":
            info["double"] = True
            info["width_pt"] = 1.6
        elif t_low == "smooth":
            info["smooth"] = True
        elif t_low == "ultra thin":
            info["width_pt"] = 0.4
        elif t_low == "very thin":
            info["width_pt"] = 0.6
        elif t_low == "thin":
            info["width_pt"] = 0.8
        elif t_low == "semithick":
            info["width_pt"] = 1.1
        elif t_low == "thick":
            info["width_pt"] = 1.35
        elif t_low == "very thick":
            info["width_pt"] = 1.8
        elif t_low == "ultra thick":
            info["width_pt"] = 2.2
        elif t_low == "midway":
            info["pos"] = 0.5
        elif t_low in (
            "above", "below", "left", "right",
            "above left", "above right", "below left", "below right",
        ):
            info["anchor"] = t_low
        elif "=" in tok:
            k, v = [x.strip() for x in tok.split("=", 1)]
            kl = k.lower()
            if kl == "domain":
                parts = v.split(":")
                if len(parts) == 2:
                    try:
                        info["domain"] = (float(parts[0]), float(parts[1]))
                    except ValueError:
                        pass
            elif kl == "samples":
                try:
                    info["samples"] = max(20, min(200, int(v)))
                except ValueError:
                    pass
            elif kl == "scale":
                try:
                    info["scale"] = float(v)
                except ValueError:
                    pass
            elif kl == "pos":
                try:
                    info["pos"] = float(v)
                except ValueError:
                    pass
            elif kl == "xshift":
                m_num = re.search(r"([+-]?\d+(?:\.\d+)?)", v)
                if m_num:
                    info["xshift_pt"] = float(m_num.group(1))
            elif kl == "yshift":
                m_num = re.search(r"([+-]?\d+(?:\.\d+)?)", v)
                if m_num:
                    info["yshift_pt"] = float(m_num.group(1))
            elif kl == "color" or kl == "draw":
                info["color"] = resolve_tikz_color(v)
            elif kl == "fill":
                info["fill"] = resolve_tikz_color(v)
            elif kl == "font":
                if "bf" in v:
                    info["bold"] = True
        else:
            base_c = t_low.split("!")[0].strip()
            if base_c in COLOR_PALETTE:
                resolved = resolve_tikz_color(t_low)
                info["color"] = resolved
                info["fill"] = resolved

    return info


def parse_pgfplots_axis_to_scene(axis_tex: str, scale: float = 1.0) -> Scene2D:
    """Parse a PGFPlots \\begin{axis}[...] ... \\end{axis} environment into a 2D vector Scene."""
    scene = Scene2D(scale=scale)
    m_opt = re.search(r"\\begin\{axis\}\s*(?:\[(.*?)\])?", axis_tex, re.S)
    opts_raw = m_opt.group(1) if (m_opt and m_opt.group(1)) else ""
    body_start = m_opt.end() if m_opt else 0
    body_end = axis_tex.rfind(r"\end{axis}")
    body = axis_tex[body_start:body_end] if body_end != -1 else axis_tex[body_start:]

    width_cm = 7.2
    height_cm = 4.6
    xmin: float | None = None
    xmax: float | None = None
    ymin: float | None = None
    ymax: float | None = None
    xlabel = "$x$"
    ylabel = "$y$"
    title = ""
    axis_lines = "box"
    has_grid = False
    default_domain = (-3.0, 3.0)
    default_samples = 80

    tokens: list[str] = []
    cur: list[str] = []
    depth = 0
    for ch in opts_raw:
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth = max(0, depth - 1)
        if ch == "," and depth == 0:
            tokens.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        tokens.append("".join(cur).strip())

    for tok in tokens:
        if "=" not in tok:
            continue
        k, v = [p.strip() for p in tok.split("=", 1)]
        kl = k.lower()
        v_clean = v[1:-1].strip() if (v.startswith("{") and v.endswith("}")) else v
        if kl == "width":
            m_n = re.search(r"(\d+(?:\.\d+)?)", v_clean)
            if m_n:
                width_cm = float(m_n.group(1))
        elif kl == "height":
            m_n = re.search(r"(\d+(?:\.\d+)?)", v_clean)
            if m_n:
                height_cm = float(m_n.group(1))
        elif kl == "xmin":
            try:
                xmin = float(v_clean)
            except ValueError:
                pass
        elif kl == "xmax":
            try:
                xmax = float(v_clean)
            except ValueError:
                pass
        elif kl == "ymin":
            try:
                ymin = float(v_clean)
            except ValueError:
                pass
        elif kl == "ymax":
            try:
                ymax = float(v_clean)
            except ValueError:
                pass
        elif kl == "xlabel":
            xlabel = v_clean
        elif kl == "ylabel":
            ylabel = v_clean
        elif kl == "title":
            title = v_clean
        elif kl == "axis lines":
            axis_lines = v_clean.lower()
        elif kl == "grid":
            has_grid = v_clean.lower() in ("major", "both", "minor")
        elif kl == "domain":
            dp = v_clean.split(":")
            if len(dp) == 2:
                try:
                    default_domain = (float(dp[0]), float(dp[1]))
                except ValueError:
                    pass
        elif kl == "samples":
            try:
                default_samples = int(v_clean)
            except ValueError:
                pass

    addplot_re = re.compile(r"\\addplot\+?\s*(?:\[(.*?)\])?\s*(\{[^{}]*\}|[^;]+);", re.S)
    plots_raw: list[tuple[dict[str, any], list[tuple[float, float]]]] = []
    all_x: list[float] = []
    all_y: list[float] = []

    for pm in addplot_re.finditer(body):
        p_opts = _parse_style_opts(pm.group(1) or "")
        expr_raw = pm.group(2).strip()
        dom = p_opts["domain"] or default_domain
        if xmin is not None and xmax is not None and p_opts["domain"] is None:
            dom = (xmin, xmax)
        samples = p_opts["samples"] or default_samples

        pts: list[tuple[float, float]] = []
        if "coordinates" in expr_raw:
            for cm in re.finditer(r"\(\s*([+-]?\d+(?:\.\d+)?)\s*,\s*([+-]?\d+(?:\.\d+)?)\s*\)", expr_raw):
                px, py = float(cm.group(1)), float(cm.group(2))
                pts.append((px, py))
                all_x.append(px)
                all_y.append(py)
        else:
            x0, x1 = dom
            for i in range(samples + 1):
                xv = x0 + (x1 - x0) * (i / float(samples))
                try:
                    yv = _safe_eval_math(expr_raw, x_val=xv)
                    if math.isfinite(yv) and abs(yv) < 1e4:
                        pts.append((xv, yv))
                        all_x.append(xv)
                        all_y.append(yv)
                except Exception:
                    continue
        plots_raw.append((p_opts, pts))

    if xmin is None:
        xmin = min(all_x) - 0.5 if all_x else default_domain[0]
    if xmax is None:
        xmax = max(all_x) + 0.5 if all_x else default_domain[1]
    if ymin is None:
        ymin = min(all_y) - 0.8 if all_y else -4.0
    if ymax is None:
        ymax = max(all_y) + 0.8 if all_y else 4.0

    if xmax <= xmin:
        xmax = xmin + 4.0
    if ymax <= ymin:
        ymax = ymin + 4.0

    def to_cm(xv: float, yv: float) -> tuple[float, float]:
        cx = (xv - xmin) / (xmax - xmin) * width_cm
        cy = (yv - ymin) / (ymax - ymin) * height_cm
        return cx, cy

    scene.clip_rect = (-0.4, -0.4, width_cm + 0.8, height_cm + 0.8)

    if has_grid:
        for gx in range(int(math.ceil(xmin)), int(math.floor(xmax)) + 1):
            cx1, cy1 = to_cm(gx, ymin)
            cx2, cy2 = to_cm(gx, ymax)
            scene.lines.append(PrimLine(cx1, cy1, cx2, cy2, color="E2E8F0", width_pt=0.6, dashed=True))
        for gy in range(int(math.ceil(ymin)), int(math.floor(ymax)) + 1):
            cx1, cy1 = to_cm(xmin, gy)
            cx2, cy2 = to_cm(xmax, gy)
            scene.lines.append(PrimLine(cx1, cy1, cx2, cy2, color="E2E8F0", width_pt=0.6, dashed=True))

    if axis_lines in ("middle", "center", "none"):
        ox_y = 0.0 if (ymin <= 0 <= ymax) else ymin
        oy_x = 0.0 if (xmin <= 0 <= xmax) else xmin
        x_start, y_axis_pos = to_cm(xmin, ox_y)
        x_end, _ = to_cm(xmax, ox_y)
        y_axis_x, y_start = to_cm(oy_x, ymin)
        _, y_end = to_cm(oy_x, ymax)

        scene.lines.append(PrimLine(x_start, y_axis_pos, x_end, y_axis_pos, color="0F172A", width_pt=1.2, arrow_end=True))
        scene.lines.append(PrimLine(y_axis_x, y_start, y_axis_x, y_end, color="0F172A", width_pt=1.2, arrow_end=True))

        ox_cm, oy_cm = to_cm(oy_x, ox_y)
        scene.labels.append(PrimLabel(ox_cm, oy_cm, "$O$", anchor="below left", font_size_pt=10.5))
        scene.labels.append(PrimLabel(x_end, y_axis_pos, xlabel or "$x$", anchor="right", font_size_pt=11.0))
        scene.labels.append(PrimLabel(y_axis_x, y_end, ylabel or "$y$", anchor="above", font_size_pt=11.0))
    else:
        scene.paths.append(
            PrimPath(
                [(0, 0), (width_cm, 0), (width_cm, height_cm), (0, height_cm)],
                closed=True,
                stroke_color="334155",
                fill_color=None,
                width_pt=1.0,
            )
        )
        for gx in range(int(math.ceil(xmin)), int(math.floor(xmax)) + 1):
            cx, _ = to_cm(gx, ymin)
            scene.lines.append(PrimLine(cx, 0, cx, 0.12, color="334155", width_pt=0.8))
            scene.labels.append(PrimLabel(cx, -0.05, f"${gx}$", anchor="below", font_size_pt=9.5))
        y_step = max(1, int(round((ymax - ymin) / 4.0)))
        gy = int(math.ceil(ymin))
        while gy <= int(math.floor(ymax)):
            _, cy = to_cm(xmin, gy)
            scene.lines.append(PrimLine(0, cy, 0.12, cy, color="334155", width_pt=0.8))
            scene.labels.append(PrimLabel(-0.08, cy, f"${gy}$", anchor="left", font_size_pt=9.5))
            gy += y_step
        if xlabel:
            scene.labels.append(PrimLabel(width_cm / 2.0, -0.55, xlabel, anchor="below", font_size_pt=10.5))
        if ylabel:
            scene.labels.append(PrimLabel(-0.65, height_cm / 2.0, ylabel, anchor="left", font_size_pt=10.5))

    if title:
        scene.labels.append(PrimLabel(width_cm / 2.0, height_cm + 0.35, title, anchor="above", bold=True, font_size_pt=11.0))

    for p_opts, pts in plots_raw:
        stroke = p_opts["color"] or "2563EB"
        w_pt = max(1.5, p_opts["width_pt"])
        cur_segment: list[tuple[float, float]] = []
        for xv, yv in pts:
            if xmin - 0.05 <= xv <= xmax + 0.05 and ymin - 0.15 <= yv <= ymax + 0.15:
                yv_clamped = max(ymin, min(ymax, yv))
                cur_segment.append(to_cm(xv, yv_clamped))
            else:
                if len(cur_segment) >= 2:
                    scene.paths.append(
                        PrimPath(
                            cur_segment,
                            closed=False,
                            smooth=True,
                            stroke_color=stroke,
                            width_pt=w_pt,
                            dashed=p_opts["dashed"],
                        )
                    )
                cur_segment = []
        if len(cur_segment) >= 2:
            scene.paths.append(
                PrimPath(
                    cur_segment,
                    closed=False,
                    smooth=True,
                    stroke_color=stroke,
                    width_pt=w_pt,
                    dashed=p_opts["dashed"],
                )
            )

    return scene


def parse_tikzpicture_to_scene(tikz_tex: str) -> Scene2D:
    """Parse a general TikZ \\begin{tikzpicture} ... \\end{tikzpicture} into a 2D vector Scene."""
    cleaned = re.sub(r"(?<!\\)%.*$", "", tikz_tex, flags=re.M)

    if r"\begin{axis}" in cleaned:
        m_sc = re.search(r"\\begin\{tikzpicture\}\s*\[([^\]]*)\]", cleaned)
        sc = _parse_style_opts(m_sc.group(1))["scale"] if m_sc else 1.0
        return parse_pgfplots_axis_to_scene(cleaned, scale=sc)

    m_header = re.search(r"\\begin\{tikzpicture\}\s*(?:\[(.*?)\])?", cleaned, re.S)
    header_opts = _parse_style_opts(m_header.group(1) if (m_header and m_header.group(1)) else "")
    scale = header_opts["scale"]
    scene = Scene2D(scale=scale)

    body_start = m_header.end() if m_header else 0
    body_end = cleaned.rfind(r"\end{tikzpicture}")
    body = cleaned[body_start:body_end] if body_end != -1 else cleaned[body_start:]

    macros: dict[str, float] = {}
    coords: dict[str, tuple[float, float]] = {}

    for dm in re.finditer(r"\\def\s*(\\[a-zA-Z]+)\s*\{([^{}]+)\}", body):
        try:
            macros[dm.group(1)] = _safe_eval_math(dm.group(2), macros=macros)
        except Exception:
            pass

    def eval_pt(pt_str: str) -> tuple[float, float] | None:
        s = pt_str.strip()
        if s.startswith("(") and s.endswith(")"):
            s = s[1:-1].strip()
        if s in coords:
            return coords[s]
        if "," in s:
            parts = s.split(",", 1)
            try:
                return (
                    _safe_eval_math(parts[0], macros=macros),
                    _safe_eval_math(parts[1], macros=macros),
                )
            except Exception:
                return None
        if ":" in s:
            parts = s.split(":", 1)
            try:
                ang = _safe_eval_math(parts[0], macros=macros)
                rad = _safe_eval_math(parts[1], macros=macros)
                return (rad * math.cos(math.radians(ang)), rad * math.sin(math.radians(ang)))
            except Exception:
                return None
        return None

    statements = [st.strip() for st in body.split(";") if st.strip()]

    for stmt in statements:
        stmt = re.sub(r"\\def\s*\\[a-zA-Z]+\s*\{[^{}]+\}\s*", "", stmt).strip()
        if not stmt:
            continue

        m_coord = re.match(r"^\\coordinate\s*\(\s*([^()]+?)\s*\)\s*at\s*\(\s*([^()]+?)\s*\)", stmt, re.S)
        if m_coord:
            name = m_coord.group(1).strip()
            pt = eval_pt(m_coord.group(2))
            if pt is not None:
                coords[name] = pt
            continue

        m_clip = re.match(r"^\\clip\s*\(\s*([^()]+?)\s*\)\s*rectangle\s*\(\s*([^()]+?)\s*\)", stmt, re.S)
        if m_clip:
            p1 = eval_pt(m_clip.group(1))
            p2 = eval_pt(m_clip.group(2))
            if p1 and p2:
                scene.clip_rect = (min(p1[0], p2[0]), min(p1[1], p2[1]), max(p1[0], p2[0]), max(p1[1], p2[1]))
            continue

        m_node = re.match(r"^\\node\s*(?:\[(.*?)\])?\s*at\s*\(\s*([^()]+?)\s*\)\s*\{(.*)\}\s*$", stmt, re.S)
        if m_node:
            n_opts = _parse_style_opts(m_node.group(1) or "")
            pt = eval_pt(m_node.group(2))
            txt = m_node.group(3).strip()
            if pt is not None:
                scene.labels.append(
                    PrimLabel(
                        pt[0],
                        pt[1],
                        txt,
                        anchor=n_opts["anchor"],
                        xshift_pt=n_opts["xshift_pt"],
                        yshift_pt=n_opts["yshift_pt"],
                        color=n_opts["color"] or "0F172A",
                        bold=n_opts["bold"],
                    )
                )
            continue

        m_cmd = re.match(r"^\\(draw|fill|filldraw|shade)\s*(?:\[(.*?)\])?\s*(.*)$", stmt, re.S)
        if not m_cmd:
            continue

        cmd_type = m_cmd.group(1)
        opts = _parse_style_opts(m_cmd.group(2) or "")
        rest = m_cmd.group(3).strip()

        is_fill_cmd = cmd_type in ("fill", "filldraw", "shade")
        stroke_col = opts["color"] or ("0F172A" if cmd_type in ("draw", "filldraw") else None)
        fill_col = opts["fill"] or ("0F172A" if cmd_type == "fill" else None)

        m_plot = re.search(r"plot\s*\(\s*\{?\\x\}?\s*,\s*\{(.*)\}\s*\)", rest, re.S)
        if m_plot:
            expr_str = m_plot.group(1).strip()
            dom = opts["domain"] or (-3.0, 3.0)
            samples = opts["samples"] or 100
            x0, x1 = dom
            raw_pts: list[tuple[float, float]] = []
            for i in range(samples + 1):
                xv = x0 + (x1 - x0) * (i / float(samples))
                try:
                    yv = _safe_eval_math(expr_str, x_val=xv, macros=macros)
                    if math.isfinite(yv) and abs(yv) < 1e4:
                        raw_pts.append((xv, yv))
                except Exception:
                    continue

            if scene.clip_rect:
                cx1, cy1, cx2, cy2 = scene.clip_rect
                seg: list[tuple[float, float]] = []
                for px, py in raw_pts:
                    if cx1 <= px <= cx2 and cy1 <= py <= cy2:
                        seg.append((px, py))
                    else:
                        if len(seg) >= 2:
                            scene.paths.append(
                                PrimPath(
                                    seg,
                                    smooth=True,
                                    stroke_color=stroke_col or "2563EB",
                                    width_pt=opts["width_pt"],
                                    dashed=opts["dashed"],
                                )
                            )
                        seg = []
                if len(seg) >= 2:
                    scene.paths.append(
                        PrimPath(
                            seg,
                            smooth=True,
                            stroke_color=stroke_col or "2563EB",
                            width_pt=opts["width_pt"],
                            dashed=opts["dashed"],
                        )
                    )
            elif len(raw_pts) >= 2:
                scene.paths.append(
                    PrimPath(
                        raw_pts,
                        smooth=True,
                        stroke_color=stroke_col or "2563EB",
                        width_pt=opts["width_pt"],
                        dashed=opts["dashed"],
                    )
                )
            continue

        m_circ = re.match(
            r"^\(\s*([^()]+?)\s*\)\s*circle\s*(?:\(\s*([^()]+?)\s*\)|\{\s*([^{}]+?)\s*\})(.*)$",
            rest,
            re.S,
        )
        if m_circ:
            cpt = eval_pt(m_circ.group(1))
            rad_str = (m_circ.group(2) or m_circ.group(3) or "2pt").strip()
            tail = m_circ.group(4).strip()
            if cpt is not None:
                is_pt = "pt" in rad_str
                m_rn = re.search(r"(\d+(?:\.\d+)?)", rad_str)
                r_val = float(m_rn.group(1)) if m_rn else 2.0
                scene.circles.append(
                    PrimCircle(
                        cpt[0],
                        cpt[1],
                        radius_cm=r_val,
                        is_radius_pt=is_pt,
                        stroke_color=stroke_col,
                        fill_color=fill_col,
                        width_pt=opts["width_pt"],
                        dashed=opts["dashed"],
                    )
                )
                m_in_node = re.search(r"node\s*(?:\[(.*?)\])?\s*\{(.*)\}", tail, re.S)
                if m_in_node:
                    n_opts = _parse_style_opts(m_in_node.group(1) or "")
                    scene.labels.append(
                        PrimLabel(
                            cpt[0],
                            cpt[1],
                            m_in_node.group(2).strip(),
                            anchor=n_opts["anchor"],
                            xshift_pt=n_opts["xshift_pt"],
                            yshift_pt=n_opts["yshift_pt"],
                            color=n_opts["color"] or "0F172A",
                            bold=n_opts["bold"],
                        )
                    )
            continue

        m_rect = re.match(r"^\(\s*([^()]+?)\s*\)\s*rectangle\s*\(\s*([^()]+?)\s*\)$", rest, re.S)
        if m_rect:
            p1 = eval_pt(m_rect.group(1))
            p2 = eval_pt(m_rect.group(2))
            if p1 and p2:
                x1, y1 = p1
                x2, y2 = p2
                scene.paths.append(
                    PrimPath(
                        [(x1, y1), (x2, y1), (x2, y2), (x1, y2)],
                        closed=True,
                        stroke_color=stroke_col,
                        fill_color=fill_col,
                        width_pt=opts["width_pt"],
                        dashed=opts["dashed"],
                    )
                )
            continue

        _parse_complex_tikz_path(rest, opts, stroke_col, fill_col, is_fill_cmd, eval_pt, scene)

    return scene


def _parse_complex_tikz_path(
    rest: str,
    opts: dict[str, any],
    stroke_col: str | None,
    fill_col: str | None,
    is_fill_cmd: bool,
    eval_pt,
    scene: Scene2D,
) -> None:
    inline_nodes: list[tuple[dict[str, any], str, int]] = []

    def _strip_node(m: re.Match[str]) -> str:
        n_opts = _parse_style_opts(m.group(1) or "")
        n_txt = m.group(2).strip()
        inline_nodes.append((n_opts, n_txt, m.start()))
        return " "

    path_clean = re.sub(r"node\s*(?:\[(.*?)\])?\s*\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}", _strip_node, rest, flags=re.S)

    if "parabola bend" in path_clean:
        closed = "cycle" in path_clean
        pts: list[tuple[float, float]] = []
        legs = [lg.strip() for lg in path_clean.split("--") if lg.strip() and lg.strip() != "cycle"]
        for leg in legs:
            m_pb = re.search(
                r"\(\s*([^()]+?)\s*\)\s*parabola\s+bend\s*\(\s*([^()]+?)\s*\)\s*\(\s*([^()]+?)\s*\)",
                leg,
            )
            if m_pb:
                pa = eval_pt(m_pb.group(1))
                pv = eval_pt(m_pb.group(2))
                pb = eval_pt(m_pb.group(3))
                if pa and pv and pb:
                    pb_pts = _parabola_bend_points(pa, pv, pb, steps=24)
                    if pts and pts[-1] == pb_pts[0]:
                        pts.extend(pb_pts[1:])
                    else:
                        pts.extend(pb_pts)
            else:
                pt = eval_pt(leg)
                if pt:
                    pts.append(pt)
        if len(pts) >= 2:
            scene.paths.append(
                PrimPath(
                    pts,
                    closed=closed or is_fill_cmd,
                    smooth=False,
                    stroke_color=stroke_col if not (is_fill_cmd and stroke_col == fill_col) else None,
                    fill_color=fill_col if is_fill_cmd else None,
                    width_pt=opts["width_pt"],
                    dashed=opts["dashed"],
                )
            )
        return

    token_re = re.compile(r"(\(\s*[^()]+\s*\)|--|cycle)")
    toks = [t.strip() for t in token_re.findall(path_clean) if t.strip()]

    subpaths: list[tuple[list[tuple[float, float]], bool]] = []
    cur_pts: list[tuple[float, float]] = []
    expect_conn = False
    is_closed = False

    for tok in toks:
        if tok == "--":
            expect_conn = True
        elif tok == "cycle":
            is_closed = True
        elif tok.startswith("("):
            pt = eval_pt(tok)
            if pt is None:
                continue
            if not cur_pts or expect_conn:
                cur_pts.append(pt)
                expect_conn = False
            else:
                if len(cur_pts) >= 2:
                    subpaths.append((cur_pts, is_closed))
                cur_pts = [pt]
                is_closed = False

    if len(cur_pts) >= 2:
        subpaths.append((cur_pts, is_closed))

    for sp_pts, sp_closed in subpaths:
        if len(sp_pts) == 2 and not sp_closed and not is_fill_cmd:
            p1, p2 = sp_pts[0], sp_pts[1]
            if scene.clip_rect:
                cx1, cy1, cx2, cy2 = scene.clip_rect
                p1 = (max(cx1, min(cx2, p1[0])), max(cy1, min(cy2, p1[1])))
                p2 = (max(cx1, min(cx2, p2[0])), max(cy1, min(cy2, p2[1])))
            scene.lines.append(
                PrimLine(
                    p1[0],
                    p1[1],
                    p2[0],
                    p2[1],
                    color=stroke_col or "0F172A",
                    width_pt=opts["width_pt"],
                    dashed=opts["dashed"],
                    dotted=opts["dotted"],
                    double_line=opts["double"],
                    arrow_start=opts["arrow_start"],
                    arrow_end=opts["arrow_end"],
                )
            )
            for n_opts, n_txt, _ in inline_nodes:
                pos_t = n_opts["pos"] if "midway" in rest or "pos=" in rest else 1.0
                nx = p1[0] + (p2[0] - p1[0]) * pos_t
                ny = p1[1] + (p2[1] - p1[1]) * pos_t
                scene.labels.append(
                    PrimLabel(
                        nx,
                        ny,
                        n_txt,
                        anchor=n_opts["anchor"],
                        xshift_pt=n_opts["xshift_pt"],
                        yshift_pt=n_opts["yshift_pt"],
                        color=n_opts["color"] or stroke_col or "0F172A",
                        bold=n_opts["bold"],
                    )
                )
        else:
            scene.paths.append(
                PrimPath(
                    sp_pts,
                    closed=sp_closed or is_fill_cmd,
                    stroke_color=stroke_col,
                    fill_color=fill_col if is_fill_cmd else None,
                    width_pt=opts["width_pt"],
                    dashed=opts["dashed"],
                    dotted=opts["dotted"],
                    arrow_start=opts["arrow_start"],
                    arrow_end=opts["arrow_end"],
                )
            )


def _catmull_rom_to_bezier_segments(
    pts: list[tuple[int, int]]
) -> list[tuple[tuple[int, int], tuple[int, int], tuple[int, int]]]:
    n = len(pts)
    if n < 2:
        return []
    if n == 2:
        return [(pts[0], pts[1], pts[1])]
    beziers: list[tuple[tuple[int, int], tuple[int, int], tuple[int, int]]] = []
    for i in range(n - 1):
        p0 = pts[max(0, i - 1)]
        p1 = pts[i]
        p2 = pts[i + 1]
        p3 = pts[min(n - 1, i + 2)]
        c1x = int(round(p1[0] + (p2[0] - p0[0]) / 6.0))
        c1y = int(round(p1[1] + (p2[1] - p0[1]) / 6.0))
        c2x = int(round(p2[0] - (p3[0] - p1[0]) / 6.0))
        c2y = int(round(p2[1] - (p3[1] - p1[1]) / 6.0))
        beziers.append(((c1x, c1y), (c2x, c2y), p2))
    return beziers


def _build_textbox_paragraph_xml(text: str, color_hex: str, bold: bool, font_size_pt: float) -> str:
    sz_half_pt = int(round(font_size_pt * 2))
    t = text.strip()
    t = re.sub(r"\\(?:scriptsize|footnotesize|small|normalsize|large|Large|bfseries)\b\s*", "", t).strip()

    if t.startswith("$") and t.endswith("$") and t.count("$") == 2:
        omml_str = latex_to_omml_xml_string(t[1:-1], display=False)
        return (
            f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
            f"{omml_str}</w:p>"
        )

    parts = re.split(r"(\$[^$]+\$)", t)
    runs_xml: list[str] = []
    for p in parts:
        if not p:
            continue
        if p.startswith("$") and p.endswith("$"):
            runs_xml.append(latex_to_omml_xml_string(p[1:-1], display=False))
        else:
            clean_p = p.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            b_tag = "<w:b/>" if bold else ""
            runs_xml.append(
                f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                f'{b_tag}<w:color w:val="{color_hex}"/><w:sz w:val="{sz_half_pt}"/></w:rPr>'
                f'<w:t xml:space="preserve">{clean_p}</w:t></w:r>'
            )
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
        + "".join(runs_xml)
        + "</w:p>"
    )


def scene_to_drawingml_element(scene: Scene2D, group_id: int = 100):
    """Compile a 2D vector Scene into a native Microsoft Word <w:r><w:drawing>...<wpg:wgp>...</w:r> element."""
    xs: list[float] = []
    ys: list[float] = []

    if scene.clip_rect:
        cx1, cy1, cx2, cy2 = scene.clip_rect
        xs.extend([cx1, cx2])
        ys.extend([cy1, cy2])
    else:
        for ln in scene.lines:
            xs.extend([ln.x1, ln.x2])
            ys.extend([ln.y1, ln.y2])
        for pa in scene.paths:
            for px, py in pa.points:
                xs.append(px)
                ys.append(py)
        for ci in scene.circles:
            r_u = 0.08 if ci.is_radius_pt else ci.radius_cm
            xs.extend([ci.cx - r_u, ci.cx + r_u])
            ys.extend([ci.cy - r_u, ci.cy + r_u])
        for lb in scene.labels:
            xs.append(lb.x)
            ys.append(lb.y)

    if not xs or not ys:
        xs, ys = [-2.0, 2.0], [-2.0, 2.0]

    min_x, max_x = min(xs), max(xs)
    min_y, max_y = min(ys), max(ys)
    span_x = max(max_x - min_x, 1.0)
    span_y = max(max_y - min_y, 1.0)

    raw_w_cm = span_x * scene.scale
    raw_h_cm = span_y * scene.scale

    max_w_cm = 11.2
    max_h_cm = 6.8
    fit_ratio = min(1.0, max_w_cm / max(raw_w_cm, 0.1), max_h_cm / max(raw_h_cm, 0.1))
    if raw_w_cm * fit_ratio < 4.5 and raw_h_cm * fit_ratio < 3.5:
        fit_ratio = min(4.8 / max(raw_w_cm, 0.1), max_h_cm / max(raw_h_cm, 0.1))

    eff_scale_x = scene.scale * fit_ratio
    eff_scale_y = scene.scale * fit_ratio

    pad_x_cm = 0.90
    pad_y_cm = 0.70

    total_w_cm = span_x * eff_scale_x + 2 * pad_x_cm
    total_h_cm = span_y * eff_scale_y + 2 * pad_y_cm

    total_w_emu = int(round(total_w_cm * EMU_PER_CM))
    total_h_emu = int(round(total_h_cm * EMU_PER_CM))

    def to_emu(x: float, y: float) -> tuple[int, int]:
        ex = int(round((pad_x_cm + (x - min_x) * eff_scale_x) * EMU_PER_CM))
        ey = int(round((pad_y_cm + (max_y - y) * eff_scale_y) * EMU_PER_CM))
        return ex, ey

    children_xml: list[str] = []

    children_xml.append(
        f"""<wps:wsp>
          <wps:cNvSpPr/>
          <wps:spPr>
            <a:xfrm><a:off x="0" y="0"/><a:ext cx="{total_w_emu}" cy="{total_h_emu}"/></a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            <a:noFill/>
            <a:ln><a:noFill/></a:ln>
          </wps:spPr>
          <wps:bodyPr/>
        </wps:wsp>"""
    )

    for pa in scene.paths:
        if len(pa.points) < 2:
            continue
        emu_pts = [to_emu(px, py) for px, py in pa.points]
        pxs = [p[0] for p in emu_pts]
        pys = [p[1] for p in emu_pts]
        bx, by = min(pxs), min(pys)
        bw = max(max(pxs) - bx, 12700)
        bh = max(max(pys) - by, 12700)

        rel_pts = [(p[0] - bx, p[1] - by) for p in emu_pts]
        path_cmds: list[str] = [f'<a:moveTo><a:pt x="{rel_pts[0][0]}" y="{rel_pts[0][1]}"/></a:moveTo>']

        if pa.smooth and len(rel_pts) >= 3:
            for c1, c2, ep in _catmull_rom_to_bezier_segments(rel_pts):
                path_cmds.append(
                    f"<a:cubicBezTo>"
                    f'<a:pt x="{c1[0]}" y="{c1[1]}"/>'
                    f'<a:pt x="{c2[0]}" y="{c2[1]}"/>'
                    f'<a:pt x="{ep[0]}" y="{ep[1]}"/>'
                    f"</a:cubicBezTo>"
                )
        else:
            for rp in rel_pts[1:]:
                path_cmds.append(f'<a:lnTo><a:pt x="{rp[0]}" y="{rp[1]}"/></a:lnTo>')

        if pa.closed:
            path_cmds.append("<a:close/>")

        fill_xml = (
            f'<a:solidFill><a:srgbClr val="{pa.fill_color}"/></a:solidFill>'
            if pa.fill_color
            else "<a:noFill/>"
        )
        if pa.stroke_color:
            w_emu = int(round(pa.width_pt * EMU_PER_PT))
            dash_xml = (
                '<a:prstDash val="dash"/>'
                if pa.dashed
                else ('<a:prstDash val="sysDot"/>' if pa.dotted else "")
            )
            head_xml = '<a:headEnd type="triangle" w="med" len="med"/>' if pa.arrow_start else ""
            tail_xml = '<a:tailEnd type="triangle" w="med" len="med"/>' if pa.arrow_end else ""
            ln_xml = (
                f'<a:ln w="{w_emu}">'
                f'<a:solidFill><a:srgbClr val="{pa.stroke_color}"/></a:solidFill>'
                f"{dash_xml}{head_xml}{tail_xml}"
                f"</a:ln>"
            )
        else:
            ln_xml = "<a:ln><a:noFill/></a:ln>"

        children_xml.append(
            f"""<wps:wsp>
              <wps:cNvSpPr/>
              <wps:spPr>
                <a:xfrm><a:off x="{bx}" y="{by}"/><a:ext cx="{bw}" cy="{bh}"/></a:xfrm>
                <a:custGeom>
                  <a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>
                  <a:rect l="0" t="0" r="r" b="b"/>
                  <a:pathLst>
                    <a:path w="{bw}" h="{bh}">
                      {''.join(path_cmds)}
                    </a:path>
                  </a:pathLst>
                </a:custGeom>
                {fill_xml}
                {ln_xml}
              </wps:spPr>
              <wps:bodyPr/>
            </wps:wsp>"""
        )

    for ln in scene.lines:
        x1, y1 = to_emu(ln.x1, ln.y1)
        x2, y2 = to_emu(ln.x2, ln.y2)

        offsets = [(-18000, 0), (18000, 0)] if ln.double_line else [(0, 0)]
        for ox, oy in offsets:
            lx1, ly1 = x1 + ox, y1 + oy
            lx2, ly2 = x2 + ox, y2 + oy
            off_x = min(lx1, lx2)
            off_y = min(ly1, ly2)
            ext_x = max(abs(lx2 - lx1), 0)
            ext_y = max(abs(ly2 - ly1), 0)
            flip_h = ' flipH="1"' if lx2 < lx1 else ""
            flip_v = ' flipV="1"' if ly2 < ly1 else ""

            w_emu = int(round((0.85 if ln.double_line else ln.width_pt) * EMU_PER_PT))
            dash_xml = (
                '<a:prstDash val="dash"/>'
                if ln.dashed
                else ('<a:prstDash val="sysDot"/>' if ln.dotted else "")
            )
            head_xml = '<a:headEnd type="triangle" w="med" len="med"/>' if ln.arrow_start else ""
            tail_xml = '<a:tailEnd type="triangle" w="med" len="med"/>' if ln.arrow_end else ""

            children_xml.append(
                f"""<wps:wsp>
                  <wps:cNvSpPr/>
                  <wps:spPr>
                    <a:xfrm{flip_h}{flip_v}><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_x}" cy="{ext_y}"/></a:xfrm>
                    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
                    <a:ln w="{w_emu}">
                      <a:solidFill><a:srgbClr val="{ln.color}"/></a:solidFill>
                      {dash_xml}{head_xml}{tail_xml}
                    </a:ln>
                  </wps:spPr>
                  <wps:bodyPr/>
                </wps:wsp>"""
            )

    for ci in scene.circles:
        cx, cy = to_emu(ci.cx, ci.cy)
        if ci.is_radius_pt:
            r_emu = int(round(max(ci.radius_cm, 1.8) * EMU_PER_PT))
        else:
            r_emu = int(round(ci.radius_cm * eff_scale_x * EMU_PER_CM))
        off_x = cx - r_emu
        off_y = cy - r_emu
        diam = 2 * r_emu
        fill_xml = (
            f'<a:solidFill><a:srgbClr val="{ci.fill_color}"/></a:solidFill>'
            if ci.fill_color
            else "<a:noFill/>"
        )
        if ci.stroke_color:
            w_emu = int(round(ci.width_pt * EMU_PER_PT))
            ln_xml = f'<a:ln w="{w_emu}"><a:solidFill><a:srgbClr val="{ci.stroke_color}"/></a:solidFill></a:ln>'
        else:
            ln_xml = "<a:ln><a:noFill/></a:ln>"

        children_xml.append(
            f"""<wps:wsp>
              <wps:cNvSpPr/>
              <wps:spPr>
                <a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{diam}" cy="{diam}"/></a:xfrm>
                <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
                {fill_xml}
                {ln_xml}
              </wps:spPr>
              <wps:bodyPr/>
            </wps:wsp>"""
        )

    # Track placed label boxes to prevent collision when multiple labels share a point (e.g. "Vàng" and "X")
    placed_boxes: list[tuple[int, int, int, int]] = []

    for lb in scene.labels:
        lx, ly = to_emu(lb.x, lb.y)
        lx += int(round(lb.xshift_pt * EMU_PER_PT))
        ly -= int(round(lb.yshift_pt * EMU_PER_PT))

        has_tall_math = any(k in lb.text for k in (r"\frac", r"\dfrac", r"\sqrt"))
        clean_len = len(re.sub(r"\\[a-zA-Z]+|[{}$_^\s]", "", lb.text))
        box_w = max(240000, min(1450000, (clean_len + 2) * 82000))
        box_h = 390000 if has_tall_math else 230000

        anc = lb.anchor
        off_x = lx - box_w // 2
        off_y = ly - box_h // 2
        gap = 36000

        if "above" in anc:
            off_y = ly - box_h - gap
        elif "below" in anc:
            off_y = ly + gap
        if "left" in anc:
            off_x = lx - box_w - gap
        elif "right" in anc:
            off_x = lx + gap

        # De-collide against already placed label boxes
        for px1, py1, px2, py2 in placed_boxes:
            overlap_x = min(off_x + box_w, px2) - max(off_x, px1)
            overlap_y = min(off_y + box_h, py2) - max(off_y, py1)
            if overlap_x > 20000 and overlap_y > 20000:
                if "left" in anc:
                    off_x = px1 - box_w - 15000
                elif "right" in anc:
                    off_x = px2 + 15000
                else:
                    off_y = py1 - box_h - 15000

        off_x = max(0, min(total_w_emu - box_w, off_x))
        off_y = max(0, min(total_h_emu - box_h, off_y))
        placed_boxes.append((off_x, off_y, off_x + box_w, off_y + box_h))

        p_xml = _build_textbox_paragraph_xml(lb.text, lb.color, lb.bold, lb.font_size_pt)

        children_xml.append(
            f"""<wps:wsp>
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr>
                <a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{box_w}" cy="{box_h}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:noFill/>
                <a:ln><a:noFill/></a:ln>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent>
                  {p_xml}
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr">
                <a:noAutofit/>
              </wps:bodyPr>
            </wps:wsp>"""
        )

    group_xml = f"""
    <w:r xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:wps="{WPS_NS}" xmlns:wpg="{WPG_NS}" xmlns:m="{M_NS}">
      <w:drawing>
        <wp:inline distT="45720" distB="45720" distL="0" distR="0">
          <wp:extent cx="{total_w_emu}" cy="{total_h_emu}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="{group_id}" name="EditableMathShapeGroup_{group_id}"/>
          <wp:cNvGraphicFramePr/>
          <a:graphic>
            <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
              <wpg:wgp>
                <wpg:cNvGrpSpPr/>
                <wpg:grpSpPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="{total_w_emu}" cy="{total_h_emu}"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="{total_w_emu}" cy="{total_h_emu}"/>
                  </a:xfrm>
                </wpg:grpSpPr>
                {''.join(children_xml)}
              </wpg:wgp>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
    """
    return parse_xml(group_xml)


def convert_tikz_or_axis_to_drawingml(tex_block: str, group_id: int = 100):
    """High-level entry point: convert any \\begin{tikzpicture} or \\begin{axis} block to native Word DrawingML."""
    if r"\begin{tikzpicture}" in tex_block:
        scene = parse_tikzpicture_to_scene(tex_block)
    else:
        scene = parse_pgfplots_axis_to_scene(tex_block)
    return scene_to_drawingml_element(scene, group_id=group_id)
