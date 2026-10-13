"""Motor diagram calculations, data validation and Excel export.

The source-point values are kept separately from interpolated values.
No extrapolation beyond a curve's documented RPM range is performed.
"""
from __future__ import annotations

import io
import math
import re
from collections import defaultdict
from typing import Any

import plotly.graph_objects as go
import xlsxwriter

RPM_TO_TORQUE = 9549.296585513721  # P[kW] = T[Nm] * rpm / this value.

# Plot colors inspired by the user's reference image.
LOW_POWER_COLORS = {
    "S3_10": "#a92323",
    "S3_15": "#7f7f7f",
    "S3_25": "#8f8f8f",
    "S2": "#9b9b9b",
    "S1": "#6f6f6f",
    "OTHER": "#7f7f7f",
}
LOW_TORQUE_COLORS = {
    "S3_10": "#cf0000",
    "S3_15": "#8d8d8d",
    "S3_25": "#8a8a8a",
    "S2": "#7f7f7f",
    "S1": "#d10000",
    "OTHER": "#8f8f8f",
}
HIGH_POWER_COLORS = {
    "S3_10": "#1b36d9",
    "S3_15": "#2346e0",
    "S3_25": "#1237ff",
    "S2": "#7d7d7d",
    "S1": "#183cff",
    "OTHER": "#2850f0",
}
HIGH_TORQUE_COLORS = {
    "S3_10": "#c40000",
    "S3_15": "#cf4c4c",
    "S3_25": "#e52525",
    "S2": "#8f8f8f",
    "S1": "#d10c0c",
    "OTHER": "#d13d3d",
}


def number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        v = float(str(value).replace(",", "").strip())
        return v if math.isfinite(v) else None
    except (TypeError, ValueError):
        return None


def normalize_winding(value: str) -> str:
    s = str(value or "").strip().upper()
    if "LOW" in s or s in {"L", "저속", "저속권선"}:
        return "LOW"
    if "HIGH" in s or s in {"H", "고속", "고속권선"}:
        return "HIGH"
    return s or "LOW"


def normalize_rating(value: str) -> str:
    s = re.sub(r"\s+", " ", str(value or "S1").strip())
    upper = s.upper()
    if upper.startswith("S1"):
        return "S1 Cont."
    m = re.search(r"S2\s*([0-9]+)\s*MIN", upper)
    if m:
        return f"S2 {m.group(1)}min"
    m = re.search(r"S3\s*/\s*S6\s*([0-9]+)\s*%", upper)
    if m:
        return f"S3/S6 {m.group(1)}%"
    return s or "S1 Cont."


def rating_order(label: str) -> tuple[int, int, str]:
    s = label.upper()
    if s.startswith("S1"):
        return (0, 0, s)
    if s.startswith("S2"):
        return (1, 0, s)
    m = re.search(r"(\d+)\s*%", s)
    if m:
        return (2, -int(m.group(1)), s)
    return (3, 0, s)


def display_rating_order(label: str) -> tuple[int, int, str]:
    """Export layout order: S3 10/15/25%, S2, S1 (no missing curves invented)."""
    rating = normalize_rating(label).upper()
    m = re.search(r"(\d+)\s*%", rating)
    if rating.startswith('S3') and m:
        return 0, int(m.group(1)), rating
    if rating.startswith('S2'):
        minutes = re.search(r'S2\s*(\d+)', rating)
        return 1, int(minutes.group(1)) if minutes else 0, rating
    if rating.startswith('S1'):
        return 2, 0, rating
    return 3, 0, rating


def rating_bucket(label: str) -> str:
    rating = normalize_rating(label).upper()
    if rating.startswith("S1"):
        return "S1"
    if rating.startswith("S2"):
        return "S2"
    m = re.search(r"(\d+)\s*%", rating)
    if rating.startswith("S3") and m:
        return f"S3_{m.group(1)}"
    return re.sub(r"[^A-Z0-9]+", "_", rating)


def bucket_order(bucket: str) -> tuple[int, int, str]:
    if bucket.startswith("S3_"):
        try:
            return 0, int(bucket.split("_", 1)[1]), bucket
        except Exception:
            return 0, 999, bucket
    if bucket == "S2":
        return 1, 0, bucket
    if bucket == "S1":
        return 2, 0, bucket
    return 3, 0, bucket


def bucket_display(bucket: str, representative: str | None = None) -> str:
    if representative:
        return normalize_rating(representative)
    if bucket == "S1":
        return "S1 Cont."
    if bucket == "S2":
        return "S2"
    if bucket.startswith("S3_"):
        return f"S3/S6 {bucket.split('_', 1)[1]}%"
    return bucket.replace("_", " ")


def tidy_rows(rows: list[dict]) -> tuple[list[dict], list[str]]:
    clean = []
    errors = []
    for i, row in enumerate(rows, 1):
        w = normalize_winding(row.get("권선", "LOW"))
        rating = normalize_rating(row.get("운전조건", "S1"))
        rpm = number(row.get("RPM"))
        kw = number(row.get("출력(kW)"))
        nm = number(row.get("토크(N·m)"))
        if rpm is None and kw is None and nm is None:
            continue
        if rpm is None or rpm < 0:
            errors.append(f"{i}행: RPM이 없거나 음수입니다.")
            continue
        if (kw is None and nm is None) or (kw is not None and kw < 0) or (nm is not None and nm < 0):
            errors.append(f"{i}행: 출력/토크 값을 확인하십시오.")
            continue
        if rpm == 0 and kw is None:
            kw = 0.0
        clean.append({"권선": w, "운전조건": rating, "RPM": rpm,
                      "출력(kW)": kw, "토크(N·m)": nm})
    clean.sort(key=lambda d: (0 if d["권선"] == "LOW" else 1, d["권선"], rating_order(d["운전조건"]), d["RPM"]))
    seen = set()
    unique = []
    for row in clean:
        k = (row["권선"], row["운전조건"], row["RPM"])
        if k in seen:
            errors.append(f"중복 기준점: {k}")
            continue
        seen.add(k)
        unique.append(row)
    for row in unique:
        rpm, kw, nm = row["RPM"], row["출력(kW)"], row["토크(N·m)"]
        if rpm > 0 and kw is not None and nm is not None:
            derived = nm * rpm / RPM_TO_TORQUE
            if kw > .1 and abs(derived - kw) / kw > .1:
                errors.append(
                    f"{row['권선']} {row['운전조건']} {rpm:g}rpm: "
                    f"출력 {kw:g}kW / 토크 {nm:g}N·m 상호 오차가 10%를 초과합니다. "
                    "현재 그래프/계산표는 출력 값을 우선 사용합니다."
                )
    return unique, errors


def group_rows(rows: list[dict]) -> dict[tuple[str, str], list[dict]]:
    grouped = defaultdict(list)
    for row in rows:
        grouped[(row["권선"], row["운전조건"])].append(row)
    return dict(grouped)


def grouped_by_bucket(rows: list[dict]) -> dict[tuple[str, str], list[dict]]:
    raw = group_rows(rows)
    out: dict[tuple[str, str], list[dict]] = {}
    for (winding, rating), pts in raw.items():
        out[(winding, rating_bucket(rating))] = pts
    return out


def bucket_representatives(rows: list[dict]) -> dict[tuple[str, str], str]:
    reps: dict[tuple[str, str], str] = {}
    for row in rows:
        key = (row["권선"], rating_bucket(row["운전조건"]))
        reps.setdefault(key, normalize_rating(row["운전조건"]))
    return reps


def power_at_source(row: dict) -> float | None:
    kw = number(row.get("출력(kW)"))
    if kw is not None:
        return kw
    rpm, nm = number(row.get("RPM")), number(row.get("토크(N·m)"))
    if rpm is not None and nm is not None:
        return nm * rpm / RPM_TO_TORQUE
    return None


def calc_at(points: list[dict], rpm: float) -> tuple[float | None, float | None]:
    if not points or rpm < points[0]["RPM"] or rpm > points[-1]["RPM"]:
        return None, None
    for i, p in enumerate(points):
        if abs(rpm - p["RPM"]) < 1e-9:
            kw = power_at_source(p)
            if kw is None:
                return None, None
            if rpm == 0:
                return 0.0, number(p.get("토크(N·m)"))
            return kw, kw * RPM_TO_TORQUE / rpm
        if i + 1 < len(points) and p["RPM"] < rpm < points[i + 1]["RPM"]:
            q = points[i + 1]
            pkw, qkw = power_at_source(p), power_at_source(q)
            if pkw is None or qkw is None:
                return None, None
            frac = (rpm - p["RPM"]) / (q["RPM"] - p["RPM"])
            kw = pkw + frac * (qkw - pkw)
            return kw, kw * RPM_TO_TORQUE / rpm if rpm else None
    return None, None


def calc_switched_at(grouped: dict[tuple[str, str], list[dict]], bucket: str, rpm: float,
                     switch_rpm: int | None) -> tuple[float | None, float | None, str | None]:
    low_pts = grouped.get(("LOW", bucket))
    high_pts = grouped.get(("HIGH", bucket))
    prefer_high = bool(switch_rpm is not None and switch_rpm > 0 and rpm >= switch_rpm)
    order = [high_pts, low_pts] if prefer_high else [low_pts, high_pts]
    source_labels = ["HIGH", "LOW"] if prefer_high else ["LOW", "HIGH"]
    for pts, src in zip(order, source_labels):
        if pts:
            kw, nm = calc_at(pts, rpm)
            if kw is not None:
                return kw, nm, src
    return None, None, None


def rpm_grid(rows: list[dict], limit: int = 20000, step: int = 100, switch_rpm: int | None = None) -> list[float]:
    anchors = {float(r["RPM"]) for r in rows if 0 <= r["RPM"] <= limit}
    additional = {0.0, 0.6, 1.0, 10.0, 100.0, 500.0}
    regular = set(range(0, limit + 1, step))
    if switch_rpm is not None and 0 <= switch_rpm <= limit:
        anchors.add(float(switch_rpm))
    return sorted(anchors | {x for x in additional if x <= limit} | regular)


def style_color(winding: str, bucket: str, quantity: str) -> str:
    if winding == "LOW":
        return LOW_POWER_COLORS.get(bucket, LOW_POWER_COLORS["OTHER"]) if quantity == "power" else LOW_TORQUE_COLORS.get(bucket, LOW_TORQUE_COLORS["OTHER"])
    return HIGH_POWER_COLORS.get(bucket, HIGH_POWER_COLORS["OTHER"]) if quantity == "power" else HIGH_TORQUE_COLORS.get(bucket, HIGH_TORQUE_COLORS["OTHER"])


def make_plot(rows: list[dict], max_rpm: int, show_power: bool = True, show_torque: bool = True,
              switch_rpm: int | None = None, apply_switch: bool = True, show_raw_points: bool = True):
    grouped = grouped_by_bucket(rows)
    reps = bucket_representatives(rows)
    fig = go.Figure()
    show_switched = bool(apply_switch and switch_rpm is not None and switch_rpm > 0)

    fig.update_layout(
        paper_bgcolor="#ececec",
        plot_bgcolor="#efefef",
        height=700,
        hovermode="x unified",
        margin={"l": 85, "r": 80, "t": 70, "b": 90},
        xaxis={
            "title": {"text": "Spindle Speed (rpm)", "font": {"size": 28, "color": "#4b4b4b"}},
            "range": [0, max_rpm],
            "showgrid": True,
            "gridcolor": "#c9c9c9",
            "zeroline": False,
            "tickfont": {"size": 14, "color": "#444"},
        },
        yaxis={
            "title": {"text": "Torque (N·m)", "font": {"size": 18, "color": "#444"}},
            "rangemode": "tozero",
            "showgrid": True,
            "gridcolor": "#d3d3d3",
            "zeroline": False,
            "tickfont": {"size": 13, "color": "#444"},
        },
        yaxis2={
            "title": {"text": "Power (kW)", "font": {"size": 18, "color": "#444"}},
            "overlaying": "y",
            "side": "right",
            "rangemode": "tozero",
            "showgrid": False,
            "tickfont": {"size": 13, "color": "#444"},
        },
        legend={"orientation": "h", "y": -0.17, "x": 0, "font": {"size": 12}},
        title={"text": "(Electric Conversion)", "x": 0.5, "font": {"size": 18, "color": "#666"}},
    )

    if switch_rpm is not None and switch_rpm > 0:
        fig.add_vline(x=switch_rpm, line_dash="dash", line_width=1.2, line_color="#6b7280")
        fig.add_annotation(x=switch_rpm, y=0, yref="paper", text=f"{switch_rpm}", showarrow=False,
                           yshift=-35, font={"size": 14, "color": "#222"})

    if show_switched:
        buckets = sorted({b for _, b in grouped}, key=bucket_order)
        for bucket in buckets:
            label = bucket_display(bucket, reps.get(("LOW", bucket)) or reps.get(("HIGH", bucket)))
            xs = rpm_grid(rows, limit=max_rpm, step=max(25, int(max_rpm / 250) or 25), switch_rpm=switch_rpm)
            power_y, torque_y, source_winding = [], [], []
            for x in xs:
                kw, nm, src = calc_switched_at(grouped, bucket, x, switch_rpm)
                power_y.append(kw)
                torque_y.append(nm)
                source_winding.append(src)
            # determine line colors based on the dominant winding after switch behavior.
            # For reference-image-like styling: power from HIGH is blue, torque from HIGH is red,
            # LOW-side portions keep gray/red-gray styling.
            if show_power:
                colors = []
                last = None
                for src in source_winding:
                    c = style_color(src or "LOW", bucket, "power") if src else None
                    colors.append(c)
                    if c:
                        last = c
                color = style_color("HIGH" if any(s == "HIGH" for s in source_winding) else "LOW", bucket, "power")
                fig.add_trace(go.Scatter(
                    x=xs, y=power_y, mode="lines", name=f"{label} · Power",
                    line={"color": color, "width": 3}, yaxis="y2", connectgaps=False,
                    hovertemplate="RPM %{x:,.0f}<br>출력 %{y:.3f} kW<extra>%{fullData.name}</extra>"
                ))
            if show_torque:
                color = style_color("HIGH" if any(s == "HIGH" for s in source_winding) else "LOW", bucket, "torque")
                fig.add_trace(go.Scatter(
                    x=xs, y=torque_y, mode="lines", name=f"{label} · Torque",
                    line={"color": color, "width": 3}, yaxis="y", connectgaps=False,
                    hovertemplate="RPM %{x:,.0f}<br>토크 %{y:.3f} N·m<extra>%{fullData.name}</extra>"
                ))
        if show_raw_points:
            for (winding, bucket), pts in grouped.items():
                label = bucket_display(bucket, reps.get((winding, bucket)))
                xs = [p["RPM"] for p in pts]
                if show_power:
                    fig.add_trace(go.Scatter(
                        x=xs, y=[power_at_source(p) for p in pts], mode="markers", showlegend=False,
                        yaxis="y2", marker={"size": 6, "color": style_color(winding, bucket, "power")},
                        hovertemplate=f"원본 기준점 ({winding})<br>정격 {label}<br>RPM %{{x:,.0f}}<br>출력 %{{y:.3f}} kW<extra></extra>"
                    ))
                if show_torque:
                    fig.add_trace(go.Scatter(
                        x=xs, y=[calc_at(pts, p["RPM"])[1] for p in pts], mode="markers", showlegend=False,
                        yaxis="y", marker={"size": 5, "symbol": "square", "color": style_color(winding, bucket, "torque")},
                        hovertemplate=f"원본 기준점 ({winding})<br>정격 {label}<br>RPM %{{x:,.0f}}<br>토크 %{{y:.3f}} N·m<extra></extra>"
                    ))
    else:
        for (winding, bucket), pts in sorted(grouped.items(), key=lambda item: (item[0][0] != "LOW", bucket_order(item[0][1]))):
            label = bucket_display(bucket, reps.get((winding, bucket)))
            lower, upper = pts[0]["RPM"], pts[-1]["RPM"]
            xs = sorted({round(lower + (upper - lower) * j / 250, 9) for j in range(251)} | {p["RPM"] for p in pts})
            calculated = [calc_at(pts, x) for x in xs]
            source_xs = [p["RPM"] for p in pts]
            if show_power:
                fig.add_trace(go.Scatter(
                    x=xs, y=[item[0] for item in calculated], mode="lines",
                    name=f"{winding} {label} · Power", line={"color": style_color(winding, bucket, "power"), "width": 3},
                    hovertemplate="RPM %{x:,.0f}<br>출력 %{y:.3f} kW<extra>%{fullData.name}</extra>",
                    yaxis="y2", connectgaps=False))
                if show_raw_points:
                    fig.add_trace(go.Scatter(
                        x=source_xs, y=[power_at_source(p) for p in pts], mode="markers", showlegend=False,
                        yaxis="y2", marker={"size": 5, "color": style_color(winding, bucket, "power")},
                        hovertemplate="원본 기준점<br>RPM %{x:,.0f}<br>출력 %{y:.3f} kW<extra></extra>"))
            if show_torque:
                fig.add_trace(go.Scatter(
                    x=xs, y=[item[1] for item in calculated], mode="lines", name=f"{winding} {label} · Torque",
                    line={"color": style_color(winding, bucket, "torque"), "width": 2.6},
                    hovertemplate="RPM %{x:,.0f}<br>토크 %{y:.3f} N·m<extra>%{fullData.name}</extra>",
                    yaxis="y", connectgaps=False))
                if show_raw_points:
                    fig.add_trace(go.Scatter(
                        x=source_xs, y=[calc_at(pts, p["RPM"])[1] for p in pts], mode="markers", showlegend=False,
                        yaxis="y", marker={"size": 5, "color": style_color(winding, bucket, "torque"), "symbol": "square"},
                        hovertemplate="원본 기준점<br>RPM %{x:,.0f}<br>토크 %{y:.3f} N·m<extra></extra>"))
    return fig


def _a1(col: int, row: int) -> str:
    name = ""
    col += 1
    while col:
        col, part = divmod(col - 1, 26)
        name = chr(65 + part) + name
    return f"{name}{row}"


def _excel_formula(target_rpm_cell: str, anchor_start: int, anchor_end: int) -> str:
    r_range = f"'기준점 입력'!$C${anchor_start}:$C${anchor_end}"
    p_range = f"'기준점 입력'!$F${anchor_start}:$F${anchor_end}"
    a = target_rpm_cell
    m = f"MATCH({a},{r_range},1)"
    x1 = f"INDEX({r_range},{m})"
    x2 = f"INDEX({r_range},{m}+1)"
    y1 = f"INDEX({p_range},{m})"
    y2 = f"INDEX({p_range},{m}+1)"
    p_last = f"INDEX({p_range},ROWS({p_range}))"
    interpolated = f"{y1}+({a}-{x1})*({y2}-{y1})/({x2}-{x1})"
    return (f'=IF(OR({a}<MIN({r_range}),{a}>MAX({r_range})),"",'
            f'IF({a}=MAX({r_range}),{p_last},IFERROR({interpolated},"")))')


def generate_excel(rows: list[dict], model_name: str, max_rpm: int = 20000,
                   step: int = 100, switch_rpm: int | None = None, notes: str = "") -> bytes:
    rows, problems = tidy_rows(rows)
    if not rows:
        raise ValueError("유효한 모터 입력 데이터가 없습니다.")
    groups = group_rows(rows)
    bucketed = grouped_by_bucket(rows)
    reps = bucket_representatives(rows)
    grid = rpm_grid(rows, limit=max_rpm, step=step, switch_rpm=switch_rpm)
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True})
    wb.set_properties({"title": "Motor Spindle Power/Torque Diagram", "comments": "Based on user supplied motor catalog data"})
    sh_chart = wb.add_worksheet("통합 선도")
    sh_grid = wb.add_worksheet("RPM 계산표")
    sh_switch = wb.add_worksheet("권선절환 적용값")
    sh_input = wb.add_worksheet("기준점 입력")
    sh_help = wb.add_worksheet("사용 안내")

    dark = "#17253F"
    low_color, high_color = "#7a7a7a", "#2453ff"
    fmt_title = wb.add_format({"font_color": "white", "bold": True, "font_size": 17, "bg_color": dark, "valign": "vcenter"})
    fmt_head = wb.add_format({"bg_color": dark, "font_color": "#FFFFFF", "bold": True,
                              "align": "center", "valign": "vcenter", "border": 1, "border_color": "#CED6E0", "text_wrap": True})
    fmt_base = wb.add_format({"num_format": "0.00", "border": 1, "border_color": "#E2E8F0"})
    fmt_input = wb.add_format({"num_format": "0.000", "border": 1, "border_color": "#D1D5DB", "bg_color": "#FFF7E8"})
    fmt_formula = wb.add_format({"num_format": "0.000", "font_color": "#146C43", "border": 1, "border_color": "#E2E8F0"})
    fmt_rpm = wb.add_format({"num_format": "0.##", "bold": True, "bg_color": "#F1F5F9", "border": 1, "border_color": "#DFE6EE"})
    fmt_break = wb.add_format({"num_format": "0.##", "bold": True, "bg_color": "#FFF1A4", "border": 1, "border_color": "#DFE6EE"})
    fmt_power = wb.add_format({"num_format": "0.000", "border": 1, "border_color": "#DFE6EE", "font_color": "#1E40AF"})
    fmt_torque = wb.add_format({"num_format": "0.000", "border": 1, "border_color": "#DFE6EE", "font_color": "#9A3412"})
    fmt_sub = wb.add_format({"bold": True, "font_color": dark, "bg_color": "#E8EDF5", "align": "center"})
    fmt_txt = wb.add_format({"text_wrap": True, "valign": "top", "font_color": "#334155"})
    fmt_err = wb.add_format({"num_format": "0.0%", "font_color": "#B91C1C", "border": 1, "border_color": "#DFE6EE"})

    sh_input.merge_range("A1:G2", "모터 카탈로그 입력 기준점 · 수정 가능한 원본 값", fmt_title)
    sh_input.write("A4", "모터명")
    sh_input.write("B4", model_name)
    sh_input.write("D4", "축 최대 RPM")
    sh_input.write_number("E4", max_rpm)
    sh_input.write("A5", "주의: 노란 셀의 원본 출력을 수정하면 RPM 계산표와 그래프가 자동 재계산됩니다.")
    incols = ["권선", "정격", "RPM", "출력(kW) 입력", "토크(N·m) 입력", "계산용 출력(kW)", "출력/토크 상대오차"]
    for c, label in enumerate(incols):
        sh_input.write(6, c, label, fmt_head)
    sh_input.freeze_panes(7, 2)
    sh_input.set_column("A:A", 12)
    sh_input.set_column("B:B", 21)
    sh_input.set_column("C:G", 20)
    row_map = {}
    ir = 7
    for key, pts in groups.items():
        first = ir + 1
        for p in pts:
            sh_input.write(ir, 0, p["권선"], fmt_base)
            sh_input.write(ir, 1, p["운전조건"], fmt_base)
            sh_input.write_number(ir, 2, p["RPM"], fmt_input)
            if p["출력(kW)"] is not None:
                sh_input.write_number(ir, 3, p["출력(kW)"], fmt_input)
            else:
                sh_input.write_blank(ir, 3, None, fmt_input)
            if p["토크(N·m)"] is not None:
                sh_input.write_number(ir, 4, p["토크(N·m)"], fmt_input)
            else:
                sh_input.write_blank(ir, 4, None, fmt_input)
            exrow = ir + 1
            kw = power_at_source(p)
            f = f'=IF(ISNUMBER(D{exrow}),D{exrow},IF(AND(ISNUMBER(E{exrow}),C{exrow}>0),E{exrow}*C{exrow}/{RPM_TO_TORQUE},""))'
            sh_input.write_formula(ir, 5, f, fmt_formula, kw if kw is not None else "")
            df = f'=IF(AND(ISNUMBER(D{exrow}),ISNUMBER(E{exrow}),C{exrow}>0,D{exrow}>0),ABS(D{exrow}-E{exrow}*C{exrow}/{RPM_TO_TORQUE})/D{exrow},"")'
            sh_input.write_formula(ir, 6, df, fmt_err, abs((p["토크(N·m)"] or 0)*p["RPM"]/RPM_TO_TORQUE - (p["출력(kW)"] or 0)) / p["출력(kW)"] if p["출력(kW)"] and p["토크(N·m)"] and p["RPM"] else "")
            ir += 1
        row_map[key] = (first, ir)

    by_winding = {}
    for key in groups:
        by_winding.setdefault(key[0], []).append(key)
    winding_names = sorted(by_winding, key=lambda w: (w != 'LOW', w != 'HIGH', w))
    for w in winding_names:
        by_winding[w].sort(key=lambda key: display_rating_order(key[1]))

    series_meta = []
    blocks = []
    next_col = 1
    for index, w in enumerate(winding_names):
        keys = by_winding[w]
        power_begin = next_col
        power_cols = {}
        for key in keys:
            power_cols[key] = next_col
            next_col += 1
        torque_begin = next_col
        torque_cols = {}
        for key in keys:
            torque_cols[key] = next_col
            next_col += 1
        winding_end = next_col - 1
        blocks.append((w, power_begin, torque_begin, winding_end, keys))
        for key in keys:
            series_meta.append((key, power_cols[key], torque_cols[key]))
        if index < len(winding_names) - 1:
            next_col += 1

    final_col = max(1, next_col - 1)
    sh_grid.merge_range(0, 0, 1, final_col, f"{model_name} | 회전수별 출력·토크 계산표", fmt_title)
    sh_grid.write('A4', '최대 RPM')
    sh_grid.write_number('B4', max_rpm)
    sh_grid.write('C4', '계산 간격')
    sh_grid.write_number('D4', step)
    sh_grid.write('E4', '절환 RPM')
    sh_grid.write('F4', switch_rpm if switch_rpm is not None else '미지정')
    sh_grid.merge_range(5, 0, 5, final_col,
                        '각 권선은 출력(kW)을 먼저 모으고 다음에 토크(N·m)를 모았습니다. 기준점 밖은 공란이며 임의 외삽하지 않습니다.', fmt_txt)
    sh_grid.merge_range(6, 0, 7, 0, '주축 회전수', fmt_sub)
    sh_grid.write(8, 0, 'RPM', fmt_head)
    sh_grid.set_column(0, 0, 17)
    sh_grid.freeze_panes(9, 1)
    fmt_low = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': low_color, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#FFFFFF'})
    fmt_high = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': high_color, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#FFFFFF'})
    fmt_power_group = wb.add_format({'bold': True, 'font_color': '#1D4ED8', 'bg_color': '#DBEAFE', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#BACCE9'})
    fmt_torque_group = wb.add_format({'bold': True, 'font_color': '#9A3412', 'bg_color': '#FFEDD5', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#F7C9A2'})
    fmt_condition_kw = wb.add_format({'bold': True, 'font_color': '#1E40AF', 'bg_color': '#EFF6FF', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#CCD9EA', 'text_wrap': True})
    fmt_condition_nm = wb.add_format({'bold': True, 'font_color': '#9A3412', 'bg_color': '#FFF7ED', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#F2D4B5', 'text_wrap': True})
    sh_grid.set_row(6, 25)
    sh_grid.set_row(7, 25)
    sh_grid.set_row(8, 36)

    for w, pstart, tstart, end, keys in blocks:
        heading = fmt_low if w == 'LOW' else fmt_high
        sh_grid.merge_range(6, pstart, 6, end, f'{w} 권선', heading)
        if len(keys) > 1:
            sh_grid.merge_range(7, pstart, 7, tstart - 1, '출력 (kW)', fmt_power_group)
            sh_grid.merge_range(7, tstart, 7, end, '토크 (N·m)', fmt_torque_group)
        else:
            sh_grid.write(7, pstart, '출력 (kW)', fmt_power_group)
            sh_grid.write(7, tstart, '토크 (N·m)', fmt_torque_group)
        for idx, key in enumerate(keys):
            sh_grid.write(8, pstart + idx, key[1], fmt_condition_kw)
            sh_grid.write(8, tstart + idx, key[1], fmt_condition_nm)
        sh_grid.set_column(pstart, end, 17)

    if len(blocks) > 1:
        for left, right in zip(blocks, blocks[1:]):
            spacer_col = left[3] + 1
            sh_grid.set_column(spacer_col, spacer_col, 3.5)

    anchor_rpms = {r['RPM'] for r in rows}
    for n, rpm in enumerate(grid):
        excel_row = n + 10
        sh_grid.write_number(9 + n, 0, rpm, fmt_break if rpm in anchor_rpms or (switch_rpm is not None and rpm == switch_rpm) else fmt_rpm)
        for key, col_p, col_t in series_meta:
            source_points = groups[key]
            p_value, t_value = calc_at(source_points, rpm)
            a_start, a_end = row_map[key]
            formula = _excel_formula(f'$A{excel_row}', a_start, a_end)
            sh_grid.write_formula(9 + n, col_p, formula, fmt_power, p_value if p_value is not None else '')
            zero_t = next((number(pt.get('토크(N·m)')) for pt in source_points if pt['RPM'] == 0), None)
            p_cell = _a1(col_p, excel_row)
            if zero_t is not None:
                tq = f'=IF($A{excel_row}=0,{zero_t},IF(ISNUMBER({p_cell}),{p_cell}*{RPM_TO_TORQUE}/$A{excel_row},""))'
            else:
                tq = f'=IF($A{excel_row}=0,"",IF(ISNUMBER({p_cell}),{p_cell}*{RPM_TO_TORQUE}/$A{excel_row},""))'
            sh_grid.write_formula(9 + n, col_t, tq, fmt_torque, t_value if t_value is not None else '')
    sh_grid.autofilter(8, 0, 8 + len(grid), 0)

    # Switched / applied values sheet.
    sh_switch.merge_range('A1:K2', f'{model_name} | 권선절환 적용값', fmt_title)
    sh_switch.write('A4', '최대 RPM')
    sh_switch.write_number('B4', max_rpm)
    sh_switch.write('D4', '절환 RPM')
    sh_switch.write('E4', switch_rpm if switch_rpm is not None else '미지정')
    sh_switch.write('A5', '입력한 절환 RPM 기준으로 LOW/HIGH를 자동 선택한 적용값입니다. 절환 RPM이 없으면 사용 가능한 곡선을 우선 표시합니다.', fmt_txt)
    sh_switch.write(7, 0, 'RPM', fmt_head)
    sh_switch.freeze_panes(8, 1)
    sh_switch.set_column(0, 0, 16)
    buckets = sorted({bucket for _, bucket in bucketed}, key=bucket_order)
    col = 1
    switched_meta = []
    for idx, bucket in enumerate(buckets):
        label = bucket_display(bucket, reps.get(('LOW', bucket)) or reps.get(('HIGH', bucket)))
        sh_switch.merge_range(6, col, 6, col + 1, label, fmt_sub)
        sh_switch.write(7, col, '출력 (kW)', fmt_condition_kw)
        sh_switch.write(7, col + 1, '토크 (N·m)', fmt_condition_nm)
        sh_switch.set_column(col, col + 1, 16)
        switched_meta.append((bucket, label, col, col + 1))
        col += 2
    for n, rpm in enumerate(grid):
        sh_switch.write_number(8 + n, 0, rpm, fmt_break if rpm in anchor_rpms or (switch_rpm is not None and rpm == switch_rpm) else fmt_rpm)
        for bucket, label, col_p, col_t in switched_meta:
            kw, nm, src = calc_switched_at(bucketed, bucket, rpm, switch_rpm)
            sh_switch.write(8 + n, col_p, kw, fmt_power)
            sh_switch.write(8 + n, col_t, nm, fmt_torque)

    chart = wb.add_chart({"type": "scatter", "subtype": "straight"})
    if switch_rpm is not None and switch_rpm > 0:
        source_sheet_name = '권선절환 적용값'
        for bucket, label, col_p, col_t in switched_meta:
            color_power = style_color('HIGH' if ('HIGH', bucket) in bucketed else 'LOW', bucket, 'power')
            color_torque = style_color('HIGH' if ('HIGH', bucket) in bucketed else 'LOW', bucket, 'torque')
            chart.add_series({"name": f"{label} · Power",
                              "categories": [source_sheet_name, 8, 0, 8 + len(grid) - 1, 0],
                              "values": [source_sheet_name, 8, col_p, 8 + len(grid) - 1, col_p],
                              "line": {"color": color_power, "width": 2.5},
                              "marker": {"type": "none"}, "y2_axis": True})
            chart.add_series({"name": f"{label} · Torque",
                              "categories": [source_sheet_name, 8, 0, 8 + len(grid) - 1, 0],
                              "values": [source_sheet_name, 8, col_t, 8 + len(grid) - 1, col_t],
                              "line": {"color": color_torque, "width": 2.5},
                              "marker": {"type": "none"}, "y2_axis": False})
    else:
        for key, col_p, col_t in series_meta:
            bucket = rating_bucket(key[1])
            color_power = style_color(key[0], bucket, 'power')
            color_torque = style_color(key[0], bucket, 'torque')
            chart.add_series({"name": f"{key[0]} {key[1]} · Power",
                              "categories": ['RPM 계산표', 9, 0, 9 + len(grid) - 1, 0],
                              "values": ['RPM 계산표', 9, col_p, 9 + len(grid) - 1, col_p],
                              "line": {"color": color_power, "width": 2.2},
                              "marker": {"type": "none"}, "y2_axis": True})
            chart.add_series({"name": f"{key[0]} {key[1]} · Torque",
                              "categories": ['RPM 계산표', 9, 0, 9 + len(grid) - 1, 0],
                              "values": ['RPM 계산표', 9, col_t, 9 + len(grid) - 1, col_t],
                              "line": {"color": color_torque, "width": 2.2},
                              "marker": {"type": "none"}, "y2_axis": False})
    chart.set_title({"name": f"{model_name} | SPINDLE PERFORMANCE DIAGRAM"})
    chart.set_x_axis({"name": "Spindle Speed (rpm)", "min": 0, "max": max_rpm, "major_gridlines": {"visible": True}})
    chart.set_y_axis({"name": "Torque (N·m)", "min": 0, "major_gridlines": {"visible": True}})
    chart.set_y2_axis({"name": "Power (kW)", "min": 0, "major_gridlines": {"visible": False}})
    chart.set_legend({"position": "bottom", "font": {"size": 8}})
    chart.set_size({"width": 1100, "height": 600})
    chart.set_style(10)
    sh_chart.merge_range("A1:P2", "SPINDLE PERFORMANCE DIAGRAM | 출력·토크 통합 선도", fmt_title)
    sh_chart.write("A4", "모터명")
    sh_chart.write("B4", model_name)
    sh_chart.write("D4", "최대 RPM")
    sh_chart.write_number("E4", max_rpm)
    sh_chart.write("G4", "절환 RPM")
    sh_chart.write("H4", switch_rpm if switch_rpm is not None else '미지정')
    sh_chart.write("A5", "권선절환 RPM이 입력되면 [권선절환 적용값] 시트를 기준으로 그래프를 생성합니다.")
    sh_chart.insert_chart("B7", chart)
    sh_chart.set_column("A:A", 16)
    sh_chart.set_column("B:P", 12)

    instructions = [
        "원본 캡처의 RPM/출력/토크를 AI로 읽고 사람이 검토해야 합니다. OCR 결과가 맞는지 반드시 확인하십시오.",
        "[기준점 입력]의 연한 노랑 셀(RPM/출력/토크)은 카탈로그 원본 수치입니다. F 열은 그래프 계산에 쓰는 출력입니다.",
        "계산 방법: 기준점 사이에서는 출력(kW)을 선형 보간하고 T(N·m) = 9549.2965855 × P(kW) / rpm으로 토크를 계산합니다.",
        "카탈로그에 출력/토크가 둘 다 있더라도 반올림으로 다를 수 있으므로 출력(kW)을 기준으로 계산합니다. G열 오차를 확인하십시오.",
        "원본 기준점의 최소/최대 RPM을 벗어나는 값은 임의로 외삽하지 않고 비웁니다.",
        "[RPM 계산표]는 LOW 출력→LOW 토크→HIGH 출력→HIGH 토크 순서입니다. 없는 정격은 생성하지 않습니다.",
        "[권선절환 적용값]은 사용자가 입력한 절환 RPM에 따라 LOW/HIGH 중 사용 가능한 값을 자동 선택한 결과입니다.",
        "통합 선도 차트는 절환 RPM이 있으면 [권선절환 적용값], 없으면 [RPM 계산표]를 기준으로 작성됩니다.",
        "원본 행 추가/삭제 또는 정격별 데이터 행 수 변경 시에는 앱에서 재생성하는 것이 안전합니다.",
        "저장된 엑셀은 로컬 파일이며, 사용자 확인 없이 모터 성능을 보증하는 자료로 사용하지 마십시오.",
    ]
    sh_help.merge_range("A1:H2", "사용 안내 · 계산 기준", fmt_title)
    sh_help.set_column("A:A", 5)
    sh_help.set_column("B:B", 108)
    for i, line in enumerate(instructions, 4):
        sh_help.write_number(i, 0, i - 3)
        sh_help.write(i, 1, line, fmt_txt)
        sh_help.set_row(i, 35)
    if notes:
        sh_help.write(17, 1, "사용자/AI 참고 사항: " + notes, fmt_txt)
    if problems:
        sh_help.write(19, 1, "검토 경고: " + " / ".join(problems[:30]), fmt_txt)
    wb.close()
    return out.getvalue()


def make_sample() -> tuple[list[dict], str]:
    sample = {
        ("LOW", "S1 Cont."): {
            "rpms": [0,1415,1650,2200,2300,2400,4500,6000],
            "kw": [0,6.8,7.9,10.5,11,11,11,11],
            "nm": [None,45.7,45.7,45.7,45.7,43.8,23.3,17.5]},
        ("LOW", "S2 15min"): {
            "rpms": [0,1415,1650,2200,2300,2400,4500,6000],
            "kw": [0,8.8,10.3,13.8,14.4,15,15,15],
            "nm": [None,59.7,59.7,59.7,59.7,59.7,31.8,23.9]},
        ("LOW", "S3/S6 25%"): {
            "rpms": [0,1415,1650,2200,2300,2400,4500,6000],
            "kw": [0,11.9,13.9,18.5,18.5,18.5,18.5,15],
            "nm": [None,80.3,80.3,80.3,76.8,73.6,39.3,23.9]},
        ("LOW", "S3/S6 15%"): {
            "rpms": [0,1415,1650,2200], "kw": [0,15.9,18.5,18.5],
            "nm": [None,107,107,80.3]},
        ("LOW", "S3/S6 10%"): {
            "rpms": [0,1415,1650], "kw": [0,18.5,18.5], "nm": [None,125,107]},
        ("HIGH", "S1 Cont."): {
            "rpms": [0,4500,6000,10000,12000,14000,18000,20000],
            "kw": [0,11.3,15,15,15,15,11.7,10],
            "nm": [None,23.9,23.9,14.3,11.9,10.2,6.19,4.77]},
        ("HIGH", "S2 10min"): {
            "rpms": [0,4500,6000,10000,12000,14000,18000,20000],
            "kw": [0,13.9,18.5,18.5,18.5,18.5,14.8,13],
            "nm": [None,29.4,29.4,17.7,14.7,12.6,7.87,6.21]},
        ("HIGH", "S3/S6 25%"): {
            "rpms": [0,4500,6000,10000,12000,14000],
            "kw": [0,19.5,26,26,22.3,18.5],
            "nm": [None,41.4,41.4,24.8,17.7,12.6]},
    }
    rows = []
    for (w, rating), entry in sample.items():
        for rpm, kw, nm in zip(entry["rpms"], entry["kw"], entry["nm"]):
            rows.append({"권선": w, "운전조건": rating, "RPM": rpm,
                         "출력(kW)": kw, "토크(N·m)": nm})
    return rows, "이전 대화에 첨부된 LOW/HIGH 모터 카탈로그를 수동 전사한 예제. 실제 적용 전 원본 대조 필요."
