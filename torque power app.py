"""Standalone Streamlit app for motor specs, switching curves, Excel exports.
Run: streamlit run "torque power app.py"
No motor_core.py dependency: calculation and Excel code is included here.
"""
from __future__ import annotations

import base64
import io
import hashlib
import json
import re
import uuid
import random
import time
from datetime import datetime, timezone
from pathlib import Path

from streamlit_paste_button import paste_image_button
import requests
import streamlit as st


# ------------------------------------------------------------------
# Internal motor calculation and Excel module.
# Kept INSIDE this file to avoid version mismatches between
# torque power app.py and motor_core.py in GitHub deployments.
# Do not import motor_core.py: it is intentionally not required.
# ------------------------------------------------------------------
import io
import math
import re
from collections import defaultdict
from typing import Any

import plotly.graph_objects as go
import xlsxwriter

RPM_TO_TORQUE = 9549.296585513721  # P[kW] = T[Nm] * rpm / this value.

# Both the Streamlit chart and Excel chart use the SAME series and palette.
EXTREME_POWER_COLOR = "#134BFF"   # top and bottom power curves
EXTREME_TORQUE_COLOR = "#D00000"  # top and bottom torque curves
MID_GRAY_COLORS = ["#787878", "#949494", "#ACACAC", "#6F6F6F", "#B2B2B2"]


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
    if switch_rpm is None or switch_rpm <= 0:
        for winding in ("LOW", "HIGH"):
            points = grouped.get((winding, bucket))
            if points:
                kw, nm = calc_at(points, rpm)
                if kw is not None:
                    return kw, nm, winding
        return None, None, None
    winding = "LOW" if rpm < switch_rpm else "HIGH"
    points = grouped.get((winding, bucket))
    if not points:
        return None, None, None
    kw, nm = calc_at(points, rpm)
    return kw, nm, winding if kw is not None else None


def rpm_grid(rows: list[dict], limit: int = 20000, step: int = 100, switch_rpm: int | None = None) -> list[float]:
    anchors = {float(r["RPM"]) for r in rows if 0 <= r["RPM"] <= limit}
    additional = {0.0, 0.6, 1.0, 10.0, 100.0, 500.0, float(limit)}
    regular = set(range(0, limit + 1, step))
    if switch_rpm is not None and 0 <= switch_rpm <= limit:
        anchors.add(float(switch_rpm))
    return sorted(anchors | {x for x in additional if x <= limit} | regular)


def extreme_buckets(grouped: dict, winding: str, quantity: str) -> set[str]:
    """Identify highest and lowest duty curves for each winding.

    Rank using actual maxima of available source curves. The intermediate
    conditions receive gray. Works even when HIGH has fewer duties than LOW.
    """
    profiles = []
    for (w, bucket), pts in grouped.items():
        if w != winding:
            continue
        if quantity == "power":
            values = [power_at_source(p) for p in pts]
        else:
            values = [calc_at(pts, p["RPM"])[1] for p in pts]
        good = [v for v in values if v is not None and math.isfinite(v)]
        if good:
            profiles.append((max(good), bucket_order(bucket), bucket))
    if not profiles:
        return set()
    # When rated maxima tie (common in 10%/15% duty curves), prefer
    # the more severe duty for the TOP curve and less severe for BOTTOM.
    top = max(profiles, key=lambda a: (a[0], -a[1][0], -a[1][1]))
    bottom = min(profiles, key=lambda a: (a[0], -a[1][0], -a[1][1]))
    return {top[2], bottom[2]}


def style_color(winding: str, bucket: str, quantity: str, grouped: dict | None = None) -> str:
    if grouped is not None and bucket in extreme_buckets(grouped, winding, quantity):
        return EXTREME_POWER_COLOR if quantity == "power" else EXTREME_TORQUE_COLOR
    # Constant neutral palette across WEB / Excel. Duty-based gray distinctions.
    order = bucket_order(bucket)
    idx = (order[1] if order[0] == 0 else (3 if order[0] == 1 else 4)) % len(MID_GRAY_COLORS)
    return MID_GRAY_COLORS[idx]


def diagram_series(rows: list[dict], max_rpm: int, switch_rpm: int | None = None,
                   step: int = 50) -> tuple[list[float], list[dict]]:
    """The single source of chart values for Plotly and Excel.

    When switching, LOW and HIGH remain independent SERIES. Both are sampled
    at the switching RPM, but never connected by an artificial diagonal.
    Missing source points are blank. No extrapolation occurs.
    """
    grouped = grouped_by_bucket(rows)
    reps = bucket_representatives(rows)
    xs = rpm_grid(rows, limit=int(max_rpm), step=step, switch_rpm=switch_rpm)
    xs = sorted(set(xs) | {float(max_rpm)})
    switched = switch_rpm is not None and switch_rpm > 0
    series = []
    for (winding, bucket), pts in sorted(grouped.items(), key=lambda item:
                                          (item[0][0] != "LOW", bucket_order(item[0][1]))):
        rating = bucket_display(bucket, reps.get((winding, bucket)))
        values = []
        for rpm in xs:
            if switched and ((winding == "LOW" and rpm > switch_rpm) or
                             (winding == "HIGH" and rpm < switch_rpm)):
                values.append((None, None))
            else:
                values.append(calc_at(pts, rpm))
        for quantity, idx, unit in (("power", 0, "kW"), ("torque", 1, "N·m")):
            ys = [v[idx] for v in values]
            if not any(y is not None for y in ys):
                continue
            series.append({
                "winding": winding, "bucket": bucket, "rating": rating,
                "quantity": quantity, "unit": unit,
                "name": f"{winding} {rating} · {'출력' if quantity == 'power' else '토크'}",
                "color": style_color(winding, bucket, quantity, grouped),
                "x": xs, "y": ys,
            })
    return xs, series


def axis_ranges(series: list[dict]) -> tuple[int, int]:
    tq_max = max((v for s in series if s["quantity"] == "torque"
                  for v in s["y"] if v is not None), default=1)
    pw_max = max((v for s in series if s["quantity"] == "power"
                  for v in s["y"] if v is not None), default=1)
    # Keep exactly matching axis limits in Plotly and Excel.
    tq_top = max(10, math.ceil(tq_max / 50) * 50)
    pw_top = max(5, math.ceil(pw_max / 5) * 5)
    return tq_top, pw_top


def make_plot(rows: list[dict], max_rpm: int, show_power: bool = True,
              show_torque: bool = True, switch_rpm: int | None = None,
              apply_switch: bool = True, show_raw_points: bool = False):
    active_switch = int(switch_rpm) if apply_switch and switch_rpm and switch_rpm > 0 else None
    xs, series = diagram_series(rows, max_rpm=max_rpm, switch_rpm=active_switch)
    torque_axis_max, power_axis_max = axis_ranges(series)
    fig = go.Figure()
    fig.update_layout(
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", height=690,
        hovermode="x unified", margin={"l": 85, "r": 80, "t": 60, "b": 100},
        xaxis={"title": "Spindle Speed (rpm)", "range": [0, max_rpm],
               "gridcolor": "#D9D9D9", "showgrid": True,
               "tickfont": {"size": 12, "color": "#454545"}},
        yaxis={"title": "Torque (N·m)", "range": [0, torque_axis_max],
               "gridcolor": "#D9D9D9", "showgrid": True, "zeroline": False},
        yaxis2={"title": "Power (kW)", "range": [0, power_axis_max],
                "overlaying": "y", "side": "right", "showgrid": False,
                "zeroline": False},
        legend={"orientation": "h", "y": -0.17, "x": 0, "font": {"size": 11}},
        title={"text": "(Electric Conversion)", "x": .5},
    )
    for curve in series:
        quantity = curve["quantity"]
        if quantity == "power" and not show_power:
            continue
        if quantity == "torque" and not show_torque:
            continue
        fig.add_trace(go.Scatter(
            x=curve["x"], y=curve["y"], mode="lines", connectgaps=False,
            name=curve["name"],
            line={"color": curve["color"], "width": 2.7},
            yaxis="y2" if quantity == "power" else "y",
            hovertemplate=("RPM %{x:,.0f}<br>" +
                           ("출력" if quantity == "power" else "토크") +
                           " %{y:.3f} " + curve["unit"] + "<extra>%{fullData.name}</extra>"),
        ))
    if active_switch:
        fig.add_vline(x=active_switch, line_color="#606060", line_dash="dash", line_width=1)
        fig.add_annotation(x=active_switch, y=1, yref="paper", text=f"권선절환 {active_switch:,} rpm",
                           showarrow=False, xanchor="left", yshift=10)
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
    wb.set_properties({"title": APP_TITLE, "comments": "Based on user supplied motor catalog data"})
    sh_chart = wb.add_worksheet("통합 선도")
    sh_grid = wb.add_worksheet("RPM 계산표")
    sh_switch = wb.add_worksheet("권선절환 적용값")
    sh_input = wb.add_worksheet("기준점 입력")
    sh_help = wb.add_worksheet("사용 안내")
    sh_support = wb.add_worksheet("차트 공통데이터")
    chart_xs, chart_curves = diagram_series(rows, max_rpm=max_rpm, switch_rpm=switch_rpm)
    tq_axis_max, pw_axis_max = axis_ranges(chart_curves)
    sh_support.write(0, 0, "RPM")
    for cc, curve in enumerate(chart_curves, 1):
        sh_support.write(0, cc, curve["name"])
    for rr, rpm in enumerate(chart_xs, 1):
        sh_support.write_number(rr, 0, rpm)
        for cc, curve in enumerate(chart_curves, 1):
            value = curve["y"][rr - 1]
            if value is not None:
                sh_support.write_number(rr, cc, value)
    sh_support.set_column(0, 0, 15)
    sh_support.set_column(1, len(chart_curves), 22)


    dark = UI_DARK_BLUE
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
    sh_grid.write('A4', '사용 최대 RPM')
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
    sh_switch.write('A4', '사용 최대 RPM')
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

    # Chart support values are formula-linked to editable catalog inputs.
    # This preserves Excel chart updates when users adjust existing source rows.
    support_column_by_key = {
        (c["winding"], c["bucket"], c["quantity"]): idx
        for idx, c in enumerate(chart_curves, 1)
    }
    row_keys_by_bucket = {(key[0], rating_bucket(key[1])): key for key in row_map}
    for idx, curve in enumerate(chart_curves, 1):
        group_key = row_keys_by_bucket[(curve["winding"], curve["bucket"])]
        a_start, a_end = row_map[group_key]
        original_points = groups[group_key]
        zero_t = next((number(pt.get("토크(N·m)")) for pt in original_points
                       if pt["RPM"] == 0), None)
        power_col = support_column_by_key[(curve["winding"], curve["bucket"], "power")]
        for rr, (rpm, val) in enumerate(zip(chart_xs, curve["y"]), 1):
            if val is None:
                continue
            sheet_row = rr + 1
            if curve["quantity"] == "power":
                formula = _excel_formula(f'$A{sheet_row}', a_start, a_end)
            else:
                pc = _a1(power_col, sheet_row)
                if zero_t is not None:
                    formula = (f'=IF($A{sheet_row}=0,{zero_t},'
                               f'IF(ISNUMBER({pc}),{pc}*{RPM_TO_TORQUE}/$A{sheet_row},""))')
                else:
                    formula = (f'=IF($A{sheet_row}=0,"",'
                               f'IF(ISNUMBER({pc}),{pc}*{RPM_TO_TORQUE}/$A{sheet_row},""))')
            sh_support.write_formula(rr, idx, formula, None, val)

    chart = wb.add_chart({"type": "scatter", "subtype": "straight"})
    chart.show_blanks_as("gap")
    # Same series, x/y samples and colors as Streamlit, irrespective of Excel grid step.
    for col, curve in enumerate(chart_curves, 1):
        chart.add_series({
            "name": curve["name"],
            "categories": ["차트 공통데이터", 1, 0, len(chart_xs), 0],
            "values": ["차트 공통데이터", 1, col, len(chart_xs), col],
            "line": {"color": curve["color"], "width": 2.4},
            "marker": {"type": "none"},
            "y2_axis": curve["quantity"] == "power",
        })
    if switch_rpm is not None and 0 < switch_rpm <= max_rpm:
        # Match the dashed switch marker shown by Plotly.
        sx_col = len(chart_curves) + 2
        sy_col = len(chart_curves) + 3
        sh_support.write(0, sx_col, "절환선 RPM")
        sh_support.write(0, sy_col, "절환선 토크축")
        for marker_row, yval in ((1, 0), (2, tq_axis_max)):
            sh_support.write_number(marker_row, sx_col, switch_rpm)
            sh_support.write_number(marker_row, sy_col, yval)
        chart.add_series({
            "name": "권선절환", "categories": ["차트 공통데이터", 1, sx_col, 2, sx_col],
            "values": ["차트 공통데이터", 1, sy_col, 2, sy_col],
            "line": {"color": "#606060", "dash_type": "dash", "width": 1.0},
            "marker": {"type": "none"}, "y2_axis": False,
        })
    chart.set_title({"name": f"{model_name} | SPINDLE PERFORMANCE DIAGRAM"})
    chart.set_x_axis({"name": "Spindle Speed (rpm)", "min": 0, "max": max_rpm, "major_gridlines": {"visible": True}})
    chart.set_y_axis({"name": "Torque (N·m)", "min": 0, "max": tq_axis_max, "major_gridlines": {"visible": True}})
    chart.set_y2_axis({"name": "Power (kW)", "min": 0, "max": pw_axis_max, "major_gridlines": {"visible": False}})
    chart.set_legend({"position": "bottom", "font": {"size": 8}})
    chart.set_size({"width": 1100, "height": 600})
    chart.set_style(10)
    sh_chart.merge_range("A1:P2", f"{APP_TITLE} | 출력·토크 통합 선도", fmt_title)
    sh_chart.write("A4", "모터명")
    sh_chart.write("B4", model_name)
    sh_chart.write("D4", "사용 최대 RPM")
    sh_chart.write_number("E4", max_rpm)
    sh_chart.write("G4", "절환 RPM")
    sh_chart.write("H4", switch_rpm if switch_rpm is not None else '미지정')
    sh_chart.write("A5", "웹 그래프와 동일한 [차트 공통데이터] 시트로 그래프를 생성합니다. 권선절환 구간은 분리합니다.")
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
        "[통합 선도]는 웹 그래프와 동일한 [차트 공통데이터] 기준이며, LOW와 HIGH를 절환점에서 분리하여 표시합니다. 기준점의 기존 값 수정 시 차트도 자동 재계산됩니다.",
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


APP_TITLE = "스핀들 주축출력선도 자동 생성 플랫폼"
LOGO_PNG_BASE64 = "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"
st.set_page_config(page_title=APP_TITLE, page_icon="📈", layout="wide")
DATA_DIR = Path("data/motor_specs")
DATA_DIR.mkdir(parents=True, exist_ok=True)



# Readability-focused interface theme. Motor analysis logic is unchanged.
UI_DARK_BLUE = "#164879"


def apply_ui_styles():
    st.markdown(
        """
        <style>
        /* Neutral, readable internal-tool UI; retains Streamlit's functionality. */
        .stApp {
            background: #F7F9FC;
        }
        section[data-testid="stSidebar"] {
            background: #F2F5F9;
            border-right: 1px solid #DAE2EC;
        }
        /* Clean section headings and subtle divider. */
        h1, h2, h3 { letter-spacing: -0.025em; }
        h1 { color: #142B46 !important; }
        h2, h3 { color: #1D3653 !important; }
        /* Progress metrics: white cards with high-contrast values. */
        [data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #D8E1EC;
            border-radius: 10px;
            padding: 12px 14px;
        }
        [data-testid="stMetricValue"] { color: #163C68 !important; }
        [data-testid="stMetricLabel"] { color: #465D75 !important; }
        /* Filled primary actions only; secondary actions remain legible on white. */
        .stButton button[kind="primary"],
        .stDownloadButton button[kind="primary"] {
            background-color: #15599B !important;
            border: 1px solid #15599B !important;
            color: white !important;
            border-radius: 8px;
            font-weight: 650;
        }
        .stButton button[kind="primary"]:hover,
        .stDownloadButton button[kind="primary"]:hover {
            background-color: #0E477E !important;
            border-color: #0E477E !important;
            color: #FFFFFF !important;
        }
        .stButton button[kind="secondary"],
        .stDownloadButton button[kind="secondary"] {
            background: #FFFFFF !important;
            border: 1px solid #9BAFC4 !important;
            color: #183D65 !important;
            border-radius: 8px;
            font-weight: 600;
        }
        .stButton button[kind="secondary"]:hover,
        .stDownloadButton button[kind="secondary"]:hover {
            background: #EBF3FB !important;
            border-color: #15599B !important;
            color: #123D68 !important;
        }
        /* Consistent disabled state instead of white-on-pale-blue. */
        .stButton button:disabled,
        .stDownloadButton button:disabled {
            background: #E9EEF4 !important;
            border-color: #D1DAE4 !important;
            color: #657588 !important;
            opacity: 1 !important;
        }
        /* Tabs use underline rather than dark filled tiles. */
        button[role="tab"] {
            background: transparent !important;
            color: #455B72 !important;
            border-radius: 0 !important;
            font-weight: 600;
            border-bottom: 2px solid transparent;
        }
        button[role="tab"][aria-selected="true"] {
            color: #124F8E !important;
            border-bottom: 3px solid #15599B !important;
            font-weight: 750;
        }
        button[role="tab"]:hover {
            color: #124F8E !important;
            background: #EDF3FA !important;
        }
        /* Input borders and data tables. */
        .stTextInput input, .stNumberInput input, .stTextArea textarea {
            border-color: #B8C6D7 !important;
            border-radius: 7px;
        }
        [data-testid="stDataFrame"] {
            border: 1px solid #DCE4ED;
            border-radius: 8px;
            overflow: hidden;
        }
        [data-testid="stExpander"] {
            border-color: #DAE2EC;
            border-radius: 9px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


SYSTEM_PROMPT = """
You are a technical transcription assistant for industrial motor PERFORMANCE TABLES.
The user uploads one or more screen captures of motor power/torque/RPM tables and charts.
Your job is ONLY to TRANSCRIBE values that are actually printed in the source.
Extract ALL winding sections separately (LOW / HIGH, if shown) and EACH duty curve
(S1 continuous, S2 n minutes, S3/S6 XX%, etc.) separately.

IMPORTANT RULES:
1) Prefer the NUMERIC specification tables rather than approximating numbers from plotted curves.
2) Do not fill in any missing/blank/dashed cell. Use null for absent power or torque.
3) Preserve every printed RPM anchor, including 0 rpm for zero power if 0 is printed.
4) At 0 rpm torque is NOT automatically 0. Use null if the catalog doesn't specify torque.
5) Preserve the distinction among 10%, 15%, 25%, S1, S2 and LOW/HIGH.
6) Ignore maximum current (A), winding resistance (mOhm) and power factor (percent)
   as curve points. Put such supplementary information in notes if relevant.
7) Do not calculate or infer the winding changeover speed if not explicitly stated.
8) If the image is only an Excel example without a motor numeric specification table,
   return curves=[] and a note that extraction is not possible.
9) If a number is unreadable, omit that cell (null) rather than guess.
10) Keep the motor_name empty if not displayed. maximum_rpm only if explicitly stated.
Return ONLY a JSON object following this exact example shape:
{
  "motor_name": "motor designation or empty string",
  "maximum_rpm": 20000,
  "curves": [
    {
      "winding": "LOW",
      "rating": "S1 Cont.",
      "points": [
        {"rpm": 0, "power_kw": 0, "torque_nm": null},
        {"rpm": 1415, "power_kw": 6.8, "torque_nm": 45.7}
      ]
    }
  ],
  "notes": []
}
The values above are an example of FORMAT ONLY. Never copy example values unless present in image.
Use JSON null for unknown numeric values. maximum_rpm should be null when not explicitly specified.
All numbers must be JSON numbers, not formatted strings. No prose, no markdown.
"""


def secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, ""))
    except (FileNotFoundError, AttributeError, KeyError):
        return ""


def record_payload(rows, motor_name, rpm_limit, notes, record_id, company="", voltage="", switch_rpm=0) -> dict:
    return {
        "record_id": record_id,
        "motor_name": motor_name,
        "motor_company": company,
        "voltage": voltage,
        "switch_rpm": int(switch_rpm or 0),
        "max_rpm": rpm_limit,
        "operating_max_rpm": max(100, min(rpm_limit, int(st.session_state.get("operating_rpm_input", rpm_limit)))),
        "notes": notes,
        "saved_at": datetime.now(timezone.utc).isoformat(),
        "rows": rows,
        "source_method": "motor-catalog-screenshot / manual-review",
        "schema_version": 1,
    }


def slug(s: str) -> str:
    s = re.sub(r"[^0-9a-zA-Z가-힣_-]+", "_", s.strip())
    return (s[:64] or "motor")


def local_path(payload: dict) -> Path:
    return DATA_DIR / f"{slug(payload['motor_name'])}_{payload['record_id']}.json"


def save_local(payload: dict):
    local_path(payload).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def github_config() -> tuple[str, str, str]:
    return secret("GITHUB_TOKEN"), secret("GITHUB_REPO"), secret("GITHUB_BRANCH") or "main"


def github_request(method, url, token, **kwargs):
    headers = {"Authorization": f"Bearer {token}",
               "Accept": "application/vnd.github+json",
               "X-GitHub-Api-Version": "2022-11-28"}
    return requests.request(method, url, headers=headers, timeout=25, **kwargs)


def github_save(payload: dict) -> str:
    token, repo, branch = github_config()
    if not token or not repo:
        return "GitHub 자동 저장 미설정 (현재 앱 서버 및 JSON 다운로드만 가능)"
    repo_info = github_request("GET", f"https://api.github.com/repos/{repo}", token)
    if repo_info.status_code != 200:
        raise RuntimeError(f"GitHub 저장소 접근 실패 HTTP {repo_info.status_code}")
    if not repo_info.json().get("private", False):
        raise RuntimeError("모터 사양의 공개 유출을 막기 위해 Public 저장소에는 자동 저장하지 않습니다. Private 저장소를 지정하십시오.")
    relpath = f"motor_specs/{slug(payload['motor_name'])}_{payload['record_id']}.json"
    url = f"https://api.github.com/repos/{repo}/contents/{relpath}"
    before = github_request("GET", url, token, params={"ref": branch})
    existing = before.json().get("sha") if before.status_code == 200 else None
    if before.status_code not in (200, 404):
        raise RuntimeError(f"GitHub 조회 실패 HTTP {before.status_code}: {before.text[:300]}")
    raw = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    body = {"message": f"Save motor spec {payload['motor_name']}",
            "content": base64.b64encode(raw).decode("ascii"), "branch": branch}
    if existing:
        body["sha"] = existing
    result = github_request("PUT", url, token, json=body)
    if result.status_code not in (200, 201):
        raise RuntimeError(f"GitHub 저장 실패 HTTP {result.status_code}: {result.text[:300]}")
    return f"GitHub 저장 완료: {repo}/{relpath}"


def github_records() -> list[dict]:
    token, repo, branch = github_config()
    if not token or not repo:
        return []
    url = f"https://api.github.com/repos/{repo}/contents/motor_specs"
    res = github_request("GET", url, token, params={"ref": branch})
    if res.status_code == 404:
        return []
    res.raise_for_status()
    return [item for item in res.json() if item.get("name", "").endswith(".json")]


def github_load(filename: str) -> dict:
    token, repo, branch = github_config()
    url = f"https://api.github.com/repos/{repo}/contents/motor_specs/{filename}"
    res = github_request("GET", url, token, params={"ref": branch})
    res.raise_for_status()
    data = res.json()
    return json.loads(base64.b64decode(data["content"]).decode("utf-8"))


def commit_save(rows, model_name, limit, notes, record_id, company="", voltage="", switch_rpm=0) -> tuple[dict, str]:
    payload = record_payload(rows, model_name, limit, notes, record_id, company, voltage, switch_rpm)
    save_local(payload)
    if github_config()[0] and github_config()[1]:
        try:
            result = github_save(payload)
            return payload, "임시 서버 저장 완료 / " + result
        except Exception as exc:
            return payload, "임시 서버 저장 완료 / GitHub 영구 저장 실패: " + str(exc)
    return payload, "임시 서버 저장 완료 (GitHub 영구 저장 미설정)"


RETRYABLE_API_CODES = {408, 429, 500, 502, 503, 504}
DEFAULT_MODELS = [
    "gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash",
    "gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite",
]


def google_models(api_key: str) -> list[str]:
    """Returns actual models allowed to use generateContent according to listModels.

    Empty return on failures. Do not assume published models are enabled for a specific key.
    """
    if not api_key:
        return []
    names = []
    page_token = None
    for _ in range(5):
        params = {"pageSize": 100}
        if page_token:
            params["pageToken"] = page_token
        response = requests.get(
            "https://generativelanguage.googleapis.com/v1beta/models",
            headers={"x-goog-api-key": api_key}, params=params, timeout=15,
        )
        response.raise_for_status()
        data = response.json()
        for item in data.get("models", []):
            name = str(item.get("name", "")).removeprefix("models/")
            supported = item.get("supportedGenerationMethods", item.get("supported_generation_methods", []))
            if name and "generateContent" in supported and name.startswith("gemini-"):
                names.append(name)
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return list(dict.fromkeys(names))


def extract_images(uploaded, api_key: str, model: str, auto_fallback=True,
                   available_models=None, progress=None, sleep_fn=None) -> tuple[dict, str, int]:
    """Gemini REST + bounded 2/4/8 second backoff + one supported fallback model.

    No retries for invalid API keys, blocked/billing errors (400/401/402/403).
    """
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", model.strip()):
        raise ValueError("모델 이름 형식이 올바르지 않습니다.")
    if not api_key:
        raise ValueError("Gemini API 키를 설정하세요.")
    progress = progress or (lambda message: None)
    sleep_fn = sleep_fn or time.sleep
    parts = [{"text": SYSTEM_PROMPT}]
    raw_total = 0
    for f in uploaded:
        raw = f.getvalue()
        raw_total += len(raw)
        if len(raw) > 10 * 1024 * 1024 or raw_total > 12 * 1024 * 1024:
            raise ValueError("사진 1장당 10MB, 총 12MB 이하로 등록하세요.")
        if raw.startswith(b"\x89PNG\r\n\x1a\n"):
            mime = "image/png"
        elif raw.startswith(b"\xff\xd8\xff"):
            mime = "image/jpeg"
        else:
            raise ValueError("PNG 또는 JPEG 사진만 등록할 수 있습니다.")
        parts.append({"inline_data": {"mime_type": mime, "data": base64.b64encode(raw).decode("ascii")}})
    payload = {"contents": [{"role": "user", "parts": parts}],
               "generationConfig": {"temperature": 0, "responseMimeType": "application/json"}}
    if available_models is None and auto_fallback:
        try:
            available_models = google_models(api_key)
        except (requests.RequestException, ValueError, KeyError):
            available_models = []
    # Only pick from the user's reported usable list; never silently use an unknown model.
    candidates = [model]
    if auto_fallback and available_models:
        fallbacks = [m for m in DEFAULT_MODELS if m in available_models and m != model]
        if not fallbacks:
            fallbacks = [m for m in available_models if m != model]
        candidates += fallbacks[:1]
    attempt_count = 0
    last_error = ""
    for selected in candidates:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{selected}:generateContent"
        for attempt in range(4):
            attempt_count += 1
            progress(f"AI 인식: {selected} · {attempt+1}/4회 시도")
            try:
                response = requests.post(url, json=payload,
                    headers={"x-goog-api-key": api_key, "Content-Type": "application/json"},
                    timeout=90)
            except (requests.Timeout, requests.ConnectionError) as exc:
                last_error = f"연결 일시 장애 ({type(exc).__name__})"
                retryable = True
            else:
                if response.ok:
                    try:
                        response_json = response.json()
                        text_parts = response_json["candidates"][0]["content"]["parts"]
                        raw_json = "\n".join(p["text"] for p in text_parts if "text" in p).strip()
                        if raw_json.startswith("```"):
                            raw_json = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw_json).strip()
                        catalog = json.loads(raw_json)
                        if not isinstance(catalog, dict) or not isinstance(catalog.get("curves"), list):
                            raise ValueError("AI 응답의 curves 구조가 잘못되었습니다.")
                    except (ValueError, KeyError, IndexError, TypeError) as exc:
                        raise RuntimeError("AI 응답을 모터 사양 JSON으로 읽지 못했습니다. 사진과 모델을 확인하세요.") from exc
                    return catalog, selected, attempt_count
                try:
                    error_detail = str(response.json().get("error", {}).get("message", ""))[:180]
                except (ValueError, TypeError):
                    error_detail = ""
                error_detail = error_detail.replace(api_key, "[비공개]")
                last_error = f"HTTP {response.status_code}: {error_detail}"
                retryable = response.status_code in RETRYABLE_API_CODES
                if not retryable:
                    # E.g. invalid key, unsupported API, exhausted prepaid billing.
                    if response.status_code in (400, 401, 402, 403):
                        raise RuntimeError(f"{selected} API 요청 오류 ({last_error}). 키·결제·권한을 확인하세요.")
                    # 404: if selected model retired, proceed with the actual available fallback.
                    if response.status_code == 404 and selected != candidates[-1]:
                        progress(f"{selected} 사용 불가(404) · 다른 모델로 변경합니다.")
                        break
                    raise RuntimeError(f"{selected} API 요청 오류 ({last_error}).")
            if retryable and attempt < 3:
                pause = [2, 4, 8][attempt] + random.uniform(0, 0.3)
                progress(f"일시적 장애: {last_error} · 약 {int(pause)}초 후 재시도")
                sleep_fn(pause)
            elif retryable:
                progress(f"{selected} 재시도 한도 도달. " +
                         ("사용 가능한 다음 모델로 전환합니다." if selected != candidates[-1] else "더 이상 재시도하지 않습니다."))
    raise RuntimeError(f"AI 인식 실패 (총 {attempt_count}회 시도). 마지막 오류: {last_error}. 잠시 후 다시 실행하거나 모델을 변경하세요.")


def flatten(catalog: dict) -> list[dict]:
    """Normalize untrusted JSON records, letting tidy_rows reject bad numbers."""
    rows = []
    for curve in catalog.get("curves", []):
        if not isinstance(curve, dict):
            continue
        points = curve.get("points", [])
        if not isinstance(points, list):
            continue
        for p in points:
            if not isinstance(p, dict):
                continue
            rows.append({
                "권선": normalize_winding(curve.get("winding", "LOW")),
                "운전조건": normalize_rating(curve.get("rating", "S1")),
                "RPM": p.get("rpm"),
                "출력(kW)": p.get("power_kw"),
                "토크(N·m)": p.get("torque_nm"),
            })
    return rows


def use_record(payload: dict):
    st.session_state.rows = payload.get("rows", [])
    st.session_state.motor_name = payload.get("motor_name", "")
    st.session_state.motor_company = payload.get("motor_company", "")
    st.session_state.voltage = payload.get("voltage", "")
    st.session_state.pending_identity = {
        "motor_name_input": st.session_state.motor_name,
        "motor_company_input": st.session_state.motor_company,
        "motor_voltage_input": st.session_state.voltage or None,
    }
    st.session_state.switch_rpm_user = int(payload.get("switch_rpm") or 0)
    st.session_state.switch_enabled_user = False
    st.session_state.max_rpm = max(100, int(payload.get("max_rpm") or 20000))
    st.session_state.operating_rpm_input = max(100, int(payload.get("operating_max_rpm") or st.session_state.max_rpm))
    st.session_state.notes = payload.get("notes", "")
    st.session_state.record_id = payload.get("record_id", str(uuid.uuid4())[:8])
    st.session_state.editor_version += 1
    st.session_state.pop("excel_bytes", None)
    st.session_state.pop("excel_key", None)


def new_id():
    return uuid.uuid4().hex[:10]


# 클립보드 이미지를 기존 파일 업로드와 동일하게 처리하기 위한 어댑터.
# 원본 이미지 바이트는 세션 내에서만 보관하며, 자동으로 파일로 저장하지 않습니다.
class ClipboardImage:
    def __init__(self, name: str, raw: bytes):
        self.name = name
        self.type = "image/png"
        self._raw = raw

    def getvalue(self) -> bytes:
        return self._raw


def clipboard_files() -> list[ClipboardImage]:
    return [ClipboardImage(item["name"], item["data"])
            for item in st.session_state.get("clipboard_images", [])]


def add_clipboard_image(image) -> tuple[bool, str]:
    """Clipboard -> PNG; exact same bytes flow as a normal uploaded file."""
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    raw = buffer.getvalue()
    if len(raw) > 10 * 1024 * 1024:
        return False, "캡처 이미지가 10MB를 초과합니다. 캡처 범위를 줄인 뒤 다시 등록하세요."
    digest = hashlib.sha256(raw).hexdigest()
    entries = st.session_state.setdefault("clipboard_images", [])
    if any(item["sha256"] == digest for item in entries):
        return False, "같은 캡처 이미지가 이미 등록되어 있습니다."
    filename = f"clipboard_{len(entries) + 1:02d}_{digest[:8]}.png"
    entries.append({"name": filename, "sha256": digest, "data": raw})
    return True, f"캡처 등록 완료: {filename}"


# ─────────────────────────────────────────────────────────
# 화면 구성 : 이미지 등록 → 인식 → 검토/저장 → 그래프/엑셀
# ─────────────────────────────────────────────────────────
def fingerprint_uploads(files) -> str | None:
    if not files:
        return None
    digest = hashlib.sha256()
    for f in files:
        digest.update(f.name.encode("utf-8", errors="replace"))
        digest.update(f.getvalue())
    return digest.hexdigest()


def file_audit(files):
    """파일을 선택했다는 사실과 실제 이미지를 분석했다는 사실은 구분한다."""
    infos, errors = [], []
    for f in files or []:
        raw = f.getvalue()
        png = raw.startswith(b"\x89PNG\r\n\x1a\n")
        jpeg = raw.startswith(b"\xff\xd8\xff")
        valid_image = png or jpeg
        if not valid_image:
            errors.append(f"{f.name}: PNG/JPEG 이미지 형식이 아닙니다.")
        if not raw:
            errors.append(f"{f.name}: 파일 내용이 비어 있습니다.")
        if len(raw) > 10 * 1024 * 1024:
            errors.append(f"{f.name}: AI 인식은 한 장당 10MB까지 지원합니다.")
        infos.append({"파일명": f.name, "크기": f"{len(raw) / (1024*1024):.2f} MB",
                      "이미지 확인": "확인" if valid_image else "형식 오류"})
    if sum(len(f.getvalue()) for f in files or []) > 12 * 1024 * 1024:
        errors.append("전체 이미지 용량이 12MB를 초과했습니다. 나누어서 등록하세요.")
    return infos, errors


def content_signature(rows, name: str, limit: int, notes: str) -> str:
    raw = json.dumps({"rows": rows, "name": name, "max": limit, "notes": notes},
                     ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def reset_excel():
    st.session_state.pop("excel_bytes", None)
    st.session_state.pop("excel_key", None)


for k, v in {
    "rows": [], "motor_name": "", "motor_company": "", "voltage": "", "max_rpm": 20000,
    "operating_rpm_input": 20000,
    "record_id": new_id(), "notes": "", "editor_version": 0,
    "switch_rpm_user": 0, "switch_enabled_user": False, "pending_identity": None,
    "motor_name_input": "", "motor_company_input": "", "motor_voltage_input": None,
    "recognized_signature": None, "recognition_report": None,
    "recognition_error": "", "saved_signature": None,
    "saved_message": "", "data_origin": "미등록",
    "clipboard_images": [], "clipboard_feedback": "", "clipboard_button_version": 0,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# Apply identity changes scheduled by AI, sample buttons or JSON load BEFORE text widgets exist.
if st.session_state.pending_identity:
    for key, value in st.session_state.pending_identity.items():
        st.session_state[key] = value
    st.session_state.pending_identity = None

apply_ui_styles()

app_password = secret("APP_PASSWORD")
if app_password and not st.session_state.get("authenticated", False):
    st.title(APP_TITLE)
    attempt = st.text_input("앱 접속 비밀번호", type="password")
    if st.button("접속", type="primary"):
        if attempt == app_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    st.stop()

with st.sidebar:
    st.markdown("### ⚙️ 설정 및 저장 내역")
    with st.expander("AI 이미지 인식 설정", expanded=False):
        api_key = st.text_input("Gemini API 키", value=secret("GEMINI_API_KEY"),
                                type="password", help="Streamlit Secrets 설정을 권장합니다.")
        auto_fallback = st.checkbox("503/429 발생 시 자동 재시도 및 다른 모델 전환", value=True)
        # A newly entered key must not accidentally reuse another key's model list.
        current_key_fingerprint = hashlib.sha256(api_key.encode("utf-8")).hexdigest() if api_key else ""
        if st.session_state.get("models_key_fingerprint") != current_key_fingerprint:
            st.session_state.available_gemini_models = []
            st.session_state.models_key_fingerprint = current_key_fingerprint
        if st.button("현재 API 키에서 사용 가능한 모델 조회", use_container_width=True):
            try:
                st.session_state.available_gemini_models = google_models(api_key)
                st.success(f"generateContent 지원 모델 {len(st.session_state.available_gemini_models)}개 확인")
            except Exception as exc:
                st.warning(f"모델 조회 실패: {type(exc).__name__} · API 키/권한/네트워크 확인")
        discovered = st.session_state.get("available_gemini_models", [])
        if discovered:
            preset_models = list(dict.fromkeys([m for m in DEFAULT_MODELS if m in discovered] + discovered))
        else:
            preset_models = DEFAULT_MODELS[:]
        preset_models += ["직접 입력"]
        default_model = secret("GEMINI_MODEL") or "gemini-3.8-flash"
        default_index = preset_models.index(default_model) if default_model in preset_models else len(preset_models)-1
        selected_model = st.selectbox("Google Gemini 모델 선택", preset_models,
                                      index=default_index, help="목록 조회로 본인 API 키의 generateContent 지원 모델을 확인할 수 있습니다.")
        if selected_model == "직접 입력":
            model = st.text_input("직접 사용할 모델 ID", value=default_model if default_model not in preset_models else "gemini-3.8-flash")
        else:
            model = selected_model
        st.caption("현재 모델: " + model + " · 503/429는 최대 3회 재시도 후 사용 가능한 대체 모델 1개로 전환합니다. 전송되는 사진은 Google API에서 처리합니다.")
    if not app_password:
        st.warning("APP_PASSWORD 미설정: 공개 앱이라면 접근 제한 설정을 권장합니다.")
    with st.expander("저장된 모터 사양 불러오기", expanded=False):
        paths = sorted(DATA_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
        if paths:
            chosen = st.selectbox("이 서버에 저장된 사양", paths, format_func=lambda f: f.stem)
            if st.button("서버 저장본 불러오기", use_container_width=True):
                use_record(json.loads(chosen.read_text(encoding="utf-8")))
                st.session_state.data_origin = "서버 JSON"
                st.session_state.recognized_signature = None
                st.session_state.recognition_report = None
                st.session_state.saved_signature = None
                st.session_state.saved_message = "불러온 데이터를 확인해 주세요."
                st.rerun()
        else:
            st.caption("현재 서버에 저장된 사양이 없습니다.")
        if github_config()[0] and github_config()[1]:
            try:
                remote = github_records()
                if remote:
                    name = st.selectbox("GitHub 영구 저장 사양", [i["name"] for i in remote])
                    if st.button("GitHub 저장본 불러오기", use_container_width=True):
                        use_record(github_load(name))
                        st.session_state.data_origin = "GitHub JSON"
                        st.session_state.recognized_signature = None
                        st.session_state.recognition_report = None
                        st.session_state.saved_signature = None
                        st.session_state.saved_message = "불러온 데이터를 확인해 주세요."
                        st.rerun()
            except Exception as exc:
                st.warning(f"GitHub 저장 목록을 읽을 수 없습니다: {exc}")
        else:
            st.caption("영구 저장을 원하면 별도 Private GitHub 저장소 연동이 필요합니다.")
    st.caption("서버 임시 저장은 재배포 후 사라질 수 있습니다. JSON 백업 또는 GitHub 저장을 이용하세요.")

st.image(base64.b64decode(LOGO_PNG_BASE64), width=155)
st.title(APP_TITLE)

# 상단 고정 진행 상태. 업로드 파일 접수, AI 인식, 데이터 저장은 각각 별개의 상태입니다.
files_at_top = list(st.session_state.get('motor_catalog_images') or []) + clipboard_files()
current_upload_sig = fingerprint_uploads(files_at_top)
top_file_info, top_file_errors = file_audit(files_at_top)
top_uploaded = bool(files_at_top and not top_file_errors)
top_recognized = bool(top_uploaded and current_upload_sig == st.session_state.recognized_signature)
top_rows = bool(st.session_state.rows)
top_saved = bool(st.session_state.saved_signature)
with st.container(border=True):
    st.markdown("**진행 현황**")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric('① 사진 등록', f'{len(files_at_top)}장 완료' if top_uploaded else '등록 대기')
    c2.metric('② AI 인식', '완료' if top_recognized else ('재인식 필요' if top_uploaded and top_rows and st.session_state.data_origin == 'AI 인식' else '실행 대기'))
    c3.metric('③ 사양 검토', f'{len(st.session_state.rows)}개 기준점' if top_rows else '검토 대기')
    c4.metric('④ 사양 저장', '저장 이력 있음' if top_saved else '저장 대기')
    if top_file_errors:
        st.error('선택한 사진 파일에 문제가 있습니다. ① 사진 등록 탭에서 세부 내용을 확인해 주세요.')
    elif top_uploaded:
        filename_line = ' · '.join(f.name for f in files_at_top)
        st.caption(f'현재 선택된 파일: {filename_line[:240]}' + ('…' if len(filename_line) > 240 else ''))
        if not top_recognized:
            st.info('**사진 등록까지만 완료**되었습니다. ①에서 「AI 인식 시작」을 눌러야 데이터가 추출됩니다.')
        else:
            st.success('현재 사진의 AI 인식 결과가 준비되었습니다. ②에서 수치를 확인하고 저장해 주세요.')

# 네 개의 작업 화면을 분리하여 필요한 정보에 집중할 수 있게 합니다.
step_upload, step_review, step_plot, step_excel = st.tabs([
    "① 사진 등록·AI 인식", "② 추출값 검토·저장", "③ 통합 출력선도", "④ 엑셀 다운로드"
])

with step_upload:
    st.subheader("① 모터 모델명 입력 및 사양 이미지 등록")
    st.caption("모터 정보 입력 → 사진 등록 → AI 인식 → 원본 그래프 확인 → 필요할 때 권선절환 RPM을 적용합니다.")
    id_company, id_name, id_voltage = st.columns([1.1, 1.5, .85])
    with id_company:
        st.text_input("모터 회사명", key="motor_company_input", placeholder="예: SIEMENS / ATE")
    with id_name:
        st.text_input("모터 모델명", key="motor_name_input", placeholder="예: 1128/20000B")
    with id_voltage:
        st.selectbox("정격 전압", options=["200 V", "400 V"], key="motor_voltage_input",
                     index=None, placeholder="전압 선택")
    st.session_state.motor_name = st.session_state.motor_name_input.strip()
    st.session_state.motor_company = st.session_state.motor_company_input.strip()
    st.session_state.voltage = st.session_state.motor_voltage_input or ""
    if not st.session_state.motor_name or not st.session_state.motor_company or not st.session_state.voltage:
        st.caption("모터 회사명·모델명·전압을 입력해 두면 저장된 사양과 엑셀에 함께 반영됩니다.")
    capture_col, upload_col = st.columns(2, gap="large")
    with capture_col:
        st.markdown("#### 📋 A. 화면 캡처 붙여넣기")
        st.markdown("**Windows:** `Win + Shift + S` → 모터 사양표 영역 드래그 → 아래 버튼 클릭")
        st.caption("클립보드에 복사된 이미지를 바로 받아옵니다. 캡처 파일을 저장할 필요가 없습니다.")
        paste_result = paste_image_button(
            "📋 복사한 캡처 이미지 등록",
            key=f"motor_clipboard_paste_button_{st.session_state.clipboard_button_version}",
            errors="raise",
        )
        if paste_result.image_data is not None:
            added, feedback = add_clipboard_image(paste_result.image_data)
            st.session_state.clipboard_feedback = feedback
            if added:
                st.rerun()
        if st.session_state.get("clipboard_feedback"):
            st.caption(st.session_state.clipboard_feedback)
        st.caption("Chrome 또는 Edge의 HTTPS 환경을 권장합니다. 브라우저가 요청하면 클립보드 읽기 권한을 허용하세요.")
        pasted_count = len(st.session_state.clipboard_images)
        if pasted_count:
            st.success(f"캡처 {pasted_count}장 등록됨")
            if st.button("캡처 등록 목록 비우기", key="clear_clipboard_images"):
                st.session_state.clipboard_images = []
                st.session_state.clipboard_button_version += 1  # Reset the component's stale paste result.
                st.session_state.clipboard_feedback = "캡처 이미지 목록을 비웠습니다."
                st.rerun()

    with upload_col:
        st.markdown("#### 📁 B. 사진 파일 업로드")
        st.caption("컴퓨터에 저장된 JPG/JPEG/PNG 파일을 선택하거나 끌어다 놓으세요.")
        uploads = st.file_uploader(
            "모터 카탈로그 사진 파일 선택",
            type=["png", "jpg", "jpeg"],
            accept_multiple_files=True,
            key="motor_catalog_images",
            help="사진 등록 후 아래 AI 인식 버튼을 눌러야 수치가 분석됩니다.",
        )
        if uploads:
            st.success(f"파일 {len(uploads)}장 업로드됨")

    screenshots = list(uploads or []) + clipboard_files()
    st.divider()
    st.markdown("#### ✅ 실제 등록된 이미지 확인")
    st.caption("아래 미리보기에 표시된 이미지가 AI가 분석할 사진입니다. 파일 업로드와 캡처 등록을 합쳐 인식합니다.")
    upload_sig = fingerprint_uploads(screenshots)
    uploads_info, upload_errors = file_audit(screenshots)
    matched = bool(upload_sig and upload_sig == st.session_state.recognized_signature)
    if not screenshots:
        st.info("**사진 미등록** · 캡처 붙여넣기 버튼을 누르거나 사진 파일을 선택해 주세요.")
    else:
        if upload_errors:
            st.error("파일을 선택했지만 처리할 수 없는 이미지가 있습니다.")
            for error in upload_errors:
                st.warning(error)
        else:
            st.success(f"**사진 등록 확인: {len(screenshots)}장** · 파일을 정상적으로 받았습니다.")
            if matched:
                st.success("**AI 분석 완료** · 현재 선택된 사진의 수치를 추출했습니다. ② 사양 검토·저장 탭을 확인하세요.")
            else:
                st.info("**AI 분석 대기** · 사진 등록만 완료되었습니다. 아래 'AI 인식 시작' 버튼을 눌러야 숫자를 읽습니다.")
        st.dataframe(uploads_info, use_container_width=True, hide_index=True)
        with st.expander(f"📷 등록 사진 미리보기 ({len(screenshots)}장)", expanded=True):
            image_cols = st.columns(min(3, len(screenshots)))
            for i, file in enumerate(screenshots):
                with image_cols[i % len(image_cols)]:
                    raw = file.getvalue()
                    if raw.startswith(b"\x89PNG\r\n\x1a\n") or raw.startswith(b"\xff\xd8\xff"):
                        st.image(raw, caption=file.name, use_container_width=True)
                    else:
                        st.error(f"이미지 열기 실패: {file.name}")
        st.caption("※ 사진 파일 자체는 영구 저장되지 않습니다. 인식된 숫자는 ②에서 검토 후 저장할 수 있습니다.")

    if screenshots and not matched and st.session_state.rows and st.session_state.recognized_signature:
        st.warning("선택한 사진이 이전 분석 사진과 다릅니다. 아래 인식 버튼을 다시 눌러야 새 사진의 수치가 반영됩니다.")

    recognize_label = "🔄 다시 인식하기" if matched else "🔎 사진에서 모터 사양 자동 인식 시작"
    if st.button(recognize_label, type="primary", use_container_width=True,
                 disabled=not screenshots or bool(upload_errors)):
        if not api_key:
            st.session_state.recognition_error = "Gemini API 키가 없습니다. 왼쪽 'AI 이미지 인식 설정' 또는 Streamlit Secrets에 GEMINI_API_KEY를 지정해 주세요."
        else:
            st.session_state.recognition_error = ""
            try:
                with st.status("이미지를 분석하고 있습니다…", expanded=True) as progress:
                    st.write(f"1/3 · 업로드한 사진 {len(screenshots)}장 확인")
                    st.write("2/3 · Gemini AI에서 RPM·출력·토크 표 추출")
                    catalog, actual_model, attempt_total = extract_images(
                        screenshots, api_key, model, auto_fallback=auto_fallback,
                        available_models=st.session_state.get("available_gemini_models") or None,
                        progress=progress.write,
                    )
                    parsed, issues = tidy_rows(flatten(catalog))
                    if not parsed:
                        raise ValueError("인식 가능한 RPM / 출력 / 토크 값이 없습니다. 숫자가 보이는 표 사진을 사용해 주세요.")
                    st.write(f"3/3 · {len(parsed)}개 기준점 · {len(group_rows(parsed))}개 운전곡선 분류 완료")
                    progress.update(label="AI 사양 인식 완료", state="complete", expanded=False)
                st.session_state.rows = parsed
                detected_name = catalog.get("motor_name") or ""
                if not st.session_state.motor_name_input.strip() and detected_name:
                    st.session_state.motor_name = detected_name
                    st.session_state.pending_identity = {
                        "motor_name_input": detected_name,
                        "motor_company_input": st.session_state.motor_company_input,
                        "motor_voltage_input": st.session_state.motor_voltage_input,
                    }
                st.session_state.max_rpm = max(100, int(catalog.get("maximum_rpm") or max(x["RPM"] for x in parsed)))
                st.session_state.operating_rpm_input = st.session_state.max_rpm
                st.session_state.notes = " / ".join(map(str, catalog.get("notes", [])))
                st.session_state.record_id = new_id()
                st.session_state.editor_version += 1
                st.session_state.recognized_signature = upload_sig
                st.session_state.recognition_report = {
                    "files": [f.name for f in screenshots],
                    "points": len(parsed),
                    "curves": len(group_rows(parsed)),
                    "warnings": issues,
                    "ai_model": actual_model,
                    "api_attempts": attempt_total,
                    "at": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S"),
                }
                st.session_state.saved_signature = None
                st.session_state.saved_message = ""
                st.session_state.data_origin = "AI 인식"
                reset_excel()
                st.rerun()
            except Exception as exc:
                st.session_state.recognition_error = str(exc)
    if st.session_state.recognition_error:
        st.error(f"AI 인식 실패: {st.session_state.recognition_error}")
    report = st.session_state.recognition_report
    if report and matched:
        st.markdown("##### 인식 결과 요약")
        k1, k2, k3 = st.columns(3)
        k1.metric("인식 기준점", f"{report['points']}개")
        k2.metric("운전 곡선", f"{report['curves']}개")
        k3.metric("등록 사진", f"{len(report['files'])}장")
        st.caption(f"마지막 인식: {report['at']} · AI 모델: {report.get('ai_model', '알 수 없음')} · API 요청 {report.get('api_attempts', 1)}회 · 인식 완료는 저장 완료를 의미하지 않습니다.")
        st.markdown('**사진에서 추출한 운전조건 (빠른 확인용)**')
        extracted = group_rows(st.session_state.rows)
        check_rows = []
        for key in sorted(extracted, key=lambda k: (k[0] != 'LOW', k[0] != 'HIGH', display_rating_order(k[1]))):
            pts = extracted[key]
            p_values = [v['출력(kW)'] for v in pts if v['출력(kW)'] is not None]
            t_values = [v['토크(N·m)'] for v in pts if v['토크(N·m)'] is not None]
            check_rows.append({'권선': key[0], '운전조건': key[1], 'RPM 범위': f"{pts[0]['RPM']:g}~{pts[-1]['RPM']:g}",
                               '기준점': len(pts), '최대 출력(kW)': max(p_values) if p_values else None,
                               '최대 토크(N·m)': max(t_values) if t_values else None})
        st.dataframe(check_rows, hide_index=True, use_container_width=True)
        st.warning('AI가 인식한 값은 오독 가능성이 있습니다. 아래 ② 탭에서 원본 사진과 비교해 검토·저장해야 합니다.')
        if report["warnings"]:
            with st.expander(f"AI 수치 확인 필요 ({len(report['warnings'])}건)"):
                for item in report["warnings"][:25]:
                    st.warning(item)
        if st.button("② 추출된 데이터 검토하러 가기", use_container_width=True):
            st.info("바로 위의 **② 사양 검토·저장** 탭을 클릭하세요.")
    st.divider()
    shortcut1, shortcut2 = st.columns(2)
    with shortcut1:
        if st.button("예제 데이터로 기능 확인", use_container_width=True):
            sample, note = make_sample()
            st.session_state.rows = sample
            st.session_state.motor_name = "LOW-HIGH 모터 예제"
            st.session_state.motor_company = ""
            st.session_state.voltage = ""
            st.session_state.pending_identity = {
                "motor_name_input": st.session_state.motor_name,
                "motor_company_input": "", "motor_voltage_input": None,
            }
            st.session_state.switch_rpm_user = 0
            st.session_state.switch_enabled_user = False
            st.session_state.max_rpm = 20000
            st.session_state.operating_rpm_input = 20000
            st.session_state.notes = note
            st.session_state.record_id = new_id()
            st.session_state.editor_version += 1
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.recognition_error = ""
            st.session_state.saved_signature = None
            st.session_state.data_origin = "예제"
            reset_excel()
            st.rerun()
    with shortcut2:
        if st.button("빈 입력표로 시작", use_container_width=True):
            st.session_state.rows = []
            st.session_state.motor_name = ""
            st.session_state.motor_company = ""
            st.session_state.voltage = ""
            st.session_state.pending_identity = {
                "motor_name_input": "", "motor_company_input": "", "motor_voltage_input": None,
            }
            st.session_state.switch_rpm_user = 0
            st.session_state.switch_enabled_user = False
            st.session_state.max_rpm = 20000
            st.session_state.operating_rpm_input = 20000
            st.session_state.notes = ""
            st.session_state.record_id = new_id()
            st.session_state.editor_version += 1
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.recognition_error = ""
            st.session_state.saved_signature = None
            st.session_state.data_origin = "수동 입력"
            reset_excel()
            st.rerun()

# The operating limit is independent of the catalog maximum. It is a per-user setting.
operating_rpm = max(100, min(int(st.session_state.operating_rpm_input), int(st.session_state.max_rpm), 100000))

with step_review:
    st.subheader("② 인식된 모터 사양 확인 및 저장")
    if screenshots:
        with st.expander(f"📷 인식에 사용한 원본 이미지 다시 확인 ({len(screenshots)}장)", expanded=False):
            st.caption("①에 등록된 원본과 아래 인식값을 비교하세요.")
            reference_cols = st.columns(min(2, len(screenshots)))
            for idx, reference_image in enumerate(screenshots):
                with reference_cols[idx % len(reference_cols)]:
                    image_bytes = reference_image.getvalue()
                    if image_bytes.startswith(b"\x89PNG\r\n\x1a\n") or image_bytes.startswith(b"\xff\xd8\xff"):
                        st.image(image_bytes, caption=reference_image.name, use_container_width=True)
    if st.session_state.data_origin == "AI 인식":
        if matched:
            st.success("현재 사진의 AI 인식 데이터입니다. 원본 사진과 비교 후 저장하세요.")
        elif screenshots:
            st.warning("현재 등록된 사진과 표시된 인식값이 다릅니다. ①에서 다시 인식하세요.")
        else:
            st.info("이전 AI 인식 데이터입니다. 필요하면 새 사진을 등록하여 다시 인식하세요.")
    elif st.session_state.rows:
        st.info(f"현재 데이터 출처: {st.session_state.data_origin}")
    else:
        st.info("① 사진 등록·인식에서 사진을 분석하거나 예제/수동 입력을 선택하세요.")

    model_name = st.session_state.motor_name_input.strip()
    company = st.session_state.motor_company_input.strip()
    voltage = st.session_state.motor_voltage_input or ""
    meta_1, meta_2 = st.columns([1.4, .6])
    with meta_1:
        st.info(f"제조사: {company or '미입력'} | 모델명: {model_name or '미입력'} | 정격전압: {voltage or '미선택'}")
        st.caption("제조사·모델명·전압은 ① 사진 등록 탭에서 수정할 수 있습니다.")
    with meta_2:
        rpm_limit = st.number_input("카탈로그 최고 RPM", min_value=100,
                                    max_value=100000, value=int(st.session_state.max_rpm), step=100)
    # Catalog limit may change here independently from the usage limit.
    operating_rpm = max(100, min(int(st.session_state.operating_rpm_input), int(rpm_limit)))
    remark = st.text_area("기타 사양 및 참고사항", value=st.session_state.notes, height=75)
    imported = st.file_uploader("저장해 둔 모터 사양 JSON 불러오기 (선택)", type=["json"], key="json_import")
    if st.button("JSON 데이터 불러오기", disabled=imported is None):
        try:
            payload = json.loads(imported.getvalue().decode("utf-8-sig"))
            use_record(payload)
            st.session_state.recognized_signature = None
            st.session_state.recognition_report = None
            st.session_state.saved_signature = None
            st.session_state.data_origin = "JSON 불러오기"
            st.rerun()
        except Exception as exc:
            st.error(f"JSON 파일을 읽지 못했습니다: {exc}")

    editable = st.session_state.rows or [
        {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 0, "출력(kW)": 0.0, "토크(N·m)": None},
        {"권선": "LOW", "운전조건": "S1 Cont.", "RPM": 1500, "출력(kW)": None, "토크(N·m)": None},
    ]
    st.caption("아래는 AI가 읽은 값입니다. 반드시 원본과 대조한 후 수정·저장하세요. 빈칸은 미확인 수치입니다.")
    updated = st.data_editor(
        editable, hide_index=True, num_rows="dynamic", use_container_width=True,
        key=f"motor_editor_{st.session_state.editor_version}",
        column_config={
            "권선": st.column_config.TextColumn("권선 LOW/HIGH"),
            "운전조건": st.column_config.TextColumn("운전조건 S1/S2/S3"),
            "RPM": st.column_config.NumberColumn("RPM", min_value=0, format="%.1f"),
            "출력(kW)": st.column_config.NumberColumn("출력 (kW)", min_value=0, format="%.3f"),
            "토크(N·m)": st.column_config.NumberColumn("토크 (N·m)", min_value=0, format="%.3f"),
        },
    )
    valid, warnings = tidy_rows(updated)
    groups = group_rows(valid)
    if valid:
        st.markdown("##### 추출 조건별 검토 요약")
        summary = []
        for (winding, rating), points in sorted(groups.items(), key=lambda item: (item[0][0] != 'LOW', item[0][0] != 'HIGH', display_rating_order(item[0][1]))):
            kw_values = [v["출력(kW)"] for v in points if v["출력(kW)"] is not None]
            nm_values = [v["토크(N·m)"] for v in points if v["토크(N·m)"] is not None]
            summary.append({"권선": winding, "운전조건": rating, "기준점 수": len(points),
                            "최소 RPM": points[0]["RPM"], "최대 RPM": points[-1]["RPM"],
                            "최대 출력(kW)": max(kw_values) if kw_values else None,
                            "최대 토크(N·m)": max(nm_values) if nm_values else None})
        st.dataframe(summary, hide_index=True, use_container_width=True)
        st.caption(f"인식/입력값 {len(valid)}개 기준점 · {len(groups)}개 곡선")
    if warnings:
        with st.expander(f"⚠️ 검토할 데이터 ({len(warnings)}건)", expanded=True):
            for warning in warnings[:30]:
                st.warning(warning)
    selected_switch = int(st.session_state.switch_rpm_user or 0) if st.session_state.switch_enabled_user else 0
    pending_sig = content_signature(valid, model_name, int(rpm_limit),
                                    remark + f"|{company}|{voltage}|switch={selected_switch}|use_rpm={operating_rpm}") if valid else None
    saved_current = bool(pending_sig and st.session_state.saved_signature == pending_sig)
    stale_photo = bool(screenshots and st.session_state.recognized_signature
                       and not matched and st.session_state.data_origin == "AI 인식")
    if saved_current:
        st.success(f"✅ **수치 저장 완료** · {st.session_state.saved_message}")
    else:
        st.info("**검토·저장 대기** · 아래 버튼을 눌러야 수정한 수치가 저장됩니다.")
    save_col, json_col = st.columns(2)
    with save_col:
        if st.button("✅ 검토한 모터 사양 저장", type="primary", use_container_width=True,
                     disabled=not valid or stale_photo):
            try:
                st.session_state.rows = valid
                st.session_state.motor_name = model_name
                st.session_state.max_rpm = int(rpm_limit)
                st.session_state.notes = remark
                st.session_state.editor_version += 1
                payload, sync = commit_save(valid, model_name, int(rpm_limit), remark,
                                            st.session_state.record_id, company=company,
                                            voltage=voltage, switch_rpm=st.session_state.switch_rpm_user if st.session_state.switch_enabled_user else 0)
                st.session_state.saved_signature = content_signature(valid, model_name, int(rpm_limit),
                    remark + f"|{company}|{voltage}|switch={selected_switch}|use_rpm={operating_rpm}")
                st.session_state.saved_message = sync
                reset_excel()
                st.rerun()
            except Exception as exc:
                st.error(f"저장 실패: {exc}")
    with json_col:
        current_payload = record_payload(valid, model_name, int(rpm_limit), remark,
                                         st.session_state.record_id, company=company,
                                         voltage=voltage,
                                         switch_rpm=st.session_state.switch_rpm_user if st.session_state.switch_enabled_user else 0)
        st.download_button("📥 모터 사양 JSON 백업", data=json.dumps(current_payload, ensure_ascii=False, indent=2),
                           file_name=f"{slug(model_name)}_motor_spec.json", mime="application/json",
                           use_container_width=True, disabled=not valid)
    if stale_photo:
        st.warning("사진이 변경되어 저장을 잠시 막았습니다. ①에서 새 사진을 인식하거나 빈 입력표/예제로 시작하세요.")

with step_plot:
    st.subheader("③ 출력·토크 통합 선도")
    if not valid:
        st.info("먼저 ①에서 사진을 등록하고 ②에서 데이터를 확인하세요.")
    else:
        if not saved_current:
            st.warning("그래프는 현재 검토표의 수치로 그려집니다. 저장이 아직 완료되지 않았다면 ②에서 저장하세요.")
        st.markdown("#### 사용 최대 회전수 설정")
        st.number_input("사용 최대 RPM (그래프·엑셀 공통 적용)", min_value=100,
                        max_value=100000, step=100, key="operating_rpm_input",
                        help="카탈로그 최고 RPM과 별도입니다. 예: 카탈로그 15,000 RPM → 장비 사용 12,000 RPM")
        operating_rpm = max(100, min(int(st.session_state.operating_rpm_input), int(rpm_limit)))
        st.caption(f"카탈로그 최고: {int(rpm_limit):,} RPM / 그래프·엑셀에 적용: {operating_rpm:,} RPM")
        if int(st.session_state.operating_rpm_input) > int(rpm_limit):
            st.warning("사용 RPM이 카탈로그 최고 RPM보다 높아, 카탈로그 최고 RPM에서 제한했습니다.")
        st.markdown("#### 1. 원본 LOW / HIGH 특성 먼저 확인")
        st.caption("원본 모터 사양을 설정한 사용 최대 RPM까지만 표시합니다. 권선절환 RPM은 자동 가정하지 않습니다.")
        show_mode = st.radio("표시 항목", ["출력 + 토크", "출력만", "토크만"], horizontal=True)
        show_raw_points = st.checkbox("기준점 점 표시", value=False)
        fig_raw = make_plot(valid, max_rpm=int(operating_rpm),
                            show_power=show_mode != "토크만", show_torque=show_mode != "출력만",
                            switch_rpm=None, apply_switch=False, show_raw_points=show_raw_points)
        st.plotly_chart(fig_raw, use_container_width=True, key="original_motor_plot")
        st.markdown("#### 2. 원본 그래프 확인 후 권선절환 포인트 지정")
        st.checkbox("권선절환 RPM 적용하기 (LOW → HIGH)", key="switch_enabled_user")
        apply_switch = bool(st.session_state.switch_enabled_user)
        if apply_switch:
            if int(st.session_state.switch_rpm_user) > int(operating_rpm):
                st.session_state.switch_rpm_user = int(operating_rpm)
            st.number_input("권선절환 RPM", min_value=0, max_value=int(operating_rpm),
                            step=100, key="switch_rpm_user",
                            help="입력 즉시 LOW(절환 전) / HIGH(절환 후) 그래프와 엑셀 적용값이 변경됩니다. 0은 미적용입니다.")
        switch_rpm = int(st.session_state.switch_rpm_user) if apply_switch else 0
        apply_switch = bool(apply_switch and switch_rpm > 0)
        if apply_switch:
            st.success(f"권선절환 적용: {switch_rpm:,} RPM · 왼쪽 구간 LOW, 오른쪽 구간 HIGH")
            fig = make_plot(valid, max_rpm=int(operating_rpm),
                            show_power=show_mode != "토크만", show_torque=show_mode != "출력만",
                            switch_rpm=switch_rpm, apply_switch=True,
                            show_raw_points=False)
            st.plotly_chart(fig, use_container_width=True, key="switched_motor_plot")
            st.caption("웹·엑셀 공통데이터: 절환점에서 LOW와 HIGH 선을 따로 그립니다. 임의의 연결선/외삽은 하지 않습니다.")
        else:
            fig = fig_raw
            st.info("현재 권선절환 미적용 — 위 원본 선도를 먼저 확인하고 필요할 때 RPM을 입력하세요.")
        st.download_button("현재 선택된 그래프 HTML 다운로드", data=fig.to_html(include_plotlyjs=True, full_html=True),
                           file_name=f"{slug(model_name)}_diagram.html", mime="text/html")

        st.markdown("##### 권선절환 적용값 미리보기")
        grouped_bucket = grouped_by_bucket(valid)
        buckets = sorted({b for _, b in grouped_bucket}, key=bucket_order)
        preview_rows = []
        preview_grid = rpm_grid(valid, limit=int(operating_rpm), step=500, switch_rpm=int(switch_rpm) if apply_switch else None)[:80]
        for rpm in preview_grid:
            row = {"RPM": rpm}
            for bucket in buckets:
                label = bucket_display(bucket)
                if apply_switch:
                    kw, nm, src = calc_switched_at(grouped_bucket, bucket, rpm, int(switch_rpm))
                    row[f"{label} 출력(kW)"] = round(kw, 3) if kw is not None else None
                    row[f"{label} 토크(N·m)"] = round(nm, 3) if nm is not None else None
                    row[f"{label} 적용권선"] = src
                else:
                    for winding in ["LOW", "HIGH"]:
                        pts = grouped_bucket.get((winding, bucket))
                        if pts:
                            kw, nm = calc_at(pts, rpm)
                            row[f"{winding} {label} 출력(kW)"] = round(kw, 3) if kw is not None else None
                            row[f"{winding} {label} 토크(N·m)"] = round(nm, 3) if nm is not None else None
            preview_rows.append(row)
        st.dataframe(preview_rows, hide_index=True, use_container_width=True)

with step_excel:
    st.subheader("④ RPM별 계산표 및 엑셀 다운로드")
    if not valid:
        st.info("먼저 데이터를 입력한 후 엑셀을 생성할 수 있습니다.")
    else:
        st.caption("엑셀 시트: 통합 선도 / RPM 계산표 / 권선절환 적용값 / 기준점 입력 / 차트 공통데이터 / 사용 안내")
        st.success("**RPM 계산표 배열:** LOW 출력(S3 10%→15%→25%→S2→S1) → LOW 토크(같은 순서) → HIGH 출력 → HIGH 토크. 카탈로그에 없는 조건은 추가하지 않습니다.")
        current_switch = int(st.session_state.switch_rpm_user or 0) if st.session_state.switch_enabled_user else 0
        if current_switch > 0:
            st.info(f"현재 권선절환 포인트: {current_switch} rpm · 엑셀의 [권선절환 적용값] 시트와 [통합 선도]에 즉시 반영됩니다.")
        else:
            st.info("권선절환 포인트가 0이므로 LOW/HIGH 원본 계산표 기준으로 엑셀이 생성됩니다.")
        rpm_step = st.selectbox("RPM 계산 간격", [10, 25, 50, 100, 200, 500], index=3)
        stamp = content_signature(valid, model_name, int(operating_rpm), remark + f"|{company}|{voltage}") + f"|{rpm_step}|{current_switch}"
        if st.session_state.get("excel_key") != stamp:
            st.session_state.pop("excel_bytes", None)
        if st.button("📊 엑셀 계산표 생성", type="primary", use_container_width=True):
            try:
                with st.spinner("엑셀 차트와 RPM 계산표를 생성하는 중입니다…"):
                    st.session_state.excel_bytes = generate_excel(
                        valid, f"{company} / {model_name} ({voltage})".strip(),
                        max_rpm=int(operating_rpm), step=int(rpm_step),
                        switch_rpm=current_switch if current_switch > 0 else None,
                        notes=remark,
                    )
                    st.session_state.excel_key = stamp
                st.success("엑셀 생성 완료 · 아래 다운로드 버튼으로 파일을 저장하세요.")
            except Exception as exc:
                st.error(f"엑셀 생성 실패: {exc}")
        if st.session_state.get("excel_bytes"):
            st.download_button("📥 스핀들 주축출력선도 엑셀 다운로드 (.xlsx)",
                               data=st.session_state.excel_bytes,
                               file_name=f"{slug(company + '_' + model_name + '_' + voltage)}_spindle_diagram.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", use_container_width=True)
        with st.expander("RPM 계산값 미리보기", expanded=False):
            preview = []
            winding_names = sorted({k[0] for k in groups}, key=lambda w: (w != 'LOW', w != 'HIGH', w))
            for rpm in rpm_grid(valid, limit=int(operating_rpm), step=int(rpm_step), switch_rpm=current_switch if current_switch > 0 else None)[:200]:
                line = {'RPM': rpm}
                for winding in winding_names:
                    keys = sorted((k for k in groups if k[0] == winding), key=lambda k: display_rating_order(k[1]))
                    for quantity in ('출력(kW)', '토크(N·m)'):
                        for key in keys:
                            kw, nm = calc_at(groups[key], rpm)
                            value = kw if quantity == '출력(kW)' else nm
                            line[f'{winding} / {quantity} / {key[1]}'] = round(value, 3) if value is not None else None
                preview.append(line)
            st.dataframe(preview, hide_index=True, use_container_width=True)

