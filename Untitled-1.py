"""
=====================================================================
GAS TURBINE THERMODYNAMIC ANALYSIS TOOLKIT
=====================================================================
DESCRIPTION:
Complete thermodynamic analysis tool for gas turbine cycles including:
- Single-stage and multi-stage compression/expansion
- Combustion chamber and reheater modeling
- Heat exchanger (intercooler, regenerator) analysis
- Complete TS and PV diagram generation with ideal vs actual overlays
- Open-cycle architecture formatting
- Detailed calculation table terminal outputs
- Advanced Excel Data Export Engine utilizing live formulas and styling

REQUIREMENTS:
- numpy, matplotlib, pandas, openpyxl
- AUTHOR: Rishabh Kumar
- DATE: 26 MAY, 2026
=====================================================================
"""
import os
import time
import math
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Set high-resolution rendering defaults for crisp visualization plots
plt.rcParams["figure.dpi"] = 110
plt.rcParams["savefig.dpi"] = 300


# =====================================================================
# CORE UTILITY HELPERS
# =====================================================================
def get_terminal_inputs(prompts, defaults):
    """Gathers configuration inputs cleanly via console line parameters."""
    print("-" * 50)
    answers = []
    for prompt, default in zip(prompts, defaults):
        user_input = input(f"{prompt} (Default: {default}): ").strip()
        answers.append(user_input if user_input else default)
    return answers


def preserve_unique_legend(ax):
    """Filters out duplicate display labels from matplotlib axis legends."""
    handles, labels = ax.get_legend_handles_labels()
    unique = [
        (h, l)
        for i, (h, l) in enumerate(zip(handles, labels))
        if l not in labels[:i] and l != ""
    ]
    if unique:
        h_un, l_un = zip(*unique)
        ax.legend(h_un, l_un, loc="best", fontsize=9)


# =====================================================================
# FUNCTION: Single-Stage Cycle Analysis
# =====================================================================
def analyze_single_stage():
    print("\n" + "=" * 70)
    print("                  SINGLE STAGE COMPRESSION-EXPANSION ANALYSIS")
    print("                  ========================================")

    # Input Processing Blocks
    prompts1 = [
        "Inlet Pressure (bar)",
        "Inlet Temperature (K)",
        "Max Cycle Temperature (K)",
        "Relative Humidity (%)",
        "Mass Flow Rate (kg/s)",
        "Pressure Ratio",
        "Turbine Exit Pressure (bar)",
    ]
    defaults1 = ["1.013", "300", "1073", "0.0", "1", "6", "1.013"]
    ans1 = get_terminal_inputs(prompts1, defaults1)
    p1, t1, t3 = float(ans1[0]), float(ans1[1]), float(ans1[2])
    rh, m, pr, p4 = float(ans1[3]), float(ans1[4]), float(ans1[5]), float(ans1[6])

    prompts2 = [
        "Compressor Efficiency (0-1)",
        "Turbine Efficiency (0-1)",
        "Combustion Efficiency (0-1)",
        "Regenerator Efficiency (0-1)",
        "Calorific Value (kJ/kg)",
    ]
    defaults2 = ["0.8", "0.85", "0.95", "0.8", "42000"]
    ans2 = get_terminal_inputs(prompts2, defaults2)
    nc, nt, ncomb, epselon, cv = (
        float(ans2[0]),
        float(ans2[1]),
        float(ans2[2]),
        float(ans2[3]),
        float(ans2[4]),
    )

    prompts3 = [
        "gamma (Compression)",
        "Cp (kJ/kg-K) (compression)",
        "gamma (Expansion)",
        "Cp (kJ/kg-K) (expansion)",
    ]
    defaults3 = ["1.4", "1.005", "1.33", "1.148"]
    ans3 = get_terminal_inputs(prompts3, defaults3)
    gc, cpc, ge, cpe = float(ans3[0]), float(ans3[1]), float(ans3[2]), float(ans3[3])

    prompts4 = [
        "Pressure loss at combustion exit (%)",
        "Pressure loss at turbine exit (%)",
        "Resolution Mesh-points",
    ]
    defaults4 = ["3", "3", "1000"]
    ans4 = get_terminal_inputs(prompts4, defaults4)
    dp2, dp3, resln = float(ans4[0]), float(ans4[1]), int(ans4[2])
    r = 0.287

    print("\n" + "=" * 70)
    print("CALCULATION PHASE: Processing Thermodynamic Cycle")
    print("=" * 70 + "\n")

    # Core Station Milestones (Actual vs Ideal Parallel Engine)
    es = 1.7526e8 * math.exp(-5315.56 / t1)
    density = 3.4848 * (p1 * 1e2 - 0.00376960 * rh * es) * 1e-3 / t1
    v1 = m / density

    # Compression State boundary points
    p2 = p1 * pr
    t2dash = t1 * (p2 / p1) ** ((gc - 1) / gc)
    t2 = t1 + (t2dash - t1) / nc
    wc = cpc * (t2 - t1)

    # Combustion State boundary points
    af = (ncomb * cv - cpc * t3) / (cpe * (t3 - t2))
    qs = (1 + 1 / af) * cpe * (t3 - t2)
    p3 = p2 * (1 - dp2 / 100)
    t3dash = t3

    # Turbine State boundary points
    t4dash = t3 * (p4 / p3) ** ((ge - 1) / ge)
    t4 = t3 - nt * (t3 - t4dash)
    wt = cpe * (t3 - t4)

    # Ideal Turbine drop from peak ideal pressure P2 down to P1
    t4dash_pure = t3dash * (p1 / p2) ** ((ge - 1) / ge)

    t5 = t3 - epselon * (t3 - t4)
    af_reg = (ncomb * cv - cpe * t3) / (cpe * (t3 - t5))
    qs_reg = (1 + 1 / af_reg) * cpe * (t3 - t5)

    # =====================================================================
    # DIRECT THERMODYNAMIC VECTOR GENERATION (Zero-Index Dependency)
    # =====================================================================
    p12_act = np.linspace(p1, p2, resln)
    t12_id = t1 * (p12_act / p1) ** ((gc - 1) / gc)
    t12_act = t1 + (t12_id - t1) / nc

    s12_act = cpc * np.log(t12_act / t1) - r * np.log(p12_act / p1) + 0.0
    s12_id = np.repeat(0.0, resln)

    v12_act = (r * t12_act) / (p12_act * 100)
    v12_id = (r * t12_id) / (p12_act * 100)

    s2_act = s12_act[-1]
    s2_id = 0.0
    t2_id = t12_id[-1]

    p23_act = np.linspace(p2, p3, resln)
    p23_id = np.repeat(p2, resln)
    t23_act = np.linspace(t2, t3, resln)
    t23_id = np.linspace(t2_id, t3dash, resln)

    s23_act = cpe * np.log(t23_act / t2) - r * np.log(p23_act / p2) + s2_act
    s23_id = cpe * np.log(t23_id / t2_id) - r * np.log(p23_id / p2) + s2_id

    v23_act = (r * t23_act) / (p23_act * 100)
    v23_id = (r * t23_id) / (p23_id * 100)

    s3_act = s23_act[-1]
    s3_id = s23_id[-1]

    p34_act = np.linspace(p3, p4, resln)
    p34_id = np.linspace(p2, p1, resln)
    t34_id_raw = t3 * (p34_act / p3) ** ((ge - 1) / ge)
    t34_act = t3 - nt * (t3 - t34_id_raw)
    t34_id = t3dash * (p34_id / p2) ** ((ge - 1) / ge)

    s34_act = cpe * np.log(t34_act / t3) - r * np.log(p34_act / p3) + s3_act
    s34_id = np.repeat(s3_id, resln)

    v34_act = (r * t34_act) / (p34_act * 100)
    v34_id = (r * t34_id) / (p34_id * 100)

    s4_act = s34_act[-1]
    s4_id = s34_id[-1]
    t4_id = t34_id[-1]

    # =================================================================
    # TERMINAL DATA CALCULATION TABLE DISPLAYS
    # =================================================================
    print(f"{' STATE SUMMARY MATRIX (SINGLE-STAGE) ':=^70}")
    print(
        f"{'Node':<6}{'Condition Type':<24}{'P (bar)':<10}{'T (K)':<10}{'v (m³/kg)':<12}{'Δs (kJ/kg-K)'}"
    )
    print("-" * 70)
    print(f"{'1':<6}{'Actual Inlet':<24}{p1:<10.3f}{t1:<10.2f}{v1:<12.4f}{0.0:<10.4f}")
    print(
        f"{'2':<6}{'Actual Compressor Out':<24}{p2:<10.3f}{t2:<10.2f}{v12_act[-1]:<12.4f}{s2_act:<10.4f}"
    )
    print(
        f"{'3':<6}{'Actual Combustor Out':<24}{p3:<10.3f}{t3:<10.2f}{v23_act[-1]:<12.4f}{s3_act:<10.4f}"
    )
    print(
        f"{'4':<6}{'Actual Turbine Exhaust':<24}{p4:<10.3f}{t4:<10.2f}{v34_act[-1]:<12.4f}{s4_act:<10.4f}"
    )
    print("-" * 70)
    print(
        f"{'2s':<6}{'Isentropic Comp Out':<24}{p2:<10.3f}{t2_id:<10.2f}{v12_id[-1]:<12.4f}{0.0:<10.4f}"
    )
    print(
        f"{'3s':<6}{'Isentropic Comb Out':<24}{p2:<10.3f}{t3dash:<10.2f}{v23_id[-1]:<12.4f}{s3_id:<10.4f}"
    )
    print(
        f"{'4s':<6}{'Isentropic Turb Exhaust':<24}{p1:<10.3f}{t4_id:<10.2f}{v34_id[-1]:<12.4f}{s4_id:<10.4f}"
    )
    print("=" * 70 + "\n")

    wn = wt - wc
    n = wn / qs
    n_reg = wn / qs_reg
    print(f"{' PERFORMANCE EVALUATION ':=^70}")
    print(f"  Compressor Work Input:    {wc:.2f} kJ/kg")
    print(f"  Turbine Work Output:      {wt:.2f} kJ/kg")
    print(f"  Net Work Yield Output:    {wn:.2f} kJ/kg")
    print(f"  Standard Cycle Efficiency: {n*100:.2f}%")
    print(f"  Regenerated Efficiency:   {n_reg*100:.2f}%")
    print("=" * 70 + "\n")

    # --- PLOTTING CANVAS ---
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6.5))
    fig.canvas.manager.set_window_title("Single-Stage Performance Diagrams")

    ax1.plot(s12_act, t12_act, "-b", linewidth=2.5, label="Compression (Actual)")
    ax1.plot(s12_id, t12_id, ":b", linewidth=1.8, label="Compression (Ideal)")
    ax1.plot(s23_act, t23_act, "-r", linewidth=2.5, label="Combustion (Actual)")
    ax1.plot(s23_id, t23_id, ":r", linewidth=1.8, label="Combustion (Ideal)")
    ax1.plot(s34_act, t34_act, "-m", linewidth=2.5, label="Expansion (Actual)")
    ax1.plot(s34_id, t34_id, ":m", linewidth=1.8, label="Expansion (Ideal)")

    ax1.text(0.0, t1, " (1)", color="black", fontweight="bold")
    ax1.text(s2_act, t2, " (2)", color="blue", fontweight="bold")
    ax1.text(s3_act, t3, " (3)", color="red", fontweight="bold")
    ax1.text(s4_act, t4, " (4)", color="magenta", fontweight="bold")
    ax1.text(0.0, t2_id, " (2s)", color="blue", alpha=0.8, fontweight="bold")
    ax1.text(s3_id, t3dash, " (3s)", color="red", alpha=0.8, fontweight="bold")
    ax1.text(s4_id, t4_id, " (4s)", color="magenta", alpha=0.8, fontweight="bold")

    ax1.set_xlabel("Entropy Change Δs (kJ/kg-K)", fontweight="bold")
    ax1.set_ylabel("Absolute Temperature T (K)", fontweight="bold")
    ax1.set_title("Temperature-Entropy (T-S) Space")
    ax1.grid(True)
    ax1.legend(loc="best")

    ax2.plot(v12_act, p12_act, "-b", linewidth=2.5)
    ax2.plot(v12_id, p12_act, ":b", linewidth=1.8)
    ax2.plot(v23_act, p23_act, "-r", linewidth=2.5)
    ax2.plot(v23_id, p23_id, ":r", linewidth=1.8)
    ax2.plot(v34_act, p34_act, "-m", linewidth=2.5)
    ax2.plot(v34_id, p34_id, ":m", linewidth=1.8)

    ax2.text(v12_act[0], p1, " (1)", fontweight="bold")
    ax2.text(v12_act[-1], p2, " (2)", fontweight="bold")
    ax2.text(v23_act[-1], p3, " (3)", fontweight="bold")
    ax2.text(v34_act[-1], p4, " (4)", fontweight="bold")
    ax2.text(v12_id[-1], p2, " (2s)", color="blue", alpha=0.7, fontweight="bold")
    ax2.text(v23_id[-1], p2, " (3s)", color="red", alpha=0.7, fontweight="bold")
    ax2.text(
        v34_id[-1], p34_id[-1], " (4s)", color="magenta", alpha=0.7, fontweight="bold"
    )

    ax2.set_xlabel("Specific Volume v (m³/kg)", fontweight="bold")
    ax2.set_ylabel("Pressure P (bar)", fontweight="bold")
    ax2.set_title("Pressure-Volume (P-V) Space")
    ax2.grid(True)
    plt.tight_layout()
    plt.draw()
    plt.show(block=False)

    # =================================================================
    # EXCEL EXPORT ENGINE WITH FORMULAS & CELL STYLING
    # =================================================================
    if (
        input("\nExport single-stage workspace metrics to Excel? (y/n): ")
        .strip()
        .lower()
        == "y"
    ):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Single Stage Workspace"
        ws.views.sheetView[0].showGridLines = True

        # Style Presets
        font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_bold = Font(name="Segoe UI", size=11, bold=True)
        font_regular = Font(name="Segoe UI", size=11)
        fill_title = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        fill_header = PatternFill(
            start_color="2C3E50", end_color="2C3E50", fill_type="solid"
        )
        fill_accent = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )

        thin = Side(border_style="thin", color="D9D9D9")
        double = Side(border_style="double", color="000000")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        border_total = Border(top=thin, bottom=double)

        # Write Title block
        ws.merge_cells("A1:F1")
        ws["A1"] = "Single-Stage Gas Turbine Thermodynamic State Space"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_title
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        # State Table Data Headers
        headers = [
            "State Index",
            "Description Type",
            "Pressure (bar)",
            "Temperature (K)",
            "Spec. Volume (m³/kg)",
            "Entropy Change (kJ/kg-K)",
        ]
        for col_idx, text in enumerate(headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 25

        # Populate Raw Node Metrics
        state_rows = [
            ["1", "Actual Compressor Inlet", p1, t1, v1, 0.0],
            ["2", "Actual Compressor Outlet", p2, t2, v12_act[-1], s2_act],
            ["3", "Actual Combustor Outlet", p3, t3, v23_act[-1], s3_act],
            ["4", "Actual Turbine Exhaust", p4, t4, v34_act[-1], s4_act],
            ["2s", "Isentropic Compressor Outlet", p2, t2_id, v12_id[-1], 0.0],
            ["3s", "Isentropic Combustor Outlet", p2, t3dash, v23_id[-1], s3_id],
            ["4s", "Isentropic Turbine Exhaust", p1, t4_id, v34_id[-1], s4_id],
        ]

        for r_idx, r_data in enumerate(state_rows, start=4):
            for c_idx, val in enumerate(r_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.font = font_regular
                cell.border = border_all
                if c_idx in [1, 2]:
                    cell.alignment = Alignment(horizontal="left")
                else:
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "0.000" if c_idx == 3 or c_idx == 5 else "0.00"

        # Performance Block Header
        ws.cell(
            row=13, column=1, value="Performance Analysis Summaries"
        ).font = font_bold
        perf_headers = [
            "Performance Metric",
            "Algebraic Formula Framework",
            "Unit Dimension",
        ]
        for col_idx, text in enumerate(perf_headers, start=1):
            cell = ws.cell(row=14, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center")

        # Write Excel formulas that compute dynamically based on state data cells above!
        perf_rows = [
            ["Compressor Work Input (wc)", f"={cpc}*(D5-D4)", "kJ/kg"],
            ["Turbine Work Output (wt)", f"={cpe}*(D6-D7)", "kJ/kg"],
            ["Net Work Output Yield (wn)", "=B16-B15", "kJ/kg"],
            ["Combustor Heat Addition (qs)", qs, "kJ/kg"],
            ["Cycle Efficiency (eta)", "=(B17/B18)*100", "%"],
            ["Regenerated Heat Added (qs_reg)", qs_reg, "kJ/kg"],
            ["Regenerated Efficiency (eta_reg)", "=(B17/B20)*100", "%"],
        ]

        for r_idx, r_data in enumerate(perf_rows, start=15):
            for c_idx, val in enumerate(r_data, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = border_all
                if c_idx == 1:
                    cell.font = font_regular
                    cell.alignment = Alignment(horizontal="left")
                elif c_idx == 2:
                    if str(val).startswith("="):
                        cell.font = font_bold
                        cell.fill = fill_accent
                    else:
                        cell.font = font_regular
                    cell.alignment = Alignment(horizontal="right")
                    cell.number_format = "0.00"
                else:
                    cell.font = font_regular
                    cell.alignment = Alignment(horizontal="center")

        # Apply summary double bottom borders to final performance rows
        for col in range(1, 4):
            ws.cell(row=21, column=col).border = Border(
                bottom=double, left=thin, right=thin, top=thin
            )

        # Autofit Column Widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        filename = "Gas_Turbine_SingleStage_Analysis.xlsx"
        wb.save(filename)
        print(
            f"✓ Success! Data file with structural formulas generated at: {os.path.abspath(filename)}"
        )


# =====================================================================
# MODULE 2: Multi-Stage Cycle Analysis
# =====================================================================
def analyze_multi_stage():
    print("\n" + "=" * 70)
    print("                    MULTI-STAGE CYCLE ANALYSIS")
    print("                    ==========================")

    prompts = [
        "Relative Humidity (%)",
        "Mass flow Rate (kg/s)",
        "Ambient Temperature (K)",
        "Inlet Pressure (bar)",
        "Resolution Grid Count",
    ]
    defaults = ["0.0", "1", "300", "1.013", "1000"]
    ans = get_terminal_inputs(prompts, defaults)
    rh, m, t_amb, p_in, resln = (
        float(ans[0]),
        float(ans[1]),
        float(ans[2]),
        float(ans[3]),
        int(ans[4]),
    )
    r_const = 0.287

    print("\nDEVICE CONFIGURATION SEQUENCE")
    print("-" * 40)
    dev_seq = []
    while True:
        try:
            device_num = int(
                input(
                    "Enter choice device step index (1.Comp, 2.Cooler, 3.Combustor, 4.Turb, 5.Reheat, 6.Execute): "
                )
            )
            if device_num == 6:
                break
            if 1 <= device_num <= 5:
                dev_seq.append(device_num)
                names = [
                    "Compressor",
                    "Inter-Cooler",
                    "Combustion Chamber",
                    "Turbine",
                    "Reheater",
                ]
                print(f"  ✓ Appended Zone Block: {names[device_num-1]}")
        except ValueError:
            print("✕ Pass integer configuration values.")

    if not dev_seq:
        return

    use_reg = (
        input("\nInclude Regenerator element array? (y/n): ").strip().lower() == "y"
    )
    epselon = (
        float(input("Regenerator Effectiveness (0-1) [Default 0.8]: ") or "0.8")
        if use_reg
        else 0.0
    )

    states = [{"P": p_in, "T": t_amb, "type": "Inlet", "cp": 1.005}]
    states_id = [{"P": p_in, "T": t_amb, "type": "Inlet", "cp": 1.005}]

    total_comp_work = 0.0
    total_turb_work = 0.0
    total_heat_added_no_reg = 0.0
    compressor_exit_idx = None
    turbine_exit_idx = None

    # Store information required to formulate dynamic dynamic excel cell mapping rows downstream
    excel_meta = []

    for idx, dev in enumerate(dev_seq):
        in_state = states[-1]
        in_state_id = states_id[-1]
        new_state = {}
        new_state_id = {}
        stage_num = idx + 1

        if dev == 1:  # COMPRESSOR
            ans_c = get_terminal_inputs(
                [
                    f"Stage {stage_num} Pressure Ratio",
                    "Isentropic Eff (0-1)",
                    "Gamma",
                    "Cp",
                ],
                ["3", "0.82", "1.4", "1.005"],
            )
            pr, nc, gam, cp = (
                float(ans_c[0]),
                float(ans_c[1]),
                float(ans_c[2]),
                float(ans_c[3]),
            )

            new_state["P"] = in_state["P"] * pr
            t_is = in_state["T"] * (pr) ** ((gam - 1) / gam)
            new_state["T"] = in_state["T"] + (t_is - in_state["T"]) / nc
            new_state["type"] = "Compressor"
            new_state["cp"] = cp
            total_comp_work += cp * (new_state["T"] - in_state["T"])
            compressor_exit_idx = len(states)
            excel_meta.append(
                {
                    "dev": "Compressor",
                    "cp": cp,
                    "in_row": len(states) + 3,
                    "out_row": len(states) + 4,
                }
            )

            new_state_id["P"] = in_state_id["P"] * pr
            new_state_id["T"] = in_state_id["T"] * (pr) ** ((gam - 1) / gam)
            new_state_id["type"] = "Compressor"
            new_state_id["cp"] = cp

        elif dev == 2:  # INTER-COOLER
            ans_i = get_terminal_inputs(
                [f"Stage {stage_num} Target Temp (K)", "Pressure Drop (bar)"],
                ["310", "0.05"],
            )
            t_target, p_drop = float(ans_i[0]), float(ans_i[1])

            new_state["P"] = in_state["P"] - p_drop
            new_state["T"] = t_target
            new_state["type"] = "Intercooler"
            new_state["cp"] = 1.005
            new_state_id["P"] = in_state_id["P"]
            new_state_id["T"] = t_target
            new_state_id["type"] = "Intercooler"
            new_state_id["cp"] = 1.005
            excel_meta.append({"dev": "Intercooler"})

        elif dev == 3:  # COMBUSTION CHAMBER
            ans_cc = get_terminal_inputs(
                [
                    f"Stage {stage_num} Firing Temp (K)",
                    "Pressure Drop (%)",
                    "Combustion Eff (0-1)",
                    "Fuel CV",
                    "Gas Cp",
                ],
                ["1200", "3", "0.95", "42000", "1.148"],
            )
            t_ex, dp_pct, n_cc, cv_f, cpe = (
                float(ans_cc[0]),
                float(ans_cc[1]),
                float(ans_cc[2]),
                float(ans_cc[3]),
                float(ans_cc[4]),
            )

            new_state["P"] = in_state["P"] * (1 - dp_pct / 100)
            new_state["T"] = t_ex
            new_state["type"] = "Combustor"
            new_state["cp"] = cpe
            q_add = cpe * (new_state["T"] - in_state["T"]) / n_cc
            total_heat_added_no_reg += q_add
            excel_meta.append(
                {
                    "dev": "Combustor",
                    "q_add": q_add,
                    "in_row": len(states) + 3,
                    "out_row": len(states) + 4,
                    "cp": cpe,
                    "eff": n_cc,
                }
            )

            new_state_id["P"] = in_state_id["P"]
            new_state_id["T"] = t_ex
            new_state_id["type"] = "Combustor"
            new_state_id["cp"] = cpe

        elif dev == 4:  # TURBINE
            ans_t = get_terminal_inputs(
                [
                    f"Stage {stage_num} Exit Pressure (bar)",
                    "Isentropic Eff (0-1)",
                    "Gamma",
                    "Cp",
                ],
                ["1.013", "0.86", "1.33", "1.148"],
            )
            p_ex, nt, gam_exp, cpe = (
                float(ans_t[0]),
                float(ans_t[1]),
                float(ans_t[2]),
                float(ans_t[3]),
            )

            new_state["P"] = p_ex
            pr_exp = in_state["P"] / p_ex
            t_is = in_state["T"] / (pr_exp ** ((gam_exp - 1) / gam_exp))
            new_state["T"] = in_state["T"] - nt * (in_state["T"] - t_is)
            new_state["type"] = "Turbine"
            new_state["cp"] = cpe
            total_turb_work += cpe * (in_state["T"] - new_state["T"])
            turbine_exit_idx = len(states)
            excel_meta.append(
                {
                    "dev": "Turbine",
                    "cp": cpe,
                    "in_row": len(states) + 3,
                    "out_row": len(states) + 4,
                }
            )

            new_state_id["P"] = p_ex
            pr_exp_id = in_state_id["P"] / p_ex
            new_state_id["T"] = (
                in_state_id["T"] / (pr_exp_id ** ((gam_exp - 1) / gam_exp))
                if pr_exp_id > 0
                else in_state_id["T"]
            )
            new_state_id["type"] = "Turbine"
            new_state_id["cp"] = cpe

        elif dev == 5:  # REHEATER
            ans_rh = get_terminal_inputs(
                [f"Stage {stage_num} Reheat Temp (K)", "Pressure Drop (%)", "Gas Cp"],
                ["1150", "2", "1.148"],
            )
            t_rh, dp_pct, cpe = float(ans_rh[0]), float(ans_rh[1]), float(ans_rh[2])

            new_state["P"] = in_state["P"] * (1 - dp_pct / 100)
            new_state["T"] = t_rh
            new_state["type"] = "Reheater"
            new_state["cp"] = cpe
            q_add = cpe * (new_state["T"] - in_state["T"])
            total_heat_added_no_reg += q_add
            excel_meta.append(
                {
                    "dev": "Reheater",
                    "q_add": q_add,
                    "in_row": len(states) + 3,
                    "out_row": len(states) + 4,
                    "cp": cpe,
                }
            )

            new_state_id["P"] = in_state_id["P"]
            new_state_id["T"] = t_rh
            new_state_id["type"] = "Reheater"
            new_state_id["cp"] = cpe

        states.append(new_state)
        states_id.append(new_state_id)

    # Performance calculation reductions
    total_heat_added_actual = total_heat_added_no_reg
    if use_reg and compressor_exit_idx and turbine_exit_idx:
        if states[turbine_exit_idx]["T"] > states[compressor_exit_idx]["T"]:
            t_preheat = states[compressor_exit_idx]["T"] + epselon * (
                states[turbine_exit_idx]["T"] - states[compressor_exit_idx]["T"]
            )
            total_heat_added_actual = total_heat_added_no_reg - 1.148 * (
                t_preheat - states[compressor_exit_idx]["T"]
            )

    net_work = total_turb_work - total_comp_work

    # =================================================================
    # STATE SUMMARY MATRIX (MULTI-STAGE TERMINAL SUMMARY)
    # =================================================================
    print(f"\n{' RECONFIGURED MULTI-STAGE CYCLE STATION PROFILE ':=^70}")
    print(
        f"{'Node':<6}{'Inlet Component Zone':<25}{'Pressure (bar)':<16}{'Temperature (K)'}"
    )
    print("-" * 70)
    for index, st in enumerate(states, start=1):
        print(f"{index:<6}{st['type']:<25}{st['P']:<16.3f}{st['T']:.2f}")
    print("=" * 70 + "\n")
    print(f"  Aggregate Compressor Load: {total_comp_work:.2f} kJ/kg")
    print(f"  Aggregate Turbine Output:   {total_turb_work:.2f} kJ/kg")
    print(f"  System Combined Net Yield: {net_work:.2f} kJ/kg")
    print(
        f"  Thermal Net Efficiency:    {(net_work / total_heat_added_actual)*100:.2f}%"
    )
    print("=" * 70 + "\n")

    # --- INTERACTIVE PLOTTING ENGINE (DISSOCIATED OVERLAYS) ---
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6.5))
    fig.canvas.manager.set_window_title("Multi-Stage Performance Overlays")
    s_b_act = np.zeros(len(states))
    s_b_id = np.zeros(len(states_id))

    for k in range(len(states) - 1):
        in_p, out_p = states[k]["P"], states[k + 1]["P"]
        in_t, out_t = states[k]["T"], states[k + 1]["T"]
        type_str = states[k + 1]["type"]
        cp_l, gam_l = (
            (1.148, 1.33)
            if type_str in ["Combustor", "Turbine", "Reheater"]
            else (1.005, 1.4)
        )
        p_path_act = np.linspace(in_p, out_p, resln)

        if type_str == "Compressor":
            t_is_p = in_t * (p_path_act / in_p) ** ((gam_l - 1) / gam_l)
            denom = t_is_p[-1] - in_t
            # Safeguard against zero-pressure-rise or flat-temperature configurations
            if abs(denom) < 1e-6 or abs(out_t - in_t) < 1e-6:
                t_path_act = np.linspace(in_t, out_t, resln)
            else:
                t_path_act = in_t + (t_is_p - in_t) * ((out_t - in_t) / denom)
            t_path_act[0] = in_t
            color_s, disp_act = "b", "Compression (Actual)"

        elif type_str == "Turbine":
            t_is_p = in_t / (in_p / p_path_act) ** ((gam_l - 1) / gam_l)
            denom = in_t - t_is_p[-1]
            # Safeguard against zero-pressure-drop or flat-temperature configurations
            if abs(denom) < 1e-6 or abs(in_t - out_t) < 1e-6:
                t_path_act = np.linspace(in_t, out_t, resln)
            else:
                t_path_act = in_t - ((in_t - out_t) / denom) * (in_t - t_is_p)
            t_path_act[0] = in_t
            color_s, disp_act = "m", "Expansion (Actual)"
        else:
            t_path_act = np.linspace(in_t, out_t, resln)
            color_s = "g" if type_str == "Intercooler" else "r"
            disp_act = "Heat Exchange (Actual)"

        ds_act = cp_l * np.log(t_path_act / in_t) - r_const * np.log(p_path_act / in_p)
        s_p_act = s_b_act[k] + ds_act
        s_b_act[k + 1] = s_p_act[-1]
        v_p_act = (r_const * t_path_act) / (p_path_act * 100)

        ax1.plot(s_p_act, t_path_act, "-" + color_s, linewidth=2.5, label=disp_act)
        # Shift actual labels to the top-right of the point
        ax1.text(
            s_b_act[k], in_t, f" ({k+1})", fontweight="bold", ha="left", va="bottom"
        )

        ax2.plot(v_p_act, p_path_act, "-" + color_s, linewidth=2.5)
        # Shift actual labels to the top-right of the point
        ax2.text(
            v_p_act[0], in_p, f" ({k+1})", fontweight="bold", ha="left", va="bottom"
        )

        in_p_id, out_p_id = states_id[k]["P"], states_id[k + 1]["P"]
        in_t_id, out_t_id = states_id[k]["T"], states_id[k + 1]["T"]

        if type_str == "Compressor":
            p_path_id = np.linspace(in_p_id, out_p_id, resln)
            t_path_id = in_t_id * (p_path_id / in_p_id) ** ((gam_l - 1) / gam_l)
            s_p_id = np.repeat(s_b_id[k], resln)
            disp_id = "Compression (Ideal)"
        elif type_str == "Turbine":
            p_path_id = np.linspace(in_p_id, out_p_id, resln)
            t_path_id = in_t_id * (p_path_id / in_p_id) ** ((gam_l - 1) / gam_l)
            s_p_id = np.repeat(s_b_id[k], resln)
            disp_id = "Expansion (Ideal)"
        else:
            p_path_id = (
                np.repeat(in_p_id, resln)
                if type_str in ["Combustor", "Reheater"]
                else np.linspace(in_p_id, out_p_id, resln)
            )
            t_path_id = np.linspace(in_t_id, out_t_id, resln)
            ds_id = cp_l * np.log(t_path_id / in_t_id) - r_const * np.log(
                p_path_id / in_p_id
            )
            s_p_id = s_b_id[k] + ds_id
            disp_id = "Heat Exchange (Ideal)"

        s_b_id[k + 1] = s_p_id[-1]
        v_p_id = (r_const * t_path_id) / (p_path_id * 100)

        ax1.plot(s_p_id, t_path_id, ":" + color_s, linewidth=1.8, label=disp_id)
        # Shift ideal labels to the bottom-left of the point
        ax1.text(
            s_p_id[0],
            in_t_id,
            f" ({k+1}s)",
            color=color_s,
            alpha=0.8,
            fontweight="bold",
            ha="right",
            va="top",
        )

        ax2.plot(v_p_id, p_path_id, ":" + color_s, linewidth=1.8)
        # Shift ideal labels to the bottom-left of the point
        ax2.text(
            v_p_id[0],
            in_p_id,
            f" ({k+1}s)",
            color=color_s,
            alpha=0.8,
            fontweight="bold",
            ha="right",
            va="top",
        )

    # Final Boundary Vertex Node Labels
    ax1.text(
        s_b_act[-1],
        states[-1]["T"],
        f" ({len(states)})",
        fontweight="bold",
        ha="left",
        va="bottom",
    )
    ax1.text(
        s_b_id[-1],
        states_id[-1]["T"],
        f" ({len(states_id)}s)",
        fontweight="bold",
        ha="right",
        va="top",
    )
    ax1.set_xlabel("Entropy Change Δs (kJ/kg-K)", fontweight="bold")
    ax1.set_ylabel("Absolute Temperature T (K)", fontweight="bold")
    ax1.grid(True)
    preserve_unique_legend(ax1)

    ax2.text(
        v_p_act[-1],
        states[-1]["P"],
        f" ({len(states)})",
        fontweight="bold",
        ha="left",
        va="bottom",
    )
    ax2.text(
        v_p_id[-1],
        states_id[-1]["P"],
        f" ({len(states_id)}s)",
        fontweight="bold",
        ha="right",
        va="top",
    )
    ax2.set_xlabel("Specific Volume v (m³/kg)", fontweight="bold")
    ax2.set_ylabel("Pressure P (bar)", fontweight="bold")
    ax2.grid(True)
    plt.tight_layout()
    plt.draw()
    plt.show(block=False)

    # =================================================================
    # EXCEL EXPORT ENGINE FOR MULTI-STAGE DYNAMIC ARRAYS
    # =================================================================
    if (
        input("\nExport multi-stage state matrix nodes worksheet to Excel? (y/n): ")
        .strip()
        .lower()
        == "y"
    ):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MultiStage States Tracking"
        ws.views.sheetView[0].showGridLines = True

        font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
        font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        font_bold = Font(name="Segoe UI", size=11, bold=True)
        font_regular = Font(name="Segoe UI", size=11)
        fill_title = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        fill_header = PatternFill(
            start_color="2C3E50", end_color="2C3E50", fill_type="solid"
        )
        fill_accent = PatternFill(
            start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
        )

        thin = Side(border_style="thin", color="D9D9D9")
        double = Side(border_style="double", color="000000")
        border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

        # Header Row Title Banner
        ws.merge_cells("A1:E1")
        ws["A1"] = "Multi-Stage Complex Gas Turbine Engineering Simulation Spreadsheet"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_title
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40

        m_headers = [
            "State No",
            "Upstream Connected Component Inlet Type",
            "Pressure P (bar)",
            "Temperature T (K)",
            "Gas Specific Heat Cp",
        ]
        for col_idx, text in enumerate(m_headers, start=1):
            cell = ws.cell(row=3, column=col_idx, value=text)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[3].height = 25

        # Populate dynamic sequence layout row items
        for idx, st in enumerate(states, start=1):
            curr_row = idx + 3
            ws.cell(row=curr_row, column=1, value=idx).alignment = Alignment(
                horizontal="center"
            )
            ws.cell(row=curr_row, column=2, value=st["type"]).alignment = Alignment(
                horizontal="left"
            )
            ws.cell(row=curr_row, column=3, value=st["P"]).number_format = "0.000"
            ws.cell(row=curr_row, column=4, value=st["T"]).number_format = "0.00"
            ws.cell(row=curr_row, column=5, value=st["cp"]).number_format = "0.000"

            for col in range(1, 6):
                c_item = ws.cell(row=curr_row, column=col)
                c_item.font = font_regular
                c_item.border = border_all
                if col in [3, 4, 5]:
                    c_item.alignment = Alignment(horizontal="right")

        # Inject dynamic Excel Formulas down based on metadata tracking rows
        start_summary_row = len(states) + 5
        ws.cell(
            row=start_summary_row,
            column=1,
            value="Multi-Stage Energy Balance Summaries",
        ).font = font_bold

        ws.cell(
            row=start_summary_row + 1, column=1, value="Component Metric"
        ).font = font_header
        ws.cell(row=start_summary_row + 1, column=1).fill = fill_header
        ws.cell(
            row=start_summary_row + 1, column=2, value="Excel Calculated Formula"
        ).font = font_header
        ws.cell(row=start_summary_row + 1, column=2).fill = fill_header
        ws.cell(row=start_summary_row + 1, column=2).alignment = Alignment(
            horizontal="right"
        )
        ws.cell(row=start_summary_row + 1, column=3, value="Unit").font = font_header
        ws.cell(row=start_summary_row + 1, column=3).fill = fill_header
        ws.cell(row=start_summary_row + 1, column=3).alignment = Alignment(
            horizontal="center"
        )

        c_formulas = []
        t_formulas = []
        q_formulas = []

        for meta in excel_meta:
            if meta["dev"] == "Compressor":
                c_formulas.append(
                    f"E{meta['out_row']}*(D{meta['out_row']}-D{meta['in_row']})"
                )
            elif meta["dev"] == "Turbine":
                t_formulas.append(
                    f"E{meta['out_row']}*(D{meta['in_row']}-D{meta['out_row']})"
                )
            elif meta["dev"] in ["Combustor", "Reheater"]:
                if "eff" in meta:
                    q_formulas.append(
                        f"(E{meta['out_row']}*(D{meta['out_row']}-D{meta['in_row']}))/{meta['eff']}"
                    )
                else:
                    q_formulas.append(
                        f"E{meta['out_row']}*(D{meta['out_row']}-D{meta['in_row']})"
                    )

        comp_total_formula = "=" + "+".join(c_formulas) if c_formulas else "=0.0"
        turb_total_formula = "=" + "+".join(t_formulas) if t_formulas else "=0.0"

        # Build base heat addition formula
        heat_total_formula = "+".join(q_formulas) if q_formulas else "0.0"

        # FIX: If regenerator is active, append the live subtraction formula using cell references
        if use_reg and compressor_exit_idx and turbine_exit_idx:
            comp_out_row = compressor_exit_idx + 4
            turb_out_row = turbine_exit_idx + 4
            # Subtraction formula: Q_saved = Cp_gas * effectiveness * (T_turbine_exhaust - T_compressor_discharge)
            heat_total_formula = f"=({heat_total_formula}) - (1.148 * {epselon} * (D{turb_out_row} - D{comp_out_row}))"
        else:
            heat_total_formula = "=" + heat_total_formula

        r_c = start_summary_row + 2
        ws.cell(
            row=r_c, column=1, value="Total Compressor Work Input (W_c)"
        ).font = font_regular
        c_cell = ws.cell(row=r_c, column=2, value=comp_total_formula)
        c_cell.font = font_bold
        c_cell.number_format = "0.00"
        c_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=r_c, column=3, value="kJ/kg").alignment = Alignment(
            horizontal="center"
        )

        r_t = r_c + 1
        ws.cell(
            row=r_t, column=1, value="Total Turbine Work Output (W_t)"
        ).font = font_regular
        t_cell = ws.cell(row=r_t, column=2, value=turb_total_formula)
        t_cell.font = font_bold
        t_cell.number_format = "0.00"
        t_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=r_t, column=3, value="kJ/kg").alignment = Alignment(
            horizontal="center"
        )

        r_n = r_t + 1
        ws.cell(
            row=r_n, column=1, value="Cycle Net Work Yield output (W_net)"
        ).font = font_regular
        n_cell = ws.cell(row=r_n, column=2, value=f"=B{r_t}-B{r_c}")
        n_cell.font = font_bold
        n_cell.fill = fill_accent
        n_cell.number_format = "0.00"
        n_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=r_n, column=3, value="kJ/kg").alignment = Alignment(
            horizontal="center"
        )

        r_q = r_n + 1
        ws.cell(
            row=r_q, column=1, value="Total Cycle Thermal Heat Addition (Q_in)"
        ).font = font_regular
        q_cell = ws.cell(row=r_q, column=2, value=heat_total_formula)
        q_cell.font = font_bold
        q_cell.number_format = "0.00"
        q_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=r_q, column=3, value="kJ/kg").alignment = Alignment(
            horizontal="center"
        )

        r_e = r_q + 1
        ws.cell(
            row=r_e, column=1, value="Cycle Overall Efficiency (eta)"
        ).font = font_regular
        e_cell = ws.cell(row=r_e, column=2, value=f"=(B{r_n}/B{r_q})*100")
        e_cell.font = font_bold
        e_cell.fill = fill_accent
        e_cell.number_format = "0.00"
        e_cell.alignment = Alignment(horizontal="right")
        ws.cell(row=r_e, column=3, value="%").alignment = Alignment(horizontal="center")

        for row in range(r_c, r_e + 1):
            for col in range(1, 4):
                ws.cell(row=row, column=col).border = border_all
                if col == 3:
                    ws.cell(row=row, column=col).font = font_regular
        for col in range(1, 4):
            ws.cell(row=r_e, column=col).border = Border(
                bottom=double, left=thin, right=thin, top=thin
            )

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        filename = "Gas_Turbine_MultiStage_Analysis.xlsx"
        wb.save(filename)
        print(
            f"✓ Success! Multi-stage workspace tracking sheet written to: {os.path.abspath(filename)}"
        )


# =====================================================================
# SYSTEM MAIN ENGINE SELECTION LOOP
# =====================================================================
if __name__ == "__main__":
    while True:
        print("\n" + "=" * 70)
        print("   GAS TURBINE THERMODYNAMIC ANALYSIS TOOLKIT")
        print("   Integrated Single-File Analysis Tool (Python)")
        print("=" * 70)
        print(
            "MAIN ANALYSIS OPTIONS:\n  1: Single-Stage Cycle Analysis\n  2: Multi-Stage Cycle Analysis\n  3: Exit Program"
        )
        print("-" * 70)

        choice = input("Select analysis type (1-3): ").strip()
        if choice == "1":
            analyze_single_stage()
        elif choice == "2":
            analyze_multi_stage()
            time.sleep(1)
        elif choice == "3":
            print("\nToolkit terminated successfully.\n")
            break
        else:
            print("\nInvalid choice. Input 1, 2, or 3.")

# =====================================================================
# END OF GAS TURBINE THERMODYNAMIC ANALYSIS TOOLKIT
# =====================================================================