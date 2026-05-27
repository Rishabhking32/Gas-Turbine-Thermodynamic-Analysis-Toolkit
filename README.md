
### Gas Turbine Thermodynamic Analysis Toolkit

```
A high-fidelity thermodynamic simulation and analysis engine written in Python for evaluating single-stage and complex multi-stage gas turbine cycles. This toolkit allows users to model real-world aerothermal behaviors by accounting for component efficiencies, pressure drops, fluid properties variations ($C_p$, $\gamma$), and environmental conditions.

The application generates clear Temperature-Entropy (T-S) and Pressure-Volume (P-V) diagrams overlaid with an isolated, continuous Isentropic Reference Cycle. It also features an automated, formula-driven Excel data export utility to generate professional worksheets for engineering documentation.

| Key Features

1. Single-Stage Cycle Analysis (Brayton Cycle)
* Real vs. Ideal Modeling: Simulates loss-driven actual cycles alongside parallel, isolated isentropic traces.
* Open-Cycle Architecture: Graphically models true open-cycle configurations by omitting artificial condenser/rejection return lines.
* Regenerative Evaluation: Computes alternative cycle efficiencies assuming a localized heat exchanger matrix is active.

2. Multi-Stage Cycle Customization
* Dynamic Stack Building: Assemble complex industrial architectures on the fly. Supported components include Compressors, Intercoolers, Combustion Chambers, Turbines, and Reheaters.
* Friction & Loss Matching: Integrates localized component isentropic efficiencies, pressure decays, and variable fluid heat capacities ($C_p$ and $\gamma$) unique to cold or hot gas zones.
* Safeguarded Numerical Engine: Equipped with isentropic denominator protections to gracefully handle idling stages or pressure ratios of 1.0 without throwing calculation math warnings.

3. High-Resolution Visual Tracking
* Collision-Free Labeling: Side-by-side T-S and P-V subplots feature split text-offset alignments (ha/va anchors) to cleanly separate real states `(1, 2...)` from pure states `(1s, 2s...)`.
* True Isentropes: Ideal compressor and turbine traces maintain perfectly vertical projections (Delta s = 0) regardless of external component friction values.

4. Interactive Spreadsheet Writer
* Live Excel Formulas: The export subsystem uses `openpyxl` to write active Excel equations (e.g., product/difference cell metrics) rather than printing flat numerical strings.
* Professional Typography: Features matching color banners, auto-fitting grid column widths, cell dimension padding, and explicit double-underline format summaries.


| Installation & Dependencies

This toolkit requires Python 3.8+ and relies on standard scientific and formatting libraries.

### Prerequisites
Install the required packages using `pip`:

->bash
pip install numpy matplotlib pandas openpyxl

```

*Alternatively, if you manage your sandbox via `uv`, you can execute the script inline without installing packages globally:*

```bash
uv run --with numpy --with matplotlib --with pandas --with openpyxl Untitled-1.py

```

---

## How To Use

1. Run the script from your terminal:
```bash
python Untitled-1.py

```


2. Select your analysis option from the master control console:
* **`1`**: Single-Stage Cycle Analysis.
* **`2`**: Multi-Stage Cycle Analysis.
* **`3`**: Terminate Program.


3. Follow the terminal prompt instructions. Press **`Enter`** at any step to accept the preloaded baseline engineering default value.
4. Review the real-time data table printed in the terminal panel, inspect the Matplotlib graphic layout window, and choose `y` to write a fully styled data spreadsheet directly to your local workspace directory.

---

## Sample Multi-Stage Test Case

To evaluate a high-efficiency **Intercooled-Reheat-Regenerative Gas Turbine** configuration, use the following sequence in Option 2:

* **Device Order**: `1` (LP Compressor) $\rightarrow$ `2` (Intercooler) $\rightarrow$ `1` (HP Compressor) $\rightarrow$ `3` (Combustor) $\rightarrow$ `4` (HP Turbine) $\rightarrow$ `5` (Reheater) $\rightarrow$ `4` (LP Turbine) $\rightarrow$ `6` (Execute Process).
* **Regenerator Integration**: Select `y` with an effectiveness index of `0.8`.

### Baseline Checkpoints

* Compressor Stages: Pressure Ratio = `3.5`, Isentropic Efficiency = `0.82`.
* Main Combustor: Firing Limit = `1300 K`, Pressure Loss = `3%`.
* Turbine Stages: Isentropic Efficiency = `0.86`.
* Reheater Zone: Reheat Limit = `1250 K`, Pressure Loss = `2%`.

---

## Program Architecture

```
Untitled-1.py
│
├── get_terminal_inputs()        # Safe console configuration parsing
├── preserve_unique_legend()     # Matplotlib label deduplication handler
│
├── analyze_single_stage()       # Closed-form single expansion loops
│   └── [Excel Export Subsystem] # Openpyxl formula injection mapping
│
└── analyze_multi_stage()        # Dynamic state array stacking machine
    ├── [Actual State Engine]    # Real-fluid loss path evaluation
    ├── [Ideal Isentropic Core]  # Parallel standalone reference loop
    └── [Advanced Excel Writer]  # Variable row-indexing worksheet generator

```

---

## Author

* **Developer**: Rishabh
* **Release Date**: May 26, 2026

```

```
