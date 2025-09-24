# Insurance-Claims-Reserving-VBA-Tool
This Excel VBA tool automates insurance claims reserving using Chain Ladder (CL) and Bornhuetter-Ferguson (BF) methods. It also performs Scenario Analysis to test reserve sensitivity with ±10% changes in development factors. The tool is designed for actuaries, students, or professionals learning claims reserving techniques.

Features

Claims Triangle Setup

Creates an n × m matrix for accident years and development years.

Upper table accepts non-incremental claim data.

Computes cumulative claims automatically.

Chain Ladder (CL) Method

Calculates Link Ratios / Development Factors (LDFs) from cumulative claims.

Projects future claims and computes reserves for each accident year.

Bornhuetter-Ferguson (BF) Method

Uses Expected Loss Ratio and premiums for each accident year.

Combines prior knowledge (premium × loss ratio) with development pattern (LDFs).

Calculates emerging liabilities and reserves.

Scenario Analysis

Tests reserve sensitivity by adjusting LDFs by +10% and -10%.

Recalculates reserves for both stressed scenarios.

Outputs a summary table comparing base reserve and stressed reserves.

How to Use

Open the Excel workbook and enable macros.

Run SetupClaimsTriangle to create the input triangle. Enter:

Number of accident years (n)

Number of development years (m)

Start year

Non-incremental claim amounts in the upper triangle

Run ChainLadder to compute CL reserves.

Run BornhuetterFerguson to compute BF reserves. Enter:

Expected Loss Ratio

Premiums for each accident year

Run ScenarioAnalysis to perform ±10% LDF stress tests and compare reserves.

Output

Chain Ladder Sheet: Shows projected claims and reserves per accident year.

Bornhuetter-Ferguson Sheet: Shows emerging liabilities and reserves using expected loss ratio.

Scenario Analysis Sheet: Shows stressed reserves for +10% and -10% LDFs with a summary comparison.

Notes

The tool assumes the upper triangle of claims is non-incremental.

Loss ratio in BF is assumed uniform across accident years.

Scenario analysis helps assess reserve sensitivity but does not replace professional judgment.
