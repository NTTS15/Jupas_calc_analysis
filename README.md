========================================
 JUPAS Score Calculator & Analysis Tool
========================================

**Version:** Beta 0.1 (Initial Release for Testing)
**Last Updated:** [Insert Date You Last Updated the Code/Data]

**--- Purpose ---**

This tool aims to help students estimate their potential JUPAS admission scores for various programs based on their DSE results and compare them against historical admission data. It also provides a D-Day report feature using uploaded student choices.

**--- Beta Version Disclaimer ---**

**IMPORTANT:** This is a **BETA TEST** version.
*   **Accuracy Not Guaranteed:** Score calculations are based on publicly available information and interpretations of admission formulas. These formulas can be complex and subject to change by universities. Official university calculators (if available) and JUPAS information should always be considered the primary source.
*   **Limited Calculation Rules:** Currently, the calculator implements logic for common patterns like "Best 5", "Best 6", "Best 4", "3C+2X", and specific rules identified for certain HKU (BBA/Econ), HKUST (Business, Science, Engineering), and CUHK (Science) programs based on 2024/2025 admission documents reviewed up to [Insert Date You Last Reviewed Docs]. **Many other specific program rules are NOT yet implemented.** Calculations for those programs will fail or be inaccurate.
*   **Data Cutoff Date:** The JUPAS program information, historical scores (LQ/M/UQ), intake quotas, and scoring methods included in this version are based on data extracted from `jupas_programs.xlsx`, which was last updated/converted on **2025-04-02**. This data reflects the 2024 admission cycle statistics and projected 2025 intake where available. Information may not reflect the most current updates from JUPAS or individual institutions.
*   **Bugs Expected:** As a beta version, there may be bugs, calculation errors, or UI issues.

**--- Features ---**

*   **Load JUPAS Data:** Automatically loads program information and historical scores (from embedded `jupas_programs.xlsx`).
*   **Load Student Choices:** Allows uploading a CSV file (exported from Google Forms or similar) containing student choices (expects columns like `RegNo`, `Class`, `ClassNo`, `Name`, `Choice1`, ..., `Choice20`). Use the sidebar uploader.
*   **View Program Info:** Search for programs by exact ID, program name keyword, or institution (using dropdown). Displays details including scoring method, weightings, and 2024 LQ/M/UQ scores.
*   **Enter DSE Results:** Input DSE grades using dropdowns for core and major elective subjects. Results are stored temporarily for calculations within the app session.
*   **Calculate Single Score:** Estimates the admission score for a single program based on entered DSE results and the specific university's point scale. Compares the result to 2024 LQ/M/UQ data and provides a suggestion (High Risk, Good Chance, etc.).
*   **Generate D-Day Report:** For a selected student (by Class and Class Number from the uploaded choices file), generates a report showing the calculated score and suggestion for each of their program choices, based on the entered DSE results. Includes a summary table and highlights the top 3 likely programs.
*   **Print Report:** A button is available on the D-Day report page to trigger the browser's print function.

**--- How to Use ---**

1.  The application should load the main JUPAS program data automatically.
2.  **(Optional)** Use the sidebar "Load Student Choices" button to upload your CSV file if you want to use the D-Day Report feature.
3.  Go to the "Enter DSE Results" tab (Tab 2) to input the relevant DSE grades. Click "Save / Update All DSE Results".
4.  Use "View Program Info" (Tab 1) to explore programs and their requirements/scores.
5.  Use "Calculate Single Score" (Tab 3) to check the estimated score for specific programs using the saved DSE results.
6.  Use "Generate D-Day Report" (Tab 4) to get a full analysis for a specific student (identified by Class/ClassNo) from the uploaded choices file, using the saved DSE results.

**--- Known Limitations / TODO ---**

*   Score calculation logic does not cover all possible JUPAS program variations. Accuracy is limited to implemented rules.
*   The mapping between full subject names and weighting dictionary keys (`SUBJECT_NAME_MAP_TO_WEIGHT_KEY`) might be incomplete.
*   DSE Elective entry is currently limited to a fixed number of slots.
*   Limited error handling for invalid data formats in uploaded CSV files.
*   User sessions are temporary; entered DSE results and uploaded choices are lost when the browser tab is closed or the app times out.
*   Print function uses browser default; formatting may not be perfect.

**--- Feedback ---**

Please report any bugs, calculation discrepancies (especially if you compare with official calculators), UI issues, or suggestions for improvement to [Your Name/Contact Method or GitHub Issues Link]. When reporting calculation issues, please include:
*   Program ID (e.g., JS1001)
*   The DSE results you entered
*   The score your app calculated
*   The score you expected (and why, e.g., from official calculator)

Thank you for testing!
