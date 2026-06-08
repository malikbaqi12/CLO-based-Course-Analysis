╔══════════════════════════════════════════════════════════════════════════════╗
║   IQRA University – Academic Result Analytics Platform                     ║
║   README.txt                                                               ║
╚══════════════════════════════════════════════════════════════════════════════╝

VERSION:   3.0
TESTED ON: AIC-211 (Intro to AI) · AIC-211L (Intro to AI Lab)
           AIN-375 (Data Visualization) · BAN-114 (Intro to CS)
           CMC-112 (Object Oriented Programming)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. WHAT THIS SYSTEM DOES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Upload any course result workbook that follows the standard template and get:

  ✅ Auto-detection of course type (Theory / Lab)
  ✅ Auto-detection of pass mark (60% / 50% / any)
  ✅ Auto-detection of assessment weightages
  ✅ Auto-detection of CLO count (3, 4, or more)
  ✅ GA attainment from Quantized Result sheet
  ✅ 11+ interactive Plotly charts
  ✅ Risk analysis (at-risk, borderline, top performers)
  ✅ Downloadable HTML dashboard, CSV, and Excel reports

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
2. FOLDER STRUCTURE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  academic_dashboard/
  ├── app.py                          ← Streamlit frontend
  ├── engine.py                       ← Analytics backend (import this)
  ├── Academic_Result_Analytics.ipynb ← Google Colab notebook
  ├── requirements.txt                ← Python dependencies
  └── README.txt                      ← This file

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
3. DEPENDENCIES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  openpyxl >= 3.1.0     # Excel parsing
  pandas   >= 2.0.0     # Data analysis
  numpy    >= 1.24.0    # Statistics
  plotly   >= 5.18.0    # Interactive charts
  streamlit>= 1.32.0    # Web dashboard (optional – Colab works without it)
  kaleido  >= 0.2.1     # Chart image export (optional)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
4. GOOGLE COLAB (RECOMMENDED – ZERO SETUP)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Step 1: Open Google Colab → https://colab.research.google.com
  Step 2: File → Upload notebook → select Academic_Result_Analytics.ipynb
  Step 3: Runtime → Run all
  Step 4: When prompted, upload your .xlsx result sheet
  Step 5: All charts appear inline; HTML + CSV auto-download at the end

  That's it. No installation needed on your machine.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
5. STREAMLIT APP (LOCAL OR CLOUD)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  LOCAL INSTALLATION:
  -------------------
  pip install -r requirements.txt
  streamlit run app.py

  Then open http://localhost:8501 in your browser.

  STREAMLIT CLOUD DEPLOYMENT:
  ---------------------------
  1. Push this folder to a GitHub repository
  2. Go to https://share.streamlit.io
  3. Click "New app" → select your repo → set "Main file path" to app.py
  4. Click Deploy

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
6. USING ENGINE.PY IN YOUR OWN CODE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  from engine import ResultEngine

  with open('my_result_sheet.xlsx', 'rb') as f:
      engine = ResultEngine(f.read())

  print(engine.meta)           # course metadata
  print(engine.stats)          # pass/fail, avg, high, low, std
  print(engine.clo_stats)      # CLO attainment per CLO
  print(engine.ga_stats)       # GA attainment per GA
  print(engine.df.head())      # full student DataFrame
  print(engine.is_lab)         # True / False
  print(engine.pass_mark)      # auto-detected pass %

  # Build standalone HTML dashboard:
  html = engine.build_html_dashboard(clo_thresh=50, gc_fn=..., C=..., GRADE_ORDER=...)
  with open('dashboard.html', 'w') as f:
      f.write(html)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
7. WORKBOOK REQUIREMENTS (TEMPLATE COMPATIBILITY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  The workbook must contain these sheets (names are flexible as long as they
  contain these keywords):
    • "Final Combined Marks Sheet"   – main marks sheet
    • "Quantized Result"             – CLO/GA attainment sheet

  Optional (auto-used if present):
    • "Weighted Theory Marks Sheet"  – for weightage extraction
    • "Award List (Theory)"          – grade distribution reference

  Row structure (Final Combined):
    • Rows 3–9:  Course metadata (label in col A, value in col C or col H)
    • Row 10:    Column headers (Sr, Reg, Name, Assg, Quiz, ..., Grade)
    • Row 11:    Max marks per component
    • Row 12+:   Student data rows

  The pass mark hint cell: "Having marks less than 59 i.e. ..." → auto-parsed
  If no hint found, the fallback from config is used.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
8. DASHBOARD CHARTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  1.  Pass/Fail donut + Grade distribution donut
  2.  Grade distribution bar chart
  3.  Marks distribution histogram (with avg and pass-mark lines)
  4.  Assessment component performance (avg obtained vs max, and %)
  5.  CLO attainment bar chart vs threshold
  6.  CLO pass/fail grouped bar chart
  7.  CLO radar chart (requires ≥ 2 CLOs)
  8.  CLO attainment heatmap per student
  9.  GA attainment bar chart (if GA data present)
  10. Top N / Bottom N students bar chart
  11. Box plot by grade
  12. Mid Term vs End Term scatter (Theory only)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
9. TROUBLESHOOTING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Error: "Final Combined Marks Sheet not found"
  → Check that your workbook has a sheet whose name contains "Final Combined"

  Error: "Quantized Result sheet not found"
  → Check that your workbook has a sheet whose name contains "Quantized"

  CLO count is 0
  → Check that the Quantized sheet has cells starting with "CLO 1", "CLO 2" etc.
    They must appear in at least 2 columns to be detected.

  Pass mark is wrong
  → The "Having marks less than N" hint cell was not found.
    Override the pass mark in the sidebar (Streamlit) or Cell 3 (Colab).

  Student count seems wrong
  → The engine stops at the first non-numeric value in the Sr. column.
    Check for blank rows or "Class Average" rows mid-table.

  Lab course not detected
  → The engine looks for "Lab Report", "Lab Project", "Viva", "Open Ended"
    in the header row. If your lab sheet uses different labels, the engine
    will fall back to Theory mode (still works; just different column names).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
10. SAMPLE WORKFLOW
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  Semester ends → Faculty exports result sheet from LMS
        ↓
  Open Colab → Upload .ipynb → Run all → Upload .xlsx
        ↓
  Dashboard generated in ~10 seconds
        ↓
  Download HTML dashboard (share with HOD / accreditation team)
  Download CSV report (for record-keeping)
  Download Excel report (CLO + GA sheets for OBE compliance)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONTACT / SUPPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Faculty: Mr. Abdul Baqi Malik
  Department: Computing & Technology
  IQRA University Chak Shahzad Campus, Islamabad
