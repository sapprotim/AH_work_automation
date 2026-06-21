# AH Work Automation

End-to-end automation suite for **ActiveHealth (AH) study data management**, covering health metric collection, participant alert/nudge assignment, and data quality validation.

---

## Project Structure

```
AH_work_automation/
├── Data Management/          # Selenium scripts to extract health metrics from the dashboard
├── Alerts & Nudge/           # Selenium-based alert/nudge assignment automation
├── Alerts_and_Nudge_POM/     # Playwright + Page Object Model rewrite of alert/nudge assignment
├── Data Checking/            # Google Apps Script for automated data validation in Google Sheets
└── kt/                       # Pytest test directory
```

---

## Modules

### 1. Data Management

Extracts 20+ health metrics per user from the GetHealthier dashboard and writes them to `user_id.xlsx`.

| File | Purpose |
|------|---------|
| `Automation.py` | Main data extraction script — loops through all users |
| `User_Search.py` | User lookup with OTP handling and retry logic |
| `Height_Data.py` | Focused extraction of height and weight data |
| `user_id.xlsx` | Input/output Excel file with one row per user per date |

**Metrics collected per user:**

| Category | Metrics |
|----------|---------|
| Activity | Step Count, Total Sedentary Time (mins), Time Wearing Device (%) |
| Vitals | Heart Rate (Avg BPM), SpO2 (%), Blood Pressure (Systolic & Diastolic mmHg) |
| Sleep | Sleep Duration (hours + minutes) |
| Metabolic | Blood Glucose (mmol/L), HbA1c (%), Weight (kg), Height (cm), Waist Circumference (cm), BMI, Blood Cholesterol |
| Wellness | Wellness Score, Health Report status |
| App Usage | App Opened Count, Time Spent in App (mins) |

**Workflow:**
1. Read user names from `user_id.xlsx`
2. Log into GetHealthier platform
3. Search each user → navigate to Overview for the target date
4. Extract metrics via XPath/CSS selectors
5. Navigate to Analysis page for sedentary time
6. Write collected values back to Excel
7. Handle re-login and network recovery automatically

---

### 2. Alerts & Nudge (Selenium)

Automates assignment of health alerts and nudges to study participants via the GetHealthier dashboard.

| File | Purpose |
|------|---------|
| `Alerts and Nudge.py` | Main Selenium WebDriver script |
| `user list.xlsx` | User IDs segmented by study arm and language |
| `Alerts Nudge list.xlsx` | Alert and nudge template definitions |

**Assignment Rules:**
- **Nudges** → Study Arms 2 and 3 (English & Chinese)
- **Alerts** → Study Arm 3 only (English & Chinese)

**Workflow:**
1. Read user list and alert/nudge templates from Excel
2. Log into the GetHealthier Alerts/Nudges Automation Dashboard
3. Navigate to Programme Management
4. Filter users by Study Arm and Language Preference
5. Assign appropriate nudges and alerts per user
6. Confirm success messages after each assignment

---

### 3. Alerts & Nudge POM (Playwright)

Playwright-based rewrite of the alert/nudge module using the **Page Object Model** pattern for better maintainability.

| File | Purpose |
|------|---------|
| `test.py` | Playwright implementation with shared `assign_alerts_nudges()` function |
| `user list.xlsx` | Same user list as the Selenium module |
| `Alerts Nudge list.xlsx` | Same alert/nudge templates |

**Improvements over the Selenium version:**
- Playwright for more reliable async browser control
- Page Object Model for cleaner, maintainable code
- More robust locator strategies
- Better exception handling with specific error types

---

### 4. Data Checking (Google Apps Script)

Two Google Apps Scripts that run inside Google Sheets to automate data validation and external data sync.

| File | Purpose |
|------|---------|
| `Data Checking.txt` | Validates health metrics against rules; highlights errors in yellow |
| `Mix_Panel.txt` | Syncs Mixpanel analytics data into the Google Sheet |

#### Validation Rules

**Consistency Checks:**
- If Time Wearing Device = 100% → Heart Rate, Sleep, and Step data must be present
- If Time Wearing Device = NULL/0% → Heart Rate, Sleep, and Step data must also be NULL
- SpO2 data recorded during sleep requires corresponding sleep data
- App Opened Count and Time Spent in App must both be present or both be absent (neither can be 0 while the other has data)
- BMI = Weight (kg) ÷ Height (m)²

**Maximum Value Checks:**

| Metric | Maximum Allowed |
|--------|----------------|
| User ID | 9 digits |
| Study Arm | 1, 2, or 3 only |
| Heart Rate | 150 bpm |
| Step Count | 40,000 steps |
| Sedentary Time | 1,500 mins |
| Time Wearing Device | 100% |
| Wellness Score | 100% |
| SpO2 | 100% |
| App Opened Count | If double-digit, divide by 3 |
| Systolic BP | 180 mmHg |
| Diastolic BP | 120 mmHg |
| Blood Glucose | 200 mg/dL |
| HbA1c | 30% |
| Weight | 180 kg |
| Height | 200 cm |
| Waist Circumference | 120 cm |
| BMI | 50 |
| Blood Cholesterol | 240 mg/dL |

Cells failing any check are highlighted **yellow**; error details are written to the Comments column.

---

## Technology Stack

| Component | Technology |
|-----------|-----------|
| Browser Automation | Selenium WebDriver, Playwright |
| Language | Python 3.12 |
| Data Processing | Pandas, OpenPyXL |
| Data Validation | Google Apps Script (JavaScript) |
| Testing Framework | Pytest |
| IDE | PyCharm |
| Analytics Sync | Mixpanel → Google Sheets |
| Target Platform | GetHealthier web application |

---

## Prerequisites

```
selenium
playwright
pandas
openpyxl
pytest
```

Install Python dependencies:
```bash
pip install selenium playwright pandas openpyxl pytest
playwright install
```

---

## End-to-End Workflow

```
Excel Input (user_id.xlsx / user list.xlsx)
        ↓
  Selenium / Playwright Automation
        ↓
GetHealthier Dashboard  →  Extract Metrics / Assign Alerts & Nudges
        ↓
  Write Results to Excel
        ↓
  Upload to Google Sheets
        ↓
  Apps Script Validation  →  Yellow Highlights + Error Comments
        ↓
  Manual Review & Correction
        ↓
  Final Submission
```

---

## Study Design Reference

| Study Arm | Nudges | Alerts |
|-----------|--------|--------|
| Arm 1 | No | No |
| Arm 2 | Yes | No |
| Arm 3 | Yes | Yes |

Language support: **English** and **Chinese** for both alerts and nudges.
