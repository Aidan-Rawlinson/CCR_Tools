import openpyxl
import random
import os

# ── Valid response values per question (matching Drop downs sheet exactly) ──

DT_DATES = ["01/07/2026", "15/07/2026", "30/06/2026", "10/08/2026", "22/07/2026"]

YN = ["Yes", "No"]

LS_390861 = ["Admitted to acute care", "Service user refused/declined",
             "No rehabilitation potential/no goals identified",
             "Lack of capacity/no bed available", "Assessment only",
             "Spontaneous recovery", "Palliative care", "Other (please specify)"]

LS_390863 = ["Male", "Female", "Other gender identity", "Not known"]

LS_62631  = ["White", "Asian/Asian British", "Black/Black British", "Mixed", "Other", "Unknown"]

LS_390864 = ["Acute hospital inpatient ward", "NHS 111", "NHS 999", "SDEC",
             "Outpatient department", "Emergency department", "GP / primary care",
             "Mental Health service", "Community health service",
             "Urgent Community Response", "Other"]

LS_151166 = ["Yes", "No", "Unknown", "Not relevant"]

LS_62635  = ["Tier 1", "Tier 2", "Tier 3", "Tier 4", "Tier 5"]

LS_151167 = ["Face to face only", "Non face to face only",
             "Remote monitoring only", "A combination"]

LS_151168 = ["Face to face visits", "Non face to face visits", "Remote monitoring"]

LS_62638  = ["Yes - in a previous service", "Yes - in this service",
             "No", "NA", "Don't know"]

LS_151171 = ["1", "2", "3", "4", "5", "6", "7", "8", "9",
             "The Rockwood Clinical Frailty Scale is not used as an indication"]

LS_151172 = ["Yes - in a previous service", "Yes - in this service",
             "No", "NA", "Don't know"]

LS_151173 = ["Yes - in a previous service", "Yes - in this service",
             "No", "NA", "Don't know"]

LS_151174 = ["Yes - in a previous service", "Yes - in this service",
             "No", "NA", "Don't know"]

LS_390867 = ["Yes", "No", "NA", "Don't know"]

LS_62641  = ["Yes", "No", "Partially", "No goals set on admission"]

LS_62643  = ["Usual place of residence", "Temporary place of residence",
             "Care home", "Acute hospital", "Patient died", "Unknown", "Other"]

LS_62652  = ["Increase", "Decrease", "No change", "NA"]

LS_151175 = ["Care home", "Community nursing", "Palliative care",
             "Home care package", "GP", "Reablement/Discharge to assess",
             "Hospital admission", "Specialist teams", "Informal care", "Other"]


def random_age():
    return random.randint(45, 95)


def random_los():
    return random.randint(1, 30)


def random_days_after():
    return random.randint(0, 10)


def patient_data(n):
    """Return a dict of row -> value for one patient column."""
    return {
        6:  random.choice(DT_DATES),
        7:  random.choice(YN),
        8:  random.choice(YN),
        9:  random.choice(LS_390861),
        10: "Other reason text " + str(n),
        12: random.choice(LS_390863),
        13: random.choice(LS_62631),
        14: random_age(),
        15: random.choice(LS_390864),
        16: "Admitted from other " + str(n),
        17: random.choice(LS_151166),
        18: random.choice(LS_62635),
        19: random.choice(LS_151167),
        20: random.choice(LS_151168),
        21: random.choice(YN),
        22: random.choice(YN),
        23: random.choice(YN),
        25: random.choice(YN),
        26: random.choice(YN),
        27: random.choice(LS_62638),
        28: random.choice(LS_151171),
        29: random.choice(LS_151172),
        30: random.choice(LS_151173),
        31: random.choice(LS_151174),
        32: random.choice(LS_390867),
        34: random_los(),
        35: random.choice(LS_62641),
        36: random_days_after(),
        37: random.choice(LS_62643),
        38: "Discharge other " + str(n),
        39: random.choice(LS_62652),
        40: random.choice(LS_151175),
        41: "Care package other " + str(n),
    }


def build_file(template_path, out_path, org_name, vw_name, num_patients):
    wb = openpyxl.load_workbook(template_path)
    ws_sub = wb["Submission Details"]
    ws_ccr = wb["CCR"]

    ws_sub["B6"] = org_name
    ws_sub["B9"] = vw_name

    for p in range(1, num_patients + 1):
        col = 4 + p  # Patient 1 = col E = col index 5
        data = patient_data(p)
        for row, val in data.items():
            ws_ccr.cell(row=row, column=col, value=val)

    wb.save(out_path)
    print(f"Saved: {out_path} ({num_patients} patients)")


template = r"C:\mcp_projects\CCR_Tools\new_questionnaires\NHSBN Virtual Ward Clinical Case Review Specification 2026 FINAL.xlsx"
out_dir  = r"C:\mcp_projects\CCR_Tools\test_inputs"

random.seed(42)

build_file(template,
           os.path.join(out_dir, "VW_Test_EssexPartnership_75.xlsx"),
           "Essex Partnership University NHS Foundation Trust",
           "MSE Virtual Hospital",
           75)

build_file(template,
           os.path.join(out_dir, "VW_Test_BromleyHealthcare_10.xlsx"),
           "Bromley Healthcare CIC Ltd",
           "Bromley Healthcare CIC Ltd",
           10)

build_file(template,
           os.path.join(out_dir, "VW_Test_CornwallPartnership_10.xlsx"),
           "Cornwall Partnership NHS Foundation Trust",
           "Cornwall Partnership NHS Foundation Trust",
           10)

print("Done.")
