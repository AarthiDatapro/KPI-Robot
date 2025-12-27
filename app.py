from flask import Flask, render_template, request, send_file
import openpyxl
import random
from io import BytesIO
import os

app = Flask(__name__)

# ---------------- FIXED PROBABILITIES ----------------
V_PROBABILITY = 0.90
MARKS_PROBABILITY = 0.85
ANSWER_PROBABILITY = 0.90

# ---------------- FIXED PROJECTS ----------------
projects = {
    1: "Exploratory Analysis of Iris Flower Measurements",
    2: "Chemical Composition Analysis of Italian Wines",
    3: "Descriptive Statistics of Tumor Cell Nuclei Characteristics",
    4: "Profiling Health Indicators in Diabetes Patients",
    5: "Physical Exercise and Physiological Measurements: A Correlation Study",
    6: "Socioeconomic and Housing Trends in 1990 California Census Data",
    7: "Demographic and Survival Trends of Titanic Passengers",
    8: "Restaurant Billing and Tipping Behavior Analysis",
    9: "Demographic and Economic Profile of US Census Data (1994)",
    10: "Bank Loan Applicant Profiles of Credit Data"
}

# ---------------- HELPERS ----------------
def col_letter_to_index(letter):
    return ord(letter.upper()) - 64

def parse_columns(text):
    return [col_letter_to_index(c.strip()) for c in text.split(",") if c.strip()]

def parse_rows(text):
    rows = set()
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-")
            rows.update(range(int(start), int(end) + 1))
        else:
            rows.add(int(part))
    return sorted(rows)

# ---------------- ROUTE ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        # Read uploaded Excel directly (NO saving)
        file = request.files["file"]
        wb = openpyxl.load_workbook(file)

        option = request.form["option"]
        columns = parse_columns(request.form["columns"])
        rows = parse_rows(request.form["row_ranges"])

        sheets = wb.sheetnames[1:]  # skip first sheet

        # -------- OPTION 1: MARK "v" --------
        if option == "1":
            for sheet in sheets:
                ws = wb[sheet]
                for col in columns:
                    for row in rows:
                        if random.random() < V_PROBABILITY:
                            ws.cell(row=row, column=col).value = "v"

        # -------- OPTION 2: MARKS + TEAM + PROJECT --------
        elif option == "2":
            marks = int(request.form["marks"])

            sheet_index = 0
            team_number = 1

            while sheet_index < len(sheets) and team_number <= 10:
                for _ in range(5):  # 5 members per team
                    if sheet_index >= len(sheets):
                        break

                    ws = wb[sheets[sheet_index]]
                    ws["D4"] = team_number
                    ws["D5"] = projects[team_number]

                    for col in columns:
                        for row in rows:
                            if random.random() < MARKS_PROBABILITY:
                                ws.cell(row=row, column=col).value = marks
                            else:
                                ws.cell(row=row, column=col).value = random.randint(
                                    marks - 3, marks - 1
                                )

                    sheet_index += 1
                team_number += 1

        # -------- OPTION 3: ANSWERS --------
        elif option == "3":
            ref_column = request.form.get("ref_column", "").strip()
            if not ref_column:
                return "Reference column is required for Option 3", 400

            ref_col = col_letter_to_index(ref_column)

            for sheet in sheets:
                ws = wb[sheet]
                for col in columns:
                    for row in rows:
                        correct = ws.cell(row=row, column=ref_col).value
                        if correct not in [1, 2, 3, 4]:
                            continue

                        if random.random() < ANSWER_PROBABILITY:
                            ws.cell(row=row, column=col).value = correct
                        else:
                            wrong = [x for x in [1, 2, 3, 4] if x != correct]
                            ws.cell(row=row, column=col).value = random.choice(wrong)

        # -------- SEND FILE (IN MEMORY) --------
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="updated.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    return render_template("index.html")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))
