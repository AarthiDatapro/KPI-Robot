from flask import Flask, render_template, request, send_file
import openpyxl
import random
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Fixed probabilities
V_PROBABILITY = 0.98
MARKS_PROBABILITY = 0.85
ANSWER_PROBABILITY = 0.95

# Fixed projects
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
    return [col_letter_to_index(c.strip()) for c in text.split(",")]

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
        file = request.files["file"]
        option = request.form["option"]

        columns = parse_columns(request.form["columns"])
        rows = parse_rows(request.form["row_ranges"])

        filepath = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
        file.save(filepath)

        wb = openpyxl.load_workbook(filepath)
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
                for _ in range(5):
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
            ref_col = col_letter_to_index(request.form["ref_column"])

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

        output_path = filepath.replace(".xlsx", "_updated.xlsx")
        wb.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
