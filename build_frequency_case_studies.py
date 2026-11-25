"""Build the five frequency-table case study submissions as one Word file.

Student: Dwayne Aitken
Course: Probability and Statistics for AI

Each frequency table is calculated here with collections.Counter, every
written answer is checked against that calculation, and each Python
program is run so its real output is printed under its code.
"""

import subprocess
import sys
from collections import Counter
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# Where the programs live and where the finished paper is saved.
PROJECT_FOLDER = Path(__file__).resolve().parent
PROGRAM_FOLDER = PROJECT_FOLDER / "programs"
OUTPUT_FOLDER = Path.home() / "Desktop" / "Frequency Table Case Studies"
OUTPUT_FILE = OUTPUT_FOLDER / "Frequency Table Case Studies - Dwayne Aitken.docx"

# Fonts and the grey used behind header and label cells.
BODY_FONT = "Times New Roman"
CODE_FONT = "Courier New"
SHADE_COLOUR = "E7E6E6"

# The five case studies: the brief's data, the written answers and the interpretation.
CASES = [
    {
        "title": "Case Study 1 – Vegetable Market",
        "scenario": "A vendor records the number of vegetable baskets sold each day "
                    "over ten days.",
        "value_label": "Baskets sold",
        "values": [10, 15, 10, 20, 15, 10, 25, 20, 10, 15],
        "mode": 10,
        "program": "case1_vegetable_market.py",
        "tool": "collections.Counter",
        "extra": None,
        "interpretation": (
            "The vendor sold 10 baskets on four of the ten days (40%), which makes 10 "
            "the most common daily sales quantity. Sales of 15 baskets followed on three "
            "days, while 25 baskets occurred only once. The pattern suggests that a "
            "typical day sits at the lower end of the range and that high-volume days "
            "are uncommon; the vendor could therefore plan stock around 10 to 15 "
            "baskets a day, keeping a small reserve for the occasional busy day."
        ),
    },
    {
        "title": "Case Study 2 – Student Examination Results",
        "scenario": "A teacher records the grades of ten students.",
        "value_label": "Grade",
        "values": ["A", "B", "A", "C", "B", "A", "A", "C", "B", "A"],
        "mode": "A",
        "program": "case2_exam_results.py",
        "tool": "pandas",
        "extra": ("Percentage of students with grade A",
                  "Percentage with grade A = (number of A grades ÷ total) × 100 = (5 ÷ 10) × 100 = 0.5 × 100 = 50%"),
        "interpretation": (
            "Half of the class (5 of 10 students, or 50%) received grade A, which is "
            "also the most frequent grade. A further 30% received a B and 20% received "
            "a C, and no student scored below C. Overall the results point to strong "
            "performance, although the two students at grade C may benefit from "
            "additional support before the next examination."
        ),
    },
    {
        "title": "Case Study 3 – Hospital Patient Categories",
        "scenario": "A hospital records the categories of eight patients in an "
                    "illustrative dataset.",
        "value_label": "Category",
        "values": ["General", "Emergency", "General", "Outpatient",
                   "Emergency", "General", "General", "Outpatient"],
        "mode": "General",
        "program": "case3_hospital_categories.py",
        "tool": "collections.Counter",
        "extra": None,
        "interpretation": (
            "General is the most frequent category, accounting for 4 of the 8 patients "
            "(50%), while Emergency and Outpatient each account for 2 patients (25%). "
            "These counts describe only how the eight recorded patients were "
            "categorised in this small illustrative sample. They do not measure how "
            "common any disease is in the population (disease prevalence), and they "
            "should not be used to draw conclusions about illness rates."
        ),
    },
    {
        "title": "Case Study 4 – Online Shopping",
        "scenario": "An online store records the product category purchased by each of "
                    "ten customers.",
        "value_label": "Category",
        "values": ["Clothing", "Books", "Clothing", "Electronics", "Books",
                   "Clothing", "Food", "Clothing", "Books", "Clothing"],
        "mode": "Clothing",
        "program": "case4_online_shopping.py",
        "tool": "pandas",
        "extra": ("Relative frequency of each category",
                  "Clothing 5 ÷ 10 = 0.5; Books 3 ÷ 10 = 0.3; Electronics 1 ÷ 10 = 0.1; "
                  "Food 1 ÷ 10 = 0.1; sum = 0.5 + 0.3 + 0.1 + 0.1 = 1.0"),
        "interpretation": (
            "Clothing has a relative frequency of 0.5, meaning that half of the ten "
            "purchases were clothing items, and it is the most frequent category. Books "
            "follow at 0.3, while Electronics and Food each have a relative frequency of "
            "only 0.1. The relative frequencies sum to 1.0, confirming that every "
            "purchase has been counted. The store may wish to give clothing the most "
            "prominent placement and review how electronics and food are promoted."
        ),
    },
    {
        "title": "Case Study 5 – Weather Monitoring",
        "scenario": "A weather station records the temperature (°C) over ten days.",
        "value_label": "Temperature (°C)",
        "values": [28, 30, 28, 32, 30, 28, 34, 30, 28, 32],
        "mode": 28,
        "program": "case5_weather_monitoring.py",
        "tool": "collections.Counter",
        "extra": None,
        "interpretation": (
            "28 °C was recorded on four of the ten days (40%), making it the most "
            "frequently recorded temperature. Readings of 30 °C occurred on three days, "
            "32 °C on two days and 34 °C on only one day. The frequency falls as the "
            "temperature rises, which indicates that the period was generally warm, with "
            "the hottest reading being the exception rather than the rule."
        ),
    },
]


def decimal(number):
    """Return a proportion with up to two decimals and no trailing zeros, e.g. 0.25, 0.5."""
    return f"{number:.2f}".rstrip("0").rstrip(".") if number % 1 else f"{number:.1f}"


def frequency_rows(values):
    """Return (value, frequency) pairs, numbers ascending and categories most common first."""
    counts = Counter(values)
    if all(isinstance(value, int) for value in values):
        return sorted(counts.items())
    return sorted(counts.items(), key=lambda pair: (-pair[1], str(pair[0])))


def check_case(case):
    """Stop the build if the written most frequent value is not the calculated one."""
    counts = Counter(case["values"])
    top_count = max(counts.values())
    leaders = [value for value, count in counts.items() if count == top_count]
    if leaders != [case["mode"]]:
        raise ValueError(f"{case['title']}: mode is {leaders}, written {case['mode']}")


def run_program(file_name):
    """Run one case study program and return exactly what it printed."""
    result = subprocess.run(
        [sys.executable, str(PROGRAM_FOLDER / file_name)],
        capture_output=True, text=True, check=True,
    )
    return result.stdout.rstrip()


def shade(cell):
    """Shade one table cell grey so headers stand out when printed."""
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), SHADE_COLOUR)
    cell._tc.get_or_add_tcPr().append(shading)


def write_cell(cell, text, bold=False):
    """Write text into a table cell in 11 pt body font."""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.keep_with_next = True
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(11)


def add_table(document, rows, widths, header=True, label_column=False):
    """Add a bordered table; shade the header row or the label column."""
    table = document.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    for row_index, row in enumerate(rows):
        for column_index, text in enumerate(row):
            cell = table.cell(row_index, column_index)
            cell.width = Inches(widths[column_index])
            is_label = (header and row_index == 0) or (label_column and column_index == 0)
            write_cell(cell, text, bold=is_label)
            if is_label:
                shade(cell)
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def add_code_block(document, text):
    """Write program code or output in a fixed-width font inside a shaded box."""
    table = document.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    shade(cell)
    # Keep the whole box on one page so code and output never split.
    table.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(0)
    for line in text.splitlines():
        run = paragraph.add_run(line)
        run.font.name = CODE_FONT
        run.font.size = Pt(8)
        run.add_break()
    document.add_paragraph().paragraph_format.space_after = Pt(0)


def frequency_table_rows(case):
    """Return the frequency table with relative frequency and percentage columns."""
    rows = frequency_rows(case["values"])
    total = len(case["values"])
    table = [(case["value_label"], "Frequency (f)", "Relative frequency", "Percentage")]
    for value, count in rows:
        table.append((value, count, f"{decimal(count / total)}", f"{count / total:.0%}"))
    table.append(("Total", total, "1.0", "100%"))
    return table


def answer_rows(case):
    """Return the label and value rows for the total, the mode and any extra answer."""
    counts = Counter(case["values"])
    total = len(case["values"])
    rows = [
        ("Total frequency (Σf)", f"{' + '.join(str(c) for _, c in frequency_rows(case['values']))} = {total}"),
        ("Most frequent value (mode)", f"{case['mode']} (frequency {counts[case['mode']]} of {total})"),
    ]
    if case["extra"]:
        rows.append(case["extra"])
    return rows


def add_heading(document, text, level):
    """Add a heading in the paper's font, black rather than Word's default blue."""
    heading = document.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = BODY_FONT
        run.font.color.rgb = RGBColor(0, 0, 0)
    return heading


def tally_marks(count):
    """Return tally marks for a count, grouping fives as |||| with a slash."""
    groups, remainder = divmod(count, 5)
    return " ".join(["||||/"] * groups + (["|" * remainder] if remainder else []))


def add_step(document, text):
    """Write one bold step heading inside the work-out."""
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.keep_with_next = True
    paragraph.add_run(text).bold = True


def add_line(document, text):
    """Write one indented line of working."""
    paragraph = document.add_paragraph(text)
    paragraph.paragraph_format.left_indent = Inches(0.3)
    paragraph.paragraph_format.space_after = Pt(2)


def add_working(document, case):
    """Write the full by-hand work-out: order, tally, total, relative frequency and mode."""
    values = case["values"]
    rows = frequency_rows(values)
    total = len(values)
    add_heading(document, "Work-out (by hand)", 2)

    # Step 1: put the data in order so equal values sit together.
    add_step(document, "Step 1 – Arrange the data so equal values are grouped together")
    grouped = [value for value, count in rows for _ in range(count)]
    add_line(document, ", ".join(str(value) for value in grouped))

    # Step 2: tally each distinct value and count the marks.
    add_step(document, "Step 2 – Tally each value and count the tally marks")
    tally_rows = [(case["value_label"], "Tally", "Frequency (f)")]
    tally_rows += [(value, tally_marks(count), count) for value, count in rows]
    add_table(document, tally_rows, (2.2, 2.4, 1.9))

    # Step 3: add the frequencies and check against the number of observations.
    add_step(document, "Step 3 – Find the total frequency (Σf)")
    add_line(document, f"Σf = {' + '.join(str(count) for _, count in rows)} = {total}")
    add_line(document, f"Check: there are {total} values in the data, so Σf = {total} is correct.")

    # Step 4: relative frequency = f ÷ Σf, and percentage = relative frequency × 100.
    add_step(document, "Step 4 – Relative frequency = f ÷ Σf, Percentage = relative frequency × 100")
    for value, count in rows:
        add_line(document, f"{value}:  {count} ÷ {total} = {decimal(count / total)}   →   "
                           f"{decimal(count / total)} × 100 = {count / total:.0%}")
    add_line(document, f"Check: {' + '.join(f'{decimal(c / total)}' for _, c in rows)} = 1.0 (100%)")

    # Step 5: the most frequent value is the one with the largest frequency.
    add_step(document, "Step 5 – Identify the most frequent value (mode)")
    comparison = " > ".join(f"{value} ({count})" for value, count in
                            sorted(rows, key=lambda pair: -pair[1]))
    add_line(document, f"Compare the frequencies: {comparison}")
    add_line(document, f"The largest frequency is {Counter(values)[case['mode']]}, "
                       f"so the most frequent value is {case['mode']}.")
    if case["extra"]:
        add_step(document, f"Step 6 – {case['extra'][0]}")
        add_line(document, case["extra"][1])


def add_case(document, case):
    """Write one complete case study: table, total, mode, program and interpretation."""
    add_heading(document, case["title"], 1)
    document.add_paragraph(case["scenario"])
    document.add_paragraph("Data: " + ", ".join(str(value) for value in case["values"]))
    add_working(document, case)

    add_heading(document, "1. Frequency Table", 2)
    add_table(document, frequency_table_rows(case), (1.9, 1.4, 1.7, 1.5))

    add_heading(document, "2–3. Total Frequency and Most Frequent Value", 2)
    add_table(document, answer_rows(case), (2.4, 4.1), header=False, label_column=True)

    add_heading(document, f"4. Python Program ({case['tool']})", 2)
    add_code_block(document, (PROGRAM_FOLDER / case["program"]).read_text())
    document.add_paragraph("Program output:").paragraph_format.keep_with_next = True
    add_code_block(document, run_program(case["program"]))

    add_heading(document, "5. Interpretation", 2)
    document.add_paragraph(case["interpretation"])
    document.add_page_break()


def add_title_block(document):
    """Write the title, the student details and the opening paragraph."""
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run("Case Study Solutions: Frequency Tables")
    title_run.bold = True
    title_run.font.size = Pt(16)
    add_table(
        document,
        [
            ("Student", "Dwayne Aitken"),
            ("Programme", "Associate Degree in Artificial Intelligence"),
            ("Course", "Probability and Statistics for AI"),
            ("Lesson by", "Prof. Chithra"),
            ("Institution", "Amber Heart Coding Academy"),
            ("Date", "November 6, 2026"),
        ],
        (1.4, 5.1), header=False, label_column=True,
    )
    document.add_paragraph(
        "This paper presents five independent case studies on frequency tables. For "
        "each dataset, the full work-out is first shown by hand: the data are grouped, "
        "tallied and totalled, the relative frequency and percentage of each value are "
        "calculated, and the most frequent value is identified. The results are then "
        "summarised in a frequency table, followed by a Python program using "
        "collections.Counter or pandas together with its actual output, and a short "
        "interpretation."
    )
    document.add_page_break()


def new_document():
    """Create a Letter-sized document with 1-inch margins and 12 pt Times New Roman."""
    document = Document()
    section = document.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Inches(1))
    normal = document.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)
    return document


def main():
    """Check every answer, then build and save the Word file."""
    for case in CASES:
        check_case(case)
    document = new_document()
    add_title_block(document)
    for case in CASES:
        add_case(document, case)
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    document.save(OUTPUT_FILE)
    print("Saved", OUTPUT_FILE)


if __name__ == "__main__":
    main()
