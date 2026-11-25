"""Case Study 2 - Student examination results.

Builds a frequency table of grades with pandas and works out the
percentage of students who received grade A.
"""

# pandas is a data analysis library; "pd" is the short name it is usually given.
import pandas as pd

# The grades of the ten students, stored as a pandas Series (a single column of data).
grades = pd.Series(["A", "B", "A", "C", "B", "A", "A", "C", "B", "A"])

# value_counts() counts how many times each grade appears: A = 5, B = 3, C = 2.
# sort_index() puts the grades in alphabetical order (A, B, C) instead of by count.
frequency = grades.value_counts().sort_index()

# frequency.sum() adds the counts: 5 + 3 + 2 = 10 (the total frequency, Σf).
total_frequency = frequency.sum()

# Percentage = (frequency ÷ total) × 100, worked out for every grade at once.
# A: 5 ÷ 10 × 100 = 50,  B: 3 ÷ 10 × 100 = 30,  C: 2 ÷ 10 × 100 = 20.
percentage = frequency / total_frequency * 100

# Put the counts and the percentages side by side in one table (a DataFrame).
table = pd.DataFrame({
    "Frequency": frequency,
    "Percentage (%)": percentage,
})

# Give the row labels column a heading so the table reads "Grade".
table.index.name = "Grade"

# Print the frequency table.
print(table)

# Print the total frequency (10 students).
print("Total frequency:", total_frequency)

# idxmax() returns the label with the largest count, i.e. the most frequent grade (A).
print("Most frequent grade:", frequency.idxmax())

# table.loc["A", "Percentage (%)"] reads the percentage in row A (50.0).
# :.0f shows it with no decimal places, so it prints as 50%.
print(f"Percentage of students with grade A: {table.loc['A', 'Percentage (%)']:.0f}%")
