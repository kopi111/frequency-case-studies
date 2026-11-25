"""Case Study 3 - Hospital patient categories (illustrative data).

Builds a frequency table of patient categories with collections.Counter
and finds the most frequent category. The counts describe this small
sample only; they are not disease prevalence figures.
"""

# Counter is a built-in Python tool that counts how many times each item appears.
from collections import Counter

# The category of each of the eight recorded patients (one word per patient).
categories = ["General", "Emergency", "General", "Outpatient",
              "Emergency", "General", "General", "Outpatient"]

# Counter counts each category.
# Result: Counter({"General": 4, "Emergency": 2, "Outpatient": 2}).
frequency = Counter(categories)

# frequency.values() keeps only the counts: 4, 2, 2.
# sum() adds them: 4 + 2 + 2 = 8 (Σf), which equals the number of patients.
total_frequency = sum(frequency.values())

# Print the column headings of the frequency table.
print("Category   | Frequency")

# most_common() with no number lists every category from most to least frequent:
# [("General", 4), ("Emergency", 2), ("Outpatient", 2)].
# Each pair is split into category (the name) and count (how many patients).
for category, count in frequency.most_common():
    # :<10 left-aligns the name and :>9 right-aligns the count so the columns line up.
    print(f"{category:<10} | {count:>9}")

# Print the total frequency worked out above.
print("Total frequency:", total_frequency)

# most_common(1) returns only the top category as a list: [("General", 4)].
# [0] takes that pair out of the list; it is split into top_category and top_count.
top_category, top_count = frequency.most_common(1)[0]

# Print the most frequent category and its share of all patients.
print(f"Most frequent category: {top_category} ({top_count} of {total_frequency} patients)")
