"""Case Study 4 - Online shopping.

Builds a frequency table of product categories with pandas and adds the
relative frequency of each category (frequency / total frequency).
"""

# pandas is a data analysis library; "pd" is the short name it is usually given.
import pandas as pd

# The product category bought by each of the ten customers, as a pandas Series.
purchases = pd.Series(["Clothing", "Books", "Clothing", "Electronics", "Books",
                       "Clothing", "Food", "Clothing", "Books", "Clothing"])

# value_counts() counts each category, most common first:
# Clothing = 5, Books = 3, Electronics = 1, Food = 1.
frequency = purchases.value_counts()

# normalize=True divides every count by the total (10) instead of returning the count.
# This is the relative frequency: Clothing 5 ÷ 10 = 0.5, Books 0.3, Electronics 0.1, Food 0.1.
relative_frequency = purchases.value_counts(normalize=True)

# Put the counts and the relative frequencies side by side in one table (a DataFrame).
table = pd.DataFrame({
    "Frequency": frequency,
    "Relative frequency": relative_frequency,
})

# Give the row labels column a heading so the table reads "Category".
table.index.name = "Category"

# Print the frequency table.
print(table)

# frequency.sum() adds the counts: 5 + 3 + 1 + 1 = 10 (the total frequency, Σf).
print("Total frequency:", frequency.sum())

# The relative frequencies must add up to 1: 0.5 + 0.3 + 0.1 + 0.1 = 1.0.
# round(..., 2) removes tiny decimal rounding errors before printing.
print("Sum of relative frequencies:", round(relative_frequency.sum(), 2))

# idxmax() returns the category with the largest count (Clothing).
print("Most frequent category:", frequency.idxmax())
