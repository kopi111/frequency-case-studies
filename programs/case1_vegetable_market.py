"""Case Study 1 - Vegetable market.

Builds a frequency table of daily basket sales with collections.Counter
and finds the sales quantity that occurs most often.
"""

# Counter is a built-in Python tool that counts how many times each item appears.
from collections import Counter

# The number of vegetable baskets sold on each of the ten days (one number per day).
baskets_sold = [10, 15, 10, 20, 15, 10, 25, 20, 10, 15]

# Counter goes through the list and counts each value.
# Result: Counter({10: 4, 15: 3, 20: 2, 25: 1})
# i.e. 10 baskets sold on 4 days, 15 on 3 days, 20 on 2 days, 25 on 1 day.
frequency = Counter(baskets_sold)

# frequency.values() keeps only the counts: 4, 3, 2, 1.
# sum() adds them: 4 + 3 + 2 + 1 = 10 (this is Σf, the total frequency).
# It equals the number of days recorded, because every day is counted once.
total_frequency = sum(frequency.values())

# Print the column headings of the frequency table.
print("Baskets sold | Frequency")

# sorted(frequency) lists the basket amounts from smallest to largest: 10, 15, 20, 25.
for quantity in sorted(frequency):
    # frequency[quantity] looks up how many days had that amount, e.g. frequency[10] = 4.
    # :>12 and :>9 right-align the numbers so the columns line up.
    print(f"{quantity:>12} | {frequency[quantity]:>9}")

# Print the total frequency worked out above.
print("Total frequency:", total_frequency)

# most_common(1) returns the single most frequent item as a list: [(10, 4)].
# [0] takes the first (and only) pair out of the list: (10, 4).
# The pair is split into two names: most_common_quantity = 10, highest_count = 4.
most_common_quantity, highest_count = frequency.most_common(1)[0]

# Print the most frequent sales quantity (the mode) and how many days it occurred.
print(f"Most frequent sales quantity: {most_common_quantity} baskets "
      f"({highest_count} days)")
