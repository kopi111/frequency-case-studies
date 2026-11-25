"""Case Study 5 - Weather monitoring.

Builds a frequency table of daily temperatures with collections.Counter
and finds the temperature recorded most often.
"""

# Counter is a built-in Python tool that counts how many times each item appears.
from collections import Counter

# The temperature (degrees Celsius) recorded on each of the ten days.
temperatures = [28, 30, 28, 32, 30, 28, 34, 30, 28, 32]

# Counter counts each temperature.
# Result: Counter({28: 4, 30: 3, 32: 2, 34: 1})
# i.e. 28 C on 4 days, 30 C on 3 days, 32 C on 2 days, 34 C on 1 day.
frequency = Counter(temperatures)

# frequency.values() keeps only the counts: 4, 3, 2, 1.
# sum() adds them: 4 + 3 + 2 + 1 = 10 (Σf), which equals the number of days recorded.
total_frequency = sum(frequency.values())

# Print the column headings of the frequency table.
print("Temperature (C) | Frequency")

# sorted(frequency) lists the temperatures from lowest to highest: 28, 30, 32, 34.
for temperature in sorted(frequency):
    # frequency[temperature] looks up how many days had that reading, e.g. frequency[28] = 4.
    # :>15 and :>9 right-align the numbers so the columns line up.
    print(f"{temperature:>15} | {frequency[temperature]:>9}")

# Print the total frequency worked out above.
print("Total frequency:", total_frequency)

# most_common(1) returns the single most frequent reading as a list: [(28, 4)].
# [0] takes that pair out; it is split into most_common_temperature = 28, highest_count = 4.
most_common_temperature, highest_count = frequency.most_common(1)[0]

# Print the most frequently recorded temperature (the mode) and how many days it occurred.
print(f"Most frequently recorded temperature: {most_common_temperature} C "
      f"({highest_count} days)")
