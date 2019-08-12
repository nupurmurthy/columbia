import OpenSSL
import csv

csv_path = "C:\\Users\\nupur_v6\\Desktop\\budget_data.csv"
output_file = "C:\\Users\\nupur_v6\\01-Lesson-Plans\\03-Python\\3\\Python-Challenge HW\\Resources\\output file.txt"

total_months = 0
revenue_previous = 0
monthly_revenue_change = []
revenue_change = []
greatest_increase = ["", 0]
greatest_decrease = ["", 1234567890]
total_revenue = 0

with open(csv_path) as budget_data:
    reader = csv.DictReader(budget_data)

    for row in reader:
        total_months = total_months + 1
        total_revenue = total_revenue + int(row["Profit/Losses"])

        change_in_revenue = int(row["Profit/Losses"]) - revenue_previous
        revnue_previous = int(row["Profit/Losses"])
        revenue_change = revenue_change + [change_in_revenue]
        monthly_revenue_change = monthly_revenue_change + [row["Date"]]

        if (change_in_revenue > greatest_increase[1]):
            greatest_increase[0] = row["Date"]
            greatest_increase[1] = change_in_revenue

        if (change_in_revenue < greatest_decrease[1]):
            greatest_decrease[0] = row["Date"]
            greatest_decrease[1] = change_in_revenue

average_revenue = sum(revenue_change) / len(revenue_change)

output = (
    f"\nFinancial Analysis\n"
    f"----------------------------------\n"
    f"Total Months: {total_months}\n"
    f"Total: ${total_revenue}\n"
    f"Average Change: ${average_revenue}\n"
    f"Greatest Increase in Profits: {greatest_increase[0]} (${greatest_increase[1]})\n"
    f"Greatest Decrease in Profits: {greatest_decrease[0]} (${greatest_decrease[1]})\n")

print(output)

with open(output_file, "w") as txt_file:
    txt_file.write(output)