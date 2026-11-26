import csv
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

FILE_NAME = "expenses.csv"


def add_expense():
    date = input("Enter date (YYYY-MM-DD): ")
    product = input("Enter product: ")
    category = input("Enter category: ").title()
    amount = input("Enter amount: ")

    with open(FILE_NAME, "a", newline="") as file:
        writer = csv.writer(file)
        writer.writerow([date, product, category, amount])

    print("Expense added successfully!")


def view_expenses():
    with open(FILE_NAME, "r", newline="") as file:
        reader = csv.reader(file)
        for row in reader:
            print(row)


def make_report():
    totals = {}

    with open(FILE_NAME, "r", newline="") as file:
        reader = csv.DictReader(file)
        for row in reader:
            category = row["Category"]
            amount = float(row["Amount"])
            totals[category] = totals.get(category, 0) + amount

    book = Workbook()
    sheet = book.active
    sheet.title = "Expense Report"
    sheet.append(["Category", "Total Amount (INR)"])

    for category, amount in totals.items():
        sheet.append([category, amount])

    chart = BarChart()
    chart.title = "Expenses by Category"
    chart.y_axis.title = "Amount (INR)"
    chart.x_axis.title = "Category"
    chart.add_data(Reference(sheet, min_col=2, min_row=1,
                             max_row=sheet.max_row), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=2,
                                    max_row=sheet.max_row))
    sheet.add_chart(chart, "D2")

    book.save("expense_report.xlsx")
    print("Excel report created!")


while True:
    print("\nPERSONAL EXPENSE TRACKER")
    print("1. Add expense")
    print("2. View expenses")
    print("3. Create Excel report")
    print("4. Exit")

    choice = input("Choose an option: ")

    if choice == "1":
        add_expense()
    elif choice == "2":
        view_expenses()
    elif choice == "3":
        make_report()
    elif choice == "4":
        print("Goodbye!")
        break
    else:
        print("Please choose a number from 1 to 4.")
