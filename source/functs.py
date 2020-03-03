import datetime


def copyColumn(rows, sheet, col1, col2):
    for x in range(1, rows):
        a = col1 % x
        b = sheet[a].value
        c = col2 % x
        sheet[c] = b


# copy data from one column to another, moving the cells up one index position
def moveOneUp(rows, sheet, col1, col2):
    for x in range(1, rows):
        y = x + 1
        a = col1 % y
        b = sheet[a].value
        c = col2 % x
        sheet[c] = b


# copy data from one column to another, moving the cells up two index positions
def moveTwoUp(rows, sheet, col1, col2):
    for x in range(1, rows):
        y = x + 2
        a = col1 % y
        b = sheet[a].value
        c = col2 % x
        sheet[c] = b


# Function to retrieve and fill report date range
def report_date(sheet, col1):
    for cell in range(1, 5):
        a = col1 % cell
        b = sheet[a].value
        c = str(b)
        if 'From' in c:
            return datetime.datetime.strptime(c[-10:], "%m/%d/%Y").date() \
                   + datetime.timedelta(days=1)


def decInput():
    while True:
        user_input = input("Enter projected increase (e.g 10% = .1): ")
        try:
            return float(user_input)
        except ValueError:
            print("Not a valid entry")
