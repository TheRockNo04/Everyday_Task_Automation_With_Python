from datetime import datetime
import os

DATE = datetime.now()

while True:
    a = int(input("Enter RollNo:- "))
    for _ in range(3):
        os.startfile(f"D:\\Piyush\\Roll\\{DATE.year}\\{DATE.month}\\{a}.xlsx", "print")