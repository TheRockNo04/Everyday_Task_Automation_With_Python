import os
import time

while True:
    a = input("Enter RollNo:- ")
    for _ in range(3):
        os.startfile(f"D:\\Piyush\\Roll\\2022\\11\\{a}.xlsx", "print")
    