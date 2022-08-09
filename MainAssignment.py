import sys
import openpyxl
import json
from openpyxl import Workbook
from json import JSONEncoder

# Data file location
TestDataLoc = "/Users/vasanp/PycharmProjects/pythonProject/Python-Assignments/DataSheet.xlsx"

# Create a workbook Object
data_obj = openpyxl.load_workbook(TestDataLoc)

# Get Sheet
sheet_obj = data_obj["Details"]

movieDetails = ["Movie Details", "title", "genre", "length", "cast", "director", "rating", "language", "timing",
                "shows per day", "first show", "interval time", "movie gap", "capacity"]

# Store new users in a list
Users = []

# List of static timings
timings = ["10:00-12:00", "12:30-02:30", "03:00-05:00"]


# home page screen
def home():
    print("******Welcome to BookMyShow*******")
    print("1. Register New User ")
    print("2. Login ")
    print("3. Exit ")
    entry_action = int(input("choose your action: "))
    if entry_action == 1:
        add_new_user()
    # elif entry_action == 2:
    #     login()
    # elif entry_action == 3:
    #     exit()
    else:
        print("Choose a valid option")
        home()


# template for new users
class register:
    def __init__(self, username, user_email, user_phone, user_age, user_password):
        self.username = username
        self.user_email = user_email
        self.user_phone = user_phone
        self.user_age = user_age
        self.user_password = user_password


# Registering New Users
def add_new_user():
    print("****Create new Account*****")
    user_name = input("Name:")
    user_email = input("Email:")
    user_phone = input("Phone:")
    user_age = input("Age:")
    user_password = input("Password:")

    # getting last row in excel sheet
    last_row_user = sheet_obj.max_row

    # storing username and password for new users in excel(users) sheet
    sheet_obj.cell(row=last_row_user + 1, column=1).value = user_name
    sheet_obj.cell(row=last_row_user + 1, column=2).value = user_password

    # saving the excel
    data_obj.save('DataSheet.xlsx')

    # Creating instance of register class and storing in Users list
    Users.append(register(user_name, user_email, user_phone, user_age, user_password))

    # go to homepage
    home()


