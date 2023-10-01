import csv
from datetime import datetime, timedelta
import openpyxl

workbook = openpyxl.load_workbook("./Assignment_Timecard.xlsx")

sheet = workbook.active

# Assuming that the table is sorted by user id and the info of a single employee is in consecutive rows
# Also assuming that the rows that contains the data of a single employee is also sorted by time
i = 1

count1 = []  # array to count Employees  who has worked for 7 consecutive days.
count2 = []  # array to count Employees  who have less than 10 hours of time between shifts but greater than 1 hour
count3 = []  # array to count Employees  Who has worked for more than 14 hours in a single shift

current_day_streak = 1  # counting current consecutive days
max_day_streak = 1  # counts maximum consecutive days
previous_user = sheet[2][0].value
previous_date = sheet[2][3].value
flag = 0 # A flag to check if a employee is found then we don't have to check it again

# if any entry is empty so leaving that row and entering the next valid value in previous date
j = 2
while previous_date == "":
    if previous_date == "":
        previous_date = sheet[3][j + 1].value
    j += 1
# iterating each rows one by one to check the employees who has worked for 7 consecutive days
for row in sheet:
    if i == 1:
        i += 1

    current_user = row[0].value
    current_date = row[2].value

    # if any date is missing so leaving that row and moving to next row
    if current_date == "":
        continue

    # if the current user is same as previous user it means we are iterating the data of same employee
    if current_user == previous_user:
        # checking if current date is previous date + 1 day or not
        # if its true then employee worked consecutively
        # and storing these consecutive days count in current_streak variable
        # as value changes for current streak we are also updating the max streak as setting the max of current streak
        if current_date.date() == (previous_date + timedelta(days=1)).date():
            current_day_streak += 1
            max_day_streak = max(max_day_streak, current_day_streak)

        # if current date is greater them previous date + 1 or its not a consecutive date
        # then the current_streak  = 1 and counting starts again
        if current_date.date() > (previous_date + timedelta(days=1)).date():
            current_day_streak = 1
            max_day_streak = max(max_day_streak, current_day_streak)

        # as the max streak become 7 and greater than 7 we are storing that employee data in count1 array
        if max_day_streak >= 7:
            if flag == 0:
                count1.append([row[7].value, row[0].value])
                flag = 1  # changing flag value so that if we have found a user so don't store that user again and again
            print("This user has worked for more then 7 days continuously")
            print(current_user)
            print(max_day_streak)


    else:
        flag = 0  # flag changes to its initial value as we get new user detail
        current_day_streak = 1
        max_day_streak = 1

    previous_user = current_user
    previous_date = current_date

    
# checking who have less than 10 hours of time between shifts but greater than 1 hour
previous_in_time = ""
previous_out_time = ""
i = 1

previous_user = sheet[2][0]
flag = 0  # flag to check if the specified Employee is found so don't store that employee again

for row in sheet:
    if i == 1:
        i += 1

    # variable initialization
    current_user = row[0].value  # current user id
    current_in_time = row[2].value  # stores the iterating date
    current_out_time = row[3].value  # stores the iterating date

    if current_in_time == "" or current_out_time == "":
        continue

    # if the current user is same as previous user it means we are iterating the data of same employee
    if current_user == previous_user:
        if previous_out_time == "":
            previous_out_time = current_out_time
            continue
        else:
            slot_gap = current_in_time - previous_out_time

            if (slot_gap < timedelta(hours=10)) and (slot_gap > timedelta(hours=1)):
                if flag == 0:
                    count2.append([row[7].value, row[0].value])
                    flag = 1  # changing flag value so that if we have found a user so don't store that user again

                print("The slot gap for the user " + current_user + " is greater then one hour and less then 10 hour")

    else:
        previous_user = current_user
        previous_in_time = current_in_time
        previous_out_time = current_out_time
        flag = 0 # flag changes to its initial value as we get new user detail

# checking the peoples who have worked for more than 14 hours in a single shift
previous_user = sheet[2][0]
flag = 0 # A flag to check if a employee is found then we don't have to check it again
i = 1

for row in sheet:
    if i == 1:
        i += 1

    current_user = row[0].value  # current user id
    current_in_time = row[2].value  # time out
    current_out_time = row[3].value  # time in

    # if the current user is same as previous user it means we are iterating the data of same employee
    if current_user == previous_user:
        # if any of the entries is empty so just leaving it
        if current_in_time == "" or current_out_time == "":
            continue

        work_time = current_out_time - current_in_time
        if work_time > timedelta(hours=14):
            if flag == 0:
                count3.append([row[7].value, row[0].value])
                flag = 1 # changing flag value so that if we have found a user so don't store that user again and again

    else:
        flag = 0  # flag changes to its initial value as we get new user detail


if len(count1) == 0:
    print("NO employee found who has worked for 7 consecutive days")
else:
    print("The Employees who has worked for 7 consecutive days")
    for row in count1:
        print(row)

if len(count2) == 0:
    print("NO employee found who have less than 10 hours of time between shifts but greater than 1 hour")
else:
    print("The Employees who have less than 10 hours of time between shifts but greater than 1 hour")
    for row in count2:
        print(row)

if len(count3) == 0:
    print("NO employee found Who has worked for more than 14 hours in a single shift")
else:
    print("The Employees Who has worked for more than 14 hours in a single shift")
    for row in count3:
        print(row)

workbook.close()