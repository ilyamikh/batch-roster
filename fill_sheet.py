from openpyxl import load_workbook
import parse_roster
import os
import time
import datetime
from math import ceil


def process_list(file='raw.xml'):
    """
    Takes the path of the raw XML file and process the roster to generate classroom lists
    :param file: path to raw XML data
    :return: none
    """
    rles = parse_roster.get_roster(file)

    print("Starting process...")
    user_month = input("Month/Year? (Leave blank to use current date)")
    if user_month:
        month = user_month
    else:
        month = datetime.datetime.now().strftime('%B, %Y')

    output_choice = None
    while output_choice != 'q':
        output_choice = input("Enter 'm' for Meal Count or 'r' for Roster. 'q' to quit.")

        if output_choice == 'r':
            filepath = "Roster_" + get_timestamps()
            if not os.path.exists(filepath):
                os.makedirs(filepath)
            create_monthly_rosters(rles, filepath, month)
        elif output_choice == 'm':
            filepath = "Meal_Count_" + get_timestamps()
            if not os.path.exists(filepath):
                os.makedirs(filepath)
            create_meal_rosters(rles, filepath, month)


def create_meal_rosters(rles, filepath, month):
    for group in rles:
        classroom = rles[group]  # left sorted by category? Or will it be left unsorted?
        children = len(classroom)
        print(group, children, "children.")
        make_meal_sheet(group, filepath, classroom, month)

    print("Processed", len(rles), "groups.")


def create_monthly_rosters(rles, filepath, month):
    for group in rles:  # I'll try and take care of the logic to split the sheets here
        classroom = sorted(rles[group])  # sort by name takes place here
        children = len(classroom)
        print(group, children, "children.")
        if children < 14:
            make_roster_sheet(group, filepath, classroom, month)  # group is the key or name of the class
        else:
            for i in range(0, ceil(children / 13)):
                group_name = group + "_" + str(i+1)

                start = i*13
                stop = i*13 + 13
                if stop > children:
                    stop = children  # what't the one line way to do this?

                make_roster_sheet(group_name, filepath, classroom[start:stop], month)
    print("Processed", len(rles), "groups.")


def get_timestamps():
    """Returns simply formatted timestamp string"""
    return datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d-%H-%M-%S')


def make_meal_sheet(group, filepath, classroom, start_date):
    """Processes each group, creating and savin the meal count sheet"""
    print("Filling", group + ', ', len(classroom), "children...")

    book = load_workbook("meal_count_template.xlsx")
    name = filepath + "/" + group.replace('/', '') + ".xlsx"
    # date variables go here
    sheet = book.active

    # code to fill dates and year/month cell goes here

    current_cell = 10  # see template
    current_cat = 1
    # to add a blank line. What way does the input roster need to be sorted?
    # to sort the roster one way and meal counts another it might need two separate source XML files

    classroom.sort(key=lambda x: x[1])
    for child in classroom:
        if current_cat != child[1]:
            current_cell += 1  # skip a line to separate the categories visually
            current_cat = child[1]
        cellname = 'B' + str(current_cell)
        sheet[cellname] = child[0]  # name
        cellname = 'C' + str(current_cell)
        sheet[cellname] = child[1]  # category
        current_cell += 1

    book.save(name)
    book.close()
    print("Done")


def make_roster_sheet(group, filepath, classroom, in_month):
    """
    Processes each group, creating and saving the roster
    :param group: name of the group
    :param filepath: path to save the output
    :param classroom: list of children
    :return: none
    """
    print("Filling", group + ', ', len(classroom), "children...")

    book = load_workbook("roster_template.xlsx")
    name = filepath + "/" + group.replace('/', '') + ".xlsx"
    month = in_month
    sheet = book.active

    sheet['B1'] = group
    sheet['Z1'] = month

    current_cell = 5  # see template
    for child in classroom:
        cellname = 'A' + str(current_cell)
        sheet[cellname] = child[0]
        current_cell += 2

    book.save(name)
    book.close()
    print("Done")

