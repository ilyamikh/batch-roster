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
    filepath = "Roster_" + get_timestamps()
    if not os.path.exists(filepath):
        os.makedirs(filepath)
    print("Starting process...")
    user_month = input("Month/Year?")
    if user_month:
        month = user_month
    else:
        month = datetime.datetime.now().strftime('%B, %Y')

    create_monthly_rosters(rles, filepath, month)


def create_monthly_rosters(rles, filepath, month):
    for group in rles:  # I'll try and take care of the logic to split the sheets here
        classroom = sorted(rles[group])
        children = len(classroom)
        print(group, children, "children.")
        if children < 14:
            make_sheet(group, filepath, classroom, month)  # group is the key or name of the class
        else:
            for i in range(0, ceil(children / 13)):
                group_name = group + "_" + str(i+1)

                start = i*13
                stop = i*13 + 13
                if stop > children:
                    stop = children  # what't the one line way to do this?

                make_sheet(group_name, filepath, classroom[start:stop], month)
    print("Processed", len(rles), "groups.")


def get_timestamps():
    """Returns simply formatted timestamp string"""
    return datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d-%H-%M-%S')


def make_sheet(group, filepath, classroom, in_month):
    """
    Processes each group, creating and saving the roster
    :param group: name of the group
    :param filepath: path to save the output
    :param classroom: list of children
    :return: none
    """
    print("Filling", group + ', ', len(classroom), "children...")

    book = load_workbook("template.xlsx")
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

