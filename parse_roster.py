import xml.etree.ElementTree as ET


def get_roster(path="raw.xml"):
    """
    Parses the XML input at the specified location and returns a dictionary as the roster
    :param path: location of XML file
    :return: dictionary with classrooms as keys and a list of names as values
    """
    tree = ET.parse(path)
    root = tree.getroot()

    roster = {}
    children = get_child_list(root)

    name = [None, None]
    room = None

    for child in children:
        for field in child:  # The data for each child is here
            if field.attrib['Name'] == "WorkAreaName1":
                for value in field:
                    if value.tag == '{urn:crystal-reports:schemas:report-detail}Value':
                        room = value.text
            elif field.attrib['Name'] == "ChildFullName1":
                for value in field:
                    if value.tag == '{urn:crystal-reports:schemas:report-detail}Value':
                        name[0] = value.text
            elif field.attrib['Name'] == "Description1":
                for value in field:
                    if value.tag == '{urn:crystal-reports:schemas:report-detail}Value':
                        name[1] = value.text
                        roster.setdefault(room, []).append((name[0], name[1]))

    return roster


def get_child_list(root):
    """
    Takes an XML object and returns an XML object with children
    :param root: parsed XML object
    :return: XML object with a list of all children
    """
    classrooms = []
    children = []

    for group in root:
        if group.tag == "{urn:crystal-reports:schemas:report-detail}Group":
            classrooms.append(group)
            # Having a list of groups is a bit redundant if I won't be using it

    for room in classrooms:
        for child in room:
            if child.tag == '{urn:crystal-reports:schemas:report-detail}Details':
                for section in child:
                    children.append(section)
                    # So now children contains a list of all children with their classroom as one of the fields.

    return children
