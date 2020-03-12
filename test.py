import json
from openpyxl import load_workbook
wb = load_workbook(filename = 'employee-import-data.xlsx')
sheet_ranges = wb['Sheet1']

# note: I have used different programming styles in this exercise, but can use any if told that
# it's preferable, namely for comprehensions, lambdas with map and filter and mutable for loops
# also I have used python dictionaries but could also have done the class serialization etc,
# usually different teams prefer different programming styles and I am adaptable to what works
# for the team. However using different styles might make this file a bit less standard...
# also i do realize that some of the validations are missing but i figured that the format
# would be well defined so didn't think you want me to add them anyways = )
# P.S. I also used syntax that would work on python 2.7 but can use the latest features when
# needed, but finally ran on 3.8.1
def convert_to_cells_array(sheet_ranges):
    cells = [[cell.value for cell in column] for column in sheet_ranges]
    del cells[0]
    return cells

# also note I am not handling duplicate emails in a spreadsheet as it was not specified which
# row to take in that solution
def remove_blank_email_or_first_name(cells):
    return list(filter(lambda x: x['first_name'] and x['email'], cells))

def add_ids(cells):
    for i, x in enumerate(cells):
        x['id'] = i + 1
    return cells

def get_manager_id(email, cells):
    if email:
        record = list(filter(lambda x: x['email'].strip() == email.strip(), cells))
        return record[0]['id']
    return None

def add_manager_ids(cells):
    for x in cells:
        x['manager_id'] = get_manager_id(x['manager_email'], cells)
    return cells

def remove_manager_emails(cells):
    for x in cells:
        del x['manager_email']
    return cells

# some teams are fans of python ternary if's and some are not so if they are not preferred
# I would not use them
def get_first_name(x):
    cell_value = x[0]
    email_cell_value = x[2]
    name_from_email = email_cell_value.split('@', 1)[0].title() if email_cell_value.find('@') else ''
    return cell_value if cell_value!= None and len(cell_value.strip()) > 0 else name_from_email

def get_last_name(x):
    cell_value = x[1]
    return x[1] if x[1] != None and len(x[1].strip()) > 0 else ""

# could have gotten these directly or could have extracted them as named columns but as the
# spreadsheet was said to be defined just used indices here, though it would be easy to convert
def get_email(x):
    return x[2]

def get_manager_email(x):
    return x[3]

# Could combine this also to process without the temporary variables but some teams prefer
# one style over another and I would probably think of the names for the functions more if
# it was a real program and add more validations and use single style
list_of_names_and_emails = map(lambda x: {'first_name': get_first_name(x), 'last_name': get_last_name(x), 'email': get_email(x), 'manager_email': get_manager_email(x)} , convert_to_cells_array(sheet_ranges))

validated_list = remove_blank_email_or_first_name(list_of_names_and_emails)
list_with_ids = add_ids(validated_list)
list_with_manager_ids = add_manager_ids(list_with_ids)
list_without_manager_emails = remove_manager_emails(list_with_manager_ids)

print(json.dumps(list_without_manager_emails))