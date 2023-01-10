import pandas as pd
import openpyxl
from global_defs import *

def add_malfunction_to_excel(input_file, new_entry):
    wb = openpyxl.load_workbook(filename=input_file)
    ws = wb[MALEX_MALFUNCTIONS_SHEET]
    row = ws.max_row + 1

    for col, entry in enumerate(new_entry, start=1):
        ws.cell(row=row, column=col, value=entry)

    wb.save(input_file)
    return row


def get_prop_list_by_subsystem(props_list, subsystem):
    if not subsystem:
        return_list = [prop.name for prop in props_list]
        return_list.append(ELSE)

    return_list = []
    for prop in props_list:
        if prop.subsystem == subsystem:
            if prop.name:
                return_list.append(prop.name)

    return_list.append(ELSE)
    return return_list

def validate_entry(list, entry):
    if entry in list:
        return list
    if not entry:
        return list
    list.append(entry)
    return list

def get_all_systems_lists(input_file):
    systems_list = []
    subsystems_list = []
    malfunctions_list = []
    malfunctions_status_list = []
    lists_sheet = pd.read_excel(input_file, sheet_name=MALEX_LISTS_SHEET, keep_default_na=False, skiprows=[0])
    for index, row in lists_sheet.iterrows():
        systems_list    = validate_entry(systems_list ,str(row[COL_SYSTEM]))
        subsystems_list  = validate_entry(subsystems_list ,str(row[COL_SUBSYSTEM]))
        malfunctions_list = validate_entry(malfunctions_list ,str(row[COL_MALFUNCTION]))
        malfunctions_status_list = validate_entry(malfunctions_status_list ,str(row[COL_MALFUNCTION_STAUTS]))

    systems_list.append(ELSE)
    subsystems_list.append(ELSE)
    malfunctions_list.append(ELSE)
    malfunctions_status_list.append(ELSE)
    return systems_list, subsystems_list, malfunctions_list, malfunctions_status_list


def load_malfunctions_excel(input_file):
    props_list = []
    props_sheet = pd.read_excel(open(input_file, 'rb'), sheet_name=MALEX_PROPS_SHEET, keep_default_na=False)
    for index, row in props_sheet.iterrows():
        system = str(row[COL_SYSTEM_IDX])
        subsystem = str(row[COL_SUBSYSTEM_IDX])
        name = str(row[COL_NAME_IDX])
        description = str(row[COL_DESCRIPTION_IDX])

        prop = Prop(name=name,
                    system= system,
                    subsystem= subsystem,
                    description= description)

        props_list.append(prop)

    return props_list

