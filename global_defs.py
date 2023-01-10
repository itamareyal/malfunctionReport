from datetime import date, datetime

MALEX_BASENAME = "דיווח תקלות מהייצור.xlsx"
MALEX_MALFUNCTIONS_SHEET = "טופס"
MALEX_PROPS_SHEET = "רשימת אביזרים"
MALEX_LISTS_SHEET = "תקלות ומערכות"

COL_SUBSYSTEM = "תת מערכת"
COL_SYSTEM = "מערכת"
COL_NAME = "שם"
COL_DESCRIPTION = "תאור"
COL_MALFUNCTION = "תקלות"
COL_MALFUNCTION_STAUTS = "מצב תקלה"

COL_SUBSYSTEM_IDX = 3
COL_SYSTEM_IDX = 0
COL_NAME_IDX = 2
COL_DESCRIPTION_IDX = 1

PARAM_WIDTH = 60
CONSOLE_WIDTH = 100

ERR_ROW_NOT_APPENDED = -1
ERR_FILE_NOT_EXIST = 1
ERR_FILE_SHEET_MISSING = 2

ELSE = "אחר"

FONT_HEADER = 'Any 24'
FONT_PARAM = 'Any 18'
FONT_LOG = 'Any 14'

class Prop:
    def __init__(self, name, system, subsystem, description) -> None:
        self.name = name
        self.system = system
        self.subsystem = subsystem
        self.description = description
