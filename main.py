import PySimpleGUI as sg
from malrep import *


# sg.theme('BluePurple')

props_structs_list = load_malfunctions_excel(MALEX_BASENAME)

systems_list, subsystems_list, malfunctions_list, malfunctions_status_list = get_all_systems_lists(MALEX_BASENAME)
props_list = get_prop_list_by_subsystem(props_structs_list, subsystems_list[0])
sys_msg = ""
layout = [  [sg.Text('דיווח תקלות'), sg.Text(size=(15,1), key='-OUTPUT-')],

            [sg.Combo(systems_list, default_value=systems_list[0], s=(60,22), enable_events=True, readonly=True, k='-SYSTEM-'), sg.Text('מערכת:'), sg.Text(size=(15,1))],
            [sg.Combo(subsystems_list, default_value=subsystems_list[0], s=(60,22), enable_events=True, readonly=True, k='-SUBSYSTEM-'), sg.Text('תת מערכת:'), sg.Text(size=(15,1))],
            [sg.Combo(props_list, default_value=props_list[0], s=(60,22), enable_events=True, readonly=True, k='-PROP-'), sg.Text('אביזר:'), sg.Text(size=(15,1))],
            [sg.Combo(malfunctions_status_list, default_value=malfunctions_status_list[0], s=(60,22), enable_events=True, readonly=True, k='-STATUS-'), sg.Text('מצב התקלה:'), sg.Text(size=(15,1))],
            [sg.Multiline(s=(PARAM_WIDTH,5), k='-DESCRIPTION-'), sg.Text('מהות התקלה:'), sg.Text(size=(15,5))],
            [sg.Input(s=(PARAM_WIDTH,5), k='-OPERATOR-'), sg.Text('שם מפעיל:'), sg.Text(size=(15,5))],
            
            [sg.Text(sys_msg, key='-SYS_MSG-'), sg.Text(size=(15,1))],
            [sg.Button('הוסף לאקסל'), sg.Button('יציאה')]]

window = sg.Window('דיווח תקלות', layout)

while True:  # Event Loop
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED or event == 'יציאה':
        break
    if event == '-SUBSYSTEM-':
        props_list = get_prop_list_by_subsystem(props_structs_list, values['-SUBSYSTEM-'])
        window['-PROP-'].update(values=props_list)
    if event == 'הוסף לאקסל':
        now = datetime.now()
        date = now.strftime("%d/%m/%Y")
        time = now.strftime("%H:%M")
        system = values['-SYSTEM-']
        subsystem = values['-SUBSYSTEM-']
        prop = values['-PROP-']
        status = values['-STATUS-']
        description = values['-DESCRIPTION-']
        operator = values['-OPERATOR-']
        mal_num = add_malfunction_to_excel(MALEX_BASENAME,
        [
            date,
            time,
            system,
            subsystem,
            prop,
            description,
            status,
            operator,
            ''])
        sys_msg = f"תקלה נרשמה {mal_num}"
        window['-SYS_MSG-'].update(sys_msg)

window.close()