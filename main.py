import PySimpleGUI as sg
from malrep import *


current_xlsx_file = MALEX_BASENAME

props_structs_list = load_malfunctions_excel(current_xlsx_file)
systems_list, subsystems_list, malfunctions_list, malfunctions_status_list = get_all_systems_lists(current_xlsx_file)
props_list = get_prop_list_by_subsystem(props_structs_list, subsystems_list[0])

layout = [  [sg.Push(), sg.Text('דיווח תקלות', font=FONT_HEADER), sg.Push()],

            [sg.Combo(systems_list, default_value=systems_list[0], s=(60,22), enable_events=True, readonly=True, k='-SYSTEM-', font=FONT_PARAM),sg.Push(), sg.Text('מערכת:', font=FONT_PARAM)],
            [sg.Combo(subsystems_list, default_value=subsystems_list[0], s=(60,22), enable_events=True, readonly=True, k='-SUBSYSTEM-', font=FONT_PARAM),sg.Push(), sg.Text('תת מערכת:', font=FONT_PARAM)],
            [sg.Combo(props_list, default_value=props_list[0], s=(60,22), enable_events=True, readonly=True, k='-PROP-', font=FONT_PARAM),sg.Push(), sg.Text('אביזר:', font=FONT_PARAM)],
            [sg.Combo(malfunctions_status_list, default_value=malfunctions_status_list[0], s=(60,22), enable_events=True, readonly=True, k='-STATUS-', font=FONT_PARAM),sg.Push(), sg.Text('מצב התקלה:', font=FONT_PARAM)],
            [sg.Multiline(s=(PARAM_WIDTH,5), k='-DESCRIPTION-', font=FONT_PARAM),sg.Push(), sg.Text('מהות התקלה:', font=FONT_PARAM)],
            [sg.Input(s=(PARAM_WIDTH,5), k='-OPERATOR-', font=FONT_PARAM),sg.Push(), sg.Text('שם מפעיל:', font=FONT_PARAM)],
            [sg.Multiline(s=(PARAM_WIDTH,5), k='-NOTES-', font=FONT_PARAM),sg.Push(), sg.Text('הערות חופשיות:', font=FONT_PARAM)],

            [sg.Push(), sg.Input(key='-FILE-', visible=False, enable_events=True), sg.FileBrowse(button_text='בחר אקסל', font=FONT_PARAM), sg.Button('הוסף לאקסל', font=FONT_PARAM), sg.Button('יציאה', font=FONT_PARAM)],
            [sg.Push(), sg.Output(size=(CONSOLE_WIDTH,5), key='-OUTPUT-', font=FONT_LOG)],
            ]

window = sg.Window('דיווח תקלות', layout, element_justification='r')

while True:  # Event Loop
    event, values = window.read()

    if event == sg.WIN_CLOSED or event == 'יציאה':
        break
    if event == '-SUBSYSTEM-':
        props_list = get_prop_list_by_subsystem(props_structs_list, values['-SUBSYSTEM-'])
        window['-PROP-'].update(values=props_list)

    if event == '-FILE-':
        if verify_input_file(values['-FILE-']):
            current_xlsx_file = values['-FILE-']
            props_structs_list = load_malfunctions_excel(current_xlsx_file)
            systems_list, subsystems_list, malfunctions_list, malfunctions_status_list = get_all_systems_lists(current_xlsx_file)
            window['-SYSTEM-'].update(values=systems_list)
            window['-SUBSYSTEM-'].update(values=subsystems_list)
            window['-STATUS-'].update(values=malfunctions_status_list)

            props_list = get_prop_list_by_subsystem(props_structs_list, values['-SUBSYSTEM-'])
            window['-PROP-'].update(values=props_list)
            print(f"קובץ אקסל נטען {current_xlsx_file}")

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
        notes = values['-NOTES-']
        mal_num = add_malfunction_to_excel(current_xlsx_file,
        [
            date,
            time,
            system,
            subsystem,
            prop,
            description,
            status,
            operator,
            notes])

window.close()