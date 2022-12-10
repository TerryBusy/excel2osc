import xlrd
import argparse
import time
import PySimpleGUI as sg
from pythonosc import udp_client

dict_object = {'Mono': 'track', 'Stereo': 'stereo', 'Speaker': 'speaker'}
list_object = []
for i in dict_object:
    list_object.append(i)

dict_coordinate = {'XYZ': ['x', 'y', 'z'], 'AED': ['azim', 'elev', 'dist']}
list_coordinate = []
for j in dict_coordinate:
    list_coordinate.append(j)

sg.theme('LightGreen')
layout = [
    [sg.Text('Convert EXCEL file to OSC for Holophonix',  font=("微软雅黑", 10), size=(40, 1), justification='left')],
    [sg.T('Choose Excel file', font=("微软雅黑", 10), size=(40, 1))],
    [sg.In(key='input'), sg.FileBrowse(key='_BUTTON_KEY_', target='input', file_types=(("excel file", "*.xls"),))],
    [sg.Text('Choose the type your want to import',  font=("微软雅黑", 10), size=(40, 1), justification='left')],
    [sg.Drop(list_object, default_value='Speaker', key='-drop1-', size=(15, 1), enable_events=True),
     sg.Drop(list_coordinate, default_value='XYZ', key='-drop2-', size=(15, 1), enable_events=True)],
    [sg.T('Holophonix IP address', font=("微软雅黑", 10), size=(18, 1)), sg.In(key='-input1-', default_text='127.0.0.1')],
    [sg.T('Holophonix OSC port', font=("微软雅黑", 10), size=(18, 1)), sg.In(key='-input2-', default_text='4003')],
    [sg.Button('Send', k='-Button1-')],
    ]

window = sg.Window('Excel2OSC', layout)

if __name__ == "__main__":
    while True:
        event, values = window.read(timeout=100)

        if event == 'Exit' or event == sg.WIN_CLOSED:
            break

        address_type = dict_object[values['-drop1-']]
        coordinate = dict_coordinate[values['-drop2-']]
        coordinate_a = coordinate[0]
        coordinate_b = coordinate[1]
        coordinate_c = coordinate[2]


        if event == '-Button1-':
            parser = argparse.ArgumentParser()
            parser.add_argument("--ip", default=(str(values['-input1-'])),
                        help="The ip of the OSC server")
            parser.add_argument("--port", type=int, default=(int(values['-input2-'])),
                        help="The port the OSC server is listening on")
            args = parser.parse_args()
            client = udp_client.SimpleUDPClient(args.ip, args.port)

            data = xlrd.open_workbook(values['input'])
            table = data.sheets()[0]
            row_num = table.nrows

            for x in range(1, row_num):
                name = table.cell_value(x, 0)
                address_x = '/' + address_type + '/' + str(x) + '/' + coordinate_a
                address_y = '/' + address_type + '/' + str(x) + '/' + coordinate_b
                address_z = '/' + address_type + '/' + str(x) + '/' + coordinate_c
                address_name = '/' + address_type + '/' + str(x) + '/' + 'name'
                x_pos = table.cell_value(x, 1)
                y_pos = table.cell_value(x, 2)
                z_pos = table.cell_value(x, 3)
                client.send_message(address_x, x_pos)
                client.send_message(address_y, y_pos)
                client.send_message(address_z, z_pos)
                client.send_message(address_name, name)
                time.sleep(0.15)

    window.close()
