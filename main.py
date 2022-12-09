import xlrd
import argparse
import time
import PySimpleGUI as sg
from pythonosc import udp_client

dictobj = {'Mono': 'track', 'Stereo': 'stereo', 'Speaker': 'speaker'}
listobj = []
for i in dictobj:
    listobj.append(i)

dictco = {'XYZ': ['x', 'y', 'z'], 'AED': ['azim', 'elev', 'dist']}
listco = []
for j in dictco:
    listco.append(j)

sg.theme('LightGreen')
layout = [
    [sg.Text('Convert EXCEL file to OSC for Holophonix',  font=("微软雅黑", 10), size=(40, 1), justification='left')],
    [sg.T('Choose Excel file', font=("微软雅黑", 10), size=(40, 1))],
    [sg.In(key='input', default_text="d:\SPEAKER XYZ.xls"), sg.FileBrowse(key='_BUTTON_KEY_', target='input', file_types=(("excel files", "*.xls"),))],
    [sg.Text('Choose the type your want to import',  font=("微软雅黑", 10), size=(40, 1), justification='left')],
    [sg.Drop(listobj, default_value='Speaker', key='xiala1', size=(15, 1), enable_events=True),
     sg.Drop(listco, default_value='XYZ', key='xiala2', size=(15, 1), enable_events=True)],
    [sg.T('Holophonix IP address', font=("微软雅黑", 10), size=(18, 1)), sg.In(key='input1', default_text='127.0.0.1')],
    [sg.T('Holophonix OSC port', font=("微软雅黑", 10), size=(18, 1)), sg.In(key='input2', default_text='4003')],
    [sg.Button('Send', k='-Button1-')],
    ]

window = sg.Window('Excel2OSC', layout)

if __name__ == "__main__":
    while True:
        event, values = window.read(timeout=100)

        if event == 'Exit' or event == sg.WIN_CLOSED:
            break

        sendtype = dictobj[values['xiala1']]
        codin = dictco[values['xiala2']]
        codina = codin[0]
        codinb = codin[1]
        codinc = codin[2]

        data = xlrd.open_workbook(values['input'])
        table = data.sheets()[0]
        numrows = table.nrows

        if event == '-Button1-':
            parser = argparse.ArgumentParser()
            parser.add_argument("--ip", default=(str(values['input1'])),
                        help="The ip of the OSC server")
            parser.add_argument("--port", type=int, default=(int(values['input2'])),
                        help="The port the OSC server is listening on")
            args = parser.parse_args()
            client = udp_client.SimpleUDPClient(args.ip, args.port)

            for x in range(1, numrows):
                name = table.cell_value(x, 0)
                addressx = '/' + sendtype + '/' + str(x) + '/' + codina
                addressy = '/' + sendtype + '/' + str(x) + '/' + codinb
                addressz = '/' + sendtype + '/' + str(x) + '/' + codinc
                addname = '/' + sendtype + '/' + str(x) + '/' + 'name'
                xpos = table.cell_value(x, 1)
                ypos = table.cell_value(x, 2)
                zpos = table.cell_value(x, 3)
                client.send_message(addressx, xpos)
                client.send_message(addressy, ypos)
                client.send_message(addressz, zpos)
                client.send_message(addname, name)
                time.sleep(0.15)

    window.close()
