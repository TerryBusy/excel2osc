import xlrd
import argparse
import time
import PySimpleGUI as sg
from pythonosc import udp_client

dict_object = {'Mono': 'track', 'Stereo': 'stereo', 'Speaker': 'speaker'}
list_object = []
for i in dict_object:
    list_object.append(i)
# 将导入对象加入字典，并创建列表
dict_coordinate = {'XYZ': ['x', 'y', 'z'], 'AED': ['azim', 'elev', 'dist']}
list_coordinate = []
for j in dict_coordinate:
    list_coordinate.append(j)
# 将坐标轴对象加入字典，并创建列表
sg.theme()  # 界面
layout = [
    [sg.Text('Convert EXCEL file to OSC for Holophonix', size=(40, 1), justification='left')],
    [sg.T('Choose Excel file', size=(40, 1))],
    [sg.In(key='input'), sg.FileBrowse(key='_BUTTON_KEY_', target='input', file_types=(("excel file", "*.xls"),))],
    [sg.Text('Choose the type your want to import', size=(40, 1), justification='left')],
    [sg.Drop(list_object, default_value='Speaker', key='-drop1-', size=(15, 1), enable_events=True),
     sg.Drop(list_coordinate, default_value='XYZ', key='-drop2-', size=(15, 1), enable_events=True)],
    [sg.T('Holophonix IP address', size=(18, 1)), sg.In(key='-input1-', size=(33, 1), default_text='127.0.0.1')],
    [sg.T('Holophonix OSC port', size=(18, 1)), sg.In(key='-input2-', size=(33, 1), default_text='4003')],
    [sg.Button('Send', k='-Button1-')],
]

window = sg.Window('Excel2OSC', layout)

if __name__ == "__main__":
    while True:
        event, values = window.read(timeout=100)

        if event == 'Exit' or event == sg.WIN_CLOSED:
            break

        address_type = dict_object[values['-drop1-']]  # 获取地址类型
        coordinate = dict_coordinate[values['-drop2-']]  # 获取坐标类型
        coordinate_a = coordinate[0]  # 取出坐标数据1
        coordinate_b = coordinate[1]  # 取出坐标数据2
        coordinate_c = coordinate[2]  # 取出坐标数据3

        if event == '-Button1-':
            parser = argparse.ArgumentParser()
            parser.add_argument("--ip", default=(str(values['-input1-'])),
                                help="The ip of the OSC server")
            parser.add_argument("--port", type=int, default=(int(values['-input2-'])),
                                help="The port the OSC server is listening on")
            args = parser.parse_args()
            client = udp_client.SimpleUDPClient(args.ip, args.port)
            # osc客户端ip和地址获取
            data = xlrd.open_workbook(values['input'])
            table = data.sheets()[0]
            row_num = table.nrows
            # 读取excel文件第一页各行列数据
            for x in range(1, row_num):
                name = table.cell_value(x, 0)  # 第0列数据为对象名称
                address_x = '/' + address_type + '/' + str(x) + '/' + coordinate_a
                address_y = '/' + address_type + '/' + str(x) + '/' + coordinate_b
                address_z = '/' + address_type + '/' + str(x) + '/' + coordinate_c
                address_name = '/' + address_type + '/' + str(x) + '/' + 'name'
                x_pos = table.cell_value(x, 1)  # 第1列数据为坐标数据1
                y_pos = table.cell_value(x, 2)  # 第2列数据为坐标数据2
                z_pos = table.cell_value(x, 3)  # 第3列数据为坐标数据3
                client.send_message(address_x, x_pos)
                time.sleep(0.15)
                client.send_message(address_y, y_pos)
                time.sleep(0.15)
                client.send_message(address_z, z_pos)
                time.sleep(0.15)
                client.send_message(address_name, name)
                time.sleep(0.15)
                # holophonix 接收数据有限制，增加0.15秒延迟

    window.close()
