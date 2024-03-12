import time
import os
import PySimpleGUI as sg
import PySimpleGUIQt as sgqt
from PIL import Image
from io import BytesIO
from sys import platform
import ctypes
from threading import Thread
import psutil
from win32com.client import Dispatch
import subprocess


sg.user_settings_filename(filename='lib/proc.json', path='.')
sg.theme('black')
sg.set_global_icon('lib/icon.ico')

if platform.startswith('win'):
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u'CompanyName.ProductName.SubProduct'
                                                                  u'.VersionInformation')


class Process(Thread):
    def __init__(self, prog=False):
        Thread.__init__(self)
        self.running = self.detected = self.first_scan = False
        self.prog = prog
        self.pid = 0

    def run(self):
        self.running = True
        while self.running:
            processes = sg.user_settings_get_entry('proc', []) if not self.prog else \
                sg.user_settings_get_entry('pprog', '')
            detections = sum(
                proc in (i.name() for i in psutil.process_iter())
                for proc in processes
            )
            self.detected = detections > 0
            self.first_scan = True
            time.sleep(0.01)

    def stop(self):
        self.running = False
        self.join()


o = Process()
op = Process(True)
auto_start = sg.user_settings_get_entry('ai', False)


def get_img_data(f, maxsize=(1200, 850), var="PNG"):
    img = Image.open(f)
    img.thumbnail(maxsize)
    bio = BytesIO()
    img.save(bio, format=var)
    return bio.getvalue()


def manage_obj(restart=True):
    global o, op
    if restart:
        o = Process()
        op = Process(prog=True)
        o.start()
        op.start()
    elif o.is_alive():
        o.stop()
        op.stop()


def action():
    global o, op
    pprog = sg.user_settings_get_entry('pprog', '')
    prog = sg.user_settings_get_entry('prog', '')
    if o.detected and op.detected:
        for p in pprog:
            subprocess.call(f"taskkill /F /IM {p}")
    elif not op.detected and not o.detected:
        op.detected = True
        os.startfile(prog)


def start_tray(tray):
    global auto_start, o
    menu_def = ['BLANK', ['&Open', '&Stop', '---', 'E&xit']] if auto_start else \
        ['BLANK', ['&Open', '&Start' if o.is_alive() else '&Stop', '---', 'E&xit']]
    tray.Update(menu=menu_def)
    manage_obj(not o.is_alive())


def icon_tray():
    global o, op, auto_start

    menu_def = ['BLANK', ['&Open', '&Start' if not o.running else '&Stop', '---', 'E&xit']]
    tray = sgqt.SystemTray(tooltip='Miner Inspector', menu=menu_def, data=get_img_data('lib/icon.ico', var="ICO"))
    opening = False

    while True:
        menu_item = tray.read(timeout=500)
        if menu_item == 'Exit':
            break
        elif menu_item in ['Open', '__ACTIVATED__', '__DOUBLE_CLICKED__']:
            opening = True
            break
        elif menu_item in ['Start', 'Stop']:
            o.running = not o.running
            start_tray(tray)

        if auto_start:
            auto_start = False
            start_tray(tray)

        if o.is_alive() and o.first_scan:
            action()

    tray.close()
    main() if opening else manage_obj(False)


def check_changes(window, val):
    prog = sg.user_settings_get_entry('prog', '')
    pprog = ','.join(sg.user_settings_get_entry('pprog', [])).replace('.exe', '')
    proc = ','.join(sg.user_settings_get_entry('proc', [])).replace('.exe', '')
    minr = sg.user_settings_get_entry('min', False)
    ai = sg.user_settings_get_entry('ai', False)
    aiw = sg.user_settings_get_entry('aiw', False)
    if prog != val['prog'] or pprog != val['pprog'] or proc != val['proc'] or minr != val['min'] or \
            ai != val['ai'] or aiw != val['aiw']:
        window['changes'].update('Cambios sin Aplicar!')
    else:
        window['changes'].update('Cambios Aplicados.')


def set_entries(proc, proc_p, val):
    sg.user_settings_set_entry('proc', proc)
    sg.user_settings_set_entry('pprog', proc_p)
    sg.user_settings_set_entry('prog', val['prog'])
    sg.user_settings_set_entry('min', val['min'])
    sg.user_settings_set_entry('ai', val['ai'])
    sg.user_settings_set_entry('aiw', val['aiw'])


def create_shortcut(shell, path, target, wdir, icon):
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wdir
    shortcut.IconLocation = icon
    shortcut.save()


def define_shortcut(val):
    shell = Dispatch("WScript.Shell")
    dir_path = os.path.join(shell.SpecialFolders("Startup"), 'MinerInspector.lnk')
    if os.path.exists(dir_path):
        os.remove(dir_path)
    if val['aiw']:
        target = os.path.join(os.getcwd(), 'MinerInspector.exe')
        create_shortcut(shell, dir_path, target, os.getcwd(), target)


def main():
    global o, op, auto_start

    processes = ','.join(sg.user_settings_get_entry('proc', [])).replace('.exe', '')
    processes_p = ','.join(sg.user_settings_get_entry('pprog', [])).replace('.exe', '')

    layout = [[sg.Text('Program:'), sg.Input(sg.user_settings_get_entry('prog', ''), key='prog'),
               sg.FileBrowse()],
              [sg.Text('Process of Program:'), sg.Input(processes_p, key='pprog')],
              [sg.Text('Processes:'), sg.Input(processes, key='proc')],
              [sg.Checkbox('Auto-Minimizar', sg.user_settings_get_entry('min', False), key='min'),
               sg.Checkbox('Auto-Iniciar', sg.user_settings_get_entry('ai', False), key='ai'),
               sg.Checkbox('Auto-Iniciar Windows', sg.user_settings_get_entry('aiw', False), key='aiw')],
              [sg.Text('Cambios Aplicados.', key='changes')],
              [sg.Button('Apply'), sg.Button('Start' if not o.running else 'Stop', key='trigger')]]

    window = sg.Window('Miner Inspector', layout, element_justification='c')

    while True:
        ev, val = window.read(timeout=500)

        if ev == sg.WIN_CLOSED:
            closing = sg.popup_yes_no('[SALIR]', 'Elegir Si cierra, elegir No minimiza en segundo plano.')
            if closing == 'No':
                icon_tray()
            break

        elif ev == 'Apply' and val['prog'] and val['proc'] and val['pprog']:
            sg.popup_timed('Listo', auto_close_duration=0.5)
            processes_p = [f'{x}.exe' for x in val['pprog'].split(',')]
            processes = [f'{x}.exe' for x in val['proc'].split(',')]
            set_entries(processes, processes_p, val)
            define_shortcut(val)

        elif ev == 'trigger':
            window['trigger'].update('Start' if o.is_alive() else 'Stop')
            manage_obj(not o.is_alive())

        elif auto_start:
            auto_start = False
            window['trigger'].click()

        if o.is_alive() and o.first_scan:
            action()

        check_changes(window, val)

    if closing == 'Yes':
        manage_obj(False)
    window.close()


if __name__ == '__main__':
    icon_tray() if sg.user_settings_get_entry('min', False) else main()
