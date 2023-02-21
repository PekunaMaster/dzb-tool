import tkinter as tk
from tkinter import filedialog as fd
from tkinter import ttk
from tkinter.messagebox import showinfo
import os
from os import walk
import xlwings as xw
import subprocess
from AppKit import NSWorkspace
import time
from datetime import datetime

#test comment 1
#test comment 2

# load file names into memory 
kap_hint = "Kapitalertragsbericht"
ein_hint = "Einkommensbericht"
abs_hint = "Abschlussbericht"
dzb_hint = "Daten zum Bericht"
spen_hint = "Spenden"
geb_hint = "Gebührenbericht"
ver_hint = "verlorenen"
tra_hint = "Trade List Full"

# define inital states
start_state = 0

now = datetime.now()
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
print(dt_string)

def startDzb(kap_init, ein_init, abs_init, spen_init, geb_init, ver_init, tra_init):
    global start_state
    start_state = 1
    if (abs_init.get() == 1):
        absCpy(abs, dzb)
    if (ein_init.get() == 1):
        einCpy(ein, dzb)
    if (spen_init.get() == 1):
        spenCpy(spen, dzb)
    if (geb_init.get() == 1):
        gebCpy(geb, dzb)
    if (ver_init.get() == 1):
        verCpy(ver, dzb)
    if (tra_init.get() == 1):
        traCpy(tra, dzb)
    if (kap_init.get() == 1):
        kapCpy(kap, dzb)
    tk.messagebox.showinfo(title='DZB Tool Fertig', message='Dateien wurden in DZB eingefügt :)')
    print("finish run")
    quit()

def select_file():
    global folder_name
    folder_name = fd.askdirectory(title='Open a file', initialdir=os.path.abspath(os.getcwd()))
    if "/" in folder_name:
        if (isFile(dzb_hint) == False):
            tk.messagebox.showinfo(title='DZB im Ordner nicht gefunden', message='In dem Ordner, den Sie gewählt haben, ist DZB nicht vorhanden. Bitte fügen sie DZB zum Ordner hinzu!')
        else:
            guiStart(folder_name)

def updateCheck():
    # repo = git.Repo("https://github.com/alecmalloc/dzb-tool.git")
    # repo.remotes.upstream.pull('main')
    subprocess.call(['git', 'fetch', 'origin', 'main'])
    subprocess.call(['git', 'merge', 'origin/main', '-m', '"this is a message"'])
    subprocess.call(['git', 'push'])

def guiSelect():
    global first_window
    first_window = tk.Tk()
    first_window.title("DZB Tool")
    first_window.geometry('200x100')
    updateCheck()
    label = tk.Label(first_window, bg='light green', width=20, text='Bitte DZB ordner aussuchen:').pack(pady=10)
    open_button = ttk.Button(first_window, text='aussuchen', command=select_file).pack(pady=10)
    first_window.eval('tk::PlaceWindow . center')
    first_window.mainloop()

# GUI start page 
def guiStart(folder_name):
    # assign second window as global
    window = tk.Tk()
    window.title("DZB Tool")
    window.geometry('220x260')
    abs_init = tk.IntVar(window)
    ein_init = tk.IntVar(window)
    kap_init = tk.IntVar(window)
    spen_init = tk.IntVar(window)
    geb_init = tk.IntVar(window)
    ver_init = tk.IntVar(window)
    tra_init = tk.IntVar(window)

    filenames = next(walk(folder_name), (None, None, []))[2]  # [] if no file

    label = tk.Label(window, bg='light green', width=20, text='Bitte prüfen:').pack(pady=10)
    c1 = tk.Checkbutton(window, text='Kapitalertragsbericht', variable=kap_init, onvalue = 1, offvalue = 0).pack()
    c2 = tk.Checkbutton(window, text='Einkommensbericht', variable=ein_init, onvalue = 1, offvalue = 0).pack()
    c3 = tk.Checkbutton(window, text='Abschlussbericht', variable=abs_init, onvalue = 1, offvalue = 0).pack()
    c4 = tk.Checkbutton(window, text='Spenden / Schenkungen', variable=spen_init, onvalue = 1, offvalue = 0).pack()
    c6 = tk.Checkbutton(window, text='Gestohlen / Verloren', variable=ver_init, onvalue = 1, offvalue = 0).pack()
    c5 = tk.Checkbutton(window, text='Gebühren', variable=geb_init, onvalue = 1, offvalue = 0).pack()
    c6 = tk.Checkbutton(window, text='Transaktionsliste', variable=tra_init, onvalue = 1, offvalue = 0).pack()
    start_button = tk.Button(window, text="Start", command= lambda: startDzb(kap_init, ein_init, abs_init, spen_init, geb_init, ver_init, tra_init)).pack(pady=10)

    if isFile(kap_hint):
        global kap
        kap = folder_name + '/' + str([filename for filename in filenames if kap_hint in filename][0])
        kap_init.set(1)
    if isFile(ein_hint):
        global ein
        ein = folder_name + '/' + str([filename for filename in filenames if ein_hint in filename][0])
        ein_init.set(1)
    if isFile(abs_hint):
        global abs
        abs = folder_name + '/' + str([filename for filename in filenames if abs_hint in filename][0])
        abs_init.set(1)
    if isFile(dzb_hint):
        global dzb
        dzb = folder_name + '/' + str([filename for filename in filenames if dzb_hint in filename][0])
    if isFile(spen_hint):
        global spen
        spen = folder_name + '/' + str([filename for filename in filenames if spen_hint in filename][0])
        spen_init.set(1)
    if isFile(geb_hint):
        global geb
        geb = folder_name + '/' + str([filename for filename in filenames if geb_hint in filename][0])
        geb_init.set(1)
    if isFile(ver_hint):
        global ver
        ver = folder_name + '/' + str([filename for filename in filenames if ver_hint in filename][0])
        ver_init.set(1)
    if isFile(tra_hint):
        global tra
        tra = folder_name + '/' + str([filename for filename in filenames if tra_hint in filename][0])
        tra_init.set(1)

    window.eval('tk::PlaceWindow . center')
    window.mainloop()

    if (start_state != 1):
        quit()

def kapCpy(kap, dzb):
    dzb_xl = xw.Book(dzb)
    kap_xl = xw.Book(kap)
    dzb_sheet = dzb_xl.sheets["Kapitalertragsbericht - Verkäuf"]
    kap_sheet = kap_xl.sheets["Report"]
    kap_range = regionMax(kap_sheet, 'A')
    kap_sheet.range(f'A3:K{kap_range.row}').copy()
    dzb_sheet.range('A8').paste(paste='all')
    xw.Book(dzb).save()
    xw.Book(dzb).close()
    subprocess.call(['open', dzb])
    macro("Makro1")

def einCpy(ein, dzb):
    dzb_xl = xw.Book(dzb)
    ein_xl = xw.Book(ein)
    dzb_sheet = dzb_xl.sheets["Einkommensbericht"]
    ein_sheet = ein_xl.sheets["Report"]
    ein_range = regionMax(ein_sheet, 'B')
    word_max = whereIs(ein_sheet, "Gesamt (alle Währungen)", 'A', ein_range)
    ein_sheet.range(f'A3:G{word_max}').copy()
    dzb_sheet.range('A8').paste(paste='all')
    xw.Book(dzb).save()
    xw.Book(dzb).close()
    subprocess.call(['open', dzb])
    macro("Einkommensbericht") 

def absCpy(abs, dzb):
    dzb_xl = xw.Book(dzb)
    abs_xl = xw.Book(abs)
    dzb_sheet = dzb_xl.sheets["Nicht verkaufte Positionen"]
    abs_sheet = abs_xl.sheets["Report"]
    abs_range = regionMax(abs_sheet, 1)
    word_max = whereIs(abs_sheet, "Gesamt", 'A', abs_range)
    abs_sheet.range(f'A3:I{word_max}').copy()
    dzb_sheet.range('A8').paste(paste='all')
    absCpyGrp(abs, dzb, abs_range, word_max)
    xw.Book(dzb).save()
    xw.Book(dzb).close()
    subprocess.call(['open', dzb])
    macro("CleanAssets")
    subprocess.call(['open', dzb])
    macro("NichtVerkauftDatum")
    subprocess.call(['open', dzb])
    macro("NichtVerkaufteGruppiert")


def absCpyGrp(abs, dzb, abs_range, word_max):
    dzb_xl = xw.Book(dzb)
    abs_xl = xw.Book(abs)
    dzb_sheet = dzb_xl.sheets["Nicht verkaufte gruppiert"]
    abs_sheet = abs_xl.sheets["Report"]
    abs_sheet.range(f'A{word_max + 2}:I{rangeToNum(abs_range)}').copy()
    dzb_sheet.range('A8').paste(paste='all')

def spenCpy(spen, dzb):
    dzb_xl = xw.Book(dzb)
    spen_xl = xw.Book(spen)
    dzb_sheet = dzb_xl.sheets["Spenden_Schenkungen"]
    spen_sheet = spen_xl.sheets["Report"]
    spen_range = regionMax(spen_sheet, 1)
    word_max = whereIs(spen_sheet, "Summe gesamt:", 'G', spen_range)
    spen_sheet.range(f'A3:I{word_max}').copy()
    dzb_sheet.range('A8').paste(paste='all')
    xw.Book(dzb).save()
    xw.Book(dzb).close()
    subprocess.call(['open', dzb])
    macro("SpendenSchenkung") 

def gebCpy(geb, dzb):
    dzb_xl = xw.Book(dzb)
    geb_xl = xw.Book(geb)
    dzb_sheet = dzb_xl.sheets["Gebühren"]
    geb_sheet = geb_xl.sheets["Report"]
    geb_range = regionMax(geb_sheet, 1)
    geb_sheet.range(f'A3:H{geb_range.row}').copy()
    dzb_sheet.range('A8').paste(paste='all')

def verCpy(ver, dzb):
    dzb_xl = xw.Book(dzb)
    ver_xl = xw.Book(ver)
    dzb_sheet = dzb_xl.sheets["Gestohlen_Verloren"]
    ver_sheet = ver_xl.sheets["Report"]
    ver_range = regionMax(ver_sheet, 1)
    word_max = whereIs(ver_sheet, "Gesamt (alle Währungen)", 'A', ver_range)
    ver_sheet.range(f'A3:H{word_max - 1}').copy()
    dzb_sheet.range('A8').paste(paste='all')

def traCpy(tra, dzb):
    dzb_xl = xw.Book(dzb)
    tra_xl = xw.Book(tra)
    dzb_sheet = dzb_xl.sheets["Transaktionsliste"]
    tra_sheet = tra_xl.sheets["Sheet1"]
    tra_range = regionMax(tra_sheet, 'A')
    tra_sheet.range(f'A3:K{tra_range.row}').copy()
    dzb_sheet.range('A8').paste(paste='all')
    xw.Book(dzb).save()
    xw.Book(dzb).close()
    subprocess.call(['open', dzb])
    macro("Transaktionsliste") 

def regionMax(kap_sheet, startCell):
    cell = kap_sheet.range(f"{startCell}1:{startCell}2").current_region.end("down")
    return cell

def whereIs(sheet, cell_text, collumn, max):
    max = (str(max))[::-1]
    mk1 = max.find('>') + 1
    mk2 = max.find('$', mk1)
    max = int((max[ mk1 : mk2 ])[::-1])
    for row in range(1, max):
        if sheet.range(f'{collumn}{row}').value == cell_text:
            return (row - 1)

def rangeToNum(max):
    max = (str(max))[::-1]
    mk1 = max.find('>') + 1
    mk2 = max.find('$', mk1)
    max = int((max[ mk1 : mk2 ])[::-1])
    return (max)

def isFile(file):
    filenames = next(walk(folder_name), (None, None, []))[2]  # [] if no file
    if (next((True for filename in filenames if file in filename), False) == True):
        return (True)
    else:
        return (False)

def macro(macroname):
    dzb_xl = xw.Book(dzb)
    time.sleep(4)
    active_app_name = NSWorkspace.sharedWorkspace().frontmostApplication().localizedName()
    while active_app_name != "Microsoft Excel":
        time.sleep(2)
    wb = dzb_xl
    app = wb.app
    macro_vba = app.macro(f"'{dzb}'!{macroname}") 
    macro_vba()
    xw.Book(dzb).save()
    xw.Book(dzb).close()

# execution logic
guiSelect()