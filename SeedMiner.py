import keyboard
import mouse
import time
from tkinter import *
import tkinter as tk
from os.path import expanduser
import os.path
import glob
import os
import json
from PIL import ImageGrab
from colorthief import ColorThief
from win32com.client import Dispatch
import win32com.client
import win32gui
import win32process
import tempfile
from win32gui import GetWindowText, GetForegroundWindow
from global_hotkeys import *
import pygetwindow as gw
import pyautogui
import mss
import mss.tools
from PIL import Image
from global_hotkeys.keycodes import vk_key_names
import win32con
import tkinter.font as tkFont
import d3dshot
from ahk import AHK
from ahk.window import Window
import playsound
import random
import string
import win32api
import win32con
import requests


ahk = AHK(executable_path='.\AutoHotkey\AutoHotkey.exe')

speak = Dispatch("SAPI.SpVoice")

Enabled = True
version = "3.3"

root = tk.Tk()
windowed = tk.IntVar()
easyDiff = tk.IntVar()
resetHotkey = tk.StringVar()
borderHotkey = tk.StringVar()
savesPath = StringVar()
currentWorldName = ''
lastCheckedWorld = ''
waitingForQuit = False
speechText = tk.StringVar()
fsMode = tk.IntVar()
fps = tk.IntVar()
multiInstance = IntVar()
fps.set(30)
cachedWindow = 0
jojoeScenes = False

mcPid = 0
mcActualPid = 0

sid = len(gw.getWindowsWithTitle('SeedMiner v'))
globalCapture = d3dshot.create(capture_output="pil")
if sid > 0:
    Label(root, text="SeedMiner ID: " + str(sid), fg="green").grid(row=16, sticky=E)

def setDefaults():
    volSlider.set(50)
    windowed.set(0)
    fsMode.set(0)
    resetHotkey.set('end')
    savesPath.set(expanduser("~") + "\AppData\Roaming\.minecraft\saves")
    speechText.set('Seed')
    borderHotkey.set('delete')
    multiInstance.set(0)
    fps.set(30)
    easyDiff.set(1)
    
def enumHandler(mcWin, ctx):
    title = win32gui.GetWindowText(mcWin)
    if title.startswith('Minecraft') and (title[-1].isdigit() or title.endswith('Singleplayer') or title.endswith('Multiplayer (LAN)')):
        style = win32gui.GetWindowLong(mcWin, -16)
        if style == 369623040:
            if fsMode.get() == 2:
                rect = win32gui.GetWindowRect(mcWin)
                if rect[3] == 1080:
                    win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 0, 1920, 1027, 0x0004)
                    return False
                else:
                    win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 0, 1920, 1080, 0x0004)
                    return False
            style = 382664704
            win32gui.SetWindowLong(mcWin, win32con.GWL_STYLE, style)
            if fsMode.get() == 3:
                win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 650, 0, 700, 1050, 0x0004)
            elif fsMode.get() == 1:
                win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 320, 1920, 400, 0x0004)
            else:
                win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 530, 250, 900, 550, 0x0004)
        else:
            style &= ~(0x00800000 | 0x00400000 | 0x00040000 | 0x00020000 | 0x00010000 | 0x00800000)
            win32gui.SetWindowLong(mcWin, win32con.GWL_STYLE, style)
            win32gui.SetWindowPos(mcWin, win32con.HWND_TOP, 0, 0, 1920, 1080, 0x0004)
        return False
    
def loadConfig():
    global forestIs, beach, desert, plains, tundra, speechText, savesPath, volSlider, Jojoe, sid
    if sid == 0:
        fileConfig = expanduser("~") + "/.mcResetSettings.json"
    else:
        fileConfig = expanduser("~") + "/.mcResetSettings" + str(sid) + ".json"
    if os.path.isfile(fileConfig):
        try:
            readFile = open(fileConfig, 'r')
            contents = readFile.read()
            readFile.close()
            settings = json.loads(contents)
        except:
            setDefaults()
            saveConfig()
            readFile = open(fileConfig, 'r')
            contents = readFile.read()
            readFile.close()
            settings = json.loads(contents)
        if "savesPath" in settings:
            savesPath.set(settings["savesPath"])
        if "speechText" in settings:
            speechText.set(settings["speechText"])
        if "volume" in settings:
            volSlider.set(settings["volume"])
        else:
            volSlider.set(50)
        if "windowed" in settings:
            windowed.set(settings['windowed'])
        if "resetHotkey" in settings:
            resetHotkey.set(settings['resetHotkey'])
        if "borderHotkey" in settings:
            borderHotkey.set(settings['borderHotkey'])
        if "fsMode" in settings:
            fsMode.set(settings["fsMode"])
        if "fps" in settings:
            fps.set(settings['fps'])
        if "multiInstance" in settings:
            multiInstance.set(settings['multiInstance'])
        if "jojoeScenes" in settings:
            jojoeScenes = True
        if "easyDiff" in settings:
            easyDiff.set(settings['easyDiff'])
    else:
        setDefaults()
        saveConfig()

def saveConfig():
    global forestIs, beach, desert, plains, tundra, speechText, savesPath, sid
    if sid == 0:
        fileConfig = expanduser("~") + "/.mcResetSettings.json"
    else:
        fileConfig = expanduser("~") + "/.mcResetSettings" + str(sid) + ".json"
    writeFile = open(fileConfig, 'w')
    
    settings = {
    'savesPath':savesPath.get(),
    'speechText':speechText.get(),
    'volume':volSlider.get(),
    'windowed':windowed.get(),
    'resetHotkey':resetHotkey.get(),
    'fps':fps.get(),
    'borderHotkey':borderHotkey.get(),
    'fsMode':fsMode.get(),
    'multiInstance':multiInstance.get(),
    'easyDiff':easyDiff.get()
    }
    
    writeFile.write(json.dumps(settings))
    writeFile.close()
    root.after(1000, saveConfig)
    
def getMcWin():
    #global mcPid
    return cachedWindow
    #return ahk.find_window(id=mcPid)
    
def getMostRecentFile(dir):
    try:
        fileList = glob.glob(dir.replace('\\', "/"))
        if not fileList:
            return False
        latest = max(fileList, key=os.path.getctime)
        return latest
    except:
        return False;

def resetRun():
    global waitingForQuit
    waitingForQuit = True
    win = getMcWin()
    time.sleep(0.01)
    print('reset')
    win.send('{escape}')
    time.sleep(0.01)
    if windowed.get():
        win.send('{shift Down}{tab}{shift Up}')
    else:
        ahk.send_input("{shift Down}{tab}{shift Up}")
    time.sleep(0.03)
    win.send('{enter}')
    
def selectMC():
    global mcPid, mcActualPid, cachedWindow
    mcPid = 0
    mcLabel.config(text='Click on Minecraft',fg='red')
    root.update()
    for i in range(20):
        win = ahk.active_window
        try:
            title = win.title.decode(sys.stdout.encoding)
        except:
            title = "aaaaaaaa"
        if title.startswith('Minecraft'):
            if title[-1].isdigit() or title.endswith('Singleplayer') or title.endswith('Multiplayer (LAN)'):
                mcPid = win.id
                mcActualPid = win.pid
                cachedWindow = win
                mcLabel.config(text='Found Minecraft (' + str(win.pid) + ')',fg='green')
                print('Found MC')
                return
        time.sleep(0.2)
    mcLabel.config(text='Could not find Minecraft',fg='red')
            
fontStyle = tkFont.Font(family="TkDefaultFont", size=9)
volSlider = Scale(root, from_=0, to=100, orient=HORIZONTAL, font=fontStyle)
volSlider.grid(row=5, sticky=W, padx=80)
volLabel = Label(root, text="TTS Volume:", font=fontStyle).grid(row=5, sticky=W, pady=1)
loadConfig()
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)

#Checkbutton(root, text="Multi-instance mode", variable=multiInstance, font=fontStyle).grid(row=14, padx=0, sticky=E)
Checkbutton(root, text="Unfocused mode", variable=windowed, font=fontStyle).grid(row=15, padx=0, sticky=E)
Radiobutton(root, text="Window", padx=7, variable=fsMode, value=0, font=fontStyle).grid(row=13, padx=150, sticky=W)
Radiobutton(root, text="Abuse Planar", padx=7, variable=fsMode, value=1, font=fontStyle).grid(row=16, padx=150, sticky=W)
Radiobutton(root, text="Perfect Travel", padx=7, variable=fsMode, value=2, font=fontStyle).grid(row=15, padx=150, sticky=W)
Radiobutton(root, text="Microlensing", padx=7, variable=fsMode, value=3, font=fontStyle).grid(row=14, padx=150, sticky=W)

Checkbutton(root, text="Easy difficulty", variable=easyDiff, font=fontStyle).grid(row=14, sticky=E)

statusLabel = Label(root, text="Running", fg='green', font=fontStyle)
statusLabel.grid(row=1, padx=0, sticky=E)
latestSeedLabel = Label(root, text="", font=fontStyle, fg='blue')
latestSeedLabel.grid(row=2, padx=0, sticky=E)
savesLabel = Label(root, text="Saves folder:", font=fontStyle)
savesLabel.grid(row=1, sticky=W)
savesPathEntry = Entry(root, width=32, exportselection=0, textvariable=savesPath, font=fontStyle).grid(row=2, padx=(5, 0), sticky=W)
speechLabel = Label(root, text="Text to speak:", font=fontStyle).grid(row=3, sticky=W)
speechTextEntry = Entry(root, width=30, exportselection=0, textvariable=speechText, font=fontStyle).grid(row=4, padx=(5, 0), sticky=W)
hotkeyLabel = Label(root, text="Reset World Hotkey:", font=fontStyle)
hotkeyLabel.grid(row=13, sticky=W)
hotkeyEntry = Entry(root, width=20, exportselection=0, textvariable=resetHotkey, font=fontStyle).grid(row=14, padx=5, sticky=W)
warningLabel = Label(root, text="", font=fontStyle)
warningLabel.grid(row=21, sticky=W)
Label(root, text="FPS you record at:", font=fontStyle).grid(row=3, sticky=E)
Radiobutton(root, text="30", padx=7, variable=fps, value=30, font=fontStyle).grid(row=4, sticky=E)
Radiobutton(root, text="60", padx=7, variable=fps, value=60, font=fontStyle).grid(row=5, sticky=E)

borderHotkeyLabel = Label(root, text="Toggle Borderless Hotkey:", font=fontStyle)
borderHotkeyLabel.grid(row=15, sticky=W)
borderHotkeyEntry = Entry(root, width=20, exportselection=0, textvariable=borderHotkey, font=fontStyle).grid(row=16, padx=5, sticky=W)
restartWarning = Label(root, text="Restart SeedMiner after changing hotkeys", font=fontStyle)
restartWarning.grid(row=17, sticky=W)

Button(root, text="Assign MC Window", command=selectMC).grid(row=17, sticky=E, padx=5)
mcLabel = Label(root, text="No Minecraft window found", fg="red", font=fontStyle)
mcLabel.grid(row=16, sticky=E)

root.resizable(False, False)
root.title("SeedMiner v" + version)

def scanForMc():
    global mcPid, mcActualPid, cachedWindow
    window = ahk.find_window(title=b'Minecraft* 1.16.1')
    if not window:
        window = ahk.find_window(title=b'Minecraft* 1.16.1 - Singleplayer')
        if not window:
            window = ahk.find_window(title=b'Minecraft* 1.16.1 - Multiplayer (LAN)')
            if not window:
                print('Can''t find MC F')
                return
    mcPid = window.id
    mcActualPid = int(window.pid)
    cachedWindow = window
    mcLabel.config(text='Found Minecraft (' + str(window.pid) + ')',fg='green')

def canCheck():
    global currentWorldName, waitingForQuit, lastCheckedWorld
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    # adv folder does not exist during early world creation
    if currentWorld == False:
        savesLabel.config(text="Saves folder: (currently invalid)", fg='red')
        return False
    if savesLabel['text'] == "Saves folder: (currently invalid)":
        savesLabel.config(text="Saves folder:", fg='black')
    if not (os.path.isdir(currentWorld + "/advancements")):
        return False
    if waitingForQuit == True:
        try:
            lockFile = open(currentWorld + "/session.lock", "r")
            # fails with error 13 (permission denied) if world is running
            lockFile.read()
            waitingForQuit = False
            makeWorld()
        except IOError as e:
            return False
    advCreation = os.stat(currentWorld + "/advancements").st_ctime
    timeElapsed = time.time() - advCreation;
    return timeElapsed < 5 and lastCheckedWorld != os.path.basename(currentWorld) and waitingForQuit == False

def switchToScene(scene):
    if not jojoeScenes:
        return
    print("Switch to scene " + scene)
    headers = {
        'Content-type': 'application/json',
    }
    data = '{"scene-name":"' + scene + '"}'
    requests.post('http://127.0.0.1:4445/emit/SetCurrentScene', headers=headers, data=data)

def reportSeed():
    global speechText, mcPid
    win32api.keybd_event(0x83, 0)
    win32api.Sleep(50)
    win32api.keybd_event(0x83, 0, win32con.KEYEVENTF_KEYUP)
    switchToScene('Stream')
    #title = ahk.active_window.title
    #win = ahk.find_window(id=mcPid)
    #win.send('{escape}')
    #time.sleep(0.05)
    #if not (title.startswith(b'Minecraft') and (str(title[-1]).isdigit() or title.endswith(b'Singleplayer') or title.endswith(b'Multiplayer (LAN)'))):
    #    win.activate()

    if random.random() > 0.9999 and os.path.isfile('CFIUUS_FFUISM_FFFHJDSJS.mp3'): # not a virus, just a special sound ;)
        playsound.playsound('CFIUUS_FFUISM_FFFHJDSJS.mp3')
        return
    if speechText.get() == '{mp3}' and os.path.isfile('seed.mp3'):
        playsound.playsound('seed.mp3')
    else:
        speak.Volume = volSlider.get()
        speak.Speak(speechText.get())

def checkBiome():
    global currentWorldName, waitingForQuit, lastCheckedWorld, currentWorldLabel
    print('Checking biome')
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    if currentWorld == False:
        print('Can''t check biome, no world')
        return
    lastCheckedWorld = os.path.basename(currentWorld)
    print(lastCheckedWorld)
    reportSeed()

def waitForColours():
    d = d3dshot.create(capture_output="pil")
    win = getMcWin()
    rect = win.rect
    if rect != (0, 0, 1920, 1080):
        rect = (rect[0] + 8, rect[1] + 24, rect[2] - 16, rect[3] - 32)
    rect = (rect[0] + 50, rect[1] + 200, rect[0] + 52, rect[1] + 202)
    root.update()
    img = d.screenshot(region=rect)
    color = img.getpixel((1, 1))
    startTime = time.time()
    while (time.time() - startTime) < 3 and not (color[0] > 13 and color[0 < 18] and color[1] > 9 and color[1] < 15 and color[2] > 5 and color[2] < 10):
        img = d.screenshot(region=rect)
        print("hmm")
        color = img.getpixel((1, 1))
        time.sleep(0.02)

def waitForWorlds():
    waitForColours()
    time.sleep(0.03)

def makeWorld():
    delay = 0.07
    if fps.get() == 60:
        delay /= 2
    elif fps.get() == 120:
        delay /= 4
    win = getMcWin()
    time.sleep(0.1)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    waitForWorlds()
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{tab}')
    time.sleep(delay)
    win.send('{enter}')
    time.sleep(delay)
    if os.path.isfile("attempts.txt"):
        countFile = open("attempts.txt", 'r+')
        counter = int(countFile.read())
        countFile.seek(0)
        countFile.write(str(counter + 1))
        countFile.close()
    time.sleep(delay)
    if easyDiff.get():
        win.send('{tab}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
        win.send('{enter}')
        time.sleep(delay)
        win.send('{enter}')
        time.sleep(delay)
        win.send('{enter}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
        win.send('{tab}')
        time.sleep(delay)
    if fps.get() == 120:
        time.sleep(0.03)
    win.send('{enter}')
    time.sleep(delay)
    time.sleep(1)
    win32api.keybd_event(0x82, 0)
    win32api.Sleep(50)
    win32api.keybd_event(0x82, 0, win32con.KEYEVENTF_KEYUP)
    switchToScene('Loading Screen')

def hotkeyReset():
    if multiInstance.get():
        print('toggle focus')
        if not str(ahk.active_window.pid) == str(mcActualPid):
            print(getMcWin().pid)
            getMcWin().activate()
            return
    win = getMcWin()
    
    currentWorld = getMostRecentFile(savesPath.get() + "/*")
    if currentWorld == False:
        print('No world found')
        return
    win32api.keybd_event(0x84, 0)
    win32api.Sleep(10)
    win32api.keybd_event(0x84, 0, win32con.KEYEVENTF_KEYUP)
    if windowed.get() and win.active:
        print('unfocus')
        root.focus_force()
        time.sleep(0.01)
    try:
        lockFile = open(currentWorld + "/session.lock", "r")
        # fails with error 13 (permission denied) if world is running
        lockFile.read()
        makeWorld()
    except IOError as e:
        if e.args[0] == 13:
            resetRun()
        else:
            print('Some generic error happened when checking world hotkey reset idk ' + e.args[1])
    
def toggleBorder():
    global mcActualPid
    if multiInstance.get():
        if not str(ahk.active_window.pid) == str(mcActualPid):
            return
        printable = set(string.printable)
        script = open("bd.txt", "r").read().replace("$PID_HERE$", str(mcActualPid))
        if fsMode.get() == 3:
            script = script.replace("$RES$", "650, 0, 700, 1050")
        elif fsMode.get() == 2:
            script = script.replace("$RES$", "0, 0, 1940, 1070")
        elif fsMode.get() == 1:
            script = script.replace("$RES$", "0, 320, 1920, 400")
        else:
            script = script.replace("$RES$", "530, 250, 900, 550")
        script = ''.join(filter(lambda x: x in printable, script))
        ahk.run_script(script, blocking=False)
    else:
        try:
            win32gui.EnumWindows(enumHandler, mcActualPid)
        except:
            return

if resetHotkey.get() != '':
    bindings = [
        [[resetHotkey.get()], None, hotkeyReset],
        [[borderHotkey.get()], None, toggleBorder]
    ]
    try:
        register_hotkeys(bindings)
    except:
        print('Error, invalid hotkey')
    start_checking_hotkeys()
    
def checkHotkeys():
    if not resetHotkey.get().lower() in vk_key_names:
        hotkeyLabel.config(text='Reset Hotkey (INVALID):', fg='red')
        restartWarning.config(fg='red')
    else:
        hotkeyLabel.config(text='Reset Hotkey:', fg='black')
    if not borderHotkey.get().lower() in vk_key_names:
        borderHotkeyLabel.config(text='Toggle Borderless Hotkey (INVALID):', fg='red')
        restartWarning.config(fg='red')
    else:
        borderHotkeyLabel.config(text='Toggle Borderless Hotkey:', fg='black')
    root.after(500, checkHotkeys)
    
scanForMc()
    
def mainLoop():
    if Enabled and canCheck() and mcPid != 0:
        checkBiome()
    root.after(50, mainLoop)
root.after(0, mainLoop)
root.after(0, checkHotkeys)
root.after(1000, saveConfig)
root.geometry('410x260')
root.mainloop()
stop_checking_hotkeys()
