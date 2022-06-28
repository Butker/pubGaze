import tkinter as tk
import tkinter.messagebox
import time
import msvcrt as m
import sys
import pandas as pd
import logging
from tkinter import filedialog as fd
import openpyxl
import tkinter.simpledialog
import numpy
ws = tk.Tk()
ws.title("Midten")

folder = fd.askdirectory()
print(folder)

x = tkinter.simpledialog.askstring("Filnavn", 'Indtast navn på log fil:')
print(x)
run = 0
ws.destroy()

logging.basicConfig(filename=folder + '/' + x + '.log', encoding='utf-8', level=logging.INFO, format='%(asctime)s.%(msecs)03d~%(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logging.info(str(run) + '~' + 'SCRIPT START')



#Capture data from excel file, and prepare variabels for later use
logging.info(str(run) + '~' + 'Reading dataFeed.xlsx...')
df = pd.read_excel("dataFeed.xlsx")
runList = df.RunNumber.unique()
maxRun = max(runList)
newMax = maxRun.item()
runObjects = {}
finalList = []
tempList = []

activeButton = False
StartUpCanvas = ["Venstre", "Højre"]
listOfCanvas = []
objectsCanvas = []

numberOfRuns = []
ListOfButtons = []

#Height and width of the windows
HEIGHT = 700
WIDTH = 500

#Loop through all the runs, and prepare a dictionary "runObjects", that contain a key value pair, consisting of the run and objects to be created
logging.info(str(run) + '~' + 'Preparing runObjects dictionary...')
for x in runList:
    tempDf = df.loc[df['RunNumber'] == x]
    tempList = []
    templist1 = []
    for row in tempDf.index:
        realTemp = []
        realTemp.extend((df['type'][row],df['Screen'][row], df['objectColor'][row],df['x0'][row],df['y0'][row],df['x1'][row],df['y1'][row],df['text'][row],df['font'][row],df['height'][row],df['width'][row],df['command'][row]))
        tempList.append(realTemp)
    runObjects[x] = tempList
logging.info(str(run) + '~' + 'Pandas operations completed! runObjects has been created')
#Definitions of functions that can be called by buttons etc.
logging.info(str(run) + '~' + 'initializing Tkinter Functions...')
def popMessage():
    tk.messagebox.showinfo("Welcome to GFG.",  "Hi I'm your message")

#Functions related to Tkinter and the 3 windows
def close(event):
    sys.exit() # call this if you want to exit the entire thing

#Not currently in use
def wait():
    m.getch()

#Not currently in use, but another way of creating new windows.
def New_Window():
    Window = tk.Toplevel()
    Window.title("left")
    canvas = tk.Canvas(Window, height=HEIGHT, width=WIDTH)
    canvas.configure(bg='white')
    listOfCanvas.append(canvas)
    canvas.pack()

#Called when the mouse cursor is entering the main screen
def enter(event):
    logging.info(str(run) + '~' + 'GAZE: Entering the main screen')
    for x in listOfCanvas[1:]:
        x.configure(bg='white')
        x.delete('overlay')
    logging.info(str(run) + '~' + 'Objects made visible')

#Called when hte mouse cursor is leaving the main screen
def leave(event):
    logging.info(str(run) + '~' + 'GAZE: Leaving the main screen')
    for x in listOfCanvas[1:]:
        x.configure(bg='black')
        temprec = x.create_rectangle(0, 0, 800, 800, fill='white', tags='overlay')
    logging.info(str(run) + '~' + 'Objects hidden')

#Called on startup
def startUp():
    for x in StartUpCanvas:
        Window = tk.Toplevel()
        Window.title(x)
        canvas = tk.Canvas(Window, height=HEIGHT, width=WIDTH)
        canvas.configure(bg='white')
        listOfCanvas.append(canvas)
        canvas.pack()
    #If anything needs to be present on startup, it can be added here. create rectangle as example
    #listOfCanvas[0].create_rectangle(0, 0, 400, 400, fill='green')
    listOfCanvas[0].create_text(200, 500, text="For at begynde og skifte side, tryk a", fill="black", font=('Helvetica 15 bold'))
    logging.info(str(run) + '~' + '*****Startup complete and ready for testing!*****')

#Function is called when the "A" key is pressed.
#Creates new objects such as buttons, text or rectangles, and removes all the old elements.
def createObjects(xd):
    global run
    logging.info(str(run) + '~' + 'User pressed the "A" key. Moving from run ' + str(run) + " to run " + str(run+1))
    #Preparing global variables, to be used in the following loops
    global activeButton
    run = run+1
    print(run)
    print(newMax)
    for x in listOfCanvas:
        x.delete('all')
        if activeButton:
            ListOfButtons[0].place_forget()
    if run > newMax:
        logging.info(str(run) + '~' + 'Last run is over, exiting')
        close("xd")
        print("HALLO")
    #For every object, that needs to be created. Check type of object, and create the appropriate
    for x in runObjects.get(run):
        print(x)
        if x[0] == "rectangle":
            listOfCanvas[int(x[1])].create_rectangle(x[3], x[4], x[5], x[6], fill=x[2], tags=run)
        elif x[0] == "text":
            listOfCanvas[int(x[1])].create_text(x[3], x[4], text=x[7], fill=x[2], font=(x[8]))
        elif x[0] == "button":
            activeButton = True
            btn = tk.Button(ws, text=x[7], width=int(x[10]),height=int(x[9]), bd='10', command=eval(x[11]))
            btn.place(x=int(x[3]), y=int(x[4]))
            ListOfButtons.append(btn)
            print(ListOfButtons)
    logging.info(str(run) + '~' + 'Next run is now being displayed!')
logging.info(str(run) + '~' + 'Initialization complete!')


#Startup procedures. A main window has to be created like tk.TK()
ws = tk.Tk()
ws.title("Midten")
canvas = tk.Canvas(ws, height=HEIGHT, width=WIDTH)
canvas.configure(bg='white')
canvas.bind("<Enter>",enter)
canvas.bind("<Leave>",leave)
canvas.master.bind('a', createObjects)
canvas.master.bind('<Escape>', close)
canvas.pack()
listOfCanvas.append(canvas)
logging.info(str(run) + '~' + 'Running Tkinter Startup...')
startUp()
ws.mainloop()
