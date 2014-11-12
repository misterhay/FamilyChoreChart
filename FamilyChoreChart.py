#!/usr/bin/env python
#code and images are released into the public domain, see unlicense.org

import gdata.spreadsheet.service #https://code.google.com/p/gdata-python-client/
import time #for the timestamp
import getpass #to hide the password when it is typed in
from Tkinter import * #for the GUI
import tkMessageBox #for the GUI

#Google Spreadsheet stuff
email = '' #value removed for privacy reasons
password = getpass.getpass()
spreadsheetKey = '' #value removed for privacy reasons
worksheetId0 = 'od6' #where the chore data will be logged
worksheetId1 = 'oqahut5' #a pivot table for adding up the points
worksheetId2 = 'oe64snu' #the sheet containing chore labels and point values

#the colors chosen by each kid (called "kids" because "children" has other meanings in programming)
kColor = 'orange'
eColor = 'purple'
jColor = 'green'

def authenticate():
    client = gdata.spreadsheet.service.SpreadsheetsService() 
    client.debug = False # feel free to toggle this 
    client.email = email 
    client.password = password 
    client.source = 'Hay Family Chore Chart' 
    client.ProgrammaticLogin()
    return client

def spreadsheetWriter(kid, choreNumber):
    timestamp = time.strftime('%Y-%m-%d %H:%M:%S %p')
    client = authenticate()
    #get the number of rows in the worksheet
    worksheetsFeed = client.GetWorksheetsFeed(key=spreadsheetKey)
    rowCount = worksheetsFeed.entry[0].row_count.text
    rowNumber = int(rowCount) + 1
    print 'We will write to row number', rowNumber
    choreFormula = '=vlookup(C' + str(rowNumber) + ', LookupPoints!A:B, 2, "FALSE")'
    pointsFormula = '=vlookup(D' + str(rowNumber) + ', LookupPoints!B:C, 2, "FALSE")'
    row = {} #column headings in the first row, capitals and spaces disregarded
    row['timestamp'] = timestamp
    row['kid'] = kid
    row['chorenumber'] = choreNumber
    row['chore'] = choreFormula
    row['pointsawarded'] = pointsFormula
    client.InsertRow(row, spreadsheetKey, worksheetId0) #write the new row to spreadsheet[0]
    #write a new row to the CSV file
    choreLog = open('choreLog.csv', 'a') #open/create the local log file for appending
    spreadsheetRow = timestamp + ',' + kid + ',' + choreNumber
    choreLog.write(spreadsheetRow)
    choreLog.write('\n')
    choreLog.close() #close the CSV file
    return spreadsheetRow

def choreDone(kid, choreNumber):
    #print kid, choreNumber
    points = client.GetListFeed(spreadsheetKey, worksheetId1) #this is where the pivot table is
    totalPointsLabel['text'] = points.entry[3].content.text.strip('points: ')
    ePointsLabel['text'] = points.entry[0].content.text.strip('points: ')
    jPointsLabel['text'] = points.entry[1].content.text.strip('points: ')
    kPointsLabel['text'] = points.entry[2].content.text.strip('points: ')
    print spreadsheetWriter(kid, str(choreNumber))

#get some stuff from the Google Spreadsheet
client = authenticate()
worksheetsFeed = client.GetWorksheetsFeed(key = spreadsheetKey)
choreListValues = client.GetListFeed(spreadsheetKey, worksheetId2)
points = client.GetListFeed(spreadsheetKey, worksheetId1) #this is where the pivot table is

#get the current points
ePoints = points.entry[0].content.text.strip('points: ')
jPoints = points.entry[1].content.text.strip('points: ')
kPoints = points.entry[2].content.text.strip('points: ')
totalPoints = points.entry[3].content.text.strip('points: ')

#initialise the GUI window
choreWindow = Tk()
choreWindow.title('Hay Family Chore Chart')

#load the images
chore1image = PhotoImage(file='chore1.gif')
chore2image = PhotoImage(file='chore2.gif')
chore3image = PhotoImage(file='chore3.gif')
chore4image = PhotoImage(file='chore4.gif')
chore5image = PhotoImage(file='chore5.gif')
chore6image = PhotoImage(file='chore6.gif')
chore7image = PhotoImage(file='chore7.gif')
chore8image = PhotoImage(file='chore8.gif')
chore9image = PhotoImage(file='chore9.gif')
chore10image = PhotoImage(file='chore10.gif')
chore11image = PhotoImage(file='chore11.gif')

#column labels
Label(choreWindow, font=(16), text='K').grid(column=1, row=0)
Label(choreWindow, font=(16), text='E').grid(column=2, row=0)
Label(choreWindow, font=(16), text='J').grid(column=3, row=0)
Label(choreWindow, text='Total Points').grid(column=4, row=0)
totalPointsLabel = Label(choreWindow, font=(72), text=totalPoints)
totalPointsLabel.grid(column=4, row=1)
Label(choreWindow, text='K points:').grid(column=4, row=3)
kPointsLabel = Label(choreWindow, text=kPoints)
kPointsLabel.grid(column=5, row=3)
Label(choreWindow, text='E points:').grid(column=4, row=4)
ePointsLabel = Label(choreWindow, text=ePoints)
ePointsLabel.grid(column=5, row=4)
Label(choreWindow, text='J points:').grid(column=4, row=5)
jPointsLabel = Label(choreWindow, text=jPoints)
jPointsLabel.grid(column=5, row=5)

#chore buttons
Button(choreWindow, text=choreListValues.entry[0].content.text.strip('chore: '), image=chore1image, compound=TOP, background=kColor, command=lambda: choreDone('K', 1)).grid(row=1, column=1)
Button(choreWindow, text=choreListValues.entry[1].content.text.strip('chore: '), image=chore2image, compound=TOP, background=kColor, command=lambda: choreDone('K', 2)).grid(row=2, column=1)
Button(choreWindow, text=choreListValues.entry[2].content.text.strip('chore: '), image=chore3image, compound=TOP, background=kColor, command=lambda: choreDone('K', 3)).grid(row=3, column=1)
Button(choreWindow, text=choreListValues.entry[3].content.text.strip('chore: '), image=chore4image, compound=TOP, background=kColor, command=lambda: choreDone('K', 4)).grid(row=4, column=1)
Button(choreWindow, text=choreListValues.entry[4].content.text.strip('chore: '), image=chore5image, compound=TOP, background=kColor, command=lambda: choreDone('K', 5)).grid(row=5, column=1)
Button(choreWindow, text=choreListValues.entry[5].content.text.strip('chore: '), image=chore6image, compound=TOP, background=kColor, command=lambda: choreDone('K', 6)).grid(row=6, column=1)
Button(choreWindow, text=choreListValues.entry[6].content.text.strip('chore: '), image=chore7image, compound=TOP, background=kColor, command=lambda: choreDone('K', 7)).grid(row=7, column=1)
Button(choreWindow, text=choreListValues.entry[7].content.text.strip('chore: '), image=chore8image, compound=TOP, background=kColor, command=lambda: choreDone('K', 8)).grid(row=8, column=1)
Button(choreWindow, text=choreListValues.entry[8].content.text.strip('chore: '), image=chore9image, compound=TOP, background=kColor, command=lambda: choreDone('K', 9)).grid(row=9, column=1)
Button(choreWindow, text=choreListValues.entry[9].content.text.strip('chore: '), image=chore10image, compound=TOP, background=kColor, command=lambda: choreDone('K', 10)).grid(row=10, column=1)
Button(choreWindow, text=choreListValues.entry[10].content.text.strip('chore: '), image=chore11image, compound=TOP, background=kColor, command=lambda: choreDone('K', 11)).grid(row=11, column=1)
Button(choreWindow, text=choreListValues.entry[0].content.text.strip('chore: '), image=chore1image, compound=TOP, background=eColor, command=lambda: choreDone('E', 1)).grid(row=1, column=2)
Button(choreWindow, text=choreListValues.entry[1].content.text.strip('chore: '), image=chore2image, compound=TOP, background=eColor, command=lambda: choreDone('E', 2)).grid(row=2, column=2)
Button(choreWindow, text=choreListValues.entry[2].content.text.strip('chore: '), image=chore3image, compound=TOP, background=eColor, command=lambda: choreDone('E', 3)).grid(row=3, column=2)
Button(choreWindow, text=choreListValues.entry[3].content.text.strip('chore: '), image=chore4image, compound=TOP, background=eColor, command=lambda: choreDone('E', 4)).grid(row=4, column=2)
Button(choreWindow, text=choreListValues.entry[4].content.text.strip('chore: '), image=chore5image, compound=TOP, background=eColor, command=lambda: choreDone('E', 5)).grid(row=5, column=2)
Button(choreWindow, text=choreListValues.entry[5].content.text.strip('chore: '), image=chore6image, compound=TOP, background=eColor, command=lambda: choreDone('E', 6)).grid(row=6, column=2)
Button(choreWindow, text=choreListValues.entry[6].content.text.strip('chore: '), image=chore7image, compound=TOP, background=eColor, command=lambda: choreDone('E', 7)).grid(row=7, column=2)
Button(choreWindow, text=choreListValues.entry[7].content.text.strip('chore: '), image=chore8image, compound=TOP, background=eColor, command=lambda: choreDone('E', 8)).grid(row=8, column=2)
Button(choreWindow, text=choreListValues.entry[8].content.text.strip('chore: '), image=chore9image, compound=TOP, background=eColor, command=lambda: choreDone('E', 9)).grid(row=9, column=2)
Button(choreWindow, text=choreListValues.entry[9].content.text.strip('chore: '), image=chore10image, compound=TOP, background=eColor, command=lambda: choreDone('E', 10)).grid(row=10, column=2)
Button(choreWindow, text=choreListValues.entry[10].content.text.strip('chore: '), image=chore11image, compound=TOP, background=eColor, command=lambda: choreDone('E', 11)).grid(row=11, column=2)
Button(choreWindow, text=choreListValues.entry[0].content.text.strip('chore: '), image=chore1image, compound=TOP, background=jColor, command=lambda: choreDone('J', 1)).grid(row=1, column=3)
Button(choreWindow, text=choreListValues.entry[1].content.text.strip('chore: '), image=chore2image, compound=TOP, background=jColor, command=lambda: choreDone('J', 2)).grid(row=2, column=3)
Button(choreWindow, text=choreListValues.entry[2].content.text.strip('chore: '), image=chore3image, compound=TOP, background=jColor, command=lambda: choreDone('J', 3)).grid(row=3, column=3)
Button(choreWindow, text=choreListValues.entry[3].content.text.strip('chore: '), image=chore4image, compound=TOP, background=jColor, command=lambda: choreDone('J', 4)).grid(row=4, column=3)
Button(choreWindow, text=choreListValues.entry[4].content.text.strip('chore: '), image=chore5image, compound=TOP, background=jColor, command=lambda: choreDone('J', 5)).grid(row=5, column=3)
Button(choreWindow, text=choreListValues.entry[5].content.text.strip('chore: '), image=chore6image, compound=TOP, background=jColor, command=lambda: choreDone('J', 6)).grid(row=6, column=3)
Button(choreWindow, text=choreListValues.entry[6].content.text.strip('chore: '), image=chore7image, compound=TOP, background=jColor, command=lambda: choreDone('J', 7)).grid(row=7, column=3)
Button(choreWindow, text=choreListValues.entry[7].content.text.strip('chore: '), image=chore8image, compound=TOP, background=jColor, command=lambda: choreDone('J', 8)).grid(row=8, column=3)
Button(choreWindow, text=choreListValues.entry[8].content.text.strip('chore: '), image=chore9image, compound=TOP, background=jColor, command=lambda: choreDone('J', 9)).grid(row=9, column=3)
Button(choreWindow, text=choreListValues.entry[9].content.text.strip('chore: '), image=chore10image, compound=TOP, background=jColor, command=lambda: choreDone('J', 10)).grid(row=10, column=3)
Button(choreWindow, text=choreListValues.entry[10].content.text.strip('chore: '), image=chore11image, compound=TOP, background=jColor, command=lambda: choreDone('J', 11)).grid(row=11, column=3)

choreWindow.mainloop() #start the window
