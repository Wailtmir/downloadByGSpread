from ReportsCollector import reportsCollector
from tkinter import *
from DateParser import *


def Main():
	bd = bottomDateEntry.get()
	td = topDateEntry.get()
	rc.collectReports(dateParse(bd),dateParse(td))
	rc.createReport(dateParse(bd),dateParse(td))
	rc.currentMoment = str(datetime.datetime.today()).replace(":"," ").split(".")[0]

def getAll():
	rc.getAll()
	rc.currentMoment = str(datetime.datetime.today()).replace(":"," ").split(".")[0]

def getListReport():
	rc.getListReport()
	rc.currentMoment = str(datetime.datetime.today()).replace(":"," ").split(".")[0]
	
rc = reportsCollector()
root = Tk()
root.title("Авто сбор отчётов")
root.geometry('230x145+500+250')
root.resizable(False,False)

start = Button(root,text = "Получить отчёт за указанный интервал", command = Main)

startAll = Button(root,text = "Полная выгрузка", command = getAll)

startListReport = Button(root,text = "Список отчётов", command = getListReport)

dateLabel = Label(root,text="Формат даты: дд.мм.гггг")
bottomDateLabel = Label(root, text = "Нижняя дата")
topDateLabel = Label(root, text = "Верхняя дата")

bottomDateEntry = Entry(root,text="Date")
topDateEntry = Entry(root)

dateLabel.place(x=26,y=1)

topDateLabel.place(x=1, y=55)
bottomDateLabel.place(x=1,y=30)

topDateEntry.place(x=100, y=55)
bottomDateEntry.place(x=100,y=30)

start.place(x=1,y=85)
startAll.place(x=1,y = 115)
startListReport.place(x=120,y = 115)

root.mainloop()