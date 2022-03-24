from openpyxl import *
from tkinter import *

badwb = load_workbook('badTeams.xlsx')
goodwb = load_workbook('goodTeams.xlsx')
badSheet  = badwb.active
goodSheet = goodwb.active
def badExcel():
    badSheet.column_dimensions['A'].width = 10
    badSheet.column_dimensions['B'].width = 10
    badSheet.column_dimensions['C'].width = 10
    badSheet.column_dimensions['D'].width = 10
    badSheet.column_dimensions['E'].width = 20
    badSheet.column_dimensions['F'].width = 20
    badSheet.column_dimensions['G'].width = 10
    badSheet.cell(row=1, column=1).value = "Team #"
    badSheet.cell(row=1, column=2).value = "ballsScoredLower"
    badSheet.cell(row=1, column=3).value = "ballsScoredUpper"
    badSheet.cell(row=1, column=4).value = "Highest Climb Level"
    badSheet.cell(row=1, column=5).value = "School"
    badSheet.cell(row=1, column=6).value = "Student or Parent Led"
    badSheet.cell(row=1, column=7).value = "bad or Good"
def goodExcel():
    goodSheet.column_dimensions['A'].width = 10
    goodSheet.column_dimensions['B'].width = 10
    goodSheet.column_dimensions['C'].width = 10
    goodSheet.column_dimensions['D'].width = 10
    goodSheet.column_dimensions['E'].width = 20
    goodSheet.column_dimensions['F'].width = 20
    goodSheet.column_dimensions['G'].width = 10
    goodSheet.cell(row=1, column=1).value = "Team #"
    goodSheet.cell(row=1, column=2).value = "ballsScoredLower"
    goodSheet.cell(row=1, column=3).value = "ballsScoredUpper"
    goodSheet.cell(row=1, column=4).value = "Highest Climb Level"
    goodSheet.cell(row=1, column=5).value = "School"
    goodSheet.cell(row=1, column=6).value = "Student or Parent Led"
    goodSheet.cell(row=1, column=7).value = "bad or Good"
def focus1(event):
    ballsScoredLower.focus_set()
def focus2(event):
    ballsScoredUpper.focus_set()
def focus3(event):
    climbLevel.focus_set()
def focus4(event):
    school.focus_set()
def focus5(event):
    studentOrParentLed.focus_set()
def focus6(event):
    badOrGood.focus_set()
def clear():
    teamNum.delete(0, END)
    ballsScoredLower.delete(0, END)
    ballsScoredUpper.delete(0, END)
    climbLevel.delete(0, END)
    school.delete(0, END)
    studentOrParentLed.delete(0, END)
    badOrGood.delete(0, END)
def bad():
    if (teamNum.get() == "" and
        ballsScoredLower.get() == "" and
        ballsScoredUpper.get() == "" and
        climbLevel.get() == "" and
        school.get() == "" and
        studentOrParentLed.get() == "" and
        badOrGood.get() == ""):
        print("empty input")
    else:
        current_row = badSheet.max_row
        current_column = badSheet.max_column
        badSheet.cell(row=current_row + 1, column=1).value = teamNum.get()
        badSheet.cell(row=current_row + 1, column=2).value = lowerCount.get()
        badSheet.cell(row=current_row + 1, column=3).value = upperCount.get()
        badSheet.cell(row=current_row + 1, column=4).value = climbLevel.get()
        badSheet.cell(row=current_row + 1, column=5).value = school.get()
        badSheet.cell(row=current_row + 1, column=6).value = studentOrParentLed.get()
        badSheet.cell(row=current_row + 1, column=7).value = badOrGood.get()
        badwb.save('badTeams.xlsx')
        teamNum.focus_set()
        clear()
def good():
    if (teamNum.get() == "" and
        ballsScoredLower.get() == "" and
        ballsScoredUpper.get() == "" and
        climbLevel.get() == "" and
        school.get() == "" and
        studentOrParentLed.get() == ""):
        print("empty input")
    else:
        current_row = badSheet.max_row
        current_column = badSheet.max_column
        goodSheet.cell(row=current_row + 1, column=1).value = teamNum.get()
        goodSheet.cell(row=current_row + 1, column=2).value = lowerCount
        goodSheet.cell(row=current_row + 1, column=3).value = upperCount
        goodSheet.cell(row=current_row + 1, column=4).value = climbLevel.get()
        goodSheet.cell(row=current_row + 1, column=5).value = school.get()
        goodSheet.cell(row=current_row + 1, column=6).value = studentOrParentLed.get()
        goodwb.save('goodTeams.xlsx')
        teamNum.focus_set()
        clear()
def upperButtonPress():
    global upperCount
    upperCount = 1 + upperCount
def lowerButtonPress():
    global lowerCount
    lowerCount = 1 + lowerCount

if __name__ == "__main__":
    lowerCount = 0
    upperCount = 0
    root = Tk()
    root.configure(background='blue')
    root.title("FRC Scouting")
    # set the configuration of GUI window
    root.geometry("400x700")
    badExcel()
    goodExcel()
    heading = Label(root, text="Scouting Form", bg="white")
    num= Label(root, text="Team Number", bg="white")
    lower = Label(root, text="Balls Scored Lower", bg="White")
    upper = Label(root, text="Balls Scored Upper", bg="White")
    climb = Label(root, text="Climb Level", bg="White")
    whatSchool = Label(root, text="School", bg="White")
    studentOrParent = Label(root, text="Student Or Parent Led", bg="White")
    theDecision = Label(root, text="Are they bad or good?", bg="white")
    heading.grid(row=0, column=0)
    num.grid(row=1, column=0)
    lower.grid(row=3, column=0)
    upper.grid(row=5, column=0)
    climb.grid(row=7, column=0)
    whatSchool.grid(row=9, column=0)
    studentOrParent.grid(row=11, column=0)
    theDecision.grid(row=13, column=0)
    teamNum = Entry(root)
    ballsScoredLower = Entry(root)
    ballsScoredUpper = Entry(root)
    climbLevel = Entry(root)
    school = Entry(root)
    studentOrParentLed = Entry(root)
    badOrGood = Entry(root) 
    teamNum.bind("<Return>", focus1)
    ballsScoredLower.bind("<Return>", focus2)
    ballsScoredUpper.bind("<Return>", focus3)
    climbLevel.bind("<Return>", focus4)
    school.bind("<Return>", focus5)
    studentOrParentLed.bind("<Return>", focus6)
    teamNum.grid(row=2, column=0, ipadx="25")
    lowerButton = Button(root, text="+1 Lower", fg="White", bg="Black", command=lowerButtonPress)
    lowerButton.grid(row=4, column=0, ipadx="25")
    upperButton = Button(root, text="+1 Upper", fg="White", bg="Black", command=upperButtonPress)
    upperButton.grid(row=6, column=0, ipadx="25")
    climbLevel.grid(row=8, column=0, ipadx="25")
    school.grid(row=10, column=0, ipadx="25")
    studentOrParentLed.grid(row=12, column=0, ipadx="25") 
    badExcel()
    goodExcel()
    bad = Button(root, text="bad", fg="Black", bg="Red", command=bad)
    bad.grid(row=14, column=0)
    good = Button(root, text="good", fg="Black", bg="Green", command=good)
    good.grid(row=15, column=0)
    root.mainloop()