from future.moves import tkinter as tk
from tkinter import *
from tkinter import ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.animation as animation
from datetime import datetime
from keithley2600 import Keithley2600
from tkinter import messagebox
import xlsxwriter
import time
import os
from os import path

#Testing for update to git
k = Keithley2600('TCPIP0::192.168.0.4::INSTR')

class View:
    def __init__(self, root, model):                                    # initialising the GUI window, was supposed to seperate view and controller but is now view/controller

        self.frame = tk.Frame(root, borderwidth=0)
        self.model = model                                              #initialising model class to store and retrieve user input

        clear = open('iv_data.txt', 'r+')
        clear.truncate(0)                                              #clear stored .txt data
        clear1 = open('iv_data1.txt', 'r+')
        clear1.truncate(0)
        clear1 = open('iv_data2.txt', 'r+')
        clear1.truncate(0)
        clear1 = open('iv_data3.txt', 'r+')
        clear1.truncate(0)
        clear1 = open('iv_data4.txt', 'r+')
        clear1.truncate(0)


        self.testCounter = 0                                            #initialising run counter to for filename to avoid overwriting previous data

        self.tabControl = ttk.Notebook(root)                           #creating tabs
        self.noteStyle = ttk.Style()                                   #notebook style
        self.noteStyle.configure('TNotebook.Tab', background='#0096D7', borderwidth=0, font=("Open Sans", 12, "normal"))    #configuring tabs
        self.noteStyle.configure('TNotebook', background='#0096D7', borderwidth=0)                                          #configuring behind tabs
        self.tabControl.pack(expand=True, fill="both")

        def tabChange(e):                                               #changing widget implementation depending on tab
            if self.tabControl.tab(self.tabControl.select(), "text") == "Student":
                self.model.resetWidgets()
                self.vMin(self.Student)
                self.vMax(self.Student)
                self.steps(self.Student)
                self.sweep(self.Student)
                self.delay(self.Student)
                self.runs(self.Student)
                self.setVoltAndCurLimStud(self.Student)
                self.sampleName(self.Student)

            elif self.tabControl.tab(self.tabControl.select(), "text") == "Sensei":
                self.model.resetWidgets()
                self.vMin(self.Sensei)
                self.vMax(self.Sensei)
                self.steps(self.Sensei)
                self.sweep(self.Sensei)
                self.delay(self.Sensei)
                self.runs(self.Sensei)
                self.setVoltAndCurLimSens(self.Sensei)
                self.sampleName(self.Sensei)



        self.tabControl.bind("<<NotebookTabChanged>>", tabChange)                   #binding on tabs

        self.Student = ttk.Frame(self.tabControl)                                   # initialising tabs
        self.Student.pack_propagate(0)
        self.tabControl.add(self.Student, text='Student')
        self.student_layout = MainWindow(master=self.Student)
        self.student_layout.pack(fill="both", expand=True)
        self.student_layout.pack_propagate(0)

        self.Sensei = ttk.Frame(self.tabControl)
        self.Sensei.pack_propagate(0)
        self.tabControl.add(self.Sensei, text='Sensei')
        self.sensei_layout = MainWindow(master=self.Sensei)
        self.sensei_layout.pack(fill="both", expand=True)
        self.sensei_layout.pack_propagate(0)

        # initialising main buttons/widgets
        self.runButton(root)
        self.stopButton(root)
        self.saveButton(root)

        # --------------------------------------------------------------------------------------------------------------

    def sampleName(self, master):   #Sample/file name entry widget implementation
        myFrame = Frame(master)
        myFrame.pack()

        self.sampleNameLabel = Label(master, text="Sample name", bg='#FFFDC1')  #Sample name heading
        self.sampleNameLabel.place(bordermode=INSIDE, relx=0.025, rely=0.005)
        self.sampleNameLabel.config(font=("Montserrat", 13, "normal"))

        def keypressed(event):          #setting event binding to button press which sets the model Sample name string to user input
            x = self.nam.get()
            self.model.setSampleName(x)

        self.nam = tk.Entry(master, width = 20)         #initialising entry widget for user input of sample name
        self.nam.place(bordermode=INSIDE, relx=0.07, rely=0.05)
        self.nam.bind("<KeyRelease>", keypressed)       #binding widget to kepressed function (above)

        #================== save destination directory =================================================================================

        self.sampleNameLabel = Label(master, text="Destination Folder Directory", bg='#FFFDC1')  # Sample name heading
        self.sampleNameLabel.place(bordermode=INSIDE, relx=0.27, rely=0.005)
        self.sampleNameLabel.config(font=("Montserrat", 12, "normal"))

        def keypressed1(event):  # setting event binding to button press which sets the model Sample name string to user input
            x = self.nam1.get()
            self.model.set_destination_folder(x)

        self.nam1 = tk.Entry(master, width=20)  # initialising entry widget for user input of sample name
        self.nam1.place(bordermode=INSIDE, relx=0.3, rely=0.05)
        self.nam1.bind("<KeyRelease>", keypressed1)


    # --------------------------------------------------------------------------------------------------------------
    def vMin(self,master):                  #Bias minimum entry widget implementation, same process as 'def sampleName'

        myFrame = Frame(master)
        myFrame.pack()

        self.biasLabel = Label(master, text="Bias", bg='#FFFDC1')
        self.biasLabel.place(bordermode=INSIDE, relx=0.025, rely=0.125)
        self.biasLabel.config(font=("Montserrat", 13, "normal"))

        self.initiallabel = Label(master, text="Min. Voltage", bg='#FFFDC1')
        self.initiallabel.place(bordermode=INSIDE, relx=0.08, rely=0.15)
        self.initiallabel.config(font=("Open Sans", 10, "normal"))

        def keypressed(event):
            self.model.setVMin(self.initialVoltage.get())

        self.initialVoltage = tk.Entry(master, width = 20)
        self.initialVoltage.place(bordermode=INSIDE, relx=0.07, rely=0.2)
        self.initialVoltage.bind("<KeyRelease>", keypressed)

    # --------------------------------------------------------------------------------------------------------------
    def vMax(self,master):              #Bias maximum entry widget implementation, same process as 'def sampleName'
        myFrame = Frame(master)
        myFrame.pack()

        self.finallabel = Label(master, text="Max. Voltage", bg='#FFFDC1')
        self.finallabel.place(bordermode=INSIDE, relx=0.315, rely=0.15)
        self.finallabel.config(font=("Open Sans", 10, "normal"))

        def keypressed(event):
            self.model.setVMax(self.finalVoltage.get())

        self.finalVoltage = tk.Entry(master, width = 20)
        self.finalVoltage.place(bordermode=INSIDE, relx=0.3, rely=0.2)
        self.finalVoltage.bind("<KeyRelease>", keypressed)

    # --------------------------------------------------------------------------------------------------------------
    def steps(self, master):                 #interval/step size entry widget implementation, same process as 'def sampleName'
        myFrame = Frame(master)
        myFrame.pack()

        self.stepEntryLabel = Label(master, text="Step Size (V)", bg='#FFFDC1')
        self.stepEntryLabel.place(bordermode=INSIDE, relx=0.025, rely=0.4)
        self.stepEntryLabel.config(font=("Montserrat", 13, "normal"))

        def keypressed(event):
            x = self.stepEntry.get()
            self.model.setSteps(x)


        self.stepEntry = tk.Entry(master)
        self.stepEntry.place(bordermode=INSIDE, relx=0.185, rely=0.445)
        self.stepEntry.bind("<KeyRelease>", keypressed)

    # --------------------------------------------------------------------------------------------------------------
    def sweep(self,master):         # implementing sweep radiobuttons
        myFrame = Frame(master)
        myFrame.pack()

        self.sweeplabel = Label(master, text="Sweep", bg='#FFFDC1')         #generating sweep label
        self.sweeplabel.place(bordermode=INSIDE, relx=0.025, rely=0.275)
        self.sweeplabel.config(font=("Montserrat", 13, "normal"))
        self.sweeplabel.option_add("*font", ("Open Sans", 10, "normal"))

        var = tk.IntVar()

        self.positiveSweepButton = Radiobutton(master, text="Positive", variable=var, value=1,
                                               command=lambda: self.posSweepFunction(), bg='#FFFDC1',
                                               activebackground='#FFFDC1')                              #generating positive radiobutton and attaching to negative radiobutton with 'value'
        self.positiveSweepButton.place(bordermode=INSIDE, relx=0.075, rely=0.325)                       #placing radio button in root frame

        self.negativeSweepButton = Radiobutton(master, text="Negative", variable=var, value=2,
                                               command=lambda: self.negSweepFunction(), bg='#FFFDC1',
                                               activebackground='#FFFDC1')                              #as above
        self.negativeSweepButton.place(bordermode=INSIDE, relx=0.3, rely=0.325)

        self.queryButton = Button(master, text="Save", command=self.sweep_query)
        self.imgQ = PhotoImage(
            file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/query.png")  # make sure to add "/" not "\" ====== This needs to be set to the location of the Save button picture
        self.queryButton.config(image=self.imgQ)  # configuring the queryButton to be the picture
        self.queryButton.place(relx=0.075, rely=0.275)  # positioning the query button

    def sweep_query(self):  # creating query button function that shows how to use
        novi = Toplevel()
        novi.title("How To Use SourceMeter Interface")
        novi.resizable(width=False, height=False)
        canvas = Canvas(novi, width = 777, height = 378)
        canvas.pack()
        expImg= PhotoImage(file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/QImg.png")
        canvas.create_image(777/2, 378/2, image = expImg, anchor = CENTER)
        canvas.expImg = expImg

    def posSweepFunction(self):                                                     #setting result of radiobutton activation
        self.model.setSweep(True)

    def negSweepFunction(self):
        self.model.setSweep(False)

#--------------------------------------------------------------------------------------------------------------
    def delay(self, master):                                                        #delay time entry widget implementation, same process as 'def sampleName'
        myFrame = Frame(master)
        myFrame.pack()

        self.delayTimeLabel = Label(master, text="Delay Time (ms)", bg='#FFE3D3')
        self.delayTimeLabel.place(bordermode=INSIDE, relx=0.025, rely=0.5025)
        self.delayTimeLabel.config(font=("Montserrat", 13, "normal"))

        def keypressed(event):
            x = self.delayTime.get()
            self.model.setDelayTime(x)

        self.delayTime = tk.Entry(master)
        self.delayTime.place(bordermode=INSIDE, relx=0.07, rely=0.55)
        self.delayTime.bind("<KeyRelease>", keypressed)


#================================================ Unique initial delay time widget creator

        def keypressedOp(event):                        #event action retrieve value from model
            x = self.opDelayTime.get()
            self.model.setOpDelayTime(x)                #set 'optional' delay to retrieved value from entry

        self.opDelayTime = tk.Entry(master)
        self.opDelayTime.place(bordermode=INSIDE, relx=0.3, rely=0.55)
        self.opDelayTime.bind("<KeyRelease>", keypressedOp)
        self.opDelayTime.configure(state='disabled')

        self.var= IntVar()

                                        #
        self.cb = Checkbutton(master,
                            text="Distinct Initial Point Delay",
                            bg='#FFE3D3',
                            activebackground='#FFE3D3',
                            command=lambda e=self.opDelayTime, v=self.var: self.naccheck(e,v))
        self.cb.place(bordermode=INSIDE, relx=0.28, rely=0.5025)


    def naccheck(self, entry, var):             #set response to check tick i.e. active/inactive
        if var.get() == 0:
            entry.configure(state='normal')
            self.model.setInitialDelay(TRUE)       #Set tracker to true
            var.set(1)
        else:
            entry.configure(state='disabled')
            self.model.setInitialDelay(False)
            var.set(0)


# ------------------------------------------------------------------------------------------------------------
    def runs(self, master):      #runs widget implementation, same process as 'def sampleName' but using combobox to restrict possibilities
        myFrame = Frame(master)
        myFrame.pack()

        variable = IntVar(master)
        variable.set(1)  # default value

        self.noOfRunsLabel = Label(master, text="Runs", bg='#FFE3D3')
        self.noOfRunsLabel.place(bordermode=INSIDE, relx=0.025, rely=0.6)
        self.noOfRunsLabel.config(font=("Montserrat", 13, "normal"))

        def setRuns(e):
            x =self.runs_combo.get()
            self.model.setRuns(x)

        self.runs_combo = ttk.Combobox(master, value=[1, 2, 3, 4, 5])
        self.runs_combo.current(0)
        self.runs_combo.bind("<<ComboboxSelected>>", setRuns)
        self.runs_combo.place(bordermode=INSIDE, relx=0.18, rely=0.65)

    # --------------------------------------------------------------------------------------------------------------
    def setVoltAndCurLimSens(self, master):             #organising limitations for sensei/student with voltage and current limitations.
            myFrame = Frame(master)
            myFrame.pack()

            self.Voltage2LimLabel = Label(master, text="Set Voltage Limit (V)", bg='#FFE3D3')       #Voltage limit label
            self.Voltage2LimLabel.place(bordermode=INSIDE, relx=0.025, rely=0.725)
            self.Voltage2LimLabel.config(font=("Montserrat", 13, "normal"))

            self.Current2LimLabel = Label(master, text="Set Current Limit (A)", bg='#FFE3D3')       #current limit label
            self.Current2LimLabel.place(bordermode=INSIDE, relx=0.025, rely=0.85)
            self.Current2LimLabel.config(font=("Montserrat", 13, "normal"))
            # creating a list
            self.voltageLim2 = ["0.2", "2", "20", "200"]                                            #voltage limit combobox options
            self.currentLimFree2 = ["0.001", "0.01", "0.1", "1", "1.5"]                             #current limit combobox options if unlimited
            self.currentLimClosed2 = ["0.001"]                                                      #current limit combobox options if limited

            def voltageLim2(e):                                                                     #making the comboboxes dependent on tabs and limits
                if self.voltage_combo2.get() == "0.2":
                    self.model.setVoltageLim(0.2)
                    self.current_combo2.config(value=self.currentLimFree2)
                    self.current_combo2.current(0)
                elif self.voltage_combo2.get() == "2":
                    self.current_combo2.config(value=self.currentLimFree2)
                    self.model.setVoltageLim(2)
                elif self.voltage_combo2.get() == "20":
                    self.current_combo2.config(value=self.currentLimFree2)
                    self.model.setVoltageLim(20)
                elif self.voltage_combo2.get() == "200":
                    self.current_combo2.config(value=self.currentLimFree2)
                    self.model.setVoltageLim(200)

            def currentLim2(e):
                if self.current_combo2.get() == "0.001":
                    self.model.setCurrentLim(0.001)
                elif self.current_combo2.get() == "0.01":
                    self.model.setCurrentLim(0.01)
                elif self.current_combo2.get() == "0.1":
                    self.model.setCurrentLim(0.1)
                elif self.current_combo2.get() == "1":
                    self.model.setCurrentLim(1)
                elif self.current_combo2.get() == "1.5":
                    self.model.setCurrentLim(1.5)

            self.voltage_combo2 = ttk.Combobox(master, value=self.voltageLim2)                          #voltage combobox for Sensei tab
            self.voltage_combo2.current(0)                                                              #setting automatic setting to lowest option
            self.voltage_combo2.place(bordermode=INSIDE, relx=0.18, rely=0.79)                          #positioning box in root window

            self.voltage_combo2.bind("<<ComboboxSelected>>", voltageLim2)                               #binding voltage combobox to choice made

            self.current_combo2 = ttk.Combobox(master, value=self.currentLimFree2)                      #as above
            self.current_combo2.current(0)
            self.current_combo2.place(bordermode=INSIDE, relx=0.18, rely=0.915)
            self.current_combo2.bind("<<ComboboxSelected>>", currentLim2)
    # --------------------------------------------------------------------------------------------------------------
    def setVoltAndCurLimStud(self, master):                                                             #same as Sensei tab but changing condition for 200V limit chosen
            myFrame = Frame(master)
            myFrame.pack()

            self.VoltageLimLabel = Label(master, text="Set Voltage Limit (V)", bg='#FFE3D3')
            self.VoltageLimLabel.place(bordermode=INSIDE, relx=0.025, rely=0.725)
            self.VoltageLimLabel.config(font=("Montserrat", 13, "normal"))

            self.CurrentLimLabel = Label(master, text="Set Current Limit (A)", bg='#FFE3D3')
            self.CurrentLimLabel.place(bordermode=INSIDE, relx=0.025, rely=0.85)
            self.CurrentLimLabel.config(font=("Montserrat", 13, "normal"))
            # creating a list
            self.voltageLim = ["0.2", "2", "20", "200"]
            self.currentLimFree = ["0.001", "0.01", "0.1", "1", "1.5"]
            self.currentLimClosed = ["0.001"]

            def voltageLim(e):
                if self.voltage_combo.get() == "0.2":
                    self.model.setVoltageLim(0.2)
                    self.current_combo.config(value=self.currentLimFree)
                    self.current_combo.current(0)
                elif self.voltage_combo.get() == "2":
                    self.current_combo.config(value=self.currentLimFree)
                    self.model.setVoltageLim(2)
                elif self.voltage_combo.get() == "20":
                    self.current_combo.config(value=self.currentLimFree)
                    self.model.setVoltageLim(20)
                elif self.voltage_combo.get() == "200":
                    self.current_combo.config(value=self.currentLimClosed)
                    self.model.setVoltageLim(200)

            def currentLim(e):
                if self.current_combo.get() == "0.001":
                    self.model.setCurrentLim(0.001)
                elif self.current_combo.get() == "0.01":
                    self.model.setCurrentLim(0.01)
                elif self.current_combo.get() == "0.1":
                    self.model.setCurrentLim(0.1)
                elif self.current_combo.get() == "1":
                    self.model.setCurrentLim(1)
                elif self.current_combo.get() == "1.5":
                    self.model.setCurrentLim(1.5)

            self.voltage_combo = ttk.Combobox(master, value=self.voltageLim)
            self.voltage_combo.current(0)
            self.voltage_combo.place(bordermode=INSIDE, relx=0.18, rely=0.79)

            self.voltage_combo.bind("<<ComboboxSelected>>", voltageLim)

            self.current_combo = ttk.Combobox(master, value=self.currentLimFree)
            self.current_combo.current(0)
            self.current_combo.place(bordermode=INSIDE, relx=0.18, rely=0.915)
            self.current_combo.bind("<<ComboboxSelected>>", currentLim)
# -------------------------------------------------------------------------------------------------------------
                                                                                                #creating a save button
    def saveButton(self, master):
        self.myFrame = Frame(master)
        self.myFrame.pack()
        self.master = master


        self.saveButton = Button(master, text="Save", command=lambda: [self.dataCheck()])
        self.img2 = PhotoImage(file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/Save.png")    # make sure to add "/" not "\" ====== This needs to be set to the location of the Save button picture
        self.saveButton.config(image=self.img2)                                                                 #configuring the saveButton to be the picture
        self.saveButton.place(relx=0.47, rely=0.065)                                                            #positioning the save button

    def dataCheck(self):
        try:
            self.saveTest()
        except:
            messagebox.showinfo('Save error', 'User has excel file of same name open \n \t OR \n User has no data to save')

    def saveTest(self):                                                                                             #saving to excel under sample name
        filename = self.model.getSampleName()+str(self.testCounter)+'.xlsx'                                         #naming the filename with a counter to limit files being overwritten, only works for multiple runs in the same sitting
        file_destination = self.model.get_destination_folder()
        data_amount = (float(self.model.getVMax())-float(self.model.getVMin()))/float(self.model.getSteps())+3      #calculating points taken to write into excel (adding 3 to cover the time stamp and headings)

        newPath = file_destination.replace('\\', '/')


        if path.exists(newPath) == False:
            try:
                os.mkdir(newPath)
            except OSError:
                messagebox.showinfo("Save error", "Unknown directory error, file saved to C:/Users/Public/Documents")
                newPath = "C:/Users/Public/Documents/"+str(self.model.getSampleName())+"Test"

        if path.exists(newPath) == True:
            pass

        if self.model.getSweep() == True:
            sweep_dir = 'Positive'
        else:
            sweep_dir = 'Negative'

        workbook = xlsxwriter.Workbook(newPath+ '/' + filename)                              #creating excel workbook called filename

        worksheet = workbook.add_worksheet()                                                 #adding worksheet to workbook

        name_time = ['Sample:', str(self.model.getSampleName()), 'Date/Time:', str(self.current_datetime), '', '', 'Bias (V):', str(self.model.getVMin()), str(self.model.getVMax()),'Delay (ms):', str(self.model.getDelayTime()), 'Sweep:', sweep_dir] #creating array for first row
                                                                                            #creating arrays for column headings
        headings5 = ['Voltage R1', 'Current R1','Voltage R2', 'Current R2','Voltage R3', 'Current R3','Voltage R4', 'Current R4','Voltage R5', 'Current R5']
        headings4 = ['Voltage R1', 'Current R1','Voltage R2', 'Current R2','Voltage R3', 'Current R3','Voltage R4', 'Current R4']
        headings3 = ['Voltage R1', 'Current R1','Voltage R2', 'Current R2','Voltage R3', 'Current R3']
        headings2 = ['Voltage R1', 'Current R1','Voltage R2', 'Current R2']
        headings1 = ['Voltage R1', 'Current R1']
        worksheet.write_row('A1', name_time)                                                #placing timestamp in cell A1
        chart1 = workbook.add_chart({'type': 'scatter'})                                    #creating scatter chart for data
        chart1.set_title({'name': 'IV characteristics'})                                    #naming scatter chart
        chart1.set_x_axis({'name': 'Voltage (V)'})                                          #Axis headings
        chart1.set_y_axis({'name': 'Current (A)'})

        if int(self.model.getRuns()) == 1:                                                                                #implementing data input with condition of runs the user implemented
            worksheet.write_row('A2', headings1)
            worksheet.write_column('A3', self.model.getVoltageReadings1())                  #input voltage readings in column A
            worksheet.write_column('B3', self.model.getCurrentReadings1())                  #input current readings in column B

            chart1.add_series({                                                             #adding data series for scatter plot
                'name': 'Run 1',                                                            #series name
                'categories': '=Sheet1!$A$3:$A$' + str(data_amount),                        #adding x values
                'values': '=Sheet1!$B$3:$B$' + str(data_amount),                            #adding y values
            })

        elif int(self.model.getRuns()) == 2:                                                #as above but for 2 runs/data sets
            worksheet.write_row('A2', headings2)
            worksheet.write_column('A3', self.model.getVoltageReadings1())
            worksheet.write_column('B3', self.model.getCurrentReadings1())
            worksheet.write_column('C3', self.model.getVoltageReadings2())
            worksheet.write_column('D3', self.model.getCurrentReadings2())

            chart1.add_series({
                'name': 'Run 1',
                'categories': '=Sheet1!$A$3:$A$' + str(data_amount),
                'values': '=Sheet1!$B$3:$B$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 2',
                'categories': '=Sheet1!$C$3:$C$' + str(data_amount),
                'values': '=Sheet1!$D$3:$D$' + str(data_amount),
            })

        elif int(self.model.getRuns()) == 3:                                                #as above but for 3 runs/data sets
            worksheet.write_row('A2', headings3)
            worksheet.write_column('A3', self.model.getVoltageReadings1())
            worksheet.write_column('B3', self.model.getCurrentReadings1())
            worksheet.write_column('C3', self.model.getVoltageReadings2())
            worksheet.write_column('D3', self.model.getCurrentReadings2())
            worksheet.write_column('E3', self.model.getVoltageReadings3())
            worksheet.write_column('F3', self.model.getCurrentReadings3())

            chart1.add_series({
                'name': 'Run 1',
                'categories': '=Sheet1!$A$3:$A$' + str(data_amount),
                'values': '=Sheet1!$B$3:$B$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 2',
                'categories': '=Sheet1!$C$3:$C$' + str(data_amount),
                'values': '=Sheet1!$D$3:$D$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 3',
                'categories': '=Sheet1!$E$3:$E$' + str(data_amount),
                'values': '=Sheet1!$F$3:$F$' + str(data_amount),
            })

        elif int(self.model.getRuns()) == 4:                                                #as above but for 4 runs/data sets
            worksheet.write_row('A2', headings4)
            worksheet.write_column('A3', self.model.getVoltageReadings1())
            worksheet.write_column('B3', self.model.getCurrentReadings1())
            worksheet.write_column('C3', self.model.getVoltageReadings2())
            worksheet.write_column('D3', self.model.getCurrentReadings2())
            worksheet.write_column('E3', self.model.getVoltageReadings3())
            worksheet.write_column('F3', self.model.getCurrentReadings3())
            worksheet.write_column('G3', self.model.getVoltageReadings4())
            worksheet.write_column('H3', self.model.getCurrentReadings4())

            chart1.add_series({
                'name': 'Run 1',
                'categories': '=Sheet1!$A$3:$A$' + str(data_amount),
                'values': '=Sheet1!$B$3:$B$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 2',
                'categories': '=Sheet1!$C$3:$C$' + str(data_amount),
                'values': '=Sheet1!$D$3:$D$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 3',
                'categories': '=Sheet1!$E$3:$E$' + str(data_amount),
                'values': '=Sheet1!$F$3:$F$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 4',
                'categories': '=Sheet1!$G$3:$G$' + str(data_amount),
                'values': '=Sheet1!$H$3:$H$' + str(data_amount),
            })


        elif int(self.model.getRuns()) == 5:                                                #as above but for 5 runs/data sets
            worksheet.write_row('A2', headings5)
            worksheet.write_column('A3', self.model.getVoltageReadings1())
            worksheet.write_column('B3', self.model.getCurrentReadings1())
            worksheet.write_column('C3', self.model.getVoltageReadings2())
            worksheet.write_column('D3', self.model.getCurrentReadings2())
            worksheet.write_column('E3', self.model.getVoltageReadings3())
            worksheet.write_column('F3', self.model.getCurrentReadings3())
            worksheet.write_column('G3', self.model.getVoltageReadings4())
            worksheet.write_column('H3', self.model.getCurrentReadings4())
            worksheet.write_column('I3', self.model.getVoltageReadings5())
            worksheet.write_column('J3', self.model.getCurrentReadings5())

            chart1.add_series({
                'name': 'Run 1',
                'categories': '=Sheet1!$A$3:$A$' + str(data_amount),
                'values': '=Sheet1!$B$3:$B$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 2',
                'categories': '=Sheet1!$C$3:$C$' + str(data_amount),
                'values': '=Sheet1!$D$3:$D$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 3',
                'categories': '=Sheet1!$E$3:$E$' + str(data_amount),
                'values': '=Sheet1!$F$3:$F$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 4',
                'categories': '=Sheet1!$G$3:$G$' + str(data_amount),
                'values': '=Sheet1!$H$3:$H$' + str(data_amount),
            })

            chart1.add_series({
                'name': 'Run 5',
                'categories': '=Sheet1!$I$3:$I$' + str(data_amount),
                'values': '=Sheet1!$J$3:$J$' + str(data_amount),
            })

        chart1.set_style(11)

        # Insert the chart into the worksheet (with an offset).
        worksheet.insert_chart('K4', chart1)

        workbook.close()                #closing the workbook

        messagebox.showinfo("Saved", "Test data has been saved to "+ str(newPath.replace('\/', ': ')) +" folder with file name: " + filename)  #notifying the user of file name and location (also feedback that save has successfully completed)


    # -------------------------------------------------------------------------------------------------------------
    def runButton(self, master):                                            #implementing run button
        self.myFrame = Frame(master)
        self.myFrame.pack()
        self.master = master
        self.counter = 2
        self.voltagereading=[]
        self.currentreading=[]

        self.runButton = Button(master, text="RUN", command= lambda:[self.checkValues(), self.runTest()])
        self.img1 = PhotoImage(file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/RUN.png")  # make sure to add "/" not "\" =======Needs to be directory to Run Button image
        self.runButton.config(image=self.img1)                                            #configuring run button to image
        self.runButton.place(relx=0.463, rely=0.35)                                        #placement of the run button

    def runTest(self):
        if self.counter == 9:                                                             #Only start run if 9 (actually less but just in case) conditions are met
            self.current_datetime = datetime.now()                                        #taking date and time for start of run (used in save button)
            self.model.setStopTest(False)                                                 #setting condition for emergency stop button

            self.testCounter += 1                                                          #increment test counter for filename

            self.model.resetReadings()                                                    #reset model data arrays
            self.model.setRunCounter(0)                                                   #counter for tracking position in test
            self.model.setRunCounter(self.model.getRunCounter() + 1)                      #increment position in test to 1 (first run)

            k.smua.measure.rangei = float(self.model.getCurrentLim())                                                  #measure rangei
            k.smua.source.limiti = float(self.model.getCurrentLim())                      #implementing limits and ranges in Keithley

            k.smua.trigger.measure.action = k.smua.ENABLE
            k.smua.source.autorangei = k.smua.AUTORANGE_ON                                  #source rangei

            k.smua.measure.rangev = float(self.model.getVoltageLim())                       #range of voltage source
            k.smua.source.limitv = float(self.model.getVoltageLim())                        #voltage limit

            k.smua.measure.delay = k.smua.DELAY_AUTO
            k.smua.measure.delayfactor = 2.0
            k.display.smua.measure.func = k.display.MEASURE_DCAMPS

            #clearing data txt files
            clear = open('iv_data.txt', 'r+')
            clear.truncate(0)
            clear1 = open('iv_data1.txt', 'r+')
            clear1.truncate(0)
            clear2 = open('iv_data2.txt', 'r+')
            clear2.truncate(0)
            clear3 = open('iv_data3.txt', 'r+')
            clear3.truncate(0)
            clear4 = open('iv_data4.txt', 'r+')
            clear4.truncate(0)


            Sweep(self.master, self.model)                  #initialising first sweep

            self.counter = 2
        else:
            self.counter = 2
            pass




    def checkValues(self):                  #checking these conditions before implementing run to avoid error and keithley non-compliance
     #========================================================================================
        a = self.model.getVMin()
        try:                            #making sure user input is a number
            a = float(a)
            if -200 <= a <= 200 and abs(a)<=self.model.getVoltageLim():     #if a number make sure it is in keithley boundaries and less than user chosen limit
                self.counter += 1                                           #increment counter (need 9 for test to run [see above], starts at 3)
            else:
                self.counter = 2                                           #if failed set counter back to original state and send error box
                messagebox.showinfo("Incorrect Maximum Voltage value", "Value must be a number between -200 and 200, with magnitude less than voltage limit")

        except ValueError:                                                  #if user input not a number, send error and set counter to 3
            messagebox.showinfo("Incorrect Initial Voltage value", "Value must be a number")
            self.counter = 2
     # ========================================================================================
        b = self.model.getVMax()
        try:                                                                    #(as above but for bias maximum)
            b = float(b)
            if -200 <= b <= 200 and b <= self.model.getVoltageLim():
                if b >= a:
                    self.counter += 1
                else:
                    messagebox.showinfo("Incorrect Maximum Voltage value",
                                        "Value must be greater than or equal to voltage minimum")
            else:
                messagebox.showinfo("Incorrect Maximum Voltage value", "Value must be a number between -200 and 200, with magnitude less than voltage limit ")
                self.counter = 2

        except ValueError:
            messagebox.showinfo("Incorrect Maximum Voltage value", "Value must be a number")
            self.counter = 2

     # ========================================================================================
        c = self.model.getSteps()
        try:
            c = float(c)
            if 0 <= c <= (float(b)-float(a)):        #step size must be less than difference Vmax-Vmin or send error
                self.counter += 1
            else:
                messagebox.showinfo("Incorrect Step interval value", "Value must be greater than 0 and no larger than bias difference")
                self.counter = 2
        except ValueError:
            messagebox.showinfo("Incorrect Step interval value", "Value must be a number")
            self.counter = 2
        # ========================================================================================
        d = self.model.getDelayTime()
        try:                                #checking user has input an integer
            d = int(d)
            if 1 <= d <=1200000:             #setting limit of delay time (900000 ms is 15 minutes)
                self.counter += 1
            else:
                messagebox.showinfo("Incorrect Delay value", "Value must be an integer between 1 and 1,200,000 (ms)")
                self.counter = 2

        except ValueError:
            messagebox.showinfo("Incorrect Delay value", "Value must be an integer between 1 and 1,200,000 (ms)")
            self.counter = 2
         #============ New Addition first data point delay option ======================

        if self.model.getInitialDelay() == False:
            self.counter += 1
        else:
            newDelay = self.model.getOpDelayTime()
            try:  # checking user has input an integer
                newDelay = int(newDelay)
                if 1 <= newDelay <= 1200000:  # setting limit of delay time (900000 ms is 15 minutes)
                    self.counter += 1
                else:
                    messagebox.showinfo("Incorrect initial point Delay value", "Value must be an integer between 1 and 1,200,000 (ms)")
                    self.counter = 2
            except ValueError:
                messagebox.showinfo("Incorrect Initial Point Delay value", "Value must be an integer between 1 and 1,200,000 (ms)")
                self.counter = 2


        # ========================================================================================
            #checking that if student tab is selected and voltage limit is 200V, that current limit is minimum
        if self.tabControl.tab(self.tabControl.select(), "text") == "Student" and self.model.getVoltageLim()==200 and self.model.getCurrentLim() != 0.001:
            messagebox.showinfo("Voltage/Current Limit Error", "If voltage limit = 200V; Student must use current limit = 0.001 A")
            self.counter = 2
        else:
            self.counter +=1

            # ========================================================================================
        #Testing keithley is compliant
        error_count = k.errorqueue.count    #calling number of current errors

        if error_count > 0:
            errorCode, message, severity, errorNode = k.errorqueue.next()
            messagebox.showinfo("Keithley non-compliant",
                                    str(errorCode)+','+str(message)+','+str(severity)+','+str(errorNode))
            k.errorqueue.clear()
            self.counter = 2
        else:
            self.counter += 1

# -------------------------------------------------------------------------------------------------------------
    def stopButton(self, master):
        self.stopButton = Button(master, text="Stop!", command=self.stopFunction)
        self.img = PhotoImage(
            file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/STOPRUN1.png")  # make sure to add "/" not "\" ===== Needs to be address of stop button image
        self.stopButton.config(image=self.img)
        self.stopButton.place(relx=0.45, rely=0.43)  # Displaying the button

    def stopFunction(self):                                                 #Building emergency stop button
        k.smua.source.levelv = 0                                            #set voltage to 0
        k.smua.source.output = k.smua.OUTPUT_OFF                            #turn output off
        self.model.setStopTest(True)                                        #set stop condition to True to avoid sweep continuing

        k.display.clear()                                                   #clear keithley screen for message send
        k.display.setcursor(1, 1)                                           #set cursor on display
        k.display.settext("Emergency Stop")                                 #send message to keithley display
        time.sleep(2)                                                       #sleep to let message be read for 2 seconds
        k.display.clear()                                                   #clear display
        k.display.screen = k.display.SMUA                                   #Set screen to smua
        messagebox.showinfo("Emergency Stop", "Test ended")                 #display message in GUI that test has ended

# -------------------------------------------------------------------------------
class MainWindow(tk.Frame):                                                 #main window design
    def __init__(self, master=None):                                        #initial conditions for root window
        self.parent = master
        tk.Frame.__init__(self, self.parent, bg='white', borderwidth=0, relief="flat", width=1300, height=700)
        self.__create_layout()                                               #no idea why this is needed, just following online example

    def __create_layout(self):                                               #see last comment
        pass

    def __create_layout(self):                                               #actual layout
    #------------------- Setting frames and colours ---------------------------------
        self.FrameTopLeft = tk.Frame(self, bg = '#FFFDC1')
        self.FrameTopRight = tk.Frame(self, bg = '#FFC9F0')
        self.FrameBottomLeft = tk.Frame(self, bg = '#FFE3D3')
        self.FrameBottomRight = tk.Frame(self, bg = '#DEFDFF')              #two frames bottom right for better control of scrollable window placement
        self.FrameBottomRightData = tk.Frame(self, bg='#DEFDFF')

                                                                            #==below==positioning of the five frames
        self.FrameTopLeft.place(bordermode=INSIDE, relheight=0.5, relwidth=0.5, relx = 0, rely = 0)
        self.FrameTopRight.place(bordermode=INSIDE, relheight=0.5, relwidth=0.5, relx = 0.5, rely = 0)
        self.FrameBottomLeft.place(bordermode=INSIDE, relheight=0.5, relwidth=0.5, relx = 0, rely = 0.5)
        self.FrameBottomRight.place(bordermode=INSIDE, relheight=0.15, relwidth=0.5, relx = 0.5, rely = 0.5)
        self.FrameBottomRightData.place(bordermode=INSIDE, relheight=0.35, relwidth=0.5, relx=0.5, rely=0.65)

        #===========initialising live plots and live data=========
        self.create_display(self.FrameTopRight)
        self.show_data_title(self.FrameBottomRight)
        self.show_data(self.FrameBottomRightData)


        #=============== Creator signature lol ======================
        self.Signature = PhotoImage(file="C:/Users/mini_/PycharmProjects/Keithley2636BInterface/Images/Signature.png")
        self.label1 = Label(self.FrameBottomLeft, image= self.Signature, relief='flat', bd=-2 )
        self.label1.place(relx=0.65, rely=0.9)

    def create_display(self, frame):                    #creating live pot that fills top right frame
        self.fig = plt.figure(figsize=(4, 3), dpi=100)
        self.fig.patch.set_facecolor('#FFC9F0')
        plt.style.use('seaborn-pastel')                 #plot style, many options available, pastel fit application palette
        self.iv_graph = self.fig.add_subplot(111)

        self.chart_type = FigureCanvasTkAgg(self.fig, master=frame)
        self.chart_type.get_tk_widget().pack(fill='both', expand=True)

        #------for liveplotting-------
        # Subplot titles
        self.iv_graph.set_title('IV characteristics')

        # Subplot Axis label
        self.iv_graph.set_xlabel('Voltage (V)')
        self.iv_graph.set_ylabel('Current (A)')

        def animate(i):
            graph_data = open('iv_data.txt', 'r').read()
            graph_data1 = open('iv_data1.txt', 'r').read()
            graph_data2 = open('iv_data2.txt', 'r').read()
            graph_data3 = open('iv_data3.txt', 'r').read()
            graph_data4 = open('iv_data4.txt', 'r').read()


            lines = graph_data.split('\n')
            lines1 = graph_data1.split('\n')
            lines2 = graph_data2.split('\n')
            lines3 = graph_data3.split('\n')
            lines4 = graph_data4.split('\n')

            xs = []
            ys = []
            xs1 = []
            ys1 = []
            xs2 = []
            ys2 = []
            xs3 = []
            ys3 = []
            xs4 = []
            ys4 = []

            for line in lines:                  #for as many lines as iv_datax contains creating an array to be plotted
                if len(line) > 1:
                    x, y = line.split(',')
                    xs.append(float(x))
                    ys.append(float(y))

            for line in lines1:
                if len(line) > 1:
                    x, y = line.split(',')
                    xs1.append(float(x))
                    ys1.append(float(y))

            for line in lines2:
                if len(line) > 1:
                    x, y = line.split(',')
                    xs2.append(float(x))
                    ys2.append(float(y))

            for line in lines3:
                if len(line) > 1:
                    x, y = line.split(',')
                    xs3.append(float(x))
                    ys3.append(float(y))

            for line in lines4:
                if len(line) > 1:
                    x, y = line.split(',')
                    xs4.append(float(x))
                    ys4.append(float(y))


            # For IV curve
            self.iv_graph.clear()                       #clear  iv_graph for new plot data to be included
            self.iv_graph.scatter(xs, ys)               #plot the different newly retrieved data sets
            self.iv_graph.scatter(xs1, ys1)
            self.iv_graph.scatter(xs2, ys2)
            self.iv_graph.scatter(xs3, ys3)
            self.iv_graph.scatter(xs4, ys4)

            self.iv_graph.set_title('IV characteristics')    #plot title etc..
            self.iv_graph.set_xlabel('Voltage (V)')
            self.iv_graph.set_ylabel('Current (A)')

        ani = animation.FuncAnimation(self.fig, animate, interval=1000)         #create live plot (animate)
        plt.tight_layout()                                                      #make sure words etc are seen on plot
        self.chart_type.draw()                                                  #show plot

    def show_data_title(self, frame):
        myFrame = Frame(frame)                                                  #live data for bottom right frames
        myFrame.grid()

        frame.columnconfigure((1,2,3,4,5,6,7,8,9,10,11,12,13,14), weight=1)

        frame.rowconfigure((0, 1), weight=1)
                                                                                #headings
        self.dataFrame = Label(frame, text="Live Data Readings", bg='#DEFDFF')
        self.dataFrame.grid(row=0, column=7)
        self.dataFrame.config(font=("Montserrat", 14, "normal"))

        self.voltReading = Label(frame, text="Voltage", bg='#DEFDFF')
        self.voltReading.grid(row=1, column=6)
        self.voltReading.config(font=("Montserrat", 12, "normal"))

        self.curReading = Label(frame, text="Current", bg='#DEFDFF')
        self.curReading.grid(row=1, column=8)
        self.curReading.config(font=("Montserrat", 12, "normal"))

    def show_data(self, parent):
        self.myFrame = tk.Frame(parent)
        self.myFrame.pack()
        self.frameData = ScrollableFrame(parent)
        self.frameData.pack()  # fill="y", expand = FALSE, side="top"
        self.dataLabels = []

        self.xs = []
        self.ys = []
        self.xs1 = []
        self.ys1 = []
        self.xs2 = []
        self.ys2 = []
        self.xs3 = []
        self.ys3 = []
        self.xs4 = []
        self.ys4 = []


        self.data = open('iv_data.txt', 'r').read()
        self.lines = self.data.split('\n')

        self.data1 = open('iv_data1.txt', 'r').read()
        lines1 = self.data1.split('\n')

        self.data2 = open('iv_data2.txt', 'r').read()
        lines2 = self.data2.split('\n')

        self.data3 = open('iv_data3.txt', 'r').read()
        lines3 = self.data3.split('\n')

        self.data4 = open('iv_data4.txt', 'r').read()
        lines4 = self.data4.split('\n')


        for line in self.lines:                 #as seen in live plotting but string form to create labels
            if len(line) > 1:
                x, y = line.split(',')
                self.xs.append(x)
                self.ys.append(y)

        for line in lines1:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs1.append(x)
                self.ys1.append(y)

        for line in lines2:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs2.append(x)
                self.ys2.append(y)

        for line in lines3:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs3.append(x)
                self.ys3.append(y)

        for line in lines4:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs4.append(x)
                self.ys4.append(y)

        for i in range(len(self.xs)):                               #printing labels to scrollable frame
            self.dataReading1 = Label(self.frameData.scrollable_frame, text='Run 1:'+'\t\t'+ self.xs[i] + '\t\t\t\t' + self.ys[i],
                                      bg='#FFFFFF')
            self.dataReading1.grid(row=i, sticky='w' + 'n' +'e')
            self.dataReading1.config(font=("Roboto", 11, "normal")) #live data readings font information
            self.dataLabels.append(self.dataReading1)               #storing all created labels in an array so they can be destroyed and recreated with new information

        for i in range(len(self.xs1)):
            self.dataReading2 = Label(self.frameData.scrollable_frame, text='Run 2:'+'\t\t' + self.xs1[i] + '\t\t\t\t' + self.ys1[i],
                                      bg='#FFFFFF')
            self.dataReading2.grid(row=i + len(self.xs), sticky='w' + 'n' +'e')
            self.dataReading2.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading2)

        for i in range(len(self.xs2)):
            self.dataReading3 = Label(self.frameData.scrollable_frame, text='Run 3:'+'\t\t' + self.xs2[i] + '\t\t\t\t' + self.ys2[i],
                                      bg='#FFFFFF')
            self.dataReading3.grid(row=i + (2*len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading3.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading3)

        for i in range(len(self.xs3)):
            self.dataReading4 = Label(self.frameData.scrollable_frame, text='Run 4:'+'\t\t' + self.xs3[i] + '\t\t\t\t' + self.ys3[i],
                                      bg='#FFFFFF')
            self.dataReading4.grid(row=i + (3*len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading4.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading4)

        for i in range(len(self.xs4)):
            self.dataReading5 = Label(self.frameData.scrollable_frame, text='Run 5:'+'\t\t' + self.xs4[i] + '\t\t\t\t' + self.ys4[i],
                                      bg='#FFFFFF')
            self.dataReading5.grid(row=i + (4*len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading5.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading5)

        self.myFrame.after(5000, self.update_data)

    def update_data(self):
        """ refresh the content of the label every 5 seconds """
        self.xs.clear()
        self.ys.clear()
        self.xs1.clear()
        self.ys1.clear()
        self.xs2.clear()
        self.ys2.clear()
        self.xs3.clear()
        self.ys3.clear()
        self.xs4.clear()
        self.ys4.clear()


        for label in self.dataLabels:
            label.destroy()

        self.data = open('iv_data.txt', 'r').read()
        self.lines = self.data.split('\n')

        self.data1 = open('iv_data1.txt', 'r').read()
        lines1 = self.data1.split('\n')

        self.data2 = open('iv_data2.txt', 'r').read()
        lines2 = self.data2.split('\n')

        self.data3 = open('iv_data3.txt', 'r').read()
        lines3 = self.data3.split('\n')

        self.data4 = open('iv_data4.txt', 'r').read()
        lines4 = self.data4.split('\n')

        for line in self.lines:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs.append(x)
                self.ys.append(y)

        for line in lines1:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs1.append(x)
                self.ys1.append(y)

        for line in lines2:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs2.append(x)
                self.ys2.append(y)

        for line in lines3:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs3.append(x)
                self.ys3.append(y)

        for line in lines4:
            if len(line) > 1:
                x, y = line.split(',')
                self.xs4.append(x)
                self.ys4.append(y)

        for i in range(len(self.xs)):
            self.dataReading1 = Label(self.frameData.scrollable_frame,
                                      text='Run 1:'+'\t\t' + self.xs[i] + '\t\t\t\t' + self.ys[i],
                                      bg='#FFFFFF')
            self.dataReading1.grid(row=i, sticky='w' + 'n' +'e')
            self.dataReading1.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading1)

        for i in range(len(self.xs1)):
            self.dataReading2 = Label(self.frameData.scrollable_frame,
                                      text='Run 2:'+'\t\t' + self.xs1[i] + '\t\t\t\t' + self.ys1[i],
                                      bg='#FFFFFF')
            self.dataReading2.grid(row=i + len(self.xs), sticky='w' + 'n' +'e')
            self.dataReading2.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading2)

        for i in range(len(self.xs2)):
            self.dataReading3 = Label(self.frameData.scrollable_frame,
                                      text='Run 3:'+'\t\t' + self.xs2[i] + '\t\t\t\t' + self.ys2[i],
                                      bg='#FFFFFF')
            self.dataReading3.grid(row=i + (2 * len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading3.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading3)

        for i in range(len(self.xs3)):
            self.dataReading4 = Label(self.frameData.scrollable_frame,
                                      text='Run 4:'+'\t\t' + self.xs3[i] + '\t\t\t\t' + self.ys3[i],
                                      bg='#FFFFFF')
            self.dataReading4.grid(row=i + (3 * len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading4.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading4)

        for i in range(len(self.xs4)):
            self.dataReading5 = Label(self.frameData.scrollable_frame,
                                      text='Run 5:'+'\t\t' + self.xs4[i] + '\t\t\t\t' + self.ys4[i],
                                      bg='#FFFFFF')
            self.dataReading5.grid(row=i + (4 * len(self.xs)), sticky='w' + 'n' +'e')
            self.dataReading5.config(font=("Roboto", 11, "normal"))
            self.dataLabels.append(self.dataReading5)

        self.myFrame.after(5000, self.update_data)          #this is the callback for every 5 seconds (5000 ms)



class ScrollableFrame(ttk.Frame):                           #creating the scrollable frame for live data readings
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(container)
        container.config(bg='#DEFDFF')
        canvas.config(bg='#FFFFFF')
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all"), bg='#FFFFFF'
            )
        )
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set, bg='#FFFFFF')
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


class Sweep(ttk.Frame):                         #creating the sweep functionality as a class
    def __init__(self, parent, model):
        self.start_time = time.time()           #tracking the start time of the sweep so it can be deducted from the delay time for higher accuracy delay
        # variable storing time
        self.myFrame = tk.Frame(parent)
        self.parent = parent
        self.myFrame.pack()
        self.model = model
        self.voltagereading = []
        self.currentreading = []

        if self.model.getSweep() == True:                               #testing for positive or negative sweep
            self.sweep_voltage = float(self.model.getVMin())            #implementing first voltage
            self.max = float(self.model.getVMax())                      #setting a variable for boundary testing later


            if abs(self.sweep_voltage) < 0.2:                           #trying to create higher accuracy with variable ranging
                k.smua.measure.rangev = 0.2

            elif abs(self.sweep_voltage) < 2.0:
                k.smua.measure.rangev = 2.0

            elif abs(self.sweep_voltage) < 20.0:
                k.smua.measure.rangev = 20.0

            elif abs(self.sweep_voltage) < 200:
                k.smua.measure.rangev = 200
            else:
                pass

            k.smua.source.levelv = self.sweep_voltage                   #set the keithley voltage to the user voltage
            k.smua.source.output = k.smua.OUTPUT_ON                     #turn keithley output on before first delay and loop to allow delay before first measurement

            if self.model.getInitialDelay() == False:                   #testing for unique initial delay time
                self.initialD = int(self.model.getDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getDelayTime()))
            else:
                self.initialD = int(self.model.getOpDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getOpDelayTime()))

        # label displaying time
        # start the time
            self.myFrame.after(self.initialD, self.next_voltage_pos)         #setting the delay to the delay minus time since sweep started (includes % to make sure no time drift over subsequent loops)

        elif self.model.getSweep() == False:                                #as above but for negative sweep
            self.sweep_voltage = float(self.model.getVMax())
            self.min = float(self.model.getVMin())

            if abs(self.sweep_voltage) < 0.2:                               #trying to create higher accuracy with variable ranging
                k.smua.measure.rangev = 0.2

            elif abs(self.sweep_voltage) < 2.0:
                k.smua.measure.rangev = 2.0

            elif abs(self.sweep_voltage) < 20.0:
                k.smua.measure.rangev = 20.0

            elif abs(self.sweep_voltage) < 200:
                k.smua.measure.rangev = 200
            else:
                pass

            k.smua.source.levelv = self.sweep_voltage
            k.smua.source.output = k.smua.OUTPUT_ON


            if self.model.getInitialDelay() == False:
                self.initialD = int(self.model.getDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getDelayTime()))
            else:
                self.initialD = int(self.model.getOpDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getOpDelayTime()))

            # label displaying time
            # start the time
            self.myFrame.after(self.initialD, self.next_voltage_neg)
        else:
            messagebox.showinfo("Sweep Error", "Must choose positive or negative sweep")



    def next_voltage_pos(self):                                 #actual measurements for looping

        self.current = k.smua.measure.i()

                               #use device to measure current
        self.voltagereading.append(k.smua.measure.v())          #add voltage measurement to array for saving to excel
        self.currentreading.append(self.current)                #add current reading to array for saving to excel

        if self.model.getRunCounter() == 1:                     #using run counter to determine which file to save to (this is first run)
            self.data = open('iv_data.txt', 'a')                #open file
            self.data.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")       #write data to file
            self.model.setCurrentReadings1(self.current)                                    #save to run 1 array in model
            self.model.setVoltageReadings1(self.sweep_voltage)                              #save to run 1 array in model
        elif self.model.getRunCounter() == 2:                  #as above for run 2
            self.data1 = open('iv_data1.txt', 'a')
            self.data1.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings2(self.current)
            self.model.setVoltageReadings2(self.sweep_voltage)
        elif self.model.getRunCounter() == 3:                   #as above
            self.data2 = open('iv_data2.txt', 'a')
            self.data2.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings3(self.current)
            self.model.setVoltageReadings3(self.sweep_voltage)
        elif self.model.getRunCounter() == 4:                   #as above
            self.data3 = open('iv_data3.txt', 'a')
            self.data3.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings4(self.current)
            self.model.setVoltageReadings4(self.sweep_voltage)
        elif self.model.getRunCounter() == 5:                   #as above
            self.data4 = open('iv_data4.txt', 'a')
            self.data4.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings5(self.current)
            self.model.setVoltageReadings5(self.sweep_voltage)

        self.sweep_voltage = round(self.sweep_voltage + float(self.model.getSteps()), 3)      #increment voltage tracker

        if abs(self.sweep_voltage) < 0.2:  # trying to create higher accuracy with variable ranging
            k.smua.measure.rangev = 0.2

        elif abs(self.sweep_voltage) < 2.0:
            k.smua.measure.rangev = 2.0

        elif abs(self.sweep_voltage) < 20.0:
            k.smua.measure.rangev = 20.0

        elif abs(self.sweep_voltage) < 200:
            k.smua.measure.rangev = 200
        else:
            pass

        error_count = k.errorqueue.count    #calling number of current errors

        if error_count > 0:                 #display errors
            self.model.setStopTest(True)
            errorCode, message, severity, errorNode = k.errorqueue.next()
            messagebox.showinfo("Keithley non-compliant",
                                str(errorCode) + ',' + str(message) + ',' + str(severity) + ',' + str(
                                    errorNode))
            k.errorqueue.clear()
        else:
            pass

        if self.sweep_voltage <= self.max:                  #checking conditions for boundary
            if self.model.getStopTest() == False:           #checking emergency stop has not been pressed

                k.smua.source.levelv = self.sweep_voltage   #set instrument voltage to voltage tracker
                # increment the time

                # request tkinter to call self.refresh after (the delay is given in ms)
                k.smua.measure.delay = k.smua.DELAY_AUTO
                k.smua.measure.delayfactor = 2.0
                self.myFrame.after(int(self.model.getDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getDelayTime())), self.next_voltage_pos)

        else:                               #sweep ended, set source to 0 and off
            k.smua.source.levelv = 0
            k.smua.source.output = k.smua.OUTPUT_OFF
            self.model.setRunCounter(self.model.getRunCounter() + 1)
            if int(self.model.getRunCounter()) <= int(self.model.getRuns()):    #if runcounter is less than or equal to user input runs, repeat sweep
                Sweep(self.parent, self.model)


    def next_voltage_neg(self):                                                 #as above but for negative sweep
        self.current = k.smua.measure.i()
        self.voltagereading.append(k.smua.measure.v())
        self.currentreading.append(self.current)

        if self.model.getRunCounter() == 1:
            self.data = open('iv_data.txt', 'a')
            self.data.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings1(self.current)
            self.model.setVoltageReadings1(self.sweep_voltage)
        elif self.model.getRunCounter() == 2:
            self.data1 = open('iv_data1.txt', 'a')
            self.data1.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings2(self.current)
            self.model.setVoltageReadings2(self.sweep_voltage)
        elif self.model.getRunCounter() == 3:
            self.data2 = open('iv_data2.txt', 'a')
            self.data2.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings3(self.current)
            self.model.setVoltageReadings3(self.sweep_voltage)
        elif self.model.getRunCounter() == 4:
            self.data3 = open('iv_data3.txt', 'a')
            self.data3.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings4(self.current)
            self.model.setVoltageReadings4(self.sweep_voltage)
        elif self.model.getRunCounter() == 5:
            self.data4 = open('iv_data4.txt', 'a')
            self.data4.write(str(k.smua.measure.v()) + "," + str(self.current) + "\n")
            self.model.setCurrentReadings5(self.current)
            self.model.setVoltageReadings5(self.sweep_voltage)


        self.sweep_voltage = round(self.sweep_voltage - float(self.model.getSteps()), 3)

        if abs(self.sweep_voltage) < 0.2:  # trying to create higher accuracy with variable ranging
            k.smua.measure.rangev = 0.2

        elif abs(self.sweep_voltage) < 2.0:
            k.smua.measure.rangev = 2.0

        elif abs(self.sweep_voltage) < 20.0:
            k.smua.measure.rangev = 20.0

        elif abs(self.sweep_voltage) < 200:
            k.smua.measure.rangev = 200
        else:
            pass

        error_count = k.errorqueue.count  # calling number of current errors from keithley

        if error_count > 0:
            self.model.setStopTest(True)
            errorCode, message, severity, errorNode = k.errorqueue.next()
            messagebox.showinfo("Keithley non-compliant",
                                str(errorCode) + ',' + str(message) + ',' + str(severity) + ',' + str(
                                    errorNode))
            k.errorqueue.clear()
        else:
            pass

        if self.sweep_voltage >= self.min:
            if self.model.getStopTest() == False:
                k.smua.source.levelv = self.sweep_voltage
                # increment the time
                # request tkinter to call self.refresh after 1s (the delay is given in ms)
                k.smua.measure.delay = k.smua.DELAY_AUTO
                k.smua.measure.delayfactor = 2.0
                self.myFrame.after(int(self.model.getDelayTime())-(int((time.time()-self.start_time)*1000) % int(self.model.getDelayTime())) % int(self.model.getDelayTime()), self.next_voltage_neg)

        else:
            k.smua.source.levelv = 0
            k.smua.source.output = k.smua.OUTPUT_OFF
            self.model.setRunCounter(self.model.getRunCounter() + 1)
            if int(self.model.getRunCounter()) <= int(self.model.getRuns()):
                Sweep(self.parent, self.model)




############################################################################################################################### Created by Jordan Beard 20/21 Season ########################################