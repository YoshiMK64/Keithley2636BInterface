from future.moves import tkinter as tk
import os
from model import Model
from view import View
from keithley2600 import Keithley2600
from tkinter import messagebox

k = Keithley2600('TCPIP0::192.168.0.4::INSTR')

class Controller:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1300x720")
        self.root.configure(bg="#0096D7")
        self.root.resizable(width=False, height=False)
        self.model = Model()
        self.view = View(self.root, self.model)

        try:
            k.display.screen = k.display.SMUA
            k.display.smua.measure.func = k.display.MEASURE_DCAMPS
        except:
            messagebox.showerror('Unable to connect to Sourcemeter',
                                 'Must wait until device line frequency detected before starting application. '
                                 'Please check device is on/try restarting the device and restarting the application')

    def run(self):

        self.root.title("KEITHLEY 2636B Sourcemeter")

        def full_quit():
            self.root.quit()
            try:
                k.smua.source.levelv = 0  # set voltage to 0
                k.smua.source.output = k.smua.OUTPUT_OFF  # turn output off
            except:
                pass

        self.root.protocol("WM_DELETE_WINDOW", full_quit)
        self.root.mainloop()

############################################################################################################################### Created by Jordan Beard 20/21 Season ########################################