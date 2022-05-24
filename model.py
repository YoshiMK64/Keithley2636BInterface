from keithley2600 import Keithley2600


k = Keithley2600('TCPIP0::192.168.0.4::INSTR')

class Model:                                            #setting initial states, storing data, and getter/setter functions
    def __init__(self):
        # initialising data variables
        self.vMax = 10.0
        self.vMin = 0.0
        self.steps= 1
        self.sweep = True
        self.initialDelay = False
        self.delayTime = 1
        self.opDelayTime = 1
        self.runs = 1
        self.voltLim= 0.2
        self.currLim= 0.001
        self.sampleName= "Test"
        self.destination_folder = "C:/Users/Public/Documents/NewTest"
        self.stopTest = False
        self.runCounter = 0
        self.currReadings1 = []
        self.voltReadings1 = []
        self.currReadings2 = []
        self.voltReadings2 = []
        self.currReadings3 = []
        self.voltReadings3 = []
        self.currReadings4 = []
        self.voltReadings4 = []
        self.currReadings5 = []
        self.voltReadings5 = []

    def resetReadings(self):
        self.currReadings1 = []
        self.voltReadings1 = []
        self.currReadings2 = []
        self.voltReadings2 = []
        self.currReadings3 = []
        self.voltReadings3 = []
        self.currReadings4 = []
        self.voltReadings4 = []
        self.currReadings5 = []
        self.voltReadings5 = []

    def resetWidgets(self):
        self.vMax = 10.0
        self.vMin = 0.0
        self.steps = 1
        self.sweep = True
        self.delayTime = 1
        self.opDelayTime = 1
        self.runs = 1
        self.voltLim = 0.2
        self.currLim = 0.001
        self.sampleName = "Test"  # may need to adjust initialisation
        self.destination_folder = "C:/Users/Public/Documents/NewTest"
        self.stopTest = False
        self.runCounter = 0

    def setRunCounter(self, x):
        self.runCounter = x

    def getRunCounter(self):
        return self.runCounter

    def setVMax(self, x):
        self.vMax = x

    def getVMax(self):
        return self.vMax

    def setVMin(self, x):
        self.vMin = x

    def getVMin(self):
        return self.vMin

    def setSteps(self, x):
        self.steps = x

    def getSteps(self):
        return self.steps

    def setSweep(self, x):
        self.sweep = x

    def getSweep(self):
        return self.sweep

    def setInitialDelay(self, x):
        self.initialDelay = x

    def getInitialDelay(self):
        return self.initialDelay

    def setDelayTime(self, x):
        self.delayTime = x

    def getDelayTime(self):
        return self.delayTime

    def setOpDelayTime(self, x):
        self.opDelayTime = x

    def getOpDelayTime(self):
        return self.opDelayTime

    def setRuns(self, x):
        self.runs = x

    def getRuns(self):
        return self.runs

    def setVoltageLim(self, x):
        self.voltLim = x

    def getVoltageLim(self):
        return self.voltLim

    def setCurrentLim(self, x):
        self.currLim = x

    def getCurrentLim(self):
        return self.currLim

    def setSampleName(self, x):
        self.sampleName = x

    def getSampleName(self):
        return self.sampleName

    def set_destination_folder(self, x):
        self.destination_folder = x

    def get_destination_folder(self):
        return self.destination_folder

    def setCurrentReadings1(self, x): #may need to change for append etc
        self.currReadings1.append(x)

    def getCurrentReadings1(self):
        return self.currReadings1

    def setVoltageReadings1(self, x):  # may need to change for append etc
        self.voltReadings1.append(x)

    def getVoltageReadings1(self):
        return self.voltReadings1

    def setCurrentReadings2(self, x): #may need to change for append etc
        self.currReadings2.append(x)

    def getCurrentReadings2(self):
        return self.currReadings2

    def setVoltageReadings2(self, x):  # may need to change for append etc
        self.voltReadings2.append(x)

    def getVoltageReadings2(self):
        return self.voltReadings2

    def setCurrentReadings3(self, x): #may need to change for append etc
        self.currReadings3.append(x)

    def getCurrentReadings3(self):
        return self.currReadings3

    def setVoltageReadings3(self, x):  # may need to change for append etc
        self.voltReadings3.append(x)

    def getVoltageReadings3(self):
        return self.voltReadings3

    def setCurrentReadings4(self, x): #may need to change for append etc
        self.currReadings4.append(x)

    def getCurrentReadings4(self):
        return self.currReadings4

    def setVoltageReadings4(self, x):  # may need to change for append etc
        self.voltReadings4.append(x)

    def getVoltageReadings4(self):
        return self.voltReadings4

    def setCurrentReadings5(self, x): #may need to change for append etc
        self.currReadings5.append(x)

    def getCurrentReadings5(self):
        return self.currReadings5

    def setVoltageReadings5(self, x):  # may need to change for append etc
        self.voltReadings5.append(x)

    def getVoltageReadings5(self):
        return self.voltReadings5

    def setStopTest(self, x):  # may need to change for append etc
        self.stopTest = x

    def getStopTest(self):
        return self.stopTest

############################################################################################################################### Created by Jordan Beard 20/21 Season ########################################