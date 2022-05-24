from controller import Controller
from keithley2600 import Keithley2600
k = Keithley2600('TCPIP0::192.168.0.4::INSTR')

if __name__ == '__main__':
    c = Controller()
    c.run()



############################################################################################################################### Created by Jordan Beard 20/21 Season ########################################