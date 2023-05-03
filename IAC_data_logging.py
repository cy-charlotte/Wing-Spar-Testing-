# -*- coding: utf-8 -*-
"""
This program is used to load the data from the USB serial.
 Whenever the serial port receives data, it will be placed in
 the "line" variable. From here, it needs to be parsed and saved.
 
You can enable a "development" mode which will feed fake data 
 by setting dev = True
"""

import serial
import msvcrt
import time
from IAC_helper import port_scan, development_data
from matplotlib import pyplot as plt
import scipy.constants
import xlsxwriter

dev = True              # Development mode
usbPort = "editMe"      # Your USB port, obtain using port_scan()

calibrating = True

try:
    if not dev:
        ser = serial.Serial(usbPort, 9600)
    running = True
    print("Serial initialized succesfully!")
except:
    print("Issue with serial! Aborting...")


beginTime = time.time()

if calibrating:
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('calibrating_data.xlsx')
    worksheet = workbook.add_worksheet()
    appliedLoad = input("Applied load? ")


if dev:
    currentTime = time.time()
    x = []
    y1 = []
    y2 = []
    fig, (ax1, ax2, ax3) = plt.subplots(3)
    fig.suptitle('Vertically stacked subplots')

    row = 0  # Starting row in excel
    col = 0  # Starting column in excel

    while running:

        if msvcrt.kbhit():
            if msvcrt.getch() == b'\r':
                break

        print("continued")
        # Delay 1 second
        while currentTime + 1 > time.time():
            pass
        currentTime = time.time()
        line = development_data()[:-2].decode('utf-8')

        load_cell_data = line.split(" ")[1]
        time_of_flight = line.split(" ")[3]

        if calibrating:
            if int(load_cell_data) > 0:
                worksheet.write(row, col, int(appliedLoad))
                worksheet.write(row, col+1, int(load_cell_data))
                row += 1
        else:
            y1.append(int(load_cell_data))
            y2.append(int(time_of_flight))

            x.append(currentTime-beginTime)

            ax1.plot(x, y1)
            ax2.plot(x, y2)
            ax3.plot(y1, y2)

            plt.ylabel('y')
            plt.xlabel('x')

            plt.show(block=False)
            plt.pause(1.2)


else:
    while running:
        line = ser.readline()[:-2].decode('utf-8')
        print(line)

        ####################
        ### YOUR CODE HERE###
        ####################


if calibrating:
    worksheet.write(row, col, "gem")
    worksheet.write(row, col+1, "=AVERAGE(B1:B"+str(row)+")")

    workbook.close()
else:
    plt.show()
