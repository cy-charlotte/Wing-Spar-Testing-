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

calibrating = True      #If calibrating is true then the calibration part of the code will run which wil return values a and b with which a the test can be run

a = 0
b = 0


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
    row = 0  # Starting row in excel
    col = 2  # Starting column in excel
    all_row = 0  # The row on which the average values begin


if dev:
    currentTime = time.time()
    x = []
    y1 = []
    y2 = []
    fig, (ax1, ax2, ax3) = plt.subplots(3)
    fig.suptitle('Vertically stacked subplots')

    while calibrating:
        appliedLoad = input("Applied load? (for exit type 'exit')")
        if appliedLoad == "exit":
            break

        while running:

            if msvcrt.kbhit():
                if msvcrt.getch() == b'\r':
                    break

            # Delay 1 second
            while currentTime + 1 > time.time():
                pass
            currentTime = time.time()
            line = development_data()[:-2].decode('utf-8')
            print(line)

            load_cell_data = line.split(" ")[1]
            time_of_flight = line.split(" ")[3]

            if calibrating:
                if int(load_cell_data) > 0:                                    #load cell data with values below zero are disregarded and not put itno the excel
                    worksheet.write(row, col, int(appliedLoad))                #Applied load previously entered is put into the row and col specified  
                    worksheet.write(row, col+1, int(load_cell_data))           #Idem for the load cell data  
                    row += 1                                                   #row is increased by one so the next value is printed on the next line 
            else:

                load = load_cell_data*a + b

                y1.append(int(load))
                y2.append(int(time_of_flight))

                x.append(currentTime-beginTime)

                ax1.plot(x, y1)
                ax2.plot(x, y2)
                ax3.plot(y1, y2)

                plt.ylabel('y')
                plt.xlabel('x')

                plt.show(block=False)
                plt.pause(0.8)

        worksheet.write(all_row, 1, "=AVERAGE(" + xlsxwriter.utility.xl_col_to_name(
            col+1) + "1:"+xlsxwriter.utility.xl_col_to_name(col+1)+str(row)+")")
        worksheet.write(all_row, 0, int(appliedLoad))

        col = col+3
        row = 0
        all_row = all_row + 1


else:
    while running:
        line = ser.readline()[:-2].decode('utf-8')
        print(line)

        ####################
        ### YOUR CODE HERE###
        ####################


if calibrating:

    chart = workbook.add_chart({'type': 'scatter'})

    chart.add_series({
        'name': 'Data',
        'categories': '=Sheet1!$A$1:$A$'+str(all_row),
        'values': '=Sheet1!$B$1:$B$'+str(all_row),
        'trendline': {'type': 'linear',
                      'display_equation': True,
                      },
    })

    worksheet.insert_chart('D22', chart)

    workbook.close()
else:
    plt.show()
