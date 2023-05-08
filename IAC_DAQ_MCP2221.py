# Just in case the environment variables were not properly set
from cedargrove_nau7802 import NAU7802
import adafruit_vl53l0x
import time
import busio
import board
import os
import xlsxwriter
from matplotlib import pyplot as plt
import scipy.constants
import msvcrt


os.environ["BLINKA_MCP2221"] = "1"
os.environ["BLINKA_MCP2221_RESET_DELAY"] = "-1"


calibrating = False

a = 44.012
b = 53540

# Load cell
loadCelSensor = NAU7802(board.I2C(), address=0x2a, active_channels=1)

# Time of flight sensor
i2c = busio.I2C(board.SCL, board.SDA)
tofSensor = adafruit_vl53l0x.VL53L0X(i2c)

print("Starting measurements. \n")

beginTime = time.time()


# Perform measurements
try:
    # Create a workbook and add a worksheet.

    row = 0  # Starting row in excel
    col = 2  # Starting column in excel
    all_row = 0  # The row on which the average values begin
    if calibrating:
        workbook_name = "calibrating_data.xlsx"

    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()

    x = []
    y1 = []
    y2 = []
    fig, (ax1, ax2, ax3) = plt.subplots(3)
    fig.suptitle('Vertically stacked subplots')

    while True:
        currentTime = time.time()
        # appliedLoad = input("Applied load? (for exit type 'exit')")
        # if appliedLoad == "exit":
        #     break

        # Get sensor readings
        loadCellValue = loadCelSensor.read()
        tofValue = tofSensor.range

        print("Got data")
        if msvcrt.kbhit():
            if msvcrt.getch() == b'\r':
                break

        load_cell_data = int(loadCellValue)
        time_of_flight = int(tofValue)

        if calibrating:
            if int(load_cell_data) > 0:
                worksheet.write(row, col, int(appliedLoad))
                worksheet.write(row, col+1, int(load_cell_data))
                row += 1
        else:

            load = (load_cell_data-b)/a

            print("Load is: " + str(load))
            print("Distance is: " + str(time_of_flight))
            y1.append(int(load))
            y2.append(int(time_of_flight))

            x.append(currentTime-beginTime)

            ax1.plot(x, y1)
            ax2.plot(x, y2)
            ax3.plot(y2, y1)

            plt.ylabel('y')
            plt.xlabel('x')

            plt.show(block=False)
            plt.pause(0.5)

        # Sleep
        time.sleep(0.1)

    # worksheet.write(all_row, 1, "=AVERAGE(" + xlsxwriter.utility.xl_col_to_name(
    #     col+1) + "1:"+xlsxwriter.utility.xl_col_to_name(col+1)+str(row)+")")
    # worksheet.write(all_row, 0, int(appliedLoad))

    # col = col+3
    # row = 0
    # all_row = all_row + 1


# Exit
except KeyboardInterrupt:
    print("\nexiting...\n")


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
