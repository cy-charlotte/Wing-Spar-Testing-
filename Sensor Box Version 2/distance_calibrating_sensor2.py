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

# Load cell
loadCelSensor = NAU7802(board.I2C(), address=0x2a, active_channels=1)

# Time of flight sensor
i2c = busio.I2C(board.SCL, board.SDA)
tofSensor = adafruit_vl53l0x.VL53L0X(i2c)

print("Starting measurements. \n")

# Perform measurements
try:
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('distance_calibrating.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0  # Starting row in excel
    col = 2
    all_row = 0

    while True:

        # Get distance that will be calibrated on
        distance = input("Distance? (for exit type 'exit') ")
        if distance == "exit":
            break

        while True:
            # Get sensor readings
            loadCellValue = loadCelSensor.read()
            tofValue = tofSensor.range

            # Print readed distance data
            print("Readed data: " + str(tofValue))

            # Break out of while loop when 'Enter' is pressed
            if msvcrt.kbhit():
                if msvcrt.getch() == b'\r':
                    break

            # Check if reading is a positive value and if so write to the excel file
            if int(tofValue) > 0:
                worksheet.write(row, col, int(distance))
                worksheet.write(row, col+1, int(tofValue))
                row += 1

            # Sleep 100 milliseconds (10Hz)
            time.sleep(0.1)

        # Write the average of the data to an excel cell with a function 'AVERAGE'
        worksheet.write(all_row, 1, "=AVERAGE(" + xlsxwriter.utility.xl_col_to_name(
            col+1) + "1:"+xlsxwriter.utility.xl_col_to_name(col+1)+str(row)+")")
        worksheet.write(all_row, 0, int(tofValue))

        # Update the starting rows and columns for the next calibrating distance
        col = col+3
        row = 0
        all_row = all_row + 1


# Exit
except KeyboardInterrupt:
    print("\nexiting...\n")

# Add a chart to display the trendline
chart = workbook.add_chart({'type': 'scatter'})

chart.add_series({
    'name': 'Data',
    'categories': '=Sheet1!$A$1:$A$'+str(all_row),
    'values': '=Sheet1!$B$1:$B$'+str(all_row),
    'trendline': {'type': 'linear',
                  'display_equation': True,
                  },
})

# Insert the chart into the worksheet
worksheet.insert_chart('D22', chart)

# Close the worksheet
workbook.close()
