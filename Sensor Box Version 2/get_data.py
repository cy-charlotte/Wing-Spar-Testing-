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

# Nice colors for the terminal (not necessary, but cool)
class bcolors:
    HEADER = "\033[95m"
    OKBLUE = "\033[94m"
    OKCYAN = "\033[96m"
    OKGREEN = "\033[92m"
    WARNING = "\033[93m"
    FAIL = "\033[91m"
    ENDC = "\033[0m"
    BOLD = "\033[1m"
    UNDERLINE = "\033[4m"

os.environ["BLINKA_MCP2221"] = "1"
os.environ["BLINKA_MCP2221_RESET_DELAY"] = "-1"

# Gravitational acceleration to calculate weight in Newton
g = 9.81

# Calibration coeffiecients for distance
a_dist = 0  # (test values)
b_dist = 0

# Calibration coefficients for load
a_load = 0
b_load = 0

# Load cell
loadCelSensor = NAU7802(board.I2C(), address=0x2a, active_channels=1)

# Time of flight sensor
i2c = busio.I2C(board.SCL, board.SDA)
tofSensor = adafruit_vl53l0x.VL53L0X(i2c)

print(f"{bcolors().BOLD}Starting measurements. \n{bcolors().ENDC}")


# begin time of program to make graphs over time
beginTime = time.time()


# Perform measurements
try:

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('data.xlsx')
    worksheet = workbook.add_worksheet()

    # Starting row in excel
    row = 2
    # Starting column in excel
    col = 0

    # Add the sensor table names
    worksheet.write(row,col, "Sensor measured data")
    worksheet.write(row+1,col, "Sensor load")
    worksheet.write(row+1,col+1, "Sensor distance")

    # Add the program table names
    worksheet.write(row,col+3, "Program measured data")
    worksheet.write(row+1,col+3, "Program load")
    worksheet.write(row+1,col+4, "Program distance")


    # x value of graph
    x = []
    # distance array (the y value of the graph)
    distance_data = []
    # load array (the y value of the graph)
    load_data = []

    # Generate plots (load over time (ax1), distance over time (ax2), load over distance (ax3))
    fig, (ax1, ax2, ax3) = plt.subplots(3)
    fig.suptitle('Plots')

    # Open the csv file to get the live data
    with open('outfile.csv', 'w') as f:
        # Give table names in the csv file (will result in a excel readable file)
        f.write("Sensor load, Sensor distance, Program load, Program distance\n") 
        while True:
            # Get current time
            currentTime = time.time()

            # Get sensor readings
            loadCellValue = loadCelSensor.read()
            tofValue = tofSensor.range

            # Output sensor data
            print(f"{bcolors().OKBLUE}Sensor =>{bcolors().ENDC} Load cell: {loadCellValue}, Distance: {tofValue}")

            # To stop the while loop at any point by pressing 'Enter'
            if msvcrt.kbhit():
                if msvcrt.getch() == b'\r':
                    break

            # Calculate calibrated load and distance
            load = (int(loadCellValue)-b) / (a*1000) * g # In newton
            distance = (int(tofValue)-b) / a # in milimeters

            # Output program data
            print(f"{bcolors().OKGREEN}Program =>{bcolors().ENDC} Load cell (N): {load}, Distance (mm): {distance}")

            distance_data.append(int(load))
            load_data.append(int(distance))

            x.append(currentTime-beginTime)

            # Add data to plots
            ax1.plot(x, y1)
            ax2.plot(x, y2)
            ax3.plot(y2, y1)

            # Labels for the graph
            plt.ylabel('y')
            plt.xlabel('x')

            # Live data graph
            plt.show(block=False)
            plt.pause(0.5)

            # Write appropiate files to worksheet in excel file
            worksheet.write(row,col,loadCellValue)
            worksheet.write(row,col+1,tofValue)
            worksheet.write(row,col+3,load)
            worksheet.write(row,col+4,distance)

            row+=1

            # Also write it to the live csv file
            f.write(str(loadCellValue)+","+str(tofValue)+","+str(load)+","+str(distance)+"\n")
            f.flush()

            # Sleep for 100 milliseconds (10Hz)
            time.sleep(0.1)

# Exit
except KeyboardInterrupt:
    print("\nexiting...\n")
    workbook.close()
    plt.show()
    exit()

workbook.close()
plt.show()
