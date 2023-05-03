# -*- coding: utf-8 -*-
"""
Helper scripts
"""
import random
import serial.tools.list_ports

def port_scan():
    ports = serial.tools.list_ports.comports()
    for item in ports:
        print(item)


def development_data():
    cell1 = random.randint(-10000, 10000)
    cell2 = random.randint(-10000, 10000)
    cell3 = random.randint(-10000, 10000)
    cell4 = random.randint(-10000, 10000)
    tof = random.randint(0, 5000)
    line = f'load_cell {cell1} time_of_flight {tof}\r\n'
    line = line.encode()
    return line

