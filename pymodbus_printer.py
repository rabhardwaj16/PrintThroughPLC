import csv, os, datetime, sys, win32print
from pymodbus.client.sync import ModbusTcpClient
from pymodbus.exceptions import ModbusException
client = ModbusTcpClient('192.168.1.10')

# print command
def printer(raw_data):
    default_printer = win32print.GetDefaultPrinter()
    h = win32print.OpenPrinter(default_printer)
    win32print.StartDocPrinter(h,1,("","","RAW"))
    win32print.WritePrinter(h,raw_data)
    win32print.EndDocPrinter(h)
    win32print.ClosePrinter (h)


# 16bit word to byte conversion
def word_2_2Dbyte(x):
	c = (x >> 8) & 0xff
	b = x & 0xff
	return c,b


# byte to string conversion
def word2string(v):
	conv = [word_2_2Dbyte(i) for i in v]
	b = [j for s in conv for j in s]
	ch = [chr(int(l)) for l in b]
	st = ''
	for k in ch:
		st += k
	return st


Triggered = False

while True:
    rightnow = datetime.datetime.now()
    date = rightnow.strftime("%d/%m/%Y")
    time = rightnow.strftime("%H:%M:%S")
    try:
        #-------------------------------------------------------X
        #Data Read
        BitInput = client.read_coils(0).bits
        b_trigger = BitInput[0]
        lst = client.read_holding_registers(10, 5).registers
        
        #String(List to String)
        strint_data = w2sconversion.word2string(lst)
        #-------------------------------------------------------X
    
        if b_trigger == True:
            if Triggered == False:
                raw_d = bytes (f'''
                ! 0 200 200 203 1
                PW 575
                TONE 0
                SPEED 3
                ON-FEED IGNORE
                NO-PACE
                BAR-SENSE
                BOX 12 50 100 110 6
                T 2 0 9 10 {strint_data}
                PRINT
                ''','utf-8')
                printer(raw_d)
                Triggered = True
                print(f"{date}, {time}, Printing...")
        else:
            Triggered = False
    except AttributeError:
        print("Connection Lost")
    except ModbusException:
        print("Connection Lost, waiting for connection...")
    except KeyboardInterrupt:
        print("Quitting..")
        break