import win32print

printer = "Saturn Card Printer"

handle = win32print.OpenPrinter(printer)

info = win32print.GetPrinter(handle, 2)

devmode = info["pDevMode"]

print(devmode)

print(devmode.Orientation)
print(devmode.Copies)
print(devmode.Color)
print(devmode.DriverExtra)