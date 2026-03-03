import win32print
import win32con
import ctypes

PRINTER = "Saturn Card Printer"

def export(path: str):
    h = win32print.OpenPrinter(PRINTER)
    try:
        dev = win32print.GetPrinter(h, 2)["pDevMode"]
        dd = dev.DriverData  # driver-private bytes (DriverExtra region)
        # In some builds this may be 'bytes' or 'str'. Normalize to bytes.
        if isinstance(dd, str):
            dd = dd.encode("latin1", errors="ignore")
        with open(path, "wb") as f:
            f.write(dd)
        print("Saved DriverData bytes:", len(dd), "DriverExtra:", dev.DriverExtra)
    finally:
        win32print.ClosePrinter(h)

# export("driverdata_off.bin")   # first run (Black Panel OFF)
export("driverdata_on.bin")  # second run (Black Panel ON) -> change filename