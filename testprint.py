from PIL import Image, ImageWin
import win32print
import win32con
import win32gui
import win32ui

PRINTER = "Saturn Card Printer"
PNG_PATH = "card.png"
K_OFFSET = 1100
K_ON = 1

im = Image.open(PNG_PATH).convert("RGBA")
w, h = im.size
src = im.load()

cmy = im.copy()
cmy_px = cmy.load()

k = Image.new("1", (w, h), 1)
k_px = k.load()

for y in range(h):
    for x in range(w):
        r, g, b, a = src[x, y]
        if a and (r, g, b) == (0, 0, 0):
            k_px[x, y] = 0
            cmy_px[x, y] = (255, 255, 255, 0)

hP = win32print.OpenPrinter(PRINTER)
dev = win32print.GetPrinter(hP, 2)["pDevMode"]

dd = dev.DriverData
if isinstance(dd, str):
    dd = dd.encode("latin1", errors="ignore")
dd = bytearray(dd)
dd[K_OFFSET] = K_ON
dev.DriverData = bytes(dd)

_ = win32print.DocumentProperties(
    None, hP, PRINTER, dev, dev,
    win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
)
win32print.ClosePrinter(hP)

hdc = win32gui.CreateDC("WINSPOOL", PRINTER, None)
hdc = win32gui.ResetDC(hdc, dev)

dc = win32ui.CreateDCFromHandle(hdc)

pw = dc.GetDeviceCaps(win32con.HORZRES)
ph = dc.GetDeviceCaps(win32con.VERTRES)
s = min(pw / w, ph / h)
ow, oh = int(w * s), int(h * s)
l, t = (pw - ow) // 2, (ph - oh) // 2
dest = (l, t, l + ow, t + oh)

dc.StartDoc("Card")
dc.StartPage()
ImageWin.Dib(cmy).draw(dc.GetHandleOutput(), dest)
ImageWin.Dib(k).draw(dc.GetHandleOutput(), dest)
dc.EndPage()
dc.EndDoc()

dc.DeleteDC()
