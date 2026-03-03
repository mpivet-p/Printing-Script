a = open("driverdata_off.bin", "rb").read()
b = open("driverdata_on.bin", "rb").read()

print("Lengths:", len(a), len(b))

diffs = []
for i, (x, y) in enumerate(zip(a, b)):
    if x != y:
        diffs.append((i, x, y))

print("Different bytes:", len(diffs))
for i, x, y in diffs[:50]:
    print(i, x, "->", y)