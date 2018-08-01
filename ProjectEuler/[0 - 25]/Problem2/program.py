x1 = 1
y1 = 2
x = 0
s = 0
# for x in range(1, 500):
while x1 < 4000000 and y1 < 4000000:
    x += 1
    if x % 2 != 0:
        x1 = x1 + y1
        if x1 % 2 == 0 and x1 < 4000000:
            s += x1
    else:
        y1 = x1 + y1
        if y1 % 2 == 0 and y1 < 4000000:
            s += y1
print(s+2)