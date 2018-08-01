x1 = 1
y1 = 1
s = 0
limit = 4000000

while y1 < limit:
    if y1 % 2 == 0:
        s += y1
    h = y1
    y1 = x1 + y1
    x1 = h
print(s)
