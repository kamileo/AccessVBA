w = 0
for x in range(100, 1000):
    for y in range(100, 1000):
        i = x * y
        s = str(i)
        ln = len(s)
        if ln % 2 == 0:
            lne = int(ln/2)
            a = s[:lne]
            b = s[-lne:]
            if a == b[::-1]:
                if i > w:
                    w = i
print(w)
