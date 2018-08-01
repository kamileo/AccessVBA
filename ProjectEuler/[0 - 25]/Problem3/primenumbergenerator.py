limit = 50
s = 0
for x in range(2, limit):
    for y in range(2, x):
        if x % y == 0:
            s = 0
            break
        else:
            s = x
    if s != 0:
        print(s)
print(s)


def primer(limit):
    lst = []
    s = 0
    for x in range(2, limit):
        if x == 2:
            lst.append(2)
        for y in range(2, x):
            if x == 2 or x % y != 0:
                s = x
            elif x % y == 0:
                s = 0
                break
        if s != 0:
            lst.append(s)
    return lst