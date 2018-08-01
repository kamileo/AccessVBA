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


limit = 20
multi = 1
lsta = primer(limit)
for x in lsta:
    for y in range(1, limit):
        if x ** y > limit:
            multi = multi * x **(y-1)
            break
print(multi)