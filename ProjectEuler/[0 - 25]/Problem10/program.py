limit = 2000000
s = 0
c = 0
p = [2, 3, 5, 7, 11, 13, 17]
z = [2, 3, 5, 7, 11, 13, 17]

for y in range(3, limit+1, 2):
    if y % 3 == 0:
        pass
    elif y % 5 == 0:
        pass
    elif y % 7 == 0:
        pass
    elif y % 11 == 0:
        pass
    elif y % 13 == 0:
        pass
    elif y % 17 == 0:
        pass

    else:
        p.append(y)

print("less gooo")

for x in p:
    for y in range(17, 449, 2):
        if x % y == 0:
            s = 0
            break
        else:
            s = x
    if s != 0:
        z.append(s)
print(sum(z))