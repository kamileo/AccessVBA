limit = 1000000
count = 0
maks = 0
for x in range(limit, 4, -1):
    y = x
    count = 0
    while y > 1:
        if y % 2 == 0:
            y = y / 2
            count += 1
        else:
            y = 3 * y + 1
            count += 1
    if count > maks:
        maks = count
        maksx = x

print(maksx)