months = {1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6: 30, 7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31}

x = 7
rest = 0
z = 0
count = 0
for y in range(1900, 2001):
    if y > 1900:
        if y % 4 == 0:
            months[2] = 29
        else:
            months[2] = 28
    for i in range(1, 13):
        while x < months.get(i):
            print("in year %d, month %d, sunday was on day: %d" % (y, i, x))
            if x == 1 and y > 1900:
                count += 1
            x += 7
        rest = (months.get(i) + z) % 7
        print("days left: ", rest)
        x = -rest + 7
        z = rest
print(count)
