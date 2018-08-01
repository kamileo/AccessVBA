limit = 600851475143

s = 0
for x in range(2, limit):
    for y in range(2, x):
        if x % y == 0:
            s = 0
            break
        else:
            s = x
    if s != 0:
        if limit % s == 0:
            print(s)