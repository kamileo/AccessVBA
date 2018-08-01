c = 0
s = 0
x = 1

limit = 10000

while c < limit:
    x += 1
    for y in range(2, x):
        if x % y == 0:
            s = 0
            break
        else:
            s = x
    if s != 0:
        c += 1
        no = s

print(no)
