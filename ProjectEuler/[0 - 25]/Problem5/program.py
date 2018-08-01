s = 1
z = 20
while s > 0:
    z += 1
    s = 0
    for x in range(1, 21):
        if z % x > 0:
            s = 1
            break
        else:
            s = 0
print(z)