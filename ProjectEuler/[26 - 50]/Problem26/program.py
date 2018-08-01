q = 7
maxx = 1
for q in range(1, 1000):
    for i in range(1, 2000):
        if 10**i % q == 1:
            if i > maxx:
                maxx = i
                numba = q
            break
print(maxx)
print(numba)