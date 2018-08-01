p = []
s = 0
for i in range(1, 10000000):
    numba = str(i)
    s = 0
    for x in numba:
        s += int(x) ** 5
    if i == s:
        p.append(i)
print(p)
print(sum(p))