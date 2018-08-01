def amicable(x):
    p = []
    q = []
    for a in range(1, x):
        if x % a == 0:
            p.append(a)
    y = sum(p)
    for b in range(1, y):
        if y % b == 0:
            q.append(b)
    z = sum(q)
    if x == z and x != y:
        return [x, y]
    else:
        return 0


p = []
for i in range(1, 500):
    if amicable(i) != 0:
        p.append(amicable(i))

print(p)