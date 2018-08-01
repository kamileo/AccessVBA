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
        if x > y:
            return [x, y]
        elif x < y:
            return [y, x]
    else:
        return 0


p = []
for i in range(1, 10000):
    if amicable(i) != 0:
        p.append(amicable(i))
s = 0
b_set = set(map(tuple, p))
print(b_set)
b = list(b_set)
print(b)
for x in b:
    s += sum(x)
print(s)