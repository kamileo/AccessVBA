p = []
for a in range(2, 101):
    for b in range(2, 101):
        p.append(a ** b)
p = list(set(p))
p.sort()
print(len(p))
