p = []
le = 0
while True:
    line = input()
    if line:
        h = list(map(int, line.split()))
        p.append(h)
        le += 1
    else:
        break
for x in range(0, le):
    while len(p[x]) < le:
        p[x].append(0)

for x in p:
    print(x)
print("")

for i in range(le - 2, -1, -1):
    for j in range(0, le):
        try:
            if p[i + 1][j] > p[i + 1][j + 1]:
                p[i][j] += p[i + 1][j]
            else:
                p[i][j] += p[i + 1][j + 1]
        except IndexError:
            p[i + 1][j] += p[i + 1][j]

for x in p:
    print(x)
print(" ")
print(max(p[0]))
