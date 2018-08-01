# w, h = 8, 5;
# Matrix = [[0 for x in range(w)] for y in range(h)]
#
# Matrix[0][0] = 1
# Matrix[0][6] = 3 # valid
#
# x, y = 0, 6
# print(Matrix[x][y])
p = []
s = []
lines = ""
for i in range(15):
    g = input()+"\n"
    h = list(map(int, g.split()))
    while len(h) < 15:
        h.append(0)
    p.append(h)

for x in p:
    print(x)
print("")

for i in range(13, -1, -1):
    for j in range(0, 15):
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