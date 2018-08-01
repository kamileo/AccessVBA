import fractions


def digit_comp(nba, nbb):
    a1 = str(nba)[:1]
    a2 = str(nba)[-1:]
    b1 = str(nbb)[:1]
    b2 = str(nbb)[-1:]
    if nba != nbb:
        if a1 == b2 and int(b1) != 0:
            return int(a2) / int(b1)
        if a2 == b1 and int(b2) != 0:
            return int(a1) / int(b2)



a = []
for i in range(10, 101):
    for j in range(10, 101):
        if int(i)/int(j) == digit_comp(i, j) and i/j < 1:
            a.append([i,j])
print(a)
# b = []
# for i in range(0, len(a)):
#     b.append([a[i][0] * a[(i + 1) % 4][1] * a[(i + 2) % 4][1] * a[(i + 3) % 4][1], a[i][1] * a[(i + 1) % 4][1] * a[(i + 2) % 4][1] * a[(i + 3) % 4][1]])

mtop = 1
mbot = 1

for x in a:
        mtop *= x[0]
        mbot *= x[1]

print(mtop)
print(mbot)

for i in range(200,1,-1):
    while mtop % i == 0 and mbot % i == 0:
        mtop /= i
        mbot /= i

print(mtop)
print(mbot)