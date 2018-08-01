x = 1001
y = 1001

x0 = x + 2
y0 = y + 2

M = [[0 for x0 in range(x0)] for y in range(y0)]
W = [[0 for x0 in range(x0)] for y in range(y0)]

a = int((x+1)/2)
b = int((y+1)/2)

M[a][b] = 1
W[a][b] = 1

a = a
b = b + 1

M[a][b] = 2
W[a][b] = 1

for i in range(3, (x*y+1)):
    try:
        if W[a-1][b] == 0 and W[a][b+1] == 0 and W[a+1][b] == 0 and W[a][b-1] == 1:
            a += 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a-1][b] == 1 and W[a][b+1] == 0 and W[a+1][b] == 0 and W[a][b-1] == 0:
            b -= 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a-1][b] == 1 and W[a][b+1] == 1 and W[a+1][b] == 0 and W[a][b-1] == 0:
            b -= 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a-1][b] == 0 and W[a][b+1] == 1 and W[a+1][b] == 0 and W[a][b-1] == 0:
            a -= 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a - 1][b] == 0 and W[a][b + 1] == 1 and W[a + 1][b] == 1 and W[a][b - 1] == 0:
            a -= 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a - 1][b] == 0 and W[a][b + 1] == 0 and W[a + 1][b] == 1 and W[a][b - 1] == 0:
            b += 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a - 1][b] == 0 and W[a][b + 1] == 0 and W[a + 1][b] == 1 and W[a][b - 1] == 1:
            b += 1
            W[a][b] = 1
            M[a][b] = i
        elif W[a - 1][b] == 1 and W[a][b + 1] == 0 and W[a + 1][b] == 0 and W[a][b - 1] == 1:
            a += 1
            W[a][b] = 1
            M[a][b] = i

    except IndexError:
        continue
sumak = 0
for i in range(0, x+1):
    sumak += M[i][i]
    sumak += M[i][x-i]

print(sumak)

#
# for x in M:
#     print (x)

