p = []
ano = 0
g = 0
count = 2
while g < 1000:
    ano = sum(p)
    count += 1
    p[0] = p[1]
    p[1] = ano
    g = len(str(ano))
print(g)
print(count)