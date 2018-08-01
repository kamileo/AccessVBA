def SieveOfEratosthenes(n):
    prime = [True for i in range(n + 1)]
    p = 2
    a = []
    while p * p <= n:
        if prime[p]:
            # Update all multiples of p
            for i in range(p * 2, n + 1, p):
                prime[i] = False
        p += 1
    for p in range(2, n):
        if prime[p]:
            a.append(p)
    return a


def circulate(j):
    a = []
    j = str(j)
    for i in range(0, len(j)):
        j = j[-1:] + j[:-1]
        a.append(int(j))
    return a


n = 1000000
M = SieveOfEratosthenes(n)
d = []
for x in M:
    p = circulate(x)
    setu = True
    # for i in range(0, len(p)):
    #     p[i] = int(''.join(list(p[i])))
    for y in p:
        if y not in M:
            setu = False
            break
    if setu is True:
        d.append(x)
print(d)
print(len(d))