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


M = SieveOfEratosthenes(1000000)
M2 = [i for i in M if i >= 10]
w = []

for x in M2:
    check = False
    if x in M:
        check = True
        y = x
        z = x
        while len(str(y)) > 1:
            y = str(y)[:-1]
            if int(y) not in M:
                check = False
                break
        while len(str(z)) > 1:
            z = str(z)[1:]
            if int(z) not in M:
                check = False
                break
    if check is True:
        w.append(x)
    if len(w) == 11:
        break
print(w)
print(sum(w))