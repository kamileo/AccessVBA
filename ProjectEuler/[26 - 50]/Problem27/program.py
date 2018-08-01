def is_prime(n):
    if n == 1 or n == 4 or n <= 0:
        return False
    for i in range(3, n):
        if n % i == 0:
            return False
    return True


limita = int(1000)
limitb = int(1001)

counter = 0
for a in range(-999, 1000):
    for b in range(-1000, 1001):
        cnt = 0
        n = 0
        pn = n**2 + a*n + b
        while is_prime(pn):
            n += 1
            pn = n ** 2 + a * n + b
            cnt += 1
        if cnt > counter:
            counter = cnt
            amax = a
            bmax = b

print(counter)
print(amax)
print(bmax)