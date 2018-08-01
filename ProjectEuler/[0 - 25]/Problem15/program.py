def lattice_paths(x, y):
    z=[]
    n = x + y
    k = y
    ns = 1
    ks = 1
    nks = 1
    for i in range(1, n+1):
        ns *= i
    for j in range(1, k+1):
        ks *= j
    for g in range(1, (n-k+1)):
        nks *= g
    len = ns/(nks*ks)
    z.append(ns)
    z.append(ks)
    z.append(nks)
    return len

print(lattice_paths(20, 20))