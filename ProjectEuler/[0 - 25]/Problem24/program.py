limit = 9
count = 0

for x in range(0, limit + 1):
    for y in range(0, limit + 1):
        if y == x:
            continue
        for z in range(0, limit + 1):
            if z == y or z == x:
                continue
            for p in range (0, limit + 1):
                if p == y or p == x or p == z:
                    continue
                for q in range (0, limit + 1):
                    if q == y or q == x or q == z or q == p:
                        continue
                    for r in range (0, limit + 1):
                        if r == y or r == x or r == z or r == p or r == q:
                            continue
                        for s in range (0, limit + 1):
                            if s == y or s == x or s == z or s == p or s == q or s == r:
                                continue
                            for t in range (0, limit + 1):
                                if t == y or t == x or t == z or t == p or t == q or t == r or t == s:
                                    continue
                                for u in range (0, limit + 1):
                                    if u == y or u == x or u == z or u == p or u == q or u == r or u == s or u == t:
                                        continue
                                    for v in range (0, limit + 1):
                                        if v == y or v == x or v == z or v == p or v == q or v == r or v == s or v == t or v == u:
                                            continue
                                        count += 1
                                        if count == 1000000:
                                            print(x,y,z,p,q,r,s,t,u,v)