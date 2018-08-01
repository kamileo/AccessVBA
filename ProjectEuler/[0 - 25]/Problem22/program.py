import string
with open('names.txt', 'r') as myfile:
    data = myfile.read().replace('"', '')

p = data.split(",")
p.sort()
print(p)

s = []
for x in range(0, len(p)):
    s.append(0)
    for y in p[x]:
        w = y.lower()
        s[x] += string.ascii_lowercase.index(w) + 1
    s[x] *= (x + 1)
print(s)
print(sum(s))