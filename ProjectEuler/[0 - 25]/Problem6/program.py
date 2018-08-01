Sx = 0
Sqy = 0
for x in range(1, 101):
    Sx += x

Sqx = Sx**2

for y in range(1, 101):
    Sqy += y**2

print(Sqx - Sqy)