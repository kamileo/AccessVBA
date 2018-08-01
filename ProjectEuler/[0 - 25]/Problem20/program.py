limit = 100
sum = 1
sum2 = 0
for x in range(1, limit + 1):
    sum *= (limit + 1 - x)

for y in str(sum):
    sum2 += int(y)
print(sum2)