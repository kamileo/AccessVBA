a = []
for i in range(3, 1000000):
    facto = 1
    sumcom = 0
    for j in str(i):
        facto = 1
        for k in range(1, int(j)+1):
            facto *= k
        sumcom += facto
    if i == sumcom:
        a.append(i)
print(a)