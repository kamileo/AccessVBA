def abundant(x):
    p = []
    q = []
    for a in range(1, x):
        if x % a == 0:
            p.append(a)
    y = sum(p)
    return y


z = []
limit = 28123
limith = int(limit/2)
for i in range(1, (limit)):
    if abundant(i) > i:
        z.append(i)
print(z)
print(len(z))





# s = 0
# count = 0
# nu = []
# for i in range(1, limit):
#     s = 1
#     print(i)
#     for j in z:
#         for k in z:
#             if j + k > i:
#                 s = 0
#                 break
#             if j + k == i:
#                 s = 1
#                 break
#             else:
#                 s = 0
#         if s == 0:
#             nu.append(i)
# print(nu)
# print(sum(nu))