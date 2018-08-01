import string
with open('names.txt', 'r') as myfile:
    data = myfile.read().replace('"', '')

print(data)
p = data.split(",")
print(p)
print(len(p))

ww = 1
s = 0
for x in range(0, len(p)):
    if p[x] == "COLIN":
        print(p[x])
        print(x)

#         for y in p[x]:
# #             w = y.lower()
# #             print(w)
# #             print(string.ascii_lowercase.index(w) + 1)
# #             s += string.ascii_lowercase.index(w) + 1
# #         s *= x + 1
# #         print(x+1)
# # print(s)