# lines = ""
# for i in range(4):
#     lines+=input()+"\n"
# print(lines)
# p = lines.split()
# print(p)

p = [3, 7, 4, 2, 4, 6, 8, 5, 9, 3, 1, 2, 3, 4]
#lewa strona - zjazd
for i in range(0, 5):
    if i == 0:
        print(p[i])
        x = i
    else:
        print(p[x+i])
        x = x + i

#lewa strona - zjazd
for i in range(0, 5):
    if i == 0:
        print(p[1])
        x = i + 1
    else:
        print(p[x+i])
        x = x + i + 1
