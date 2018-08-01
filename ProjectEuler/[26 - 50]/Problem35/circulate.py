def circulate(j):
    a = []
    j = str(j)
    for i in range(0, len(j)):
        j = j[-1:] + j[:-1]
        a.append(int(j))
    return a


print(circulate(1234))

# test = "python"
# print(test[:-1])
# print(test[-1:])
# print(test[-1:] + test[:-1] )