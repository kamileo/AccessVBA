# Q = input("podaj")
# print(Q)
# z=[]
# for i in range(1,len(Q)+1,50):
#     z.append(Q[i:i+50])
# print(z)

#no_of_lines = 10
lines = ""
for i in range(100):
    lines+=input()+"\n"

p = lines.split()
print(p)
sum = 0
for i in range(0,100):
    sum += int(p[i])
print(sum)
