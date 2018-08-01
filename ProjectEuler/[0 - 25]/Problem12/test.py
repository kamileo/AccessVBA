import math

# Method to print the divisors
def printDivisors(n):
    list = []
    z = []
    # List to store half of the divisors
    for i in range (1, int (math.sqrt (n) + 1)):

        if (n % i == 0):

            # Check if divisors are equal
            if (n / i == i):
                #print (i, end=" ")
                list.append (i)
            else:
                # Otherwise print both
                #print (i, end=" ")
                list.append(i)
                list.append (int (n / i))

    # The list will be printed in reverse
    return list

p = printDivisors(100)
print(len(p))

s = []
sum = 0
count = 0
i = 0
while count < 500:
    i += 1
    sum += i
    count = len(printDivisors(sum))
print(sum)