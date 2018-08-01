# We shall say that an n-digit number is pandigital if it makes use of all the digits 1 to n exactly once; for example, the 5-digit number, 15234, is 1 through 5 pandigital.
# The product 7254 is unusual, as the identity, 39 × 186 = 7254, containing multiplicand, multiplier, and product is 1 through 9 pandigital.
# Find the sum of all products whose multiplicand/multiplier/product identity can be written as a 1 through 9 pandigital.
# HINT: Some products can be obtained in more than one way so be sure to only include it once in your sum.
a = []
for i in range(1, 9999):
    for j in range(1, 9999):
        x = i * j
        word = str(x) + str(i) + str(j)
        if len(set(word)) == len(word) == 9 and "0" not in word:
            if x in a:
                break
            else:
                a.append(x)
print(a)
print(sum(a))
