hi = 2000000

# create a set excluding even numbers
numbers = set(range(3, hi + 1, 2))

for number in range(3, int(hi ** 0.5) + 1):
    if number not in numbers:
        #number must have been removed because it has a prime factor
        continue

    num = number
    while num < hi:
        num += number
        if num in numbers:
            # Remove multiples of prime found
            numbers.remove(num)

print(2 + sum(numbers))

