target = 50
coins = [1, 2, 5, 10, 20, 50]
ways = [1] + [0]*target  # wypełniamy tablicę zerami do targetu

for coin in coins:
    for i in range(coin, target+1):
        ways[i] += ways[i-coin]

print("Ways to make change =", ways[target])
print(ways)

