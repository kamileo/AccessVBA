def is_palindrome(N):
    # base 10 palindrome
    s = str(N)
    ln = len(s)
    if ln == 1:
        return True
    a = s[:int(ln / 2)]
    b = s[-int(ln / 2):]
    if a == b[::-1]:
        return True
    else:
        return False


def binary_maker(N):
    rest = 2
    bina = ""
    while rest > 0:
        d = N % 2
        bina = str(d) + bina
        N = N - d
        N = int(N / 2)
        rest = N
    return bina


z = []
for x in range(1, 1000000):
    if is_palindrome(x) == is_palindrome(binary_maker(x)) == True:
        z.append(x)
        # print(x)
        # print(binary_maker(x))
print(sum(z))
print(z)
print(len(z))






