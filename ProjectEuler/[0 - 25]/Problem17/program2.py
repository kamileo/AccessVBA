ones = {"1":"One","2":"Two","3":"Three","4":"Four","5":"Five","6":"Six", "7":"Seven","8":"Eight","9":"Nine","10":"Ten","11":"Eleven","12":"Twelve","13":"Thirteen","14":"Fourteen","15":"Fifteen","16":"Sixteen", "17":"Seventeen","18":"Eighteen","19":"Nineteen"}
tens = {"2":"Twenty","3":"Thirty","4":"forty","5":"Fifty","6":"Sixty", "7":"Seventy","8":"Eighty","9":"Ninety"}

suma = 0
limit = 999
for x in range(1, limit + 1):
    if x < 20:
        suma += len(ones.get(str(x)))
    elif 20 <= x < 100:
        x = str(x)
        for i in range(0, 2):
            if i == 0:
                suma += len(tens.get(x[i]))
            if i == 1 and int(x[i]) > 0:
                suma += len(ones.get(x[i]))
    elif 100 <= x < 1000:
        x = str(x)
        for i in range(0, 3):
            if 0 < int(x[1:]) < 20:
                print(ones.get(str(int(x[1:]))))
                suma += len(ones.get(str(int(x[1:]))))
                suma += len(ones.get(x[0])) + 7
                if int(x[1]) > 0 or int(x[2]) > 0:
                    suma += 3
                break
            elif i == 1 and int(x[i]) > 0:
                print(tens.get(x[i]))
                suma += len(tens.get(x[i]))
            elif i == 2 and int(x[i]) > 0:
                print(ones.get(x[i]))
                suma += len(ones.get(x[i]))
            if i == 0:
                print(ones.get(x[i]))
                suma += len(ones.get(x[i])) + 7
                if int(x[1]) > 0 or int(x[2]) > 0:
                    suma += 3

print(suma)