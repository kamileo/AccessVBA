ones = {"0":"zero","1":"One","2":"Two","3":"Three","4":"Four","5":"Five","6":"Six", "7":"Seven","8":"Eight","9":"Nine","10":"Ten","11":"Eleven","12":"Twelve","13":"Thirteen","14":"Fourteen","15":"Fifteen","16":"Sixteen", "17":"Seventeen","18":"Eighteen","19":"Nineteen"}
tens = {"2":"Twenty","3":"Thirty","4":"Fourty","5":"Fifty","6":"Sixty", "7":"Seventy","8":"Eighty","9":"Ninety"}

sumnum = 0
for i in range(1, 1000):
    num = str(i)
    if i < 20:
        for x in num:
            sumnum += len(ones.get(str(i)))
    elif 19 < i < 100:
        num.split()
        for j in range(len(num) - 1, -1, -1):
            if j == 1 and num[j] > 0:
                sumnum += len(ones.get(str(num[j])))
            if j == 0 and num[j] > 0:
                sumnum += len(tens.get(str(num[j])))
    else:
        num.split()
        for j in range(len(num) - 1, -1, -1):
            if j == 2 and int(num[j]) > 0:
                sumnum += len(ones.get(str(num[j])))
            if j == 1 and int(num[j]) > 0:
                sumnum += len(tens.get(str(num[j])))
            if j == 0 and int(num[j]) > 0:
                print(num[j])
                sumnum += len(ones.get(str(num[j]))) + 7
print(sumnum)