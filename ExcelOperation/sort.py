
l = [9,8,7,6,4]
for i in range(5-1):
    for j in range(5-i-1):
        a = l[j]
        b = l[j+1]
        if a > b : 
            l[j] = b
            l[j+1] = a
print(l)