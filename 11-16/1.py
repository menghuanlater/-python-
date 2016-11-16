#亲和数
for i in range(1000,10000):
	sum1 = 0
	sum2 = 0
	for j in range(1,i):
		if(i%j==0):
			sum1 += j
	for j in range(1,sum1):
		if(sum1%j==0):
			sum2+=j
	if(sum2==i):
		print("%d与%d是亲和数"%(i,sum1))
