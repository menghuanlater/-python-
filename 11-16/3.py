#上楼梯
def main():
	count = 0
	n= int(input("How many steps:"))
	count = upStairs(n)
	print("There are",count,"Solutions")

def upStairs(n):
	if(n==1):
		return 1
	elif(n==2):
		return 2
	elif(n==3):
		return 4
	else:
		return 7+upStairs(n-1)+upStairs(n-2)+upStairs(n-3)

main()
		
