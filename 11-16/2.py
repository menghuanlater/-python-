#约瑟夫环
print("开始先采用顺时针报数")
n = int(input("输入总人数:"))
a = int(input("开始的号码:"))
m = int(input("第几个出局:"))

nameList = []
for i in range(0,n):
	nameList.append(i+1)

def Function(nameList,startWith,count,direction):
	global n
	del_index = 0
	if(n==1):
		print(nameList[0])
		return
	if(direction==1):
		del_index = (startWith + count)%n
		direction =0
	else:
		del_index = (startWith + n -count)%n
		direction =1
	print(nameList[del_index])
	for i in range(del_index,n-1):
		nameList[i] = nameList[i+1]
	if(del_index==n-1):
		startWith = n-2
	else:
		startWith = del_index
	n-=1
	Function(nameList,startWith,count,direction)
	
	
direction = 1		
Function(nameList,a-1,m,direction)
