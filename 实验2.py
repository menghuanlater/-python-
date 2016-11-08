#!/usr/bin/python
# -*-coding:utf-8 -*-
#大整数加法的实现（stack实现)
class Stack:
	def __init__(self):
		self.items = []
	def push(self,item):
		self.items.append(item)
	def pop(self):
		return self.items.pop()
	def peek(self):
		return self.items[len(self.items) - 1]
	def isEmpty(self):
		return len(self.items) == 0
	def size(self):
		return len(self.items)
def main():
	number1 = Stack()
	number2 = Stack()
	outcome = Stack()
	string1 = str(input("请输入第一个加数:"))
	string2 = str(input("请输入第二个加数:"))
	
	for i in range(0,len(string1)):
		number1.push(string1[i])
		number2.push(string2[i])
	
	cin = 0
	for i in range(0,len(string1)):
		temp1 = ord(number1.pop())
		temp2 = ord(number2.pop())
		mod = (temp1 + temp2 - 2*ord('0') +cin)%10
		cin = int((temp1 + temp2 - 2*ord('0') +cin)/10)
		outcome.push(mod)
	if(cin!=0):
		outcome.push(cin)
	print("%s与%s相加的结果是:"%(string1,string2),end='')
	while(not outcome.isEmpty()):
		print(outcome.pop(),end='')
	print("") 	
if __name__ == "__main__":
	main()
