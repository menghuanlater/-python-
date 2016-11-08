#!/usr/bin/python
# -*-coding:utf-8 -*-
#重复密钥加密法的实现
#采用倒插队列
class Queue:
	def __init__(self):
		self.items = []
	def isEmpty(self):
		return self.items == []
	def Enqueue(self,item):
		self.items.insert(0,item)
	def Dequeue(self):
		return self.items.pop()
	def size(self):
		return len(self.items)

def main():
		

if __name__ == "__main__":
	main()

