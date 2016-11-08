#!usr/bin/python
# -*-coding:utf-8 -*-
#序列求和
import math

def main():
	a = int(input("请输入a:"))
	n = int(input("请输入n:"))
	l = []
	
	temp = 0
	for i in range(1,n+1):
		l.append(temp + a*math.pow(10,i-1))
		temp = l[i-1]
	
	temp = 0	
	for i in range(0,n):
		temp += l[i]
	print(int(temp))

if __name__ == "__main__":
	main()
