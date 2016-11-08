#!/usr/bin/python
# -*-coding:utf-8 -*-
#日期计算

def LeapYear(year):
	if(year%4==0 and year%100!=0) or (year%400==0):
		return 1
	else:
		return 0

def main():
	MONTH = [31,28,31,30,31,30,31,31,30,31,30,31]
	year = int(input("请输入年份："))
	month = int(input("请输入月份："))
	day = int(input("请输入日号："))

	if(LeapYear(year)):
		MONTH[1]+=1
	out = 0
	for i in range(0,month-1):
		out += MONTH[i]
	out += day
	print(year,"年",month,"月",day,"日是这一年的第",out,"天")
if __name__ == "__main__":
	main()
