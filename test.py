import jin as j
import global_value as g
import os
import sys
import time

def ddd():
    print("KKKKKKKKKKKKKKKKKKKKK")

g.CUR_DIR=os.getcwd()
j.get_gamen_size()
#グローバル変数にはg.を付ける


##j.msg("hooho")


table=list(range(2))
table[0]="apple"
table[1]="orange"

getdata=j.popupmenu2(table)