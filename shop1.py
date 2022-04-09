# -*- coding: utf-8 -*-
"""
Created on Sat Apr  9 12:22:04 2022

@author: DhilipKumarTG
"""

import sys
import os

''' Read input from STDIN. Print your output to STDOUT '''
# Use input() to read input from STDIN and use print to write your output to STDOUT
# Write code here
Num_T_Case = int(input())
Price_Shop = []
Tot_val = []

for c in range(Num_T_Case):
    if(Num_T_Case>=1 and Num_T_Case <=10):
        Count = Num_T_Case-1
        temp = int()
        Num_N_Buy = int(input())
        Num_G_Shop = int(input())
        Price_Shop = str(input())
        if(Num_G_Shop>=1 and Num_N_Buy<=Num_G_Shop and Num_N_Buy>= 1):
            Price_Shop = Price_Shop.split(' ')
            Price_Shop = list(map(int, Price_Shop))
            if(len(Price_Shop)==Num_G_Shop):
                #print(len(Price_Shop))
                #for x in range(len(Price_Shop)):
                    #if(int(Price_Shop[x]) <= 0 or int(Price_Shop[x]) <= 10000000):
                        #print(Price_Shop)
                    Price_Shop.sort()
                    #print("Sorted",Price_Shop)
                    for y in range(0,Num_N_Buy):
                        #print("In range",Price_Shop[y])
                        if(int(Price_Shop[y]) >= 0 or int(Price_Shop[y]) <= 10000000):
                            temp += int(Price_Shop[y])
                            #for y in range(len(Price_Shop)):
                                #temp += int(Price_Shop[y])
                        #print("my temp",temp)
    Tot_val.append(temp)
                
#print(Tot_val)
for v in range(len(Tot_val)):
    print(Tot_val[v])
