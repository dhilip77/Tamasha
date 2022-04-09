# -*- coding: utf-8 -*-
"""
Created on Sat Apr  9 11:50:12 2022

@author: DhilipKumarTG
"""

import sys
import os

''' Read input from STDIN. Print your output to STDOUT '''
# Use input() to read input from STDIN and use print to write your output to STDOUT
Ltime =input()
Ttime =input()

Lhour =int(Ltime[0:2])
Lmin =int(Ltime[3:])

Thour =int(Ttime[0:2])
Tmin =int(Ttime[3:])
if((Lhour,Thour >= 00) and (Lhour,Thour >= 23) and (Lmin,Tmin >= 00) and (Lmin,Tmin <=59)):
    Tot_min = int(Lmin + Tmin)
    Tot_hour = int(Lhour + Thour)
    print(Tot_hour)
    if(Tot_min > 59):
        Tot_min -=60
        Tot_hour +=1
        if(Tot_hour >= 24):
            Tot_hour-=24
            print("inside if",Tot_hour)
            if(Tot_hour < 10 and Tot_min < 10):
                print(f'{Tot_hour:02d}',f'{Tot_min:02d}')
            elif(Tot_hour < 10 and Tot_min >=10):
                print(f'{Tot_hour:02d}',Tot_min)
            elif(Tot_hour >= 10 and Tot_min <10):
            		print(Tot_hour,f'{Tot_min:02d}')
            elif(Tot_hour >= 10 and Tot_min >= 10):
            		print(Tot_hour,Tot_min)
        else:
            print("Total hour less 24 at IF",Tot_hour)
            if(Tot_hour < 10 and Tot_min < 10):
                print(f'{Tot_hour:02d}',f'{Tot_min:02d}')
            elif(Tot_hour < 10 and Tot_min >=10):
                print(f'{Tot_hour:02d}',Tot_min)
            elif(Tot_hour >= 10 and Tot_min <10):
            		print(Tot_hour,f'{Tot_min:02d}')
            elif(Tot_hour >= 10 and Tot_min >= 10):
            		print(Tot_hour,Tot_min)
            
    else:
        Tot_min = Tot_min
        Tot_hour = Tot_hour
        if(Tot_hour >= 24):
            Tot_hour-=24
            print("inside else",Tot_hour)
            if(Tot_hour < 10 and Tot_min < 10):
            		print(f'{Tot_hour:02d}',f'{Tot_min:02d}')
            elif(Tot_hour < 10 and Tot_min >=10):
            		print(f'{Tot_hour:02d}',Tot_min)
            elif(Tot_hour >= 10 and Tot_min <10):
            		print(Tot_hour,f'{Tot_min:02d}')
            elif(Tot_hour >= 10 and Tot_min >= 10):
            		print(Tot_hour,Tot_min)
        else:
            print("Total hour less 24 at Else",Tot_hour)
            if(Tot_hour < 10 and Tot_min < 10):
                print(f'{Tot_hour:02d}',f'{Tot_min:02d}')
            elif(Tot_hour < 10 and Tot_min >=10):
                print(f'{Tot_hour:02d}',Tot_min)
            elif(Tot_hour >= 10 and Tot_min <10):
            		print(Tot_hour,f'{Tot_min:02d}')
            elif(Tot_hour >= 10 and Tot_min >= 10):
            		print(Tot_hour,Tot_min)