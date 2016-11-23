#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Tue Nov 22 16:44:50 2016

@author: MacbookRetina
"""
from itertools import combinations

import pandas as pd
import requests
from openpyxl import Workbook
import openpyxl
import os.path

def main():
    wb = openpyxl.load_workbook('FoodCombinations.xlsx')
    sheet_names = wb.get_sheet_names()
    sheet = wb.get_sheet_by_name(sheet_names[0])
    
    list1 = ['cheeseburger','fries','coke','ice cream', 'chicken nuggets']
    list2 = ['small', 'medium', 'large']

    list3 = [(y,x) for x in list1 for y in list2]
             
    output = sum([map(list, combinations(list3, i)) for i in range(len(list3) + 1)], [])
    output = [map(list, comb) for comb in combinations(list3, 5)]
    counter = 1
    for element in output:
        print counter
        duplicates = []
        flag = True
        string = "can I have"
        for tuples in element:
            if tuples[1] not in duplicates:
                duplicates.append(tuples[1])
            else:
                flag = False
                break
            string = string + " " + tuples[0] + " " + tuples[1]

        if flag:
            sheet['A' + str(counter)] = "none"
            sheet['C' + str(counter)] = string
            sheet['B' + str(counter)] = "food order"
            counter = counter + 1
            sheet['A' + str(counter)] = "none"
            sheet['C' + str(counter)] = "confirm"
            sheet['B' + str(counter)] = "complete food order"
            counter = counter + 1
            sheet['A' + str(counter)] = "none"
            sheet['C' + str(counter)] = "cancel"
            sheet['B' + str(counter)] = "cancel"
            counter = counter + 1

    wb.save('FoodCombinations.xlsx')
        
    
           
           

main()