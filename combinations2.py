#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Wed Nov 23 01:58:19 2016

@author: MacbookRetina
"""
from itertools import combinations

def main():
    list1 = ['cheeseburger','fries','coke','ice cream', 'chicken nuggets']
    list2 = ['small', 'medium', 'large']

    list3 = [(y,x) for x in list1 for y in list2]
             
    #print list3
             
    #output = [map(list, comb) for comb in combinations(list3, 5)]
    
    #print output[0][0][0]
    
    print 'cheeseburger' in list1
        
        
    
main()

