#!/usr/local/bin/python
# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import numpy as np

width = 0.3

d1 = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
c1 = [0.2, 0.1, 0.3, 0.8, 0.7, 0.4, 0.8, 0.1, 0.5, 0.4, 0.1, 0.2]

d2 = []
[d2.append(item+width) for item in d1]
c2 = [0.7, 0.5, 0.1, 0.1, 0.7, 0.4, 0.8, 0.4, 0.5, 0.4, 0.3, 0.7]

plt.bar(d1,c1, width=0.3,color='b')
#plt.bar(d2,c2, width=0.3,color='r')
#plt.legend(('2'))
plt.axis([1,12+2*width,0,1])
plt.xlabel('Month')
plt.xticks(np.arange(1+width,13+width),("1", "2","3", "4","5", "6","7", "8","9", "10","11", "12"))
plt.ylabel('Scale of the PHR/Month')
#plt.legend(('label1', 'label2'))
title = "Bar chart of the Unit PHR"
plt.title(title)
#plt.savefig('bar1.png')

plt.table(cellText= [['12    3'], [456]], cellColours=['w','w'], \
    rowLabels=[' a',' b'], cellLoc='left', \
    colColours = ['g', 'b'], colLoc = 'left')

plt.show()
