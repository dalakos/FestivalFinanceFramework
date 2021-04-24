# -*- coding: utf-8 -*-
"""
Created on Thu Apr 22 09:35:08 2021

@author: Joanne
"""
import numpy as np
import matplotlib.pyplot as plt

# set seed
np.random.seed(1)

# MODEL INPUT
NUM_SEG = 8 # How many segments to model
SEG_LEN = 15  # Length of a pick-up time segment in minutes
MAX_NUM_SEG = 50 # Number of customers/cars per segment
STD_ARRIVAL_MIN = 5 # standard deviation in minutes for arrival times to target
RES_TIME_MIN = 10 # residence time average in minutes
# stdResTime = 3 # standard deviation in minutes for residence time
MAX_WALKINS = 5

NUM_CUST = NUM_SEG * MAX_NUM_SEG # Total number of customers

carTimeStart = []
carTimeEnd = []

for i in range(NUM_SEG):
    for j in range(MAX_NUM_SEG):
        TMP_TIME = np.random.normal(i*SEG_LEN, STD_ARRIVAL_MIN, 1)
        carTimeStart.append(int(TMP_TIME))
        #carTimeEnd.append(int(TMP_TIME + RES_TIME_MIN + np.random.weibull(.7,1)))
        carTimeEnd.append(int(TMP_TIME + RES_TIME_MIN + 5*np.random.weibull(1,1)))

carStart = np.array(carTimeStart)
carEnd = np.array(carTimeEnd)

histCount = []
relTime = []
for k in range(-SEG_LEN,(NUM_SEG+1)*SEG_LEN):
    TMP_COUNTER=0
    for l in range(NUM_SEG * MAX_NUM_SEG):
        if carStart[l] <= k <= carEnd[l]:
            TMP_COUNTER = TMP_COUNTER + 1

    relTime.append(k)
    histCount.append(TMP_COUNTER)

#End of loop
plt.close('all')
plt.figure()

plt.plot(np.array(relTime),np.array(histCount), 'k', label='sales')
plt.ylabel('# of cars in parking lot')
plt.xlabel('Relative time in minutes')
plt.grid(axis='both')
plt.hlines(np.max(histCount)+MAX_WALKINS,0,NUM_SEG * SEG_LEN,'r', linestyles='--', lw=2)
plt.title('Cars in parking lot as a fxn of time. Max # capacity plan = '
          +str(np.max(histCount) + MAX_WALKINS))

plt.figure()
plt.hist(carEnd-carStart,bins=max(carEnd-carStart)-min(carEnd-carStart)+1)
plt.ylabel('# of counts')
plt.xlabel('Residence time in minutes')
plt.grid(axis='both')
plt.title('Histogram of how long cars spend in parking lot in minutes')
