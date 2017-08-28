#### 
# work for Dr. Cohen, 5_30, making surfaces and shit
####

import numpy as np
import scipy.interpolate as interp
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

## global variables
mu_B = list()
mu_E = list()
expRec = list()
prbDef = list()

inFile = open('5_30_work.csv', 'r')
for line in inFile:

    temp = line.strip().split(",")
    mu_B.append(float(temp[0]))
    mu_E.append(float(temp[1]))
    expRec.append(float(temp[2]))
    prbDef.append(float(temp[3]))

plotx, ploty = np.meshgrid(np.linspace(np.min(mu_E),np.max(mu_E),10), \
                           np.linspace(np.min(mu_B),np.max(mu_B),10))

plotz = interp.griddata((mu_E,mu_B),expRec,(plotx,ploty),method="linear")

fig = plt.figure()
ax = fig.add_subplot(111,projection='3d')
ax.plot_surface(plotx,ploty,plotz,cstride=1,rstride=1,cmap='hot')







