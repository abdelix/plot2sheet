#Copyright (c) 2022 Abdelfettah Hadij El Houati
import matplotlib.pyplot as plt
import numpy as np
import plot2sheet

x=np.linspace(0,5)
fig=plt.figure()
plt.plot(x,5*x+1,label='A',color='b')
plt.plot(x,x**2+1, label='B',marker='s',markeredgecolor='k')
plt.plot(x,x**0.5+6,linestyle=':', label='C')
plt.xlabel('x')
plt.ylabel('y')


plot2sheet.save_axes_to_xlsx(plt.gca(),'test_plot.xlsx')
