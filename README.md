# Plot2sheet
Copyright (c) 2022 Abdelfettah Hadij El Houati

*Plot2sheet* is a small python library that exports matplotlib generated plots to spreadsheet files containing both a chart and the raw datapoints. 

## Installation
TBD

## Dependencies
* matplotlib
* xlsxwriter

## Usage

```python
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
```
## Features
Exports all lines within and Axes object including some formating, namely:
* line color
* line style
* line labels
* marker 
* marker color
* marker edge color

## Limitations

It only supports line graphs. 
Only one axes at a time can be exported.

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

See LICENSE file.