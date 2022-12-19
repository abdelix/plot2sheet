#Copyright (c) 2022 Abdelfettah Hadij El Houati

import matplotlib.pyplot as plt
from matplotlib.colors import to_hex
import numpy as np
import xlsxwriter



_linestyle_map={'-'      :'solid',
               'solid'  :'solid',
               ':'      :'round_dot',
               'dotted' :'round_dot',
               '--'     :'dash',
               'dashed' :'dash',
               '-.'     : 'dash_dot',
               'dahshdot' :'dash_dot'} 

_marker_map={   'None' :'none',
               's'    :'square',
               'd'    :'diamond',
               'v'    :'triangle',
               'x'    :'x',
               '*'    :'star',
               '-'    : 'short_dash',
               '|'    :'long_dash',
               'o'    : 'circle',
               '+'     :'plus'} 

def save_axes_to_xlsx(ax:plt.Axes,xlsx_filepath:str):
    """Saves the matplotlib axes ax to a XLSX spreadsheet file located

    Args:
        ax: matplotlib Axes to save
        xlsx_filepath: destination xlsx file path.
        
    """

    workbook = xlsxwriter.Workbook(xlsx_filepath)
    chartsheet = workbook.add_chartsheet('Figure alone')

    ws_withchart= workbook.add_worksheet('Figure in sheet')
    chart1      = workbook.add_chart({'type': 'line',
                                    'subtype':'straight_with_markers'})
    chart2     = workbook.add_chart({'type': 'line',
                                    'subtype':'straight_with_markers'})
    ws_withchart.insert_chart('A1',chart2)
    # Configure the chart.

    chartsheet.set_chart(chart1)

    chartsheet.set_header('main graph')


    for idx,line in enumerate(ax.get_lines()):
        ws=workbook.add_worksheet('line_'+str(idx)+'_raw_data')
        ws.write('A1','xdata')
        ws.write('B1','ydata')
        ws.write_column('A2',line.get_xdata())
        ws.write_column('B2',line.get_ydata())
        
        length=len(line.get_xdata())
        
        dash_type=_linestyle_map.get(line.get_linestyle(),'-')
        
        serires_config={
        'name':       line.get_label(),
        'categories': [ws.name, 1, 0, length, 0],
        'values':     [ws.name, 1, 1, length, 1],
        'line':   {'color': to_hex(line.get_color()),
                'dash_type':dash_type},
        'marker':  {'type':_marker_map.get(line.get_marker(),'None'),
                    'size':line.get_markersize(),
                    'border': {'color': to_hex(line.get_markeredgecolor()),
                            'width':line.get_markeredgewidth()},
                    'fill':   {'color': to_hex(line.get_markerfacecolor())},},
    }
        chart1.add_series(serires_config)
        chart2.add_series(serires_config)
        

    for chart in [chart1,chart2]:
        chart.set_title ({'name': ax.get_title()})
        chart.set_x_axis({'name': ax.get_xlabel(),
                        'text_axis': False,
                        'date_axis': False,
                        'major_gridlines': {
                                            'visible': True,
                                            },
                        'minor_gridlines': {
                                            'visible': True,
                                            },
                        'major_tick_mark': 'inside',
                        'minor_tick_mark': 'inside'},
                        )
        chart.set_y_axis({'name': ax.get_ylabel(),
                        'major_gridlines': {
                                        'visible': True,
                                        },
                        'mminor_gridlines': {
                                        'visible': True,
                                        },
                        'major_tick_mark': 'inside',
                        'minor_tick_mark': 'inside'})    
        
    workbook.close()

