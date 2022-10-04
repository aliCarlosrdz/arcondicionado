#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Oct  1 16:34:19 2022

@author: alisson
"""

import pandas as pd
import seaborn as sns; sns.set_theme()


file = 'RESULTADO.xlsx'
clima = pd.read_excel(file)
data = pd.pivot_table(clima, index='Hour',columns = ['Month','Day'], values='Carga Térm. de Ren. [kW]')

sns.set(rc={'figure.figsize':(18,8)})
ax = sns.heatmap(data,cmap='coolwarm',center = 0,xticklabels=12,yticklabels=True)
ax.set_yticklabels(ax.get_yticklabels(),rotation=0, fontsize=10)
ax.set_xticklabels(ax.get_xticklabels(),rotation=90, fontsize=10)
figure = ax.get_figure()
figure.savefig('heatmap_6.png',dpi=500,bbox_inches='tight')           
