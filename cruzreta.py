# -*- coding: utf-8 -*-
"""
Created on Mon Oct 25 21:25:59 2021

@author: robson
"""

import openpyxl
xfile = openpyxl.load_workbook('cruzreta.xlsx')

sheet = xfile.get_sheet_by_name('Sheet1')
#valores REC3D de x, y e z
sheet['Q17'] = '=AVERAGE(O2:O3)'
sheet['R17'] = '=AVERAGE(P2,P4)'
sheet['S17'] = '=AVERAGE(P3,O4)'

#erro de x, y e z
sheet['Q19'] = '=Q18-Q17'
sheet['R19'] = '=R18-R17'
sheet['S19'] = '=S18-S17'

#erro absoluto de x, y e z
sheet['Q20'] = '=ABS(Q19)'
sheet['R20'] = '=ABS(R19)'
sheet['S20'] = '=ABS(S19)'

#média total de x, y e z
sheet['S21'] = '=AVERAGE(Q20:S20)'

#salva arquivo
xfile.save('cruzretaScript.xlsx')