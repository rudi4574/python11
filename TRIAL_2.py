# -*- coding: utf-8 -*-
"""
Created on Fri Jan  3 06:55:33 2020

@author: Rudresh
"""

import pandas as pd
import os
import numpy as np
import scipy.stats as ss

def open_excel_files(excel_files):
    number_of_excel_files = len(excel_files)
    for i in range(number_of_excel_files):
        x = None
        try:
            x = pd.read_excel('workbooks/'+str(excel_files[i]))
            a = np.mean(x)
            print(a)
# ##### Median

            b = np.median(x)
            print(b)
# ##### Mode
            c = ss.mode(x)
            print(c)
# ##### Kurtosis

            d = ss.kurtosis(x)
            print(d)

# ##### Skewness
            e = ss.skew(x)
            print(e)

# ##### Standard deviation
            f = np.std(x)
            print(f)

# ##### Variance
            g = np.var(x)
            print(g)
# ##### Standard error of measurement
            h = ss.sem(x)
            print(h)

# ##### Maximum
            i = np.amax(x)
            print(float(i))
# ##### Minimum
            j = np.amin(x)
            print(j)
# ##### Sum

            k = np.sum(x)
            print(k)
# ##### Quantiles

# Quantile 25%

            m = np.percentile(x,25)
            print(m)
# Quantile 50% (Median)

            n = np.percentile(x,50)
            print(n)

# Quantile 75%
            o = np.percentile(x,75)
            print(o)
            

            
        except Exception as e:
            print("Exception occured while opening file => ", excel_files[i], e)
    
        all_data = pd.DataFrame({'mean':[a],'Mediam':[b],'Mode':[c],'Kurtosis':[d],'skewness':[e],'Standard deviation':[f],'Variance':[g],'Standard error':[h],'Maximum':[i],'Minimum':[j],'sum':[k],'Quartile25':[m],'Quartile50':[n],'Quartile75':[o]})
        all_data =all_data.append(all_data,ignore_index=True)
        ExcelRR = pd.ExcelWriter('Excelfile5.xlsx',engine="xlsxwriter")
    print(all_data.to_excel(ExcelRR,sheet_name = "ML"))
    ExcelRR.save()
    ExcelRR.close()
    


def files_in_directory():
    return os.listdir('workbooks')


def main():
    COUNT = 172
    excel_files = files_in_directory()
    open_excel_files(excel_files)


if __name__ == "__main__":
    main()
    