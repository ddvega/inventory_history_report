"""
 Program:      Holiday Ordering Tool
 Author:       David Vega
 Date:         9/8/19
 Description:  The purpose of this program is to clean and create usable excel
               sheets out of data extracted from a pdf file converted using PDF
               converter. This program is specifically designed for pdf
               reports generated by company X. This prram is intended to work
               exclusively with reports from company X.
"""

from source.clean import cleanData
from source.reports import makeReport

""" Start message and input request """

response = int(input
               ('\nSALES FORECASTING TOOL \n'
                '\nThis tool creates a report from your history store order '
                '\nguide. Start with STEP 1. Once this step is completed, '
                '\na .xls called holidayformatted will be created. Open this '
                '\nfile and make sure that there are no empty cells or cells '
                '\nwith more than one value. It is rare for there to be no '
                '\nerrors. You will most likely find them in bananas, tomatoes '
                '\nand spinach. Once you have fixed errors, save the file and '
                '\nrun STEP 2 to create a report.\n'
                '\n[1] Fix errors'
                '\n[2] Generate report'
                '\n[3] Exit'
                '\n-->'))

if response == 1:
    cleanData()

elif response == 2:
    makeReport()

else:
    print("Have a great day :)")
