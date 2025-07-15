First you must download the Stock Analysis Workbook.xlsx and know the address it is being written to

Next you will download the Stock Analyzer.py script and change the file path inside the script to the path of the Stock Analysis Workbook.xlsx
It should look something like User/(Name)/Stock Analysis Workbook.xlsx or mini variations. It will be at the top of the script.

Next you will download the Stock Analyzer.bat 

It will contain this

@echo off
"C:\Users\(Name of you)\Documents\Scripts\Stock Analyzer.py" # Edit this portion to the location of your Stock Analyzer.py Script
pause


Everything should be all set now
Just double click the Stock Analyzer.bat file and it will run the program and you will give it a symbol and results will show in your excel file!
