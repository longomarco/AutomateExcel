# Automate Excel with Python
Python script to aggregate a series of csv files, extracted from a database, into a single excel file.

In this scenario, I've implied I've extracted 12 csv files from a database, each containing monthly sales data. 
To get started, therefore, we first need to aggregate all the data into a single file, an operation that can be tedious and time-consuming (especially if we have to work with many more files than the 12 considered in this example). Unless, of course, we can automate the process.

The script I have written basically does this: first it goes and searches for files that have been conveniently saved in the SalesData folder, after which it aggregates them and thanks to the capabilities offered by the 'pandas' library creates pivot tables to analyze the data at hand. Finally it creates an excel file with the aggregated data and spreadsheets for each pivot table.

This whole process is done automatically by running the AutoExcel.bat file.

The project could be further expanded by considering several possibilities: the script could ask the user if he or she is interested in emailing the said excel file, as well as automatically create charts to attach to the report; the possibilities are many.
