# Weekly report asset table for TWAG

## 1 - Update Sources
Within the file `Sources.xlsx` update the table with the following information:
    * Name : Name of the asset
    * Source : Webpage used for scraping 
    * Init : Level of the asset as of January 1st of the current year
    * Type : wheher it is a rate (in percent) or other (pt, USD, ...)

## 2 - Run python file
Launch the `get_data_for_nl.py` script after every market closed for the week (a good idea is to do it on Saturdays or Sundays).
The results will be saved in `results.xlsx` in the save directory. // This part should be impoved later //

## 3 - Update main Excel
Copy/paste the table in `results.xlsx` at the good place in `AP_Table_PARSE.xlsm`. 
Then run the relevant macros using "PRE-FINAL CHECK" and "FINAL COPY".

## 4 - Take the Sceenshot
Follow the instruction in `AP_Table_PARSE.xlsm` to do so.
