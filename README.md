This was a small project that is highly specialized for a specific scenario. If you have somehow come across this then it may not be useful to you at all

## Purpose
This combines the data saved in many different files and consolidate them into one file

1. Data for a specific day is saved in one sheet of one workbook
2. That workbook contains data for all days in a month
3. The workbook excel file is saved in a folder that corresponds to the year
4. All years are saved in the "books" directory
5. A specific table for daily reports needs to be extracted
6. Get this data and save it in new files using the following format 
    - Data from one day is added into a new sheet
    - Data from the next day is appended to that new sheet
    - At the start of a new month, create a new sheet
    - At the start of a new year, create a new workbook
