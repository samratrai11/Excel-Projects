Overview:

•	Dataset: https://www.kaggle.com/datasets/panwararpit/bike-sales-data

•	Tools used: Microsoft excel.

•	Functions applied =IF statement and Nested IF's.

•	Functions used: Pivot Table, slicer, bar and line charts.


Steps taken:

•	The first step is to remove duplicates by selecting the entire worksheet and then going to the data tab and selecting the remove duplicate option which removes all the duplicate values. 26 duplicate values were found, and 1000 unique values remain.

•	changing the abbreviated names for marital status and gender to their full names. Keyboard shortcut: control + h which brings out the find and replace tab to alter. Use the replace section to replace M with married and select the column option to replace the whole column with married instead of M. Repeated the same process for the gender column as well.

•	Changed the income column to currency by going to the home tab selecting the whole column and then selecting the currency option tab.

•	IF(L2>54,”Old”,IF(L2>=31,”Middle Age”,IF(L2<31,”Adolescent”,”Invalid”))), used the nested IF function to create a new age column for easier referencing. 

•	Created a new sheet called pivot table. go to the sheet containing the data and press control + a to select the entire data then create a pivot table by selecting the pivot table option and create the pivot table in a new worksheet.

•	Created three new sheets each with pivot tables.

•	Then created bar and line charts for every pivot table.


