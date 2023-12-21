Coffee Beans Project Overview:

•	Dataset link: https://www.kaggle.com/datasets/saadharoon27/coffee-bean-sales-raw-dataset

•	Tools used: Microsoft Excel.

•	Goal: Data manipulation and visualization.

Summary of cells, sheets, rows, columns, and chart manipulation:

Customer name cell:

•	=XLOOKUP (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

•	Where the lookup_value is the customer ID from the orders sheet, the lookup_array is the customer id from the customer's sheet, the return_array is the customer name from the customer's sheet, if_not _found is left empty with “” sign, match_mode is 0 which is known as the exact match and then at last close the bracket.
Email cell:

•	=XLOOKUP (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

•	Where the lookup_value is the customer ID from the orders sheet, the lookup_array is the customer id from the customer's sheet, the return_array is the email column from the customer's sheet, if_not _found is left empty with “” sign, match_mode is 0 which is known as the exact match and then at last close the bracket.

•	Since some of the email cell values didn’t have any values, we use IF (logical_test, [value_if_true], [value_if_false]) to find the missing emails.

•	The IF (logical_test, [value_if_true], [value_if_false]) is used on the =XLOOKUP (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).

•	It is used like, =IF(XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0,"",XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)). Where the logical_test is XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0)=0, the value_if_true is “” and the value_if_false is XLOOKUP(C2,customers!$A$1:$A$1001,customers!$C$1:$C$1001,,0).
Country cell:

•	=XLOOKUP (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).

•	Where the lookup_value is the customer ID from the orders sheet, the lookup_array is the customer id from the customer's sheet, the return_array is the country’s column from the customer's sheet, if_not _found is left empty with “” sign, match_mode is 0 which is known as the exact match and then at last close the bracket.
Coffee Type:

•	INDEX (array, row_num, [column_num]) and MATCH(lookup_value, lookup_array, [match_type]).

•	Where the array is the whole sheet of the products sheet, then use the match function the lookup_value is the product ID column from the orders sheet, the lookup_array is the product ID column from the products sheet, the match type is 0 which is known as the exact match(note: this match syntax was used for the row value), again use match syntax but this time for the column. For the second match syntax, the lookup_value is the coffee type column from the orders sheet, the lookup_array is only the entire header row columns of the products sheet, the exact match is 0 and then lastly close the bracket twice.

•	=INDEX(products!$A$1:$G$49,MATCH(orders!$D2,products!$A$1:$A$49,0),MATCH(orders!I$1,products!$A$1:$G$1,0)) - this is the entire index and double match function used which I have explained.
Roast type, size and unit price column:

•	Auto-populate these column cells by dragging the coffee-type cell to the right and the values will be filled. 
 
Sales column:
•	To calculate the sales column, a simple multiplication procedure is needed where the unit price cell 2 * quantity cell 2 equals the sales total.
 
 

Coffee type name cell:

•	IF (logical_test, [value_if_true], [value_if_false]) 

•	=IF(I2="Rob","Robusta",IF(I2="Exc","Excelsa",IF(I2="Ara","Arabica",IF(I2="Lib","Liberica",""))))

•	the IF function is used to correct the short abbreviation of the coffee name to the full form. For E.g.: Rob to Robusta for a clearer understanding. 

•	The logical test was filled with if the cell I2 = “Rob” then in the true condition is filled with the full name of the coffee which is “Robusta”, the full names of all the remaining coffee are stated using the IF functions again and at last in the value if false condition “” a blank value is left if the condition is false. 

•	Brackets are used to close each IF function stated.

Roast Type Name:

•	IF (logical_test, [value_if_true], [value_if_false]) 

•	=IF(J2="M","Medium",IF(J2="L","Large",IF(J2="D","Dark","")))

•	Used the same IF function used for the cell coffee type name to rename the short abbreviations to full names for better understanding purposes. 
Order date cell:

•	Formatted the date to European date where the format is dd-mmm-yyyy.

•	In this format the day comes first, the month is shown in abbreviation like sep and the year is shown in full.  

Size cell:

•	Formatted the size cell to show metrics. Previously the size cell didn’t have any metrics and was represented in single digits which are then formatted to decimal points with metrics shown. For e.g.: 1 to 1.0 Kg.
 
Unit price and sales cell:

•	Updated the two cells with $ currency values. Used the accounting format in the home tab and selected the $ currency format to update both cells. 
Duplicate values check:

•	Selected all the sheets with data and then selected the data Tab which has a remove duplicate option, when clicked it removes all the duplicate values.
 
To create tables:

•	Used keyboard shortcut: Command + T in Mac and for Windows use, control + T.

•	Created tables for ease of pivot table creation and ease of updating pivot table data.
 
 
Pivot table:

•	Use of pivot table by selecting the pivot table from the insert tab.

•	Renamed the pivot table name and sheet to TotalSales.

•	In the pivot table Fields, Dragged the order date into the row’s fields.

•	Right-click on the year rows and select the grouping option to select month and year as those are the options wanted.

•	Go to the design tab then click on the reports layout and select Show in tabular form, click grand totals select off for rows and columns and lastly click on subtotals then select do not show subtotals.

•	Drag the coffee type names to the column fields and drag the sum to the values field.

•	Format the sum cell which is dragged into the values fields. Formatted the sum cell by selecting the format of the number to have no decimal places and clicked on the 1000 comma separator option. This makes the sheet much cleaner. 

•	2D line chart visualization- click on the pivot chart sheet go to the insert tab then click on the line chart option and select 2D line chart.

•	Formatting the 2D line chart- double click on the chart and then fill the gradient with solid fill. Used the custom RGB colour option and used light purple as a gradient fill. 

•	Formatted the axis label and chart label by selecting the design tab and clicking on the add chart element option.

•	Timeline- to insert timeline go to pivot chart analyse and click on insert timeline.
 
 

 
 
Insert slicer:

•	From the pivot chart analyze tab select Insert slicer.

•	Create slicers for size, coffee roast type and loyalty card.


Loyalty card cell creation:

•	=XLOOKUP (lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])

•	Use of the xlookup function.

•	Look up value is the customer ID, look up array customer ID from the customer's sheet, return is the Loyalty card, lastly 0 for an exact match and then close the bracket.
Duplicate total sales pivot table:

•	Rename it to country bar chart.

•	Remove all the pivot table fields except for sum and drag the country to the rows field.

•	Go to the insert tab and insert a bar chart.

•	Format the pivot table sheet of the country in ascending order and use the sum of sales option to rearrange the bar charts.

•	Format the bar chart’s colour, font size, font colour and axis colour.

•	Go to the design tab and insert data labels.

•	Format the currency of the data labels to $.	
Duplicate country bar chart pivot table sheet:

•	Rename it to Top5Customers.

•	Change the country axis to the customer’s name axis which is done by first dropping the country from the axis field and then dragging the customer’s name into the axis field of the pivot table fields. 

•	Then go into the customer’s name column of the pivot table sheet click on value filters and then select top 10 and then modify it to top 5 and order by the sum of sales.


Create a new worksheet:

•	Name it Dashboard.

•	Insert a rectangle shape (press the option key and hold it while dragging the shape) fill it will colour and name it the coffee sales dashboard. This is the header of the visualisation. 

•	Go to the timeline tab click on report connections and then select all the pivot tables which will then allow for the slicer to visualize all the charts. 

•	Do the same for the slicers and select all the pivot tables by clicking on report connections.

•	Remove the gridlines by selecting the view tab.
 
Final visualisation result:
 


