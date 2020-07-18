# What is 'GetLikeness'?
This project is a Visual Basic script for Excel that adds 2 functions for processing *precise* levels of equality between two strings.

In Excel, sometimes you need to determine exact equality between two strings. This can easily be done by using the '=' operator within an Excel function or VBA statement. But what if the strings are not in the same case? What if one contains a period, one doesnt, but other than that the strings are equal?

Excel only supports direct equality tests, resulting in TRUE or FALSE values. This library allows you to return a percentage value calculated based on similar characters and concurant character patterns seen within the test subjects.


## How do you use this script in your workbooks?
1) You first have to enable the 'Developer' tab in Excel by going to File > Options > Customize Ribbon, and checking the box on the right labeled 'Developer'. 
2) Once enabled, navigate back to the primary Excel window, go to the Ribbon, select 'Developer', and click on the 'Visual Basic' button (first button from left).
3) Inside the 'Microsoft Visual Basic for Applications' window that appears, go to File > Import File... (or press CTRL + M while the window is in focus). 
4) Browse to the .bas file to load the module into the workbook. After you see the module in the 'Project' treeview on the right inside the 'Modules' folder, you're good to go!

## What do the results look like?
Once the module is placed within your file (or, placed inside your NORMAL template), you can call `=GETLIKENESS(Input1, Input2)` to run the comparison and return the result (for one string to one string comparison):

![Figure1](https://i.imgur.com/L9aYj1H.png)


You can also run the comparison of one string against a range of strings by utilizing the `=GETMAXLIKENESS(Input1, Range1)`:

![Figure2](https://i.imgur.com/E1aOH9d.png)


The final and optional to the `GETMAXLIKENESS` function called IgnoreOriginal (defaulted to FALSE) that, when enabled, allows you to scan the same range that the source test subject came from while ignoring 1 instance of the test subject:

![Figure3](https://i.imgur.com/7vIfd38.png)
