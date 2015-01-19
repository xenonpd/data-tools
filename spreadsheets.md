Use INDIRECT when you want to change the reference to a cell within a formula without changing the formula itself.

Syntax

INDIRECT(ref_text,a1)

Ref_text     is a reference to a cell that contains an A1-style reference, an R1C1-style reference, a name defined as a reference, or a reference to a cell as a text string. If ref_text is not a valid cell reference, INDIRECT returns the #REF! error value.

If ref_text refers to another workbook (an external reference), the other workbook must be open. If the source workbook is not open, INDIRECT returns the #REF! error value.

If ref_text refers to a cell range outside the row limit of 1,048,576 or the column limit of 16,384 (XFD), INDIRECT returns a #REF! error.

NOTE   This behavior is different from Excel versions earlier than Microsoft Office Excel 2007, which ignore the exceeded limit and return a value.

A1     is a logical value that specifies what type of reference is contained in the cell ref_text.

If a1 is TRUE or omitted, ref_text is interpreted as an A1-style reference.

If a1 is FALSE, ref_text is interpreted as an R1C1-style reference.

---
Example:

Use INDIRECT to dynamically refer to worksheets in Excel and Google Spreadsheets

Sometimes you want to make a reference to certain worksheets dynamically. For example if you have data in the same format split over multiple worksheets and you want to select data from different sheets dynamically. In this case, you can use the INDIRECT() function, which is available in both Excel and Google Spreadsheets.

More examples at [Microsoft's support page](https://support.office.com/en-us/article/INDIRECT-function-21f8bcfc-b174-4a50-9dc6-4dfb5b3361cd?ui=en-US&rs=en-US&ad=US).



References:



http://spreadsheetpro.net/how-to-make-a-dynamic-reference-to-a-worksheet-in-excel-and-google-spreadsheets/



Brian Connelly: Working with CSVs on the Command Line
http://bconnelly.net/working-with-csvs-on-the-command-line/


How to batch convert .csv to .xls/xlsx
http://superuser.com/questions/301431/how-to-batch-convert-csv-to-xls-xlsx

Merge Multiple CSV Files
http://www.zorbathegeek.com/156/merge-multiple-csv-files.html
