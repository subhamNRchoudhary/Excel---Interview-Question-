# Excel---Interview-Question-
Interview Question asked



### Basic Excel Questions:

1. **Q: What is the shortcut to insert a new row in Excel?**
     **A:** Press `Ctrl + Shift + "+"`.

2. **Q: How do you create a formula in Excel?**
     **A:** Start with an equal sign (`=`), followed by the formula or function you want to use, e.g., `=A1 + B1`.

3. **Q: How can you quickly sum a range of cells?**
     **A:** Use the `SUM` function, e.g., `=SUM(A1:A10)`.

4. **Q: What is the use of the `IF` function in Excel?**
     **A:** The `IF` function checks a condition and returns one value if true and another if false. Example: `=IF(A1 > 10, "Yes", "No")`.

5. **Q: How can you freeze the top row in Excel?**
     **A:** Go to the `View` tab, and click on `Freeze Panes`, then select `Freeze Top Row`.

6. **Q: What is the shortcut for copying the selected cells?**
     **A:** Press `Ctrl + C`.

7. **Q: How do you apply a filter to a column?**
     **A:** Select the column and then click on the `Data` tab, then `Filter`.

8. **Q: How can you concatenate two strings in Excel?**
     **A:** Use the `&` operator or the `CONCATENATE` function. Example: `=A1 & " " & B1` or `=CONCATENATE(A1, " ", B1)`.

9. **Q: What is the `VLOOKUP` function used for?**
     **A:** `VLOOKUP` searches for a value in the first column of a table and returns a value in the same row from another column. Example: `=VLOOKUP(A1, B:C, 2, FALSE)`.

10. **Q: How can you remove duplicates from a dataset?**
      **A:** Go to the `Data` tab and click `Remove Duplicates`.

### Intermediate Excel Questions:

11. **Q: What is the difference between `VLOOKUP` and `HLOOKUP`?**
      **A:** `VLOOKUP` searches vertically in a column, while `HLOOKUP` searches horizontally in a row.

12. **Q: How can you create a dropdown list in a cell?**
      **A:** Use the `Data Validation` feature. Go to `Data` > `Data Validation`, choose `List`, and enter your list items.

13. **Q: How do you calculate the difference between two dates in Excel?**
      **A:** Use the `DATEDIF` function. Example: `=DATEDIF(A1, B1, "D")` returns the difference in days.

14. **Q: What is the purpose of the `INDEX` function?**
      **A:** `INDEX` returns the value of a cell in a specified row and column within a range. Example: `=INDEX(A1:B5, 2, 1)` returns the value from the second row, first column.

15. **Q: How can you create a pivot table in Excel?**
      **A:** Select your data range, go to `Insert` > `PivotTable`, and choose the fields you want to analyze.

16. **Q: What does the `MATCH` function do?**
      **A:** `MATCH` returns the relative position of a value in a specified range. Example: `=MATCH("Apple", A1:A5, 0)` returns the position of "Apple" in the range A1:A5.

17. **Q: How do you use the `TEXT` function?**
      **A:** The `TEXT` function formats a number and converts it to text. Example: `=TEXT(1234.567, "0.00")` returns "1234.57".

18. **Q: How can you apply conditional formatting based on a cell value?**
      **A:** Select the cells, go to `Home` > `Conditional Formatting`, and create a new rule based on the cell value.

19. **Q: What is the `SUMIF` function?**
      **A:** `SUMIF` sums the values in a range that meets a specified criterion. Example: `=SUMIF(A1:A10, ">5")` sums values greater than 5.

20. **Q: How do you protect a worksheet in Excel?**
    **A:** Go to the `Review` tab and click `Protect Sheet`, then set a password if desired.

### Advanced Excel Questions:

21. **Q: What is the `ARRAYFORMULA` in Excel?**
      **A:** `ARRAYFORMULA` is used to apply a formula to a range of cells. Example: `{=A1:A10 * B1:B10}` (in Google Sheets, similar in Excel using `CTRL+SHIFT+ENTER`).

22. **Q: How can you create a dynamic range in Excel?**
      **A:** Use the `OFFSET` and `COUNTA` functions. Example: `=OFFSET(A1, 0, 0, COUNTA(A:A), 1)` creates a dynamic range based on the number of non-blank cells in column A.

23. **Q: What is the `INDIRECT` function used for?**
      **A:** `INDIRECT` returns the reference specified by a text string. Example: `=INDIRECT("A" & B1)` returns the value in column A, row B1.

24. **Q: How do you use the `XLOOKUP` function in Excel?**
      **A:** `XLOOKUP` searches a range or array and returns an item corresponding to the first match it finds. Example: `=XLOOKUP(A1, B1:B10, C1:C10)`.

25. **Q: What is the purpose of the `TRANSPOSE` function?**
      **A:** `TRANSPOSE` changes the orientation of a range from rows to columns or vice versa. Example: `=TRANSPOSE(A1:B2)`.

26. **Q: How do you calculate the internal rate of return (IRR) in Excel?**
      **A:** Use the `IRR` function. Example: `=IRR(A1:A5)` calculates the IRR for a series of cash flows in A1:A5.

27. **Q: What is the `GOAL SEEK` feature in Excel?**
      **A:** `GOAL SEEK` finds the input value needed to achieve a desired result in a formula. It’s located under `Data` > `What-If Analysis` > `Goal Seek`.

28. **Q: How can you use the `OFFSET` function to reference a range dynamically?**
      **A:** `OFFSET` returns a reference to a range that is a specified number of rows and columns from a cell or range. Example: `=OFFSET(A1, 2, 2, 1, 3)`.

29. **Q: What is the `Power Query` feature in Excel?**
      **A:** `Power Query` is used for data transformation and automation, allowing you to import, clean, and reshape data from various sources.

30. **Q: How do you use `Solver` in Excel?**
      **A:** `Solver` is an add-in used to find an optimal value (maximum or minimum) for a formula in one cell, subject to constraints. Activate it via `Data` > `Solver`.
