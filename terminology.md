# Terminology Excel

## Active Cell

- The cell in the spreadsheet that is currently selected. 


## Cell

- A cell is at the intersection between a row and a column. 
- A cell is referenced by the column letter and row number. 


## Fill handle

- The fill handle is the small black square on the bottom right-hand corner of the active cell.

```
* You can display or hide the fill handle by:

Click File > Options. 

* In Excel 2007 only: Click the Microsoft Office Button, and then click Excel Options.

* In the Advanced category, under Editing options, 

select or clear the Enable fill handle and cell drag-and-drop check box.
```

## Ribbon

- The ribbon is the main menu bar at the top of the Excel screen. 
- The ribbon is several tabs. 

```
* The HOME tab has some of the most frequently used tools. 
* You can collapse the ribbon to allow more space for the spreadsheet 
* in the main window by double-clicking on any of the tab labels (single-click for Mac users). 

* When you repeat the action, the ribbon will re-appear. 
* Once your ribbon is hidden, you can bring it back temporarily with a single-click, 
* use the tools you needed, and then make it disappear again 
* with another single-click on the tab or anywhere in the spreadsheet.
```

## Row

- The rows are counted in numbers. 
- There are 1,048,576 rows in an Excel spreadsheet. 
- You can read more about the specifications and limits of Excel spreadsheets

## Column

- The columns are listed in letters. 
- There are 16,384 columns in an Excel spreadsheet. 
- You can read more about the specifications and limits of Excel spreadsheets

## Quick Access Toolbar

- The Quick Access Toolbar sits above the ribbon 
- (can also be customized to sit below the ribbon). 
- Tools from any of the ribbon tabs can be added for quick access without switching between tabs.

## Status bar

- The Status Bar is below the spreadsheet. 
- It contains several useful areas. 
- The Zoom tool, access to three different view options as well as, by default, 
- several calculation results which display dynamically whenever data is selected in the spreadsheet: 

## Workbook

- The term workbook refers to the entire Excel file. 
- The file name of each workbook is at the top of the file window.

## Worksheet

- A workbook can contain several worksheets. 
- You can add worksheets at the bottom left by clicking on the plus sign next to the last worksheet tab. 
- Right-click on the worksheet tab and you can rename the worksheet and execute a range of other commands.

## Formula

- A formula is entered into a cell to perform a calculation. 
- A formula always starts with an equal sign (=) and once committed (press Enter), 
- the result is displayed in that cell. 

```
* At its most basic, formulas can be simple mathematical calculations 
  with values much like you would type into a calculator. 
 
* An example of a formula would be: =A1+B1 
  which would take whatever value was entered into cell A1 and 
  add it to the value that was typed into B1.

* After typing the formula and pressing the Enter key, 
  the resulting value will be displayed in the cell in which you entered the formula.
```

## Function

 - A function is what we referred to in the videos a 'mini-program' 
 - that you can use to make more complex calculations. 
 - Functions are used inside formulas and therefore, 
 - you need to start with an equal sign (=).
 - Formulas operate with cell references and are very powerful. 
- 
```
* One commonly used function is SUM, 
  which will add up the values in a defined range. 

* The function: =SUM(A1:A12) 
  will sum up all values contained in cells A1through to A12 
  and return the result once you commit the function by pressing the ENTER key.
```

## Formula Bar

- The formula bar is located underneath the ribbon. 
- The first edit line shows cell reference of the currently active cell 
  (this is called the Name Box)
- The second edit line provides space to enter cell content 
  and a helper tool to enter formulas:

```
* Once you enter an equal sign into the active cell, 
* frequently used functions appear in the Name Box on the left
* a drop-down menu offers more options.
```

## Value

- Values are numeric data that is entered into a cell. 

```
* When text is entered into a cell 
  without being assigned a number format, 
  we refer to them as labels. 
  
* When data is formatted as a value type, 
  it can be referred to in formulas & functions 
  and used in calculations.
```  

## Range

- A range refers to two or more cells. 

```
* When these cells are together, 
  we call this an adjacent range. 

Consider this example:

* This adjacent range covers all the cells from A1 through to C2 
or in Excel syntax this is written as A1:C2. 
The colon (:) basically stands for 'through to'. 

* Whenever we want to define a range of cells 
  that are not all in one place, 
  we talk about non-adjacent ranges:

This range includes cells A1:A2 and C1:C2. 
In Excel syntax this is written as A1:A2,C1:C2.
```

## Reference, relative

- A relative cell reference is one that changes relative to the direction in which it is copied.

```
* A2 and B2 are relative cell references. 
* When we copy the formula in C2 
  downwards into C3 and C4 with the fill handle, 
  
* then Excel will assume that you want to conduct 
  the same calculation in rows 3 and 4 as you did in row 2. 
  
* In other words, Excel will perform the calculation 
A3*B3 in C3 and A4*B4 in C4. 

* Excel effectively updates the row number 
  in each of the cell references for every row 
  that you copy your formula downwards.
```

## Reference, absolute

- Or, as we like to fondly call it, the dollar thingy. 
- A cell reference is absolute when it does not change whenever it is copied.
   
```
* To make a cell reference absolute, 
* you must include a $ before each element of the cell reference: $A$1. 
* This can be a bit cumbersome. 
* The keyboard shortcut to turn a cell reference 
  into an absolute cell reference is to press F4.
```