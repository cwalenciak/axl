# axl
Excel VBA library/class for working with arrays.

## Conventions:
* Array suffix is used for a 2D array with a starting index of 1. &nbsp;ex. dataArray(1 to 5, 1 to 3)
* List suffix is used for a 1D array with a starting index of 0.  &nbsp;ex. dataList(0 to 5)
* Max line Width: 100
* "col" or "cols" for columns

### **Variables**
* **headers**: true if dataArray has headers (defualt is true). 
* **caseMatters**: set to false if you don't care about capital letters (default is false).

### **Argument Variables**
* dataArray
* dataList
* lists
* sortArray
* sortList
* searchArray
* sortList

### **Comment Symbols**
-Section Header: //
-Funtion Header: @
-Function Comments (Outer): '-
-Function Comments (Inner): '#
-Function dependent on internal function list: '!
