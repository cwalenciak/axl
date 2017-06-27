# axl
Excel VBA library/class for working with arrays.

## Conventions:

### Misc
* **Array** suffix is used for a 2D array with a starting index of 1. &nbsp;ex. dataArray(1 to 5, 1 to 3)
* **List** suffix is used for a 1D array with a starting index of 0.  &nbsp;ex. dataList(0 to 5)
* **"col"** or **"cols"** for columns

### Variables
* **headers**: boolean to pass to a function to mark whether or not a dataArray has headers (defualt is true). 
* **caseMatters**: boolean to pass to a function to mark whether or not you care about capital letters (default is false).

### Data Variables
* dataArray - Main array that is past to an axl function that works with arrays.
* dataList - Main list that is past to an axl function that works with list.
* sortArray
* sortList
* searchArray
* sortList

### Comment Symbols
* Section Header: //
* Funtion Header: @
* Function Comments (Outer): '-
* Function Comments (Inner): '#
* Function dependent on internal function list: '!

### Warnings/Error:
* <!DATA ERROR!>
* <!NO DATA!>
* <!EMPTY ARRAY!>
* <!NO POSITION!>
* <!NULL!>
