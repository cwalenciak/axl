# axl
An Excel VBA library/class for working with arrays.

# Philosophy
To have an Excel VBA library that works with data sets you recieve in excel, so that you can quickly perform analysis or build an Excel app to process the data.

# Required References
* Microsoft Scripting Runtime
* Microsoft Forms 2.0 Object Library

# Conventions

### Misc
* **Array** suffix is used for a 2D array with a starting index of 1. &nbsp;ex. `` dataArray(1 to 5, 1 to 3) ``
* **List** suffix is used for a 1D array with a starting index of 0.  &nbsp;ex. `` dataList(0 to 5) ``
* **"col"** or **"cols"** for columns

### Variables
* **headers**- boolean to pass to a function to mark whether or not a dataArray has headers (defualt is true). 
* **caseMatters**- boolean to pass to a function to mark whether or not you care about capital letters (default is false).
* **dataArray** - Main array that is past to an axl function that works with arrays.
* **dataList** - Main list that is past to an axl function that works with list.
* **sortArray**
* **sortList**
* **searchArray**
* **sortList**

### Comment Symbols
* Section Header: '//
* Funtion Header: '@
* Function Comments (Outer): '-
* Function Comments (Inner): '#
* Function dependent on internal function list: '!

### Warnings/Error:
* <\!DATA ERROR!\>
* <\!NO DATA!\>
* <\!EMPTY ARRAY!\>
* <\!NO POSITION!\>
* <\!NULL!\>

# Groups and Functions

### Build Arrays
* arrangeArray
* bindLists
* rowToCol
* stackLists
* uniqueArray

### Build Lists
* arrangeList
* colToList
* rowToList
* uniqueList

### Cast Data
* absouluteValue
* blankToZero
* castDate
* castDbl
* castLong
* castString
* roundNum
* stringUpper
* stringLower
* trimData

### Cell Modify
* append
* concat
* copyCells
* copyCellsIf
* fillCells
* fillCellsIf
* splitData
* selectLeft
* selectMid
* selectRight
* parse

### Column Modify
* colMove
* colSelect
* colSpacers
* colSwap

### Compare
* compareData
* compareBool

### Error Arrays
* errorArray
* emptyArray

### Export Arrays
* arrayToWs
* arrayToWb
* arraytoCmbx
* fillWorksheet (*Private*)

### Extract
* extractDay
* extractMonth
* extractYear
* monthIndex
* monthName

### Format WS
* autofitCols
* hideRightCols
* hideBottomRows
* boldFirstCol
* boldTopRow
* boldLastRow
* colsToNum1
* colsToNum2
* colsToNumString
* colsToPercent
* colsToAcnt
* colsPreserveText
* workhseetReturn (*Private*)

### Get Data
* externalWbData
* internalWbData
* getData (*Private*)

### Get Value
* colSum
* countIf
* getElement
* listSum
* ncol
* nrow
* sumIf

### Header Data
* headerIndex
* list (*Private*)
* scalar (*Private*)

### Match Data
* matchDataArr
* matchMetaArr
* matchRowCol

### Math
* colMath1
* colMath2
* mathOp (*Private*)

### Row Modify
* addTotalRow
* removeTopRow

### Search Data
* binarySearch
* actualBinarySearch (*Private*)

### Sort Data
* sortArray
* sortList
* sortMonthNames
* quickSortArray (*Private)

### Statistics
* colMean
* colSD
* listMean
* listSD

### Transform
* arrayAddIndex
* filterData
* groupData
* listAddIndex
* missingCategories
* reshapeArray
* transpose

