VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "axl_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'==================================================================================================
'//CLASS VARIABLES
'==================================================================================================

Private monthNameIndex As Dictionary




'==================================================================================================
'//ENUMS
'==================================================================================================


'************************************************
'@Parse Operations
'************************************************
Public Enum axlParse
    RemoveNonAlpha
    RemoveNonNum
    RemoveNonAphaNum
    RemoveHyphens
    RemoveBlanks
End Enum


'************************************************
'@Compare Operator
'************************************************
Public Enum axlCompare
    LessThan
    LessThanOrEqualTo
    EqualTo
    GreaterThanOrEqualTo
    GreaterThan
    NotEqual
    BetweenExclusive
    BetweenInclusive
    Contains
    DoesNotContain
End Enum


'************************************************
'@Math Operations
'************************************************
Public Enum axlMath
    Add
    subtract
    Multiply
    divide
    Exponent
    Root
End Enum




'==================================================================================================
'//DICTIONARIES
'==================================================================================================

'************************************************
'@Month Index
'************************************************
Private Sub createMonthNameIndex()
    Set monthNameIndex = New Dictionary
    monthNameIndex.compareMode = vbTextCompare
    monthNameIndex.Add Key:="January", item:=1
    monthNameIndex.Add Key:="February", item:=2
    monthNameIndex.Add Key:="March", item:=3
    monthNameIndex.Add Key:="April", item:=4
    monthNameIndex.Add Key:="May", item:=5
    monthNameIndex.Add Key:="June", item:=6
    monthNameIndex.Add Key:="July", item:=7
    monthNameIndex.Add Key:="August", item:=8
    monthNameIndex.Add Key:="September", item:=9
    monthNameIndex.Add Key:="October", item:=10
    monthNameIndex.Add Key:="November", item:=11
    monthNameIndex.Add Key:="December", item:=12
End Sub




'==================================================================================================
'//BUILD ARRAYS & LIST
'==================================================================================================


'************************************************
'@arrangeArray
'************************************************
Function arrangeArray(startNumber, _
                      numberCount, _
                      Optional step = 1, _
                      Optional headers As Boolean = True, _
                      Optional headerName)
                      
    ReDim tempArray(1 To numberCount + Abs(headers), 1 To 1)

    Dim i As Long
    For i = LBound(tempArray) + Abs(headers) To UBound(tempArray)
        tempArray(i, 1) = startNumber + (i - Abs(headers) - 1) * step
    Next i

    '#Fill in headers
    If IsMissing(headerName) = False And headers Then
        tempArray(1, 1) = headerName
    End If

    arrangeArray = tempArray
    
End Function


'************************************************
'@arrangeList
'************************************************
Function arrangeList(startNumber, _
                     numberCount, _
                     Optional step = 1, _
                     Optional ascending As Boolean = True)
    
    ReDim tempArray(0 To numberCount - 1)

    Dim i As Long
    For i = LBound(tempArray) To UBound(tempArray)
        tempArray(i) = startNumber + (i * step)
    Next i
    
    arrangeList = tempArray

End Function


'************************************************
'@bindList
'************************************************
'-List become columns in new array
Function bindLists(lists)
    
    '#Loop Variables
    Dim i As Integer
    Dim j As Integer
    
    Dim maxItem As Long

    For i = LBound(lists) To UBound(lists)
        If UBound(lists(i)) > maxItem Then
            maxItem = UBound(lists(i))
        End If
    Next i
    
    ReDim listArray(1 To maxItem + 1, 1 To UBound(lists) + 1) As Variant
    
    For i = LBound(lists) To UBound(lists)
        For j = LBound(listArray) To UBound(listArray)
            
            If j - 1 <= UBound(lists(i)) Then
                listArray(j, i + 1) = lists(i)(j - 1)
            Else
                listArray(j, i + 1) = "<!NULL!>"
            End If
        Next j
    Next i
    
    bindLists = listArray
    
End Function


'************************************************
'@colToList
'************************************************
Function colToList(dataArray, _
                   sourceCol, _
                   Optional headers = True)

    sourceCol = headerIndex(sourceCol, dataArray)

    ReDim tempArray(0 To UBound(dataArray) - 1 + headers)
    
    Dim i As Integer
    For i = LBound(tempArray) To UBound(tempArray)
        tempArray(i) = dataArray(i + 1 + Abs(headers), sourceCol)
    Next i
        
    colToList = tempArray

End Function


'************************************************
'@rowToList
'************************************************
Function rowToList(dataArray, _
                   sourceRow, _
                   Optional dropFirstElem = False)

    ReDim tempArray(0 To UBound(dataArray, 2) - 1 + dropFirstElem)
    
    Dim i As Integer
    For i = LBound(tempArray) To UBound(tempArray)
        tempArray(i) = dataArray(sourceRow, i + 1 + Abs(dropFirstElem))
    Next i
        
    rowToList = tempArray

End Function


'************************************************
'@stackList
'************************************************
'-List become rows in new array
'!transpose(), bindList()
Function stackLists(lists)
    stackLists = transpose(bindList(lists))
End Function


'************************************************
'@uniqueArray
'************************************************
'!compareData(), reshapeArray()
Function uniqueArray(ByVal dataArray, _
                     ByVal filterCol, _
                     Optional headers As Boolean = True, _
                     Optional caseMatters As Boolean = True)

    filterCol = headerIndex(filterCol, dataArray)
    dataArray = sortArray(dataArray, Array(filterCol), True, headers, caseMatters)

    Dim rowIncr As Long
    rowIncr = 1
    
    Dim nextRowIncr As Long
    nextRowIncr = 1
    
    Dim tempIncr As Long
    tempIncr = 1
    
    ReDim tempArray(1 To UBound(dataArray), 1 To 1) As Variant
    
    '#Assign First Row to Temp Array
    tempArray(1, 1) = dataArray(1, filterCol)
    If headers = True Then
        tempArray(2, 1) = dataArray(2, filterCol)
        rowIncr = 2
        tempIncr = 2
    End If
    
    '#Search For Unique then Fill Array
    Do While (rowIncr + nextRowIncr <= UBound(dataArray))
        Dim nextRow As Variant
        nextRow = dataArray(rowIncr + nextRowIncr, filterCol)
        
        If compareData(dataArray(rowIncr, filterCol), nextRow, caseMatters) = 0 Then
            nextRowIncr = nextRowIncr + 1
        Else
            tempIncr = tempIncr + 1
            tempArray(tempIncr, 1) = nextRow
            rowIncr = rowIncr + nextRowIncr
            nextRowIncr = 1
        End If
    Loop
    
    '#Assing Temp Array to Final Array
    tempArray = reshapeArray(tempArray, tempIncr, ncol(tempArray))
    
    uniqueArray = tempArray
    
End Function


'************************************************
'@uniqueList
'************************************************
'!compareData()
Function uniqueList(ByVal dataList, Optional caseMatters As Boolean = True)
                    
    Dim tempArray As Variant
    tempArray = uniqueArray(bindList(Array(dataList)), 1, False, caseMatters)
    uniqueList = colToList(tempArray, 1, False)
    
End Function




'==================================================================================================
'//CAST DATA
'==================================================================================================


'************************************************
'@absoluteValue
'************************************************
Function absoluteValue(dataArray, _
                       sourceCol, _
                       Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = Abs(dataArray(i, sourceCol))
    Next i

    '#Return
    absoluteValue = dataArray
    
End Function


'************************************************
'@blankToZero
'************************************************
Function blankToZero(dataArray, _
                     colIndex, _
                     Optional headers As Boolean = True)
    
    colIndex = headerIndex(colIndex, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If Trim(dataArray(i, colIndex)) = "" Then
            dataArray(i, colIndex) = 0
        End If
    Next i
    
    blankToZero = dataArray

End Function


'************************************************
'@castDate
'************************************************
Function castDate(dataArray, _
                  sourceCol, _
                  Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsDate(dataArray(i, sourceCol)) = True Then
            dataArray(i, sourceCol) = CDate(dataArray(i, sourceCol))
        End If
    Next i
    
    castDate = dataArray
    
End Function


'************************************************
'@castDbl
'************************************************
Function castDbl(dataArray, _
                 colIndex, _
                 Optional headers As Boolean = True)
               
    colIndex = headerIndex(colIndex, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, colIndex)) = True Then
            dataArray(i, colIndex) = CDbl(dataArray(i, colIndex))
        End If
    Next i
    
    castDbl = dataArray
    
End Function


'************************************************
'@castLong
'************************************************
Function castLong(dataArray, _
                  sourceCol, _
                  Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, sourceCol)) = True Then
            dataArray(i, sourceCol) = CLng(dataArray(i, sourceCol))
        End If
    Next i
    
    castLong = dataArray
    
End Function


'************************************************
'@castString
'************************************************
Function castString(dataArray, _
                    sourceCol, _
                    Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = CStr(dataArray(i, sourceCol))
    Next i
    
    castString = dataArray
    
End Function


'************************************************
'@roundNum
'************************************************
Function roundNum(dataArray, _
                  sourceCol, _
                  Optional digitCount As Long = 0, _
                  Optional headers As Boolean = True)
                  
    sourceCol = headerIndex(sourceCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = Round(dataArray(i, sourceCol), digitCount)
    Next i
       
    '#Return
    roundNum = dataArray
    
End Function


'************************************************
'@stringUpper
'************************************************
Function stringUpper(dataArray, _
                     sourceCol, _
                     Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)

    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = UCase(dataArray(i, sourceCol))
    Next i
    
    '#Return
    stringUpper = dataArray
    
End Function


'************************************************
'@stringLower
'************************************************
Function stringLower(dataArray, _
                     sourceCol, _
                     Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
   
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = LCase(dataArray(i, sourceCol))
    Next i
       
    '#Return
    stringLower = dataArray
    
End Function


'************************************************
'@trimData
'************************************************
Function trimData(dataArray, _
                  sourceCol, _
                  Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
       
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, sourceCol) = Trim(dataArray(i, sourceCol))
    Next i
    
    '#Return
    trimData = dataArray
    
End Function




'==================================================================================================
'//CELL MODIFY
'==================================================================================================


'************************************************
'@append
'************************************************
Function append(dataArray, _
                appendCol, _
                appendItem, _
                Optional infront As Boolean = False, _
                Optional headers As Boolean = True)

    appendCol = headerIndex(appendCol, dataArray)
    
    Dim i As Long
    
    If infront = False Then
        For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
            dataArray(i, appendCol) = dataArray(i, appendCol) & appendItem
        Next i
    Else
        For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
            dataArray(i, appendCol) = appendItem & dataArray(i, appendCol)
        Next i
    End If
    
    append = dataArray

End Function


'************************************************
'@concat (Concatenate)
'************************************************
Function concat(ByVal dataArray, _
                ByVal fillCol, _
                ByVal columnList, _
                Optional ByVal seperator As String = "", _
                Optional ByVal headers As Boolean = True)
                     
    fillCol = headerIndex(fillCol, dataArray)
    columnList = headerIndex(columnList, dataArray)
    
    Dim i As Long
    Dim j As Long
    Dim cellData As Variant

    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        For j = LBound(columnList) To UBound(columnList)
            If j = LBound(columnList) Then
                cellData = dataArray(i, columnList(j))
            Else
                cellData = cellData & seperator & dataArray(i, columnList(j))
            End If
        Next j
        dataArray(i, fillCol) = cellData
    Next i

    '#Return
    concat = dataArray

End Function


'************************************************
'@copyCells
'************************************************
Function copyCells(dataArray, _
                   fillCol, _
                   copyColumn, _
                   Optional headers As Boolean = True)
    
    fillCol = headerIndex(fillCol, dataArray)
    copyColumn = headerIndex(copyColumn, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, fillCol) = dataArray(i, copyColumn)
    Next i
    
    '#Return
    copyCells = dataArray

End Function


'************************************************
'@copyCellsIf
'************************************************
Function copyCellsIf(dataArray, _
                     fillCol, _
                     copyColumn, _
                     criteria, _
                     searchCol, _
                     meetsCriteriaBy As axlCompare, _
                     Optional uboundCriteria, _
                     Optional caseMatters As Boolean = False, _
                     Optional headers As Boolean = True)
    
    fillCol = headerIndex(fillCol, dataArray)
    copyColumn = headerIndex(copyColumn, dataArray)
    searchCol = headerIndex(searchCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
    
        Dim outcome As Boolean
        outcome = compareBool(dataArray(i, searchCol), criteria, meetsCriteriaBy, uboundCriteria, _
                              caseMatters)
                                   
        If outcome = True Then
            dataArray(i, fillCol) = dataArray(i, copyColumn)
        End If
        
    Next i
    
    '#Return
    copyCellsIf = dataArray
    
End Function


'************************************************
'@fillCells
'************************************************
Function fillCells(dataArray, _
                   fillCol, _
                   fillWith, _
                   Optional headers As Boolean = True)

    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long

    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, fillCol) = fillWith
    Next i

    '#Return
    fillCells = dataArray

End Function


'************************************************
'@fillCellsIf
'************************************************
Function fillCellsIf(dataArray, _
                     fillCol, _
                     criteria, _
                     searchCol, _
                     meetsCriteriaBy As axlCompare, _
                     trueReturn, _
                     Optional uboundCriteria, _
                     Optional headers As Boolean = True, _
                     Optional caseMatters As Boolean = True)

    fillCol = headerIndex(fillCol, dataArray)
    searchCol = headerIndex(searchCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
    
        Dim outcome As Boolean
        outcome = compareBool(dataArray(i, searchCol), criteria, meetsCriteriaBy, uboundCriteria, _
                              caseMatters)
                                   
        If outcome = True Then
            dataArray(i, fillCol) = trueReturn
        End If
        
    Next i
   
    '#Return
    fillCellsIf = dataArray

End Function


'************************************************
'@splitData
'************************************************
Function splitData(dataArray, _
                   sourceCol, _
                   fillColList, _
                   delimeter, _
                   Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillColList = headerIndex(fillColList, dataArray)
    
    Dim i As Long
    Dim j As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
            
        If InStr(1, dataArray(i, sourceCol), delimeter) > 0 And IsMissing(delimeter) = False Then
        
            Dim tempArr As Variant
            tempArr = Split(dataArray(i, sourceCol), delimeter)
            
            If UBound(tempArr) = UBound(fillColList) Then
            
                For j = LBound(fillColList) To UBound(fillColList)
                    dataArray(i, fillColList(j)) = tempArr(j)
                Next j
            Else
                dataArray(i, fillColList(0)) = "AXL SPLIT ERROR"
            End If
        End If
    
    Next i

    '#Return
    splitData = dataArray

End Function


'************************************************
'@selectLeft
'************************************************
Function selectLeft(dataArray, _
                    sourceCol, _
                    fillCol, _
                    charCount As Long, _
                    Optional headers As Boolean = True)

    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)

    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, fillCol) = Left(dataArray(i, sourceCol), charCount)
    Next i
    
    selectLeft = dataArray

End Function

'************************************************
'@selectMid
'************************************************
Function selectMid(dataArray, _
                    sourceCol, _
                    fillCol, _
                    start, _
                    charCount As Long, _
                    Optional headers As Boolean = True)

    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)

    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, fillCol) = Mid(dataArray(i, sourceCol), start, charCount)
    Next i
    
    selectMid = dataArray

End Function


'************************************************
'@selectRight
'************************************************
Function selectRight(dataArray, _
                     sourceCol, _
                     fillCol, _
                     charCount As Long, _
                     Optional headers As Boolean = True)

    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)

    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, fillCol) = Right(dataArray(i, sourceCol), charCount)
    Next i
    
    selectRight = dataArray

End Function


'************************************************
'@parse (Parse String)
'************************************************
Function parse(dataArray, _
               sourceCol, _
               fillCol, _
               axlParse As axlParse, _
               Optional headers As Boolean = True)
                      
    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
    
        Dim sourceData As String
        sourceData = dataArray(i, sourceCol)
        
        Dim parsedData As Variant
        parsedData = ""
    
        For j = 1 To Len(sourceData)
                 
            Select Case axlParse
                        
                '#Remove Non Alpha
                Case RemoveNonAlpha
                    Select Case Asc(Mid(sourceData, j, 1))
                        Case 32, 65 To 90, 97 To 122
                            parsedData = parsedData & Mid(sourceData, j, 1)
                    End Select
                
                '#Removed Non Numeric
                Case RemoveNonNum
                    Select Case Asc(Mid(sourceData, j, 1))
                        Case 48 To 57
                            parsedData = parsedData & Mid(sourceData, j, 1)
                    End Select
                
                '#Remove Non Alpha Numeric
                Case RemoveNonAphaNum
                    Select Case Asc(Mid(sourceData, j, 1))
                        Case 32, 48 To 57, 65 To 90, 97 To 122
                            parsedData = parsedData & Mid(sourceData, j, 1)
                    End Select
                
                '#Remove Hyphens
                Case RemoveBlanks
                    If Asc(Mid(sourceData, j, 1)) <> 32 Then
                        parsedData = parsedData & Mid(sourceData, j, 1)
                    End If
                
                '#Remove Hyphens
                Case RemoveHyphens
                    If Asc(Mid(sourceData, j, 1)) <> 45 Then
                        parsedData = parsedData & Mid(sourceData, j, 1)
                    End If
                     
            End Select
            
        Next j
        
        dataArray(i, fillCol) = parsedData
        
    Next i
        
    parseData = dataArray
    
End Function
                 



'==================================================================================================
'//COLUMN MODIFY
'==================================================================================================


'************************************************
'@colMove
'************************************************
Function colMove(dataArray, _
                 moveColumn, _
                 newPosition)
    
    moveColumn = headerIndex(moveColumn, dataArray)
    newPosition = headerIndex(newPosition, dataArray)
    
    Dim i As Long
    
    If moveColumn < newPosition Then
        For i = moveColumn To newPosition - 1
            colSwap dataArray, i, i + 1
        Next i
    Else
        For i = moveColumn To newPosition + 1 Step -1
            colSwap dataArray, i, i - 1
        Next i
    End If
       
    '#Return
    colMove = dataArray

End Function


'************************************************
'@colSelect
'************************************************
Function colSelect(dataArray, columnList)
    
    columnList = headerIndex(columnList, dataArray)
    
    ReDim tempArray(1 To UBound(dataArray), 1 To UBound(columnList) + 1)
    
    Dim i As Long
    Dim j As Long
    
    For i = LBound(columnList) To UBound(columnList)
        For j = LBound(tempArray) To UBound(tempArray)
            tempArray(j, i + 1) = dataArray(j, columnList(i))
        Next j
    Next i
      
    '#Return
    colSelect = tempArray

End Function


'************************************************
'@colSpacers
'************************************************
Function colSpacers(dataArray)

    Dim originalCount As Long
    originalCount = UBound(dataArray, 2)
    
    dataArray = reshapeArray(dataArray, UBound(dataArray), _
                             UBound(dataArray, 2) + UBound(dataArray, 2) - 1)
                               
    Dim i As Long
    For i = LBound(dataArray, 2) To originalCount - 1
        dataArray = colMove(dataArray, UBound(dataArray, 2), i * 2)
    Next i
       
    '#Return
    colSpacers = dataArray

End Function



'************************************************
'@colSwap
'************************************************
Function colSwap(dataArray, _
                 column1, _
                 column2)

    column1 = headerIndex(column1, dataArray)
    column2 = headerIndex(column2, dataArray)
    
    Dim tempVar As Variant

    Dim i As Long
    For i = LBound(dataArray) To UBound(dataArray)
        tempVar = dataArray(i, column1)
        dataArray(i, column1) = dataArray(i, column2)
        dataArray(i, column2) = tempVar
    Next i

    '#Return
    colSwap = dataArray

End Function




'==================================================================================================
'//COMPARE
'==================================================================================================


'************************************************
'@compareData
'************************************************
Private Function compareData(ByVal item1, _
                             ByVal item2, _
                             Optional caseMatters As Boolean = False) As Long
    
    '#Case Matters
    If caseMatters = False And IsNumeric(item1) = False Then
        item1 = LCase(item1)
    End If

    If caseMatters = False And IsNumeric(item2) = False Then
        item2 = LCase(item2)
    End If
    
    '#Actual Comparison
    If item1 > item2 Then
        compareData = 1
    ElseIf item1 = item2 Then
        compareData = 0
    ElseIf item1 < item2 Then
        compareData = -1
    End If

End Function


'************************************************
'@compareBool
'************************************************
Function compareBool(item1, _
                     item2, _
                     compareOperation As axlCompare, _
                     Optional uBoundItem, _
                     Optional caseMatters As Boolean = False) As Boolean

    '#Results (Lower Bound Resulsts)
    Dim results As Integer
    results = compareData(item1, item2, caseMatters)
        
    '#Upper Bound Results
    If IsMissing(uBoundItem) = False Then
        Dim uboundResults As Integer
        uboundResults = compareData(item1, uBoundItem, caseMatters)
    End If
    
    Dim outcome As Boolean
    outcome = False

    '#Compare Method
    Dim compmeth As VbCompareMethod
    compmeth = vbBinaryCompare
    
    If caseMatters = False Then
        compmeth = vbTextCompare
    End If
    
    Select Case compareOperation
        
        '#Less than
        Case Is = LessThan
            If results = -1 Then
                outcome = True
            End If
        
        '#Less than or equal to
        Case Is = LessThanOrEqualTo
            If results <> 1 Then
                outcome = True
            End If
        
        '#Equal to
        Case Is = EqualTo
            If results = 0 Then
                outcome = True
            End If
        
        '#Greater than or equal to
        Case Is = GreaterThanOrEqualTo
            If results <> -1 Then
                outcome = True
            End If
        
        '#Greater than
        Case Is = GreaterThan
            If results = 1 Then
                outcome = True
            End If
        
        '#Does not equal
        Case Is = NotEqual
            If results <> 0 Then
                outcome = True
            End If

        '#BetweenExlusive
        Case Is = BetweenExclusive
            If results = 1 And uboundResults = -1 Then
                outcome = True
            End If
        
        '#BetweenExlusive
        Case Is = BetweenInclusive
            If (results = 1 Or results = 0) And (uboundResults = 0 Or uboundResults = -1) Then
                outcome = True
            End If

        '#Contains
        Case Is = Contains
            If InStr(1, item1, item2, compmeth) > 0 Then
                outcome = True
            End If

        '#Does not contain
        Case Is = DoesNotContain
            If InStr(1, item1, item2, compmeth) = 0 Then
                outcome = True
            End If

    End Select
    
    compareBool = outcome

End Function




'==================================================================================================
'//ERROR ARRAYS
'==================================================================================================


'************************************************
'@errorArray
'************************************************
Function errorArray()
    Dim errArr(1 To 1, 1 To 1) As Variant
    errArr(1, 1) = "<!NO DATA!>"
    errorArray = errArr
End Function


'************************************************
'@Empty Array
'************************************************
Function emptyArray()
    Dim emptyArr(1 To 1, 1 To 1) As Variant
    emptyArr(1, 1) = "<!EMPTY ARRAY!>"
    emptyArray = emptyArr
End Function




'==================================================================================================
'//EXPORT ARRAYS
'==================================================================================================


'************************************************
'@arrayToWs
'************************************************
Sub arrayToWs(ByVal dataArray, _
              Optional startRow As Integer = 1, _
              Optional startCol As Integer = 1, _
              Optional wsData, _
              Optional leadingZeroList)
    
    Dim ws As Worksheet
    
    If IsObject(wsData) = True Then
        Set ws = wsData
    ElseIf IsMissing(wsData) = False Then
        Set ws = Sheets(wsData)
    Else
        Set ws = ActiveSheet
    End If
    
    fillWorksheet dataArray, startRow, startCol, ws, leadingZeroList
       
End Sub


'************************************************
'@arrayToWb
'************************************************
Sub arrayToWb(ByVal dataArray, _
              Optional startRow As Integer = 1, _
              Optional startCol As Integer = 1, _
              Optional wsData, _
              Optional leadingZeroList)
    
    Workbooks.Add
    
    Dim ws As Worksheet
    Set ws = ActiveSheet

    fillWorksheet dataArray, startRow, startCol, ws, leadingZeroList
    
End Sub


'************************************************
'@arrayToCmbx
'************************************************
Sub arrayToCmbx(dataArray, _
                ByVal cmbx As ComboBox, _
                Optional displayItemOne As Boolean = True)

    cmbx.Clear
        
    Dim i As Long
    For i = LBound(dataArray) To UBound(dataArray)
        cmbx.AddItem dataArray(i, 1)
    Next i

    If displayItemOne = True Then
        cmbx.value = dataArray(LBound(dataArray))
    End If

End Sub


'************************************************
'@colToCmbx
'************************************************
Sub colToCmbx(dataArray, _
              colSource, _
              ByVal cmbx As ComboBox, _
              Optional displayItemOne As Boolean = True, _
              Optional headers As Boolean = True)

    cmbx.Clear
    
    colSource = headerIndex(colSource, dataArray)
    
    Dim i As Integer

    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        cmbx.AddItem dataArray(i, colSource)
    Next i

    If displayItemOne = True Then
        cmbx.value = dataArray(LBound(dataArray), colSource)
    End If

End Sub


'************************************************
'@fillWorksheet (*Private*)
'************************************************
Private Sub fillWorksheet(dataArray, _
                          Optional startRow As Integer = 1, _
                          Optional startCol As Integer = 1, _
                          Optional worksheetData, _
                          Optional leadingZeroList)
    
    Dim ws As Worksheet
    Set ws = worksheetData
    
    Dim endRow As Long
    endRow = startRow + UBound(dataArray) - 1
    
    Dim endColumn As Long
    endColumn = startCol + UBound(dataArray, 2) - 1
    
    If IsMissing(leadingZeroList) = False Then
        leadingZeroList = headerIndex(leadingZeroList, dataArray)
        Dim i As Integer
        For i = LBound(leadingZeroList) To UBound(leadingZeroList)
            Columns(leadingZeroList(i)).NumberFormat = "@"
        Next i
    End If
    
    Dim exportRange As Range
    Set exportRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endColumn))
             
    '#Export Array
    exportRange = dataArray

End Sub




'==================================================================================================
'#EXTRACT
'==================================================================================================


'************************************************
'@extractDay
'************************************************
Function extractDay(dataArray, _
                    sourceCol, _
                    fillCol, _
                    Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsDate(dataArray(i, sourceCol)) = True Then
            dataArray(i, fillCol) = Day(dataArray(i, sourceCol))
        End If
    Next i
      
    '#Return
    extractDay = dataArray

End Function


'************************************************
'@extractMonth
'************************************************
Function extractMonth(dataArray, _
                      sourceCol, _
                      fillCol, _
                      Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsDate(dataArray(i, sourceCol)) = True Then
            dataArray(i, fillCol) = Month(dataArray(i, sourceCol))
        End If
    Next i
      
    '#Return
    extractMonth = dataArray

End Function


'************************************************
'@extractYear
'************************************************
Function extractYear(dataArray, _
                     sourceCol, _
                     fillCol, _
                     Optional headers As Boolean = True)
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsDate(dataArray(i, sourceCol)) = True Then
            dataArray(i, fillCol) = Year(dataArray(i, sourceCol))
        End If
    Next i
      
    '#Return
    extractYear = dataArray

End Function


'************************************************
'@monthName
'************************************************
Function monthName(dataArray, _
                   sourceCol, _
                   fillCol, _
                   Optional headers As Boolean = True)
    
    createMonthNameIndex
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, sourceCol)) = True And dataArray(i, sourceCol) - 1 < 12 Then
            dataArray(i, fillCol) = monthNameIndex.Keys()(dataArray(i, sourceCol) - 1)
        End If
    Next i
      
    '#Return
    monthName = dataArray

End Function


'************************************************
'@monthIndex
'************************************************
'-Puts in place month name given a months index
Function monthIndex(dataArray, _
                    sourceCol, _
                    fillCol, _
                    Optional headers As Boolean = True)
    
    createMonthNameIndex
    
    sourceCol = headerIndex(sourceCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If Not IsNumeric(dataArray(i, sourceCol)) Then
            dataArray(i, fillCol) = monthNameIndex(dataArray(i, sourceCol))
        End If
    Next i
    
    '#Return
    monthIndex = dataArray

End Function




'==================================================================================================
'#FORMAT WS
'==================================================================================================


'************************************************
'@autofitCols
'************************************************
Sub autofitCols(dataArray As Variant, _
                Optional wsStartRow As Long = 1, _
                Optional wsStartCol As Long = 1, _
                Optional wsData)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)
    
    Dim endRow As Long
    endRow = wsStartRow + UBound(dataArray) - 1
    
    Dim endCol As Integer
    endCol = wsStartCol + UBound(dataArray, 2) - 1
    
    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(wsStartRow, wsStartCol), ws.Cells(endRow, endCol))
    
    formatRange.EntireColumn.AutoFit
    
End Sub


'************************************************
'@hideRightCols
'************************************************
Sub hideRightCols(dataArray As Variant, _
                  Optional hideOffset As Long = 0, _
                  Optional wsStartCol As Long = 1, _
                  Optional wsData)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim endCol As Integer
    endCol = wsStartCol + UBound(dataArray, 2) + hideOffset
    
    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(1, endCol), ws.Cells(1, Columns.count))
    
    formatRange.EntireColumn.Hidden = True
            
End Sub


'************************************************
'@hideBottomRows
'************************************************
Sub hideBottomRows(dataArray As Variant, _
                   Optional start_row As Long = 1, _
                   Optional hideOffset As Long = 0, _
                   Optional wsData)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim endRow As Integer
    endRow = start_row + UBound(dataArray) + hideOffset

    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(endRow, 1), ws.Cells(Rows.count, 1))

    formatRange.Rows.Hidden = True
    
End Sub


'************************************************
'@boldFirstCol
'************************************************
Sub boldFirstCol(dataArray As Variant, _
                 Optional wsStartRow As Integer = 1, _
                 Optional wsStartCol As Integer = 1, _
                 Optional wsData, _
                 Optional headers As Boolean = True)

    Dim ws As Worksheet
    Set ws = wsReturn(wsData)
    
    Dim endRow As Long
    endRow = wsStartRow + UBound(dataArray) - 1

    Dim startRow As Long
    startRow = wsStartRow + Abs(headers)
    
    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(startRow, wsStartCol), ws.Cells(endRow, wsStartCol))
    
    formatRange.Font.Bold = True

End Sub


'************************************************
'@boldTopRow
'************************************************
Sub boldTopRow(dataArray As Variant, _
               Optional wsStartRow As Integer = 1, _
               Optional wsStartCol As Integer = 1, _
               Optional wsData)

    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim endCol As Long
    endCol = wsStartCol + UBound(dataArray, 2) - 1
    
    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(wsStartRow, wsStartCol), ws.Cells(wsStartRow, endCol))
    
    formatRange.Font.Bold = True

End Sub


'************************************************
'@boldLastRow
'************************************************
Sub boldLastRow(dataArray As Variant, _
                Optional wsStartRow As Integer = 1, _
                Optional wsStartCol As Integer = 1, _
                Optional wsData)

    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim endRow As Long
    endRow = wsStartRow + UBound(dataArray) - 1

    Dim endCol As Long
    endCol = wsStartCol + UBound(dataArray, 2) - 1

    Dim formatRange As Range
    Set formatRange = ws.Range(ws.Cells(endRow, wsStartCol), ws.Cells(endRow, endCol))

    formatRange.Font.Bold = True

End Sub


'************************************************
'@colsToNum1
'************************************************
'-Numbers No Decimal
Sub colsToNum1(dataArray, _
               colList, _
               Optional wsStartRow As Long = 1, _
               Optional wsStartCol As Long = 1, _
               Optional wsData, _
               Optional headers As Boolean = True)


    colList = headerIndex(colList, dataArray)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim i As Long
    For i = LBound(colList) To UBound(colList)
        
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol)) '%%%% Changed to nrow from ubound %%%%
        
        numRange.NumberFormat = "#,###"
    
    Next i

End Sub


'************************************************
'@colsToNum2
'************************************************
'-Numbers Two Decimal
Sub colsToNum2(dataArray, _
              colList, _
              Optional wsStartRow As Long = 1, _
              Optional wsStartCol As Long = 1, _
              Optional wsData, _
              Optional headers As Boolean = True)


    colList = headerIndex(colList, dataArray)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)
       
    Dim i As Long
    For i = LBound(colList) To UBound(colList)
    
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol)) '%%%% Changed to nrow from ubound %%%%
        
        numRange.NumberFormat = "#,###.00"
        
    Next i

End Sub


'************************************************
'@colsToNumString
'************************************************
'-Numbers become string
Sub colsToNumString(dataArray, _
                    colList, _
                    Optional wsStartRow As Long = 1, _
                    Optional wsStartCol As Long = 1, _
                    Optional wsData, _
                    Optional headers As Boolean = True)

    colList = headerIndex(colList, dataArray)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim i As Long
    For i = LBound(colList) To UBound(colList)
    
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol)) '%%%% Changed to nrow from ubound %%%%
        
        numRange.NumberFormat = "####;"
    Next i

End Sub


'************************************************
'@colsToPercent
'************************************************
Sub colsToPercent(dataArray, _
                  colList, _
                  Optional wsStartRow As Long = 1, _
                  Optional wsStartCol As Long = 1, _
                  Optional wsData, _
                  Optional headers As Boolean = True)

    colList = headerIndex(colList, dataArray)

    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim i As Long
    For i = LBound(colList) To UBound(colList)
        
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol)) '%%%% Changed to nrow from ubound %%%%
        
        numRange.NumberFormat = "0.00%"
    
    Next i

End Sub


'************************************************
'@colsToAcnt (Accounting)
'************************************************
Sub colsToAcnt(dataArray, _
                  colList, _
                  Optional wsStartRow As Long = 1, _
                  Optional wsStartCol As Long = 1, _
                  Optional wsData, _
                  Optional headers As Boolean = True)

    colList = headerIndex(colList, dataArray)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim i As Long
    For i = LBound(colList) To UBound(colList)
        
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol)) '%%%% Changed to nrow from ubound %%%%
        
        numRange.NumberFormat = "#,##0.00_);[Red](#,##0.00); "
    
    Next i

End Sub


'************************************************
'@colsPreserveText
'************************************************
Sub colsPreserveText(dataArray, _
                     colList, _
                     Optional wsStartRow As Long = 1, _
                     Optional wsStartCol As Long = 1, _
                     Optional wsData, _
                     Optional headers As Boolean = True)

    colList = headerIndex(colList, dataArray)
    
    Dim ws As Worksheet
    Set ws = wsReturn(wsData)

    Dim i As Long
    For i = LBound(colList) To UBound(colList)
        
        Dim startRow As Long
        startRow = wsStartRow + Abs(headers)
        
        Dim numCol As Long
        numCol = colList(i) + (wsStartCol - 1)
        
        Dim numRange As Range
        Set numRange = ws.Range(Cells(startRow, numCol), Cells(nrow(dataArray) + startRow, numCol))
        
        numRange.NumberFormat = "@"
    
    Next i

End Sub


'************************************************
'@wsReturn (*Private*)
'************************************************
Function wsReturn(wsData) As Worksheet

    If IsObject(wsData) = True Then
        Set wsReturn = wsData
    ElseIf IsMissing(wsData) = False Then
        Set wsReturn = Sheets(wsData)
    Else
        Set wsReturn = ActiveSheet
    End If

End Function




'==================================================================================================
'//GET DATA
'==================================================================================================


'************************************************
'@externalWbData
'************************************************
Function externalWbData(wbPath As String, _
                        Optional startRow As Long = 1, _
                        Optional startCol As Long = 1, _
                        Optional endRow As Long, _
                        Optional endCol As Long, _
                        Optional wsData)
    
    '#Workbook Setup
    Dim wb As Workbook
    Set wb = Workbooks.Open(wbPath)
    
    externalWbData = getData(wb, wsData, startRow, startCol, endRow, endCol)
    
    wb.Close
    Set wb = Nothing
    
End Function


'************************************************
'@internalWbData
'************************************************
Function internalWbData(wsData, _
                        Optional startRow As Long = 1, _
                        Optional startCol As Long = 1, _
                        Optional endRow As Long = 0, _
                        Optional endCol As Long = 0) As Variant
    
    internalWbData = getData(ActiveWorkbook, wsData, startRow, startCol, endRow, endCol)
       
End Function


'************************************************
'@getData (*Private*)
'************************************************
Private Function getData(wb As Workbook, _
                         wsData, _
                         Optional startRow As Long = 1, _
                         Optional startCol As Long = 1, _
                         Optional endRow As Long = 0, _
                         Optional endCol As Long = 0)
    
    '#Worksheet
    Dim ws As Worksheet
    If IsObject(wsData) = True Then
        Set ws = wsData
    ElseIf IsMissing(wsData) = False Then
        Set ws = Sheets(wsData)
    Else
        Set ws = ActiveSheet
    End If
    
    '#Column Count
    Dim colCount As Long
    If endCol = 0 Then
        colCount = ws.Cells(startRow, Columns.count).End(xlToLeft).Column
    Else
        colCount = endCol
    End If
    
    '#Row Count
    Dim rowCount As Long
    If endRow = 0 Then
        Dim i As Long
        For i = 1 To colCount
            Dim tempRow As Long
            tempRow = ws.Cells(Rows.count, i).End(xlUp).row
            If tempRow > rowCount Then
                rowCount = tempRow
            End If
        Next i
    Else
        rowCount = endRow
    End If
    
    '#Create Array
    getData = ws.Range(ws.Cells(startRow, startCol), ws.Cells(rowCount, colCount))
    
End Function




'==================================================================================================
'//GET VALUE
'==================================================================================================


'************************************************
'@colSum
'************************************************
Function colSum(dataArray, _
                sumColList, _
                Optional headers As Boolean = True)
    
    sumColList = headerIndex(sumColList, dataArray)

    Dim i As Long
    Dim j As Long
    Dim total As Double
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        For j = LBound(sumColList) To UBound(sumColList)
            If IsNumeric(dataArray(i, sumColList(j))) = True Then
                total = total + dataArray(i, sumColList(j))
            End If
        Next j
    Next i
    
    colSum = total

End Function


'************************************************
'@countIf
'************************************************
Function countIf(dataArray, _
                 criteria, _
                 searchCol, _
                 meetsCriteriaBy As axlCompare, _
                 Optional uboundCriteria, _
                 Optional headers As Boolean = True, _
                 Optional caseMatters As Boolean = False)
               
    searchCol = headerIndex(searchCol, dataArray)

    Dim total As Double
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        
        Dim outcome As Boolean
        outcome = compareBool(dataArray(i, searchCol), criteria, meetsCriteriaBy, _
                              uboundCriteria, caseMatters)
        
        If outcome = True Then
            total = total + 1
        End If
        
    Next i
    
    countIf = total

End Function


'************************************************
'@getElement
'************************************************
Function getElement(dataArray, _
                    col, _
                    row As Long)

    col = headerIndex(col, dataArray)
    getElement = dataArray(row, col)

End Function


'************************************************
'@listSum
'************************************************
Function listSum(dataList)
    
    Dim total As Double
    
    Dim i As Long
    For i = LBound(dataList) To UBound(dataList)
        If IsNumeric(dataList(i)) Then
            total = total + dataList(i)
        End If
    Next i

    listSum = total
    
End Function


'************************************************
'@ncol (Column Count)
'************************************************
Function ncol(dataArray)
    ncol = UBound(dataArray, 2)
End Function


'************************************************
'@nrow (Row Count)
'************************************************
Function nrow(dataArray)
    nrow = UBound(dataArray)
End Function


'************************************************
'@sumIf
'************************************************
Function sumIf(dataArray, _
               criteria, _
               searchCol, _
               sumCol, _
               meetsCriteriaBy As axlCompare, _
               Optional uboundCriteria, _
               Optional headers As Boolean = True, _
               Optional caseMatters As Boolean = False)
               
    searchCol = headerIndex(searchCol, dataArray)
    sumCol = headerIndex(sumCol, dataArray)

    Dim total As Double
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        Dim outcome As Boolean
        outcome = compareBool(dataArray(i, searchCol), criteria, meetsCriteriaBy, _
                              uboundCriteria, caseMatters)
                              
        If outcome = True Then
            total = total + dataArray(i, sumCol)
        End If
        
    Next i
    
    sumIf = total

End Function


'==================================================================================================
'//HEADER DATA
'==================================================================================================


'************************************************
'@headerIndex
'************************************************
Function headerIndex(ByVal headerData, ByVal sourceArray)

    If IsArray(headerData) = True Then
        headerIndex = lists(headerData, sourceArray)
    Else
        headerIndex = scalar(headerData, sourceArray)
    End If
    
End Function


'************************************************
'@list (*Private*)
'************************************************
Private Function lists(headerList, sourceArray)

    ReDim tempIndexList(0 To UBound(headerList))
    
    Dim i As Integer
    For i = LBound(headerList) To UBound(headerList)
        If IsNumeric(headerList(i)) = False Then
            tempIndexList(i) = scalar(headerList(i), sourceArray)
        Else
            tempIndexList(i) = headerList(i)
        End If
    Next i
    
    List = tempIndexList

End Function


'************************************************
'@scalar (*Private*)
'************************************************
'-Find header position within given name. If header name is not a string then _
  it returns the headerName give (which is the column position)
Private Function scalar(headerName, sourceArray)

    Dim returnValue As Variant
    
    If IsNumeric(headerName) = False Then
        Dim i As Integer
        For i = LBound(sourceArray, 2) To UBound(sourceArray, 2)
            If LCase(headerName) = LCase(sourceArray(1, i)) Then
                returnValue = i
                Exit For
            ElseIf i = UBound(sourceArray, 2) Then
                returnValue = -1
            End If
        Next i
    Else
        returnValue = (headerName * 1)
    End If
    
    If returnValue = -1 Then
        MsgBox "Header <" & headerName & "> not found"
    End If
    
    scalar = returnValue
    
End Function




'==================================================================================================
'//MATCH DATA
'==================================================================================================


'************************************************
'@matchDataArr
'************************************************
'-Searches external array to see if item from self _
  array item (supplied from search column) is in it
'-Returns data from selected column
'!headerIndex(), sortArray(), binarySearch()
Function matchDataArr(dataArray, _
                      searchCol, _
                      fillCol, _
                      ByVal matchArray, _
                      matchSearchCol, _
                      Optional matchGetCol, _
                      Optional matchHeaders As Boolean = True, _
                      Optional returnBool As Boolean = False, _
                      Optional headers As Boolean = True, _
                      Optional caseMatters As Boolean = False)
                      
    '#Header Index Data Array
    searchCol = headerIndex(searchCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    '#Header Index Match Array
    matchSearchCol = headerIndex(matchSearchCol, matchArray)
    
    If IsMissing(matchGetCol) = False Then
        matchGetCol = headerIndex(matchGetCol, matchArray)
    End If
    
    '#Sort Match Array
    matchArray = sortArray(matchArray, Array(matchSearchCol), True, matchHeaders, caseMatters) '<<<changed Headers to matchHeaders
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        
        Dim rowPos As Variant
        rowPos = binarySearch(dataArray(i, searchCol), matchArray, matchSearchCol, matchHeaders, _
                              caseMatters)
        
        '#Return Selection
        If rowPos <> "<!NO POSITION!>" And returnBool = False Then
            dataArray(i, fillCol) = matchArray(rowPos, matchGetCol)
        ElseIf rowPos <> "<!NO POSITION!>" And returnBool = True Then
            dataArray(i, fillCol) = True
        ElseIf rowPos = "<!NO POSITION!>" And returnBool = True Then
            dataArray(i, fillCol) = False
        End If
        
    Next i
    
    '#Return
    matchDataArr = dataArray

End Function


'************************************************
'@matchMetaArr
'************************************************
'-Searches a array that contains the metadata of array and the array described
'-example: getting the appropriate depreciation schedule
'-metaArr: column 1: id
'          column 2: headers boolean
'          column 3: array search Col
'          column 4: Array Get Col
'          column 5: match array
'-metaArr must have no headers
'!headerIndex(), binarySearch()
Function matchMetaArr(dataArray, _
                      matchMetaCol, _
                      searchCol, _
                      fillCol, _
                      ByVal metaArr, _
                      Optional returnBool As Boolean = False, _
                      Optional headers As Boolean = True, _
                      Optional caseMatters As Boolean = False)
                          
    searchCol = headerIndex(searchCol, dataArray)
    matchMetaCol = headerIndex(matchMetaCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
      
    '#Loop Variables
    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
    
        '#Search Array List
        Dim dataIndex As Variant
        dataIndex = binarySearch(dataArray(i, matchMetaCol), metaArr, 1, False, caseMatters)
    
        If dataIndex <> "<!NO POSITION!>" Then
            
            '#Assign match Array
            Dim matchedArray As Variant
            matchedArray = metaArr(dataIndex, 5)
            
            '#Assign Search Column
            Dim colSearch As Long
            colSearch = headerIndex(metaArr(dataIndex, 3), matchedArray)
            
            '#Get Column
            Dim getCol As Variant
            getCol = headerIndex(metaArr(dataIndex, 4), matchedArray)
            
            '#Sort Match Array
            matchedArray = sortArray(metaArr(dataIndex, 5), Array(colSearch), True, _
                                     CBool(metaArr(dataIndex, 2)), caseMatters)

            Dim matchPos As Variant
            matchPos = binarySearch(dataArray(i, searchCol), matchedArray, colSearch, _
                                    CBool(metaArr(dataIndex, 2)), caseMatters)
                                           
            '#Return Selection
            If matchPos <> "<!NO POSITION!>" And returnBool = False Then
                dataArray(i, fillCol) = matchedArray(matchPos, getCol)
            ElseIf matchPos <> "<!NO POSITION!>" And returnBool = True Then
                dataArray(i, fillCol) = True
            ElseIf matchPos = "<!NO POSITION!>" And returnBool = True Then
                dataArray(i, fillCol) = False
            End If
        
        End If
        
    Next i

    matchMetaArr = dataArray

End Function


'************************************************
'@matchRowCol
'************************************************
'-rowMatchCol: The column that contains the corresponding element to match the matchArray's row
'-colMatchCol: The column that contains the corresponding element to match the matchArray's column
'!headerIndx(), binarySearch()
Function matchRowCol(dataArray, _
                     rowMatchCol, _
                     colMatchCol, _
                     fillCol, _
                     matchArray, _
                     Optional matchHeaders As Boolean = True, _
                     Optional headers As Boolean = True, _
                     Optional caseMatters As Boolean = False)
                     
    rowMatchCol = headerIndex(rowMatchCol, dataArray)
    colMatchCol = headerIndex(colMatchCol, dataArray)
    fillCol = headerIndex(fillCol, dataArray)
    
    matchArray = sortArray(matchArray, Array(1), True, matchHeaders, caseMatters)

    Dim i As Long
    Dim j As Long

    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        
        Dim rowMatch As Variant
        rowMatch = binarySearch(dataArray(i, rowMatchCol), matchArray, rowMatchCol, matchHeaders, _
                                caseMatters)
        
        Dim colMatch As Variant
        colMatch = "<!NO POSITION!>"
        
        For j = LBound(matchArray, 2) To UBound(matchArray, 2)
            If dataArray(i, colMatchCol) = matchArray(1, j) Then
                colMatch = j
                Exit For
            End If
        Next j

        If rowMatch <> "<!NO POSITION!>" And colMatch <> "<!NO POSITION!>" Then
            dataArray(i, fillCol) = matchArray(rowMatch, colMatch)
        End If

    Next i

    matchRowCol = dataArray

End Function




'==================================================================================================
'//MATH
'==================================================================================================


'************************************************
'@colMath1 (Math One Column)
'************************************************
Function colMath1(dataArray, _
                  fillCol, _
                  mathColumn, _
                  operand As Double, _
                  mathOperation As axlMath, _
                  Optional switchOperationOrder As Boolean = False, _
                  Optional headers As Boolean = True)
    
    fillCol = headerIndex(fillCol, dataArray)
    mathColumn = headerIndex(mathColumn, dataArray)
    
    Dim i As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, mathColumn)) = True Then
            Dim value
            If switchOperationOrder = False Then
                value = mathOp(dataArray(i, mathColumn), operand, mathOperation)
                dataArray(i, fillCol) = value
            Else
                value = mathOp(dataArray(i, mathColumn), operand, mathOperation)
                dataArray(i, fillCol) = value
            End If
            
        End If
    Next i

    '#Return
    colMath1 = dataArray

End Function


'************************************************
'@colMath2 (Math Multiple Columns)
'************************************************
Function colMath2(dataArray, _
                  fillCol, _
                  columnList, _
                  mathOperation As axlMath, _
                  Optional headers As Boolean = True)
                           
    fillCol = headerIndex(fillCol, dataArray)
    columnList = headerIndex(columnList, dataArray)

    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        
        Dim total As Double
        total = dataArray(i, columnList(0))

        For j = LBound(columnList) + 1 To UBound(columnList)

            If IsNumeric(dataArray(i, columnList(j))) = True Then
                total = mathOp(total, dataArray(i, columnList(j)), mathOperation)
            End If
            
            dataArray(i, fillCol) = total
        
        Next j
        
    Next i
    
    '#Return
    colMath2 = dataArray

End Function


'************************************************
'@mathOp (*Private*) (Math Operations)
'************************************************
Private Function mathOp(operand_1, _
                        operand_2, _
                        mathOperation As axlMath)
                               
    Select Case mathOperation
    
        Case Is = Add
            mathOp = operand_1 + operand_2
        
        Case Is = subtract
            mathOp = operand_1 - operand_2
            
        Case Is = Multiply
            mathOp = operand_1 * operand_2
        
        Case Is = divide
            If operand_2 <> 0 Then
                mathOp = operand_1 / operand_2
            Else
                mathOp = 0
            End If
            
        Case Is = Exponent
            mathOp = operand_1 ^ operand_2
            
        Case Is = Root
            mathOp = operand_1 ^ (1 / operand_2)
    
    End Select

End Function




'==================================================================================================
'//ROW MODIFY
'==================================================================================================


'************************************************
'@addTotalRow
'************************************************
'!reshapeArray(), nrow(), ncol()
Function addTotalRow(dataArray, _
                     columnList, _
                     Optional headers As Boolean = True)

    columnList = headerIndex(columnList, dataArray)
    dataArray = reshapeArray(dataArray, nrow(dataArray) + 1, ncol(dataArray))
    
    Dim i As Long
    Dim j As Long

    '#Sum Selected Columns
    For i = LBound(columnList) To UBound(columnList)
        
        Dim total As Double
        total = 0
        
        For j = LBound(dataArray) + Abs(headers) To UBound(dataArray) - 1
            If IsNumeric(dataArray(j, columnList(i))) = True Then
                total = total + dataArray(j, columnList(i))
            End If
        Next j
        
        dataArray(UBound(dataArray), columnList(i)) = total
    
    Next i
    
    '#Return
    addTotalRow = dataArray
    
End Function


'************************************************
'@removeTopRow
'************************************************
Function removeTopRow(dataArray)
    
    ReDim tempData(1 To UBound(dataArray) - 1, 1 To UBound(dataArray, 2))
        
    Dim i As Long
    Dim j As Long
    
    For i = LBound(dataArray) + 1 To UBound(dataArray)
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            tempData(i - 1, j) = dataArray(i, j)
        Next j
    Next i
    
    removeTopRow = tempData

End Function




'==================================================================================================
'//SEARCH DATA
'==================================================================================================

'************************************************
'@binarySearch
'************************************************
'-Returns row index position or <!NO POSITION!>
'-If no match, returns -1
'-search Array MUST be sorted ascending
'!compareData()
Function binarySearch(ByVal itemToMatch, _
                      ByVal searchArray, _
                      ByVal searchCol, _
                      Optional searchHeaders As Boolean = False, _
                      Optional caseMatters As Boolean = False)
                      
    searchCol = headerIndex(searchCol, searchArray)
    
    '#Set foundPosition to a false match
    Dim foundPosition As Variant
    foundPosition = "<!NO POSITION!>"
    
    '#Account for headers in lowerbounds
    Dim low As Long
    low = 1

    If searchHeaders = True Then
        low = 2
    End If
    
    Dim high As Long
    high = UBound(searchArray)
    
    '#Begin Loop
    Do While (low <= high)

        Dim midpointIndex As Integer
        midpointIndex = (high + low) \ 2

        Dim midpointNum As Variant
        midpointNum = searchArray(midpointIndex, searchCol)

        If compareData(itemToMatch, midpointNum, caseMatters) = 1 Then
             low = midpointIndex + 1
         ElseIf compareData(itemToMatch, midpointNum, caseMatters) = -1 Then
            high = midpointIndex - 1
        Else
            foundPosition = midpointIndex
            Exit Do
        End If

    Loop

    '#Return
    binarySearch = foundPosition

End Function


'==================================================================================================
'//SORT DATA
'==================================================================================================

'************************************************
'@sortArray
'************************************************
'!headerIndex(), concate(), quickSortArray()
Function sortArray(dataArray, _
                   sortColList, _
                   Optional ascending As Boolean = True, _
                   Optional headers As Boolean = True, _
                   Optional caseMatters As Boolean = False)
    
    sortColList = headerIndex(sortColList, dataArray)
    
    '#Add concatenate column
    ReDim Preserve dataArray(1 To UBound(dataArray), 1 To UBound(dataArray, 2) + 1)
    
    '#Concatenate sort columns
    dataArray = concat(dataArray, UBound(dataArray, 2), sortColList, "/-/", headers)

    '#Sort Array
    quickSortArray dataArray, UBound(dataArray, 2), LBound(dataArray) + Abs(headers), _
                   UBound(dataArray), ascending, caseMatters
    
    '#Remove concatenated column
    ReDim Preserve dataArray(1 To UBound(dataArray), 1 To UBound(dataArray, 2) - 1)
    
    '#Return
    sortArray = dataArray
    
End Function


'************************************************
'@sortList
'************************************************
'!quickSortData()
Function sortList(dataList, _
                  Optional ascending As Boolean = True, _
                  Optional caseMatters As Boolean = False)
                  
    '#Make list Array
    dataList = bindList(Array(dataList))

    '#Sort
    quickSortArray dataList, 1, LBound(dataList), UBound(dataList), ascending, caseMatters
    
    '#Turn list back to list
    dataList = colToList(dataList, 1, False)
    
    sortList = dataList

End Function


'************************************************
'@sortMonthNames <<< CHANGE TO ARRAY AND LIST
'************************************************
'Sorts based on months
Function sortMonthNames(sortArray, _
                        monthCol, _
                        Optional ascending As Boolean = True, _
                        Optional headers As Boolean = True, _
                        Optional caseMatters As Boolean = False)
    
    monthCol = headerIndex(monthCol, dataArray)
    
    createMonthNameIndex

    '#Add Column
    ReDim Preserve sortArray(1 To UBound(sortArray), 1 To UBound(sortArray, 2) + 1)
    
    Dim i As Integer
    For i = LBound(sortArray) + Abs(headers) To UBound(sortArray)
        Dim monthName As String
        monthName = monthNameIndex.item(sortArray(i, monthCol))
        sortArray(i, UBound(sortArray, 2)) = monthName
    Next i
    
    '#Sort Month Indexes
    quickSortArray sortArray, UBound(sortArray, 2), ascending, headers, caseMatters
    
    '#Remove Month Indexes
    ReDim Preserve sortArray(1 To UBound(sortArray), 1 To UBound(sortArray, 2) - 1)
    
    '#Return
    sortMonthNames = sortArray

End Function


'************************************************
'@quickSortArray
'************************************************
'!compareData()
Private Sub quickSortArray(sortArray, _
                           sortOnCol As Long, _
                           lowerBounds As Long, _
                           upperBounds As Long, _
                           Optional ascending As Boolean = True, _
                           Optional caseMatters As Boolean = False)
    
    '#Ascending or Descending
    Dim switch As Long
    switch = 1
    
    If ascending = False Then
        switch = -1
    End If
    
    Dim pivot As Variant
    pivot = sortArray((lowerBounds + upperBounds) \ 2, sortOnCol)
    
    '#Lower Index and Num
    Dim lowerIndex As Long
    lowerIndex = lowerBounds
    
    Dim lowerNum As Variant
    lowerNum = sortArray(lowerIndex, sortOnCol)
    
    '#Upper Index and Num
    Dim upperIndex As Long
    upperIndex = upperBounds
    
    Dim upperNum As Variant
    upperNum = sortArray(upperIndex, sortOnCol)
        
    '#Sort
    Do While (lowerIndex <= upperIndex)
    
        '#Lower partition
        Do While compareData(lowerNum, pivot, caseMatters) = -1 * switch
            lowerIndex = lowerIndex + 1
            lowerNum = sortArray(lowerIndex, sortOnCol)
        Loop
        
        '#upper partition
        Do While compareData(upperNum, pivot, caseMatters) = 1 * switch
            upperIndex = upperIndex - 1
            upperNum = sortArray(upperIndex, sortOnCol)
        Loop
    
        '#Swap Data
        If (lowerIndex <= upperIndex) Then
        
            Dim i As Long
            Dim tempItem As Variant
    
            For i = LBound(sortArray, 2) To UBound(sortArray, 2)
                tempItem = sortArray(lowerIndex, i)
                sortArray(lowerIndex, i) = sortArray(upperIndex, i)
                sortArray(upperIndex, i) = tempItem
            Next i
            
            lowerIndex = lowerIndex + 1
            upperIndex = upperIndex - 1
            
            If upperIndex = 0 Then
                upperIndex = 1
            End If
            
            If lowerIndex > UBound(sortArray) Then
                lowerIndex = lowerIndex - 1
            End If
            
        End If
        
        '#Update lowerNum and upperNum
        lowerNum = sortArray(lowerIndex, sortOnCol)
        upperNum = sortArray(upperIndex, sortOnCol)
                   
    Loop

    '#Recursive Element
    If (lowerBounds < upperIndex) Then
        quickSortArray sortArray, sortOnCol, lowerBounds, upperIndex, ascending, caseMatters
    End If
    
    If (lowerIndex < upperBounds) Then
        quickSortArray sortArray, sortOnCol, lowerIndex, upperBounds, ascending, caseMatters
    End If
    
End Sub




'==================================================================================================
'//STATISTICS
'==================================================================================================


'************************************************
'@colMean
'************************************************
Function colMean(dataArray, _
                 selectCol, _
                 Optional headers As Boolean = True)
    
    selectCol = headerIndex(selectCol, dataArray)

    '#Loop Variables
    Dim i As Long
    
    '#Function Variables
    Dim total As Double
    Dim count As Double
    
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, selectCol)) Then
            total = total + dataArray(i, selectCol)
            count = count + 1
        End If
    Next i
    
    colMean = total / count

End Function


'************************************************
'@colSD (Column Standard Deviation)
'************************************************
Function colSD(dataArray, _
               selectCol, _
               Optional population As Boolean = True, _
               Optional headers As Boolean)
               
    selectCol = headerIndex(selectCol, dataArray)
    
    Dim mean As Double
    Dim squaredErr As Double
    Dim count As Double
    Dim variance As Double
    
    mean = colMean(dataArray, selectCol, headers)
    
    '#Squared Error
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        If IsNumeric(dataArray(i, selectCol)) Then
            squaredErr = squaredErr + (dataArray(i, selectCol) - mean) ^ 2
            count = count + 1
        End If
    Next i
    
    '#Variance
    If population = True Then
        variance = squaredErr * (1 / count)
    Else
        variance = squaredErr * (1 / (count - 1))
    End If
    
    colSD = variance ^ (1 / 2)
    
End Function


'************************************************
'@listMean
'************************************************
Function listMean(dataList)
    
    '#Loop Variables
    Dim i As Long
    
    '#Function Variables
    Dim total As Double
    Dim count As Double

    For i = LBound(dataList) To UBound(dataList)
        If IsNumeric(dataList(i)) Then
            total = total + dataList(i)
            count = count + 1
        End If
    Next i
    
    listMean = total / count
    
End Function


'************************************************
'@listSD (List Standard Deviation)
'************************************************
Function listSD(dataList, Optional population As Boolean = True)
    
    Dim mean As Double
    Dim squaredErr As Double
    Dim count As Double
    Dim variance As Double
    
    mean = listMean(dataList)
    
    '#Squared Error
    Dim i As Long
    For i = LBound(dataList) To UBound(dataList)
        If IsNumeric(dataList(i)) Then
            squaredErr = squaredErr + (dataList(i) - mean) ^ 2
            count = count + 1
        End If
    Next i
    
    '#Variance
    If population = True Then
        variance = squaredErr * (1 / count)
    Else
        variance = squaredErr * (1 / (count - 1))
    End If
    
    listSD = variance ^ (1 / 2)
    
End Function

    


'==================================================================================================
'//TRANSFORM
'==================================================================================================


'************************************************
'@arrayAddIndex
'************************************************
'-Add an index to array. Good for when you need to keep the original index.
'!reshapeArray()
Function arrayAddIndex(ByVal dataArray, _
                       Optional indexColName As String = "<INDEX>", _
                       Optional headers As Boolean = True)
    
    dataArray = reshapeArray(dataArray, nrow(dataArray), ncol(dataArray) + 1, Array(indexColName))
    
    Dim i As Long
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
        dataArray(i, ncol(dataArray)) = i + CLng(headers)
    Next i
    
    arrayAddIndex = dataArray

End Function


'************************************************
'@filterData
'************************************************
Function filterData(dataArray, _
                    criteria, _
                    searchCol, _
                    meetsCriteriaBy As axlCompare, _
                    Optional headers As Boolean = True, _
                    Optional caseMatters As Boolean = False)
 
    ReDim tempArray(1 To UBound(dataArray), 1 To UBound(dataArray, 2))

    searchCol = headerIndex(searchCol, dataArray)

    Dim n As Long
    Dim i As Long
    Dim j As Long
    
    '#Fill Headers
    If headers = True Then
        n = 1
        For j = LBound(tempArray, 2) To UBound(tempArray, 2)
            tempArray(n, j) = dataArray(1, j)
        Next j
    End If

    '#Create New Body Array of Criteria Match
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)
    
        Dim outcome As Boolean
        outcome = compareBool(dataArray(i, searchCol), criteria, meetsCriteriaBy, caseMatters)
                                   
        If outcome = True Then
            n = n + 1
            For j = LBound(tempArray, 2) To UBound(tempArray, 2)
                tempArray(n, j) = dataArray(i, j)
            Next j
        End If
        
    Next i

    '#Error Handler
    If n > 1 Or (n = 1 And headers = False) Then
        '#Assign to New Array
        tempArray = reshapeArray(tempArray, n, ncol(tempArray))
        dataArray = tempArray
        filterData = tempArray
    Else
        filterData = errorArray
    End If
    
End Function


'************************************************
'@groupData
'************************************************
'-Groups data on a given column. Creates a unique list based on the group col.
'-colSumList must be numeric
'!headerIndex(), concat(), uniqueArray(), binarySearch()
Function groupData(ByVal dataArray, _
                   ByVal groupColList, _
                   Optional colSumList, _
                   Optional count As Boolean = True, _
                   Optional headers As Boolean = True, _
                   Optional caseMatters As Boolean = False)

    colSumList = headerIndex(colSumList, dataArray)
    groupColList = headerIndex(groupColList, dataArray)
    
    '#Concatenate sort columns (accounts for multip column groups)
    dataArray = concat(dataArray, groupColList(0), groupColList, "/--/", False)
        
    '#Create Unique Array
    Dim tempArray As Variant
    tempArray = uniqueArray(dataArray, groupColList(0), headers, caseMatters)
    
    '#Redimension Array
    ReDim Preserve tempArray(1 To UBound(tempArray), 1 To UBound(colSumList) + 2 + Abs(count))

    '#Sum/Fill Data
    Dim i As Long
    Dim j As Long

    '#Match sum values and count List Data
    For i = LBound(dataArray) + Abs(headers) To UBound(dataArray)

        '#Find Match
        Dim rowPos As Variant
        rowPos = binarySearch(dataArray(i, groupColList(0)), tempArray, 1, headers, caseMatters)
                                  
        '#Column Sum
        For j = LBound(colSumList) To UBound(colSumList)
            If IsNumeric(dataArray(i, colSumList(j))) = True And rowPos <> "<!NO POSITION!>" Then
                tempArray(rowPos, j + 2) = tempArray(rowPos, j + 2) + dataArray(i, colSumList(j))
            End If
        Next j

        '#Count Sum
        If count And rowPos <> "<!NO POSITION!>" Then
            tempArray(rowPos, UBound(tempArray, 2)) = tempArray(rowPos, UBound(tempArray, 2)) + 1
        End If
    
    Next i

    '#Add Column List Header
    If headers = True Then
        
        '#Add Value Headers
        If IsArray(colSumList) = True Then
            For i = LBound(colSumList) To UBound(colSumList)
                tempArray(1, 2 + i) = dataArray(1, colSumList(i))
            Next i
        End If
        
        '#Add Count Header
        If count Then
            tempArray(1, UBound(tempArray, 2)) = "Count"
        End If
    
    End If
            
    '#Split Concatenated Grouped Data
    Dim groupCount As Variant
    groupCount = arrangeList(1, UBound(groupColList) + 1)
    
    tempArray = reshapeArray(tempArray, nrow(tempArray), ncol(tempArray) + UBound(groupColList))

    For i = LBound(groupColList) To UBound(groupColList) - 1
        colMove tempArray, UBound(tempArray, 2), i + 2
    Next i
    
    tempArray = splitData(tempArray, 1, groupCount, "/--/", False)

    '#Return
    groupData = tempArray
    
End Function


'************************************************
'@listAddIndex
'************************************************
Function listAddIndex(ByVal dataList)

    Dim tempArray As Variant
    tempArray = bindLists(Array(dataList))
    tempArray = reshapeArray(tempArray, nrow(tempArray), ncol(tempArray) + 1)
    
    Dim i As Long
    For i = LBound(dataList) To UBound(dataList)
        tempArray(i + 1, 2) = i
    Next i
    
    listAddIndex = tempArray

End Function


'************************************************
'@linearYearFill
'************************************************
'-Takes in two column array
'-Column 1:Years
'-Column 2:Numeric Data
'-dataArray:Data points for the linear fill (must be sorted)
Function linearYearFill(dataArray, _
                        startYear As Long, _
                        yearCount As Long, _
                        Optional percentChange As Double = 0.1, _
                        Optional descending As Boolean = True, _
                        Optional headers As Boolean = True)
    
    Dim fillArr As Variant
    fillArr = yearRangeArray(startYear, yearCount, 1, descending)
    
    '#Sort Data Points Array Descending
    dataArray = sortArray(dataArray, Array(1), Not descending, True)
    
    Dim i As Long
    For i = LBound(fillArr) To UBound(fillArr)
        
        If fillArr(i, 1) = 1986 Then
            Debug.Print 1
        End If
    
        '#fillArr Year > then newest dataArray: forward % increase
        If fillArr(i, 1) > dataArray(1 + Abs(headers), 1) Then
           fillArr(i, 2) = dataArray(1 + Abs(headers), 2) * ((1 + percentChange) _
                                     ^ (fillArr(i, 1) - dataArray(1 + Abs(headers), 1)))
        
        '#fillArr Year < then oldest data array year: backward % decrease
        ElseIf fillArr(i, 1) < dataArray(UBound(dataArray), 1) Then
            fillArr(i, 2) = dataArray(UBound(dataArray), 2) _
                            * (1 - percentChange) _
                            ^ (dataArray(UBound(dataArray), 1) - fillArr(i, 1))
        Else
            Dim j As Long
            For j = LBound(dataArray) + Abs(headers) To UBound(dataArray)
                
                '#See if fillArr year = dataArray year
                If fillArr(i, 1) = Int(dataArray(j, 1)) Then
                   fillArr(i, 2) = dataArray(j, 2)
                   Exit For
                
                '#Check to see if fillArr falls between two data points
                ' from dataArray
                ElseIf (fillArr(i, 1) < Int(dataArray(j, 1)) _
                        And fillArr(i, 1) > Int(dataArray(j + 1, 1))) Then
                    
                    Dim costIncr As Double
                    costIncr = (dataArray(j, 2) - dataArray(j + 1, 2)) _
                               / (dataArray(j, 1) - dataArray(j + 1, 1))
                    
                    fillArr(i, 2) = costIncr _
                                    * (fillArr(i, 1) _
                                    - dataArray(j + 1, 1)) _
                                    + dataArray(j + 1, 2)
                    Exit For
                End If
            Next j
        End If
    Next i
    
    linearYearFill = fillArr
    
End Function


'************************************************
'@missingCategories (Fill Missing Categories)
'************************************************
'-Adds missing category names to array of groupd data (*must already be grouped).
'-dataArray must be sorted
'-catArray must be 2D and starting index must be 1 (not 0)
'-Example: An array grouped on months thats missing July or August.
'!headerIndex(),reshapeArray(),matchGetBool(),filterData()
Function missingCategories(ByVal dataArray, _
                           dataCol, _
                           categoryArr, _
                           categoryDataCol, _
                           Optional catArrHeaders As Boolean = False, _
                           Optional headers As Boolean = True, _
                           Optional caseMatters As Boolean = False)
   
    dataCol = headerIndex(dataCol, dataArray)
    categoryDataCol = headerIndex(categoryDataCol, categoryArr)

    '#Match Category Array
    categoryArr = reshapeArray(categoryArr, nrow(categoryArr), ncol(categoryArr) + 1)

    categoryArr = matchGetBool(categoryArr, 1, ncol(categoryArr), dataArray, dataCol, headers, _
                               catArrHeaders, caseMatters)
    
    '#Filter Out True
    categoryArr = filterData(categoryArr, False, ncol(categoryArr), EqualTo, catArrHeaders)
    
    '#Create New Array
    ReDim finalArr(1 To nrow(dataArray) + nrow(categoryArr), 1 To ncol(dataArray))
    
    '#Transfer Data Array
    Dim i As Long
    Dim j As Long
    For i = LBound(dataArray) To UBound(dataArray)
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            finalArr(i, j) = dataArray(i, j)
        Next j
    Next i

    '#Transfer Category Array
    For i = LBound(categoryArr) To UBound(categoryArr)
        finalArr(i + nrow(dataArray), dataCol) = categoryArr(i, categoryDataCol)
    Next i

    missingCategories = sortArray(finalArr, Array(dataCol), , headers, caseMatters)
    
End Function


'************************************************
'@reshapeArray
'************************************************
Function reshapeArray(dataArray, _
                      rowCount As Long, _
                      columnCount As Long, _
                      Optional addedColumnNames)
    
    Dim i As Long
    Dim j As Long
    
    '#Column Reshape
    Dim originalCount As Long
    originalCount = UBound(dataArray, 2)
    
    ReDim Preserve dataArray(1 To UBound(dataArray), 1 To columnCount)
    
    '#Fill Headers
    If IsArray(addedColumnNames) = True Then
        For i = LBound(addedColumnNames) To UBound(addedColumnNames)
            dataArray(1, originalCount + 1 + i) = addedColumnNames(i)
        Next i
    End If

    '#Row Reshape
    ReDim tempArray(1 To rowCount, 1 To UBound(dataArray, 2))
    
    Dim upperBound As Long
    upperBound = rowCount
    
    If upperBound > UBound(dataArray) Then
        upperBound = UBound(dataArray)
    End If
    
    For i = LBound(dataArray) To upperBound
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            tempArray(i, j) = dataArray(i, j)
        Next j
    Next i

    '#Return
    reshapeArray = tempArray

End Function


'************************************************
'@transpose (Transpose Array)
'************************************************
Function transpose(dataArray) As Variant

    Dim i As Integer
    Dim j As Integer
    
    ReDim tempArray(LBound(dataArray, 2) To UBound(dataArray, 2), _
                    LBound(dataArray) To UBound(dataArray))
    
    For i = LBound(dataArray) To UBound(dataArray)
        For j = LBound(dataArray, 2) To UBound(dataArray, 2)
            tempArray(j, i) = dataArray(i, j)
        Next j
    Next i

    transpose = tempArray
    
End Function





