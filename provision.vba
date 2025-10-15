Option Explicit

Public Sub BuildProvisionReports()
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsMapping As Worksheet
    Dim dictGL As Object, dictData As Object
    Dim dictMonthsGlobal As Object
    Dim GLCode As String, GLDesc As String
    Dim PC As String, PCDesc As String
    Dim PostingKey As String
    Dim Amount As Double
    Dim DocDate As Date, MonthKey As String
    Dim r As Long, lastRow As Long
    
    Set wb = ActiveWorkbook
    Set wsData = ActiveSheet
    
    ' --- GL Mapping in Personal.xlsb ---
    On Error Resume Next
    Set wsMapping = Nothing
    Set wsMapping = ThisWorkbook.Sheets("GL_Mapping")
    If wsMapping Is Nothing Then
        MsgBox "GL_Mapping sheet not found in Personal.xlsb.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Load existing GL mapping
    Set dictGL = CreateObject("Scripting.Dictionary")
    Dim mapLastRow As Long
    mapLastRow = wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row
    For r = 2 To mapLastRow
        GLCode = Trim(wsMapping.Cells(r, 1).Value)
        GLDesc = wsMapping.Cells(r, 2).Value
        If GLCode <> "" Then dictGL(GLCode) = GLDesc
    Next r
    
    ' --- Prepare in-memory data dictionary ---
    Set dictData = CreateObject("Scripting.Dictionary")
    Set dictMonthsGlobal = CreateObject("Scripting.Dictionary")
    
    ' Map headers
    Dim hdrMap As Object
    Set hdrMap = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long, c As Long
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hdrMap(LCase(Trim(wsData.Cells(1, c).Value))) = c
    Next c
    
    Dim hDocDate As Long, hPCDesc As Long, hPC As Long, hPostingKey As Long, hAmount As Long, hOffset As Long
    hDocDate = hdrMap("document date")
    hPCDesc = hdrMap("profit center: short text")
    hPC = hdrMap("profit center")
    hPostingKey = hdrMap("posting key")
    hAmount = hdrMap("company code currency value")
    hOffset = hdrMap("offsetting account")
    
    If hDocDate = 0 Or hPCDesc = 0 Or hPC = 0 Or hPostingKey = 0 Or hAmount = 0 Or hOffset = 0 Then
        MsgBox "One or more required headers not found.", vbCritical
        Exit Sub
    End If
    
    ' --- Collect unique months & read data into dictionary ---
    lastRow = wsData.Cells(wsData.Rows.Count, hDocDate).End(xlUp).Row
    Dim key As Variant
    Dim tmpGLDesc As String, tmpPC As String, tmpPCDesc As String
    Dim tmpPostingKey As String, tmpAmount As Double
    Dim tmpDocDate As Date, tmpMonthKey As String
    Dim newGLDesc As String
    
    For r = 2 To lastRow
        If Not IsDate(wsData.Cells(r, hDocDate).Value) Then GoTo NextRow
        tmpDocDate = wsData.Cells(r, hDocDate).Value
        tmpMonthKey = Format(tmpDocDate, "mm-yyyy")
        dictMonthsGlobal(tmpMonthKey) = 1
        
        GLCode = Trim(wsData.Cells(r, hOffset).Value)
        If GLCode = "" Then GoTo NextRow
        
        tmpPC = wsData.Cells(r, hPC).Value
        tmpPCDesc = wsData.Cells(r, hPCDesc).Value
        tmpPostingKey = wsData.Cells(r, hPostingKey).Value
        tmpAmount = wsData.Cells(r, hAmount).Value
        
        ' Adjust Amount by PostingKey
        Select Case tmpPostingKey
            Case "50": tmpAmount = Abs(tmpAmount)
            Case "40": tmpAmount = -Abs(tmpAmount)
            Case Else: GoTo NextRow
        End Select
        
        ' Get GL Description, prompt only if truly new
        If dictGL.exists(GLCode) Then
            tmpGLDesc = dictGL(GLCode)
        Else
            newGLDesc = InputBox("Enter description for new GL Code: " & GLCode, "New GL Code")
            If newGLDesc = "" Then newGLDesc = GLCode
            dictGL(GLCode) = newGLDesc
            Dim newRow As Long
            newRow = wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row + 1
            wsMapping.Cells(newRow, 1).Value = GLCode
            wsMapping.Cells(newRow, 2).Value = newGLDesc
            tmpGLDesc = newGLDesc
        End If
        
        ' Key = GLDesc | ProfitCenter
        key = tmpGLDesc & "|" & tmpPC
        If Not dictData.exists(key) Then
            Set dictData(key) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Sum amounts per month
        If dictData(key).exists(tmpMonthKey) Then
            dictData(key)(tmpMonthKey) = dictData(key)(tmpMonthKey) + tmpAmount
        Else
            dictData(key)(tmpMonthKey) = tmpAmount
        End If
        
NextRow:
    Next r
    
    ' --- Create GL sheets & write 3-row blocks ---
    Dim wsGL As Worksheet, pcRowPosted As Long, pcRowReversed As Long, pcRowBalance As Long
    Dim dictSheetMonths As Object
    Set dictSheetMonths = CreateObject("Scripting.Dictionary")
    
    Dim parts() As String
    Dim month As Variant, colNum As Long
    Dim monthsDict As Object
    Dim monthList As Variant
    Dim m As Variant
    
    For Each key In dictData.Keys
        parts = Split(key, "|")
        tmpGLDesc = parts(0)
        tmpPC = parts(1)
        
        ' Create or activate GL sheet
        On Error Resume Next
        Set wsGL = Nothing
        Set wsGL = wb.Sheets(tmpGLDesc)
        If wsGL Is Nothing Then
            Set wsGL = wb.Sheets.Add
            wsGL.Name = tmpGLDesc
            wsGL.Range("A1").Value = "Profit Center"
            wsGL.Range("B1").Value = "Type"
        End If
        On Error GoTo 0
        
        ' Find or add Profit Center 3-row block
        pcRowPosted = 0
        Dim foundRow As Long
        Dim lastUsedRow As Long
        lastUsedRow = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row
        If lastUsedRow < 1 Then lastUsedRow = 1
        For foundRow = 2 To lastUsedRow
            If wsGL.Cells(foundRow, 1).Value = tmpPC Then
                pcRowPosted = foundRow
                Exit For
            End If
        Next foundRow
        If pcRowPosted = 0 Then
            pcRowPosted = wsGL.Cells(wsGL.Rows.Count, 1).End(xlUp).Row + 1
            If pcRowPosted < 2 Then pcRowPosted = 2
            ' Add blank row before new profit center (except for the first one)
            If pcRowPosted > 2 Then pcRowPosted = pcRowPosted + 1
            wsGL.Cells(pcRowPosted, 1).Value = tmpPC
        End If
        ' Always ensure the Profit Center value is set for all 3 rows
        wsGL.Cells(pcRowPosted, 1).Value = tmpPC
        wsGL.Cells(pcRowPosted + 1, 1).Value = tmpPC
        wsGL.Cells(pcRowPosted + 2, 1).Value = tmpPC
        ' Always ensure the Type labels are set for all 3 rows
        wsGL.Cells(pcRowPosted, 2).Value = "Posted"
        wsGL.Cells(pcRowPosted + 1, 2).Value = "Reversed"
        wsGL.Cells(pcRowPosted + 2, 2).Value = "Balance"
        pcRowReversed = pcRowPosted + 1
        pcRowBalance = pcRowPosted + 2
        
        ' Build month columns for sheet
        If Not dictSheetMonths.exists(tmpGLDesc) Then
            Set monthsDict = CreateObject("Scripting.Dictionary")
            monthList = dictMonthsGlobal.Keys
            ' Sort months chronologically
            QuickSortMonths monthList, LBound(monthList), UBound(monthList)
            colNum = 3
            For Each m In monthList
                wsGL.Cells(1, colNum).Value = m
                monthsDict(m) = colNum
                colNum = colNum + 1
            Next m
            ' Add Total column at the end
            wsGL.Cells(1, colNum).Value = "Total"
            monthsDict("Total") = colNum
            Set dictSheetMonths(tmpGLDesc) = monthsDict
        End If
        
        Set monthsDict = dictSheetMonths(tmpGLDesc)
        
        ' Fill Posted/Reversed/Balance and track totals
        Dim totalPosted As Double, totalReversed As Double
        totalPosted = 0
        totalReversed = 0
        
        For Each month In dictData(key).Keys
            colNum = monthsDict(month)
            If dictData(key)(month) > 0 Then
                wsGL.Cells(pcRowPosted, colNum).Value = Nz(wsGL.Cells(pcRowPosted, colNum).Value) + dictData(key)(month)
                totalPosted = totalPosted + dictData(key)(month)
            Else
                wsGL.Cells(pcRowReversed, colNum).Value = Nz(wsGL.Cells(pcRowReversed, colNum).Value) + dictData(key)(month)
                totalReversed = totalReversed + dictData(key)(month)
            End If
            wsGL.Cells(pcRowBalance, colNum).Value = Nz(wsGL.Cells(pcRowPosted, colNum).Value) + Nz(wsGL.Cells(pcRowReversed, colNum).Value)
        Next month
        
        ' Fill Total column
        Dim totalColNum As Long
        totalColNum = monthsDict("Total")
        wsGL.Cells(pcRowPosted, totalColNum).Value = Nz(wsGL.Cells(pcRowPosted, totalColNum).Value) + totalPosted
        wsGL.Cells(pcRowReversed, totalColNum).Value = Nz(wsGL.Cells(pcRowReversed, totalColNum).Value) + totalReversed
        wsGL.Cells(pcRowBalance, totalColNum).Value = Nz(wsGL.Cells(pcRowPosted, totalColNum).Value) + Nz(wsGL.Cells(pcRowReversed, totalColNum).Value)
    Next key
    
    ' --- Build Summary Sheet ---
    Dim wsSummary As Worksheet
    On Error Resume Next
    Set wsSummary = Nothing
    Set wsSummary = wb.Sheets("Summary")
    If wsSummary Is Nothing Then
        Set wsSummary = wb.Sheets.Add
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Build unique lists of GL Accounts and Profit Centers
    Dim dictGLAccounts As Object, dictProfitCenters As Object
    Set dictGLAccounts = CreateObject("Scripting.Dictionary")
    Set dictProfitCenters = CreateObject("Scripting.Dictionary")
    
    For Each key In dictData.Keys
        parts = Split(key, "|")
        tmpGLDesc = parts(0)
        tmpPC = parts(1)
        dictGLAccounts(tmpGLDesc) = 1
        dictProfitCenters(tmpPC) = 1
    Next key
    
    ' Sort GL Accounts alphabetically
    Dim glArray As Variant
    glArray = dictGLAccounts.Keys
    If dictGLAccounts.Count > 0 Then
        QuickSortStrings glArray, LBound(glArray), UBound(glArray)
    End If
    
    ' Build header row: Profit Center | GL1-Posted | GL1-Reversed | GL1-Balance | GL2-Posted | ...
    wsSummary.Cells(1, 1).Value = "Profit Center"
    colNum = 2
    Dim glAccount As Variant
    Dim dictGLColumns As Object
    Set dictGLColumns = CreateObject("Scripting.Dictionary")
    
    For Each glAccount In glArray
        wsSummary.Cells(1, colNum).Value = glAccount & " - Posted"
        wsSummary.Cells(1, colNum + 1).Value = glAccount & " - Reversed"
        wsSummary.Cells(1, colNum + 2).Value = glAccount & " - Balance"
        ' Store column positions for each GL Account
        Set dictGLColumns(glAccount) = CreateObject("Scripting.Dictionary")
        dictGLColumns(glAccount)("Posted") = colNum
        dictGLColumns(glAccount)("Reversed") = colNum + 1
        dictGLColumns(glAccount)("Balance") = colNum + 2
        colNum = colNum + 3
    Next glAccount
    
    ' Build data rows - one row per Profit Center
    Dim pcArray As Variant
    pcArray = dictProfitCenters.Keys
    If dictProfitCenters.Count > 0 Then
        QuickSortStrings pcArray, LBound(pcArray), UBound(pcArray)
    End If
    
    Dim summaryRow As Long
    summaryRow = 2
    Dim profitCenter As Variant
    Dim glPostedCol As Long, glReversedCol As Long, glBalanceCol As Long
    Dim totalPostedVal As Double, totalReversedVal As Double
    
    For Each profitCenter In pcArray
        wsSummary.Cells(summaryRow, 1).Value = profitCenter
        
        ' For each GL Account, calculate aggregated Posted, Reversed, Balance
        For Each glAccount In glArray
            key = glAccount & "|" & profitCenter
            
            If dictData.exists(key) Then
                totalPostedVal = 0
                totalReversedVal = 0
                
                ' Sum all months for this GL+PC combination
                For Each month In dictData(key).Keys
                    If dictData(key)(month) > 0 Then
                        totalPostedVal = totalPostedVal + dictData(key)(month)
                    Else
                        totalReversedVal = totalReversedVal + dictData(key)(month)
                    End If
                Next month
                
                glPostedCol = dictGLColumns(glAccount)("Posted")
                glReversedCol = dictGLColumns(glAccount)("Reversed")
                glBalanceCol = dictGLColumns(glAccount)("Balance")
                
                ' Write values to Summary sheet (hyperlinks removed)
                ' glAccount parameter is the GL sheet name, kept for potential future formula reference use
                If totalPostedVal <> 0 Then
                    AddCellReferenceFormula wsSummary, summaryRow, glPostedCol, glAccount, totalPostedVal
                End If
                
                If totalReversedVal <> 0 Then
                    AddCellReferenceFormula wsSummary, summaryRow, glReversedCol, glAccount, totalReversedVal
                End If
                
                ' Balance = Posted + Reversed
                Dim balanceVal As Double
                balanceVal = totalPostedVal + totalReversedVal
                If balanceVal <> 0 Then
                    AddCellReferenceFormula wsSummary, summaryRow, glBalanceCol, glAccount, balanceVal
                End If
            End If
        Next glAccount
        
        summaryRow = summaryRow + 1
    Next profitCenter
    
    MsgBox "Provision GL processing completed."
    
End Sub

' --- Helper function to sort month strings chronologically (mm-yyyy) ---
Sub QuickSortMonths(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While CDate("01-" & arr(i)) < CDate("01-" & pivot): i = i + 1: Loop
        Do While CDate("01-" & arr(j)) > CDate("01-" & pivot): j = j - 1: Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortMonths arr, first, j
    If i < last Then QuickSortMonths arr, i, last
End Sub

' --- Helper function to sort strings alphabetically ---
Sub QuickSortStrings(arr As Variant, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As String, temp As String
    i = first
    j = last
    pivot = arr((first + last) \ 2)
    Do While i <= j
        Do While StrComp(arr(i), pivot, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(arr(j), pivot, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then QuickSortStrings arr, first, j
    If i < last Then QuickSortStrings arr, i, last
End Sub

' --- Nz helper ---
Function Nz(val As Variant) As Double
    If IsEmpty(val) Or IsNull(val) Or val = "" Then
        Nz = 0
    Else
        Nz = val
    End If
End Function

' --- Helper to set cell values without hyperlinks ---
Sub AddCellReferenceFormula(ws As Worksheet, cellRow As Long, cellCol As Long, sheetName As String, displayValue As Variant)
    ' This function sets simple cell values without hyperlinks
    ' The sheetName parameter (which receives the GL Account/Sheet name) is kept for API compatibility
    ' It could be used in the future if formula references like =SheetName!B1 are desired
    ' For now, it's unused - users can manually adjust cells to add formula references as needed
    On Error Resume Next
    ' Remove any existing hyperlink from this specific cell only
    ws.Cells(cellRow, cellCol).Hyperlinks.Delete
    ' Set the value (not a formula) - hyperlinks are removed
    ws.Cells(cellRow, cellCol).Value = displayValue
    On Error GoTo 0
End Sub
