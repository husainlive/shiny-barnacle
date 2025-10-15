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
    
    wsSummary.Cells(1, 1).Value = "GL Account"
    wsSummary.Cells(1, 2).Value = "Profit Center"
    wsSummary.Cells(1, 3).Value = "Type"
    
    ' Sort months chronologically for Summary sheet
    Dim sortedMonths As Variant
    sortedMonths = dictMonthsGlobal.Keys
    QuickSortMonths sortedMonths, LBound(sortedMonths), UBound(sortedMonths)
    
    colNum = 4
    For Each month In sortedMonths
        wsSummary.Cells(1, colNum).Value = month
        colNum = colNum + 1
    Next month
    ' Add Total column
    wsSummary.Cells(1, colNum).Value = "Total"
    Dim totalColIdx As Long
    totalColIdx = colNum
    
    Dim rowOut As Long: rowOut = 2
    Dim rowPosted As Long, rowReversed As Long, rowBalance As Long
    For Each key In dictData.Keys
        parts = Split(key, "|")
        tmpGLDesc = parts(0)
        tmpPC = parts(1)
        
        rowPosted = rowOut
        rowReversed = rowOut + 1
        rowBalance = rowOut + 2
        
        wsSummary.Cells(rowPosted, 1).Value = tmpGLDesc
        wsSummary.Cells(rowPosted, 2).Value = tmpPC
        wsSummary.Cells(rowPosted, 3).Value = "Posted"
        wsSummary.Cells(rowReversed, 1).Value = tmpGLDesc
        wsSummary.Cells(rowReversed, 2).Value = tmpPC
        wsSummary.Cells(rowReversed, 3).Value = "Reversed"
        wsSummary.Cells(rowBalance, 1).Value = tmpGLDesc
        wsSummary.Cells(rowBalance, 2).Value = tmpPC
        wsSummary.Cells(rowBalance, 3).Value = "Balance"
        
        ' Fill in the month data using sorted months and track totals
        Dim sumPosted As Double, sumReversed As Double
        sumPosted = 0
        sumReversed = 0
        
        colNum = 4
        For Each month In sortedMonths
            If dictData(key).exists(month) Then
                If dictData(key)(month) > 0 Then
                    wsSummary.Cells(rowPosted, colNum).Value = Nz(wsSummary.Cells(rowPosted, colNum).Value) + dictData(key)(month)
                    sumPosted = sumPosted + dictData(key)(month)
                Else
                    wsSummary.Cells(rowReversed, colNum).Value = Nz(wsSummary.Cells(rowReversed, colNum).Value) + dictData(key)(month)
                    sumReversed = sumReversed + dictData(key)(month)
                End If
                wsSummary.Cells(rowBalance, colNum).Value = Nz(wsSummary.Cells(rowPosted, colNum).Value) + Nz(wsSummary.Cells(rowReversed, colNum).Value)
            End If
            colNum = colNum + 1
        Next month
        
        ' Fill Total column
        wsSummary.Cells(rowPosted, totalColIdx).Value = Nz(wsSummary.Cells(rowPosted, totalColIdx).Value) + sumPosted
        wsSummary.Cells(rowReversed, totalColIdx).Value = Nz(wsSummary.Cells(rowReversed, totalColIdx).Value) + sumReversed
        wsSummary.Cells(rowBalance, totalColIdx).Value = Nz(wsSummary.Cells(rowPosted, totalColIdx).Value) + Nz(wsSummary.Cells(rowReversed, totalColIdx).Value)
        
        rowOut = rowOut + 3
        ' Add blank row after each profit center group
        rowOut = rowOut + 1
    Next key
    
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

' --- Nz helper ---
Function Nz(val As Variant) As Double
    If IsEmpty(val) Or IsNull(val) Or val = "" Then
        Nz = 0
    Else
        Nz = val
    End If
End Function
