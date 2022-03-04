Attribute VB_Name = "GTINChecker_v006"
Option Explicit

'set our datasource.  Preconfigured on users's system
'Const strCon = "DSN=Netezza"
Const strCon = "DSN=Snow"

Sub GTIN_Netezza(ByRef AWB As Workbook)
'******************************************************************************
' Originally written by Joe Hubbard 01/01/2016ish
' Check UPCs on sheet against a DB of "in-use" UPCs
' written for access first, modified to use BI's Netezza DB by MH 06/29/2016
'
'
' Read UPCs off of the AC sheet
'   "Clean" them
'   Build SQL string (ignore "internals")
'       "Select gtin, sku from table where gtin in ('row11UPC', 'row12upc',...)
'   Check for Dupes
'   Check for UPCs in use
'   Allow for timeouts
'
'   Write results to sheet.
'
'******************************************************************************

Dim cn As Object  'ADODB.Connection 'This is the direct DB connection (late Bound)
'Dim strCon As String 'our connection as a string
Dim i As Long
Dim j As Long  'holds num of upc's for loop later
Dim k As Long
Dim SourceArray As Variant  'The GTINs we want to check
Dim NumUPCs As Long
Dim UPCSheet As Worksheet
Dim rs As Object 'ADODB.Recordset 'Gives us a record to return DB results
Dim rsArray As Variant 'array to shift rows.
Dim PFULast As Long         'Last row on the PFU Sheet
Dim pfusheet As Worksheet
Dim ACSheet As Worksheet
Dim UPCsInUse As Boolean
Dim UPCsDUPES As Boolean
Dim errMsg As String
Dim rowDic As Object
Dim selSQL As String
Dim starttime As Date
Dim selectfromtimer As String

    Set ACSheet = AWB.Worksheets("Article Create")
    'Set AWB = ActiveWorkbook

    Set pfusheet = AWB.Worksheets("PFUs")
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    Set rowDic = CreateObject("Scripting.dictionary")
    
    'PFULast = pfusheet.Range("A" & Rows.Count).End(xlUp).row + 1
    NumUPCs = 0
    errMsg = "Likely Timed out checking UPCs.  Please validate them manually."
    
    'Start by pulling our UPC's from AC sheet
    'Find out last populated UPC row  -simple count loop
    
    i = 0
    Do While ACSheet.Range("Q" & 11 + i).Value <> ""
       i = i + 1
    Loop
        
    'Populate our array from the template, UPCs and a "helper"
    SourceArray = ACSheet.Range("Q11:R" & i + 10).Value
       
    'remove spaces, dashes, periods so it they will be "clean" to write back
    'to the sheet.
    For i = LBound(SourceArray) To UBound(SourceArray)
        SourceArray(i, 1) = Replace(SourceArray(i, 1), " ", "")
        SourceArray(i, 1) = Replace(SourceArray(i, 1), "-", "")
        SourceArray(i, 1) = Replace(SourceArray(i, 1), ".", "")
        SourceArray(i, 2) = ""
    Next i
    
    'write back to the sheet cleaned UPCs
    ACSheet.Range("Q11:Q" & UBound(SourceArray, 1) + 10).Value = SourceArray
    ACSheet.Range("R:R").Calculate
    
    'set our datasource.  Preconfigured on users's system
    'strCon = "DSN=Netezza"
    'strCon = "DSN=Snow"
    
    'start our "Select" sql
    selSQL = "Select ppg.gtin, ppg.sku " & vbLf & _
        "from PRODUCT.PRODUCT.GTIN_SKU as PPG " & vbLf & _
        "where PPG.M_INSERT_USER != 'MIGRATION - Conversion from Legacy Tables' " & vbLf & _
        "and PPG.GTIN_SKU_ACTIVE_FLAG = 1 and PPG.GTIN in ("
    
    
    'loop through our array.  Do lots of things.
    'Check for duplicates
    'color AC sheet for duplicates
    'Build "where in" sql string
    'Add to Dictionary for Row number memory (select blah from blah where in)
    
    'check first for duplicates by populating (i,2) of our source array
    'We'll fill in SourceArray with "DUPE" if it is in our array more than once.
    'and then immediately color those cells red.
    i = 0
    For i = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If SourceArray(i, 2) <> "DUPE" Then ' And UCase(SourceArray(i, 1)) <> "INTERNAL" Then
            'Add to row dic for select sql
            If Not rowDic.Exists(SourceArray(i, 1)) Then
                rowDic.Add SourceArray(i, 1), i + 10 'UPC and "row"
            End If
            'add to select sql string if it is not Internal
            If UCase(Trim(SourceArray(i, 1))) <> "INTERNAL" Then
                selSQL = selSQL & "'" & SourceArray(i, 1) & "', "
                NumUPCs = NumUPCs + 1
            End If
           
            'check to see if it is a duplimacated
            For k = i + 1 To UBound(SourceArray, 1)
                If SourceArray(i, 1) = SourceArray(k, 1) And _
                    UCase(SourceArray(i, 1)) <> "INTERNAL" Then
                    SourceArray(i, 2) = "DUPE"
                    ACSheet.Range("Q" & i + 10).Interior.Color = 255
                    SourceArray(k, 2) = "DUPE"
                    ACSheet.Range("Q" & k + 10).Interior.Color = 255
                    UPCsDUPES = True
                End If
            Next k
        End If '/not "DUPE"
    Next i
        
    'finish select sql
    selSQL = Trim(selSQL)
    selSQL = Left(selSQL, Len(selSQL) - 1) & ");"
    
    
    'sql should be ready.
    
    'skip query if numUPCs = 0
    If NumUPCs > 0 Then
        On Error Resume Next
        'opens connection to Netezza
        cn.Open strCon
        If Error <> "" Then GoTo Timeout
        On Error GoTo 0
    
        'start timer for select sql
        starttime = Now
    
        Set rs.ActiveConnection = cn
        rs.Open selSQL
        
        'stop timer for select sql
        selectfromtimer = Format(Now - starttime, "hh:mm:ss")
    
        ACSheet.Range("W4").Value = selectfromtimer
        ACSheet.Range("V4").Value = "Select From Timer"
    
        'convert record set to an array to limit touch time on DB to prevent records from being
        If Not rs.EOF Then rsArray = rs.GetRows
    
        'close connection
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing

        'If we had UPCs in use, Let's drop a summary out
        If Not IsEmpty(rsArray) Then
            UPCsInUse = True
            On Error Resume Next
            Set UPCSheet = Sheets("UPC Check")
            On Error GoTo 0
            If Not UPCSheet Is Nothing Then
                If Not UPCSheet.Visible = xlSheetVisible Then
                    UPCSheet.Visible = xlSheetVisible
                End If
                UPCSheet.Select
            Else
                AWB.Sheets.Add().Name = "UPC Check"
                Set UPCSheet = Sheets("UPC Check")
                AWB.Sheets("UPC Check").Select
            End If
         
        'Clear contents just in case something got left there
            UPCSheet.Cells.ClearContents
            'Create some ghetto column headers
            UPCSheet.Range("A1").Value = "Line"
            UPCSheet.Range("B1").Value = " - "
            UPCSheet.Range("C1").Value = " _ _UPC_ _ "
            UPCSheet.Range("D1").Value = " - "
            UPCSheet.Range("E1").Value = "In use on article"
            UPCSheet.Range("F1").Value = "Article on AM Tab?"
            
            'Drop out anything where we found a match
            j = 2 'row to insert to in ++'d per iteration
                      
            For i = 0 To UBound(rsArray, 2)
                If rsArray(0, i) <> "" Then
                    UPCSheet.Range("A" & j).Value = rowDic(rsArray(0, i))
                    UPCSheet.Range("B" & j).Value = " - "
                    UPCSheet.Range("C" & j).Value = CDbl(rsArray(0, i))
                    UPCSheet.Range("D" & j).Value = " - "
                    UPCSheet.Range("E" & j).Value = rsArray(1, i)
                    'Could probably make this better to see if we are actually doing UPC work for it,
                    'but we'll leave as is for now
                    UPCSheet.Range("F" & j).FormulaR1C1 = _
                        "=OR(IFERROR(COUNTIF('Maintain Article'!C[-5],'UPC Check'!RC[-1])>0,FALSE),IFERROR(COUNTIF('Maintain Article'!R9C15:R300C15,'UPC Check'!RC[-1])>0,FALSE))"
                    j = j + 1
                    ACSheet.Range("Q" & rowDic(rsArray(0, i))).Interior.Color = 255
                End If
            Next i
            UPCSheet.Range("F:F").Calculate
            UPCSheet.Range("F:F").Value = Range("F:F").Value
            UPCSheet.Cells.EntireColumn.AutoFit
        End If 'rs array not empty
    
    End If '/Num UPCS > 0 so we wanted to query netezza
    
    'Log dupes on UPC check sheet.
    If UPCsDUPES = True Then
        On Error Resume Next
            Set UPCSheet = Sheets("UPC Check")
        On Error GoTo 0
        If UPCSheet Is Nothing Then
            AWB.Sheets.Add().Name = "UPC Check"
            Set UPCSheet = Sheets("UPC Check")
        End If
        With UPCSheet
            'Create some ghetto column headers
            .Range("H1").Value = "Line"
            .Range("I1").Value = " - "
            .Range("J1").Value = " _ _UPC_ _ "
            .Range("K1").Value = " - "
            
            'sets counter i for proper output line.
            i = 2
            
            For k = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                If SourceArray(k, 2) = "DUPE" Then
                    .Range("H" & i).Value = k + 10
                    .Range("I" & i).Value = "-"
                    .Range("J" & i).Value = SourceArray(k, 1)
                    .Range("K" & i).Value = "-"
                    .Range("L" & i).Value = "Duplicate UPC"
                    i = i + 1
                End If
            Next k
        
            .Cells.EntireColumn.AutoFit
        End With
    End If '/end dupes

    PFULast = pfusheet.Range("A" & Rows.Count).End(xlUp).row + 1
    
    'Log some stuff to PFUs
    If UPCsDUPES = True Then
        UPCsInUse = True
        pfusheet.Range("F" & PFULast).Value = "Duplicate UPC's on AC Sheet"
    End If
    
    If UPCsDUPES Or UPCsInUse Then
        With UPCSheet.Sort
            .SortFields.Clear
            .SortFields.Add key:=Range("A2:A" & Rows.Count), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal
            .SetRange Columns("A:F")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

    pfusheet.Range("A" & PFULast).Value = "GTINS/UPCs"
    pfusheet.Range("B" & PFULast).Value = UPCsInUse
    If UPCsInUse = True Then
        pfusheet.Range("C" & PFULast).Value = "Check the ""UPC Check"" sheet"
        UPCSheet.Columns("B:E").NumberFormat = "0"
    End If
Timeout:
    AWB.Worksheets("Article Create").activate
    If Error <> "" Then
        MsgBox Error & vbLf & errMsg
        On Error Resume Next
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
        On Error GoTo 0
    End If
    On Error Resume Next
    Erase rsArray
    Erase SourceArray
    On Error GoTo 0
    Set rowDic = Nothing
    Set UPCSheet = Nothing
    Set pfusheet = Nothing
End Sub


