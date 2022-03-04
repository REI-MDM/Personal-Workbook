Attribute VB_Name = "Check_Future_Prices_v001"
Option Explicit
Const refLoc = "\\teamsites.rei.com\DavWWWRoot\merchandising\Article and Vendor Master Data\Shared Documents\"
Const refWB = "Future_Price_Changes.xlsm"

Dim AppSU As Boolean, AppDSB As Boolean, AppEE As Boolean
Dim APPCalc As Long
Sub checkForFuturePrices()
'******************************************************************************
' Compare an Article Maintain Sheet with the reference of future prices.
'
' Doesn't currently support multiple price changes, either on the AM sheet
' or bring back multiple future price conditions.
'
' Currently "written" for personal workbook.  If we move into AM it would be
' good to change the
' "Set AMWS = ActiveWorkbook.Worksheets("Maintain Article")"
' into
' "Set AMWS = ThisWorkbook.Worksheets("Maintain Article")"
'******************************************************************************
Dim RWB As Workbook
Dim RefWS As Worksheet
Dim AMWS As Worksheet
Dim AMRetails As Object
Dim futRetails As Object
Dim lastRow As Long
Dim i As Long
Dim dataArray As Variant
Dim temp(1 To 2) As Variant

    Toggle "Off"
'Check for retails on this sheet to "make sure" we need to continue
    Set AMWS = ActiveWorkbook.Worksheets("Maintain Article")
    
    Set AMRetails = CreateObject("Scripting.dictionary")
    
    lastRow = AMWS.Range("A" & Rows.Count).End(xlUp).row
    dataArray = AMWS.Range("A9:L" & lastRow).Value
    'Dataarray (i, 1) = Article
    'DataArray (i, 11) has price
    'Dataarray(i, 12) has price date
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        If dataArray(i, 1) <> "" And dataArray(i, 11) <> "" And _
            dataArray(i, 12) <> "" Then
            If Not AMRetails.Exists(dataArray(i, 1)) Then
                AMRetails.Add dataArray(i, 1), dataArray(i, 12)
            End If
        End If
    Next i
    
    If AMRetails.Count = 0 Then GoTo garbage    'nothing to check
    
    'Check futures!
    Set RWB = Workbooks.Open(fileName:=refLoc & refWB, _
        UpdateLinks:=False, ReadOnly:=True)
    Set RefWS = RWB.Worksheets("Sheet1")
    Set futRetails = CreateObject("Scripting.dictionary")
    lastRow = RefWS.Range("B" & Rows.Count).End(xlUp).row
    dataArray = RefWS.Range("B2:G" & lastRow).Value
    'dataarray(i, 1) = Article
    'Dataarray(i, 5) has price
    'dataarray(i, 6) has price date
    
    RWB.Close False
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        If AMRetails.Exists(dataArray(i, 1)) Then
            'this article is on our sheet!
            If AMRetails(dataArray(i, 1)) <= dataArray(i, 6) Then
                temp(1) = dataArray(i, 6)
                temp(2) = dataArray(i, 5)
                futRetails.Add dataArray(i, 1), temp
            End If '/Date Check
        End If '/Same article check
    Next i

    'Now check if we found futures, and output if so.
    If futRetails.Count = 0 Then
    'disable msgbox for "no issues"?
        MsgBox ("No retail conflicts found")
        GoTo garbage
    End If
    'We'll just cycle through the sheet directly, instead of arraying it.
    lastRow = AMWS.Range("A" & Rows.Count).End(xlUp).row
    
    For i = 9 To lastRow
        'check for retails
        If AMWS.Range("A" & i).Value <> "" And _
            AMWS.Range("K" & i).Value <> "" And _
            AMWS.Range("L" & i).Value <> "" And _
            futRetails.Exists(AMWS.Range("A" & i).Value) Then
            'has article, date, price AND future!
            'Output future
            temp(1) = futRetails(AMWS.Range("A" & i).Value)(1)
            temp(2) = futRetails(AMWS.Range("A" & i).Value)(2)
            AMWS.Range("BL" & i & ":BM" & i).Value = temp
        End If
    Next i
    
    MsgBox ("Found " & futRetails.Count & " articles with at least one " & _
        "future price in SAP.  The 'Next' one can be found in cells BL:BM" & _
        vbLf & vbLf & "All known as of last night can be found in:" & vbLf & _
        refLoc & refWB)
garbage:
    Toggle "ON"
On Error Resume Next
    Erase dataArray
    Set AMWS = Nothing
    Set AMRetails = Nothing
    Set futRetails = Nothing
    Set RWB = Nothing
    Set RefWS = Nothing
On Error GoTo 0
End Sub
Private Sub Toggle(OnOff As String)
'Capture users settings if toggling "Off", then "Enhance speed"
'If toggling "On" restore to defaults. Could be Boolean, but I like this
    On Error Resume Next
    If UCase(OnOff) = "OFF" Then
        AppSU = Application.ScreenUpdating
        AppDSB = Application.DisplayStatusBar
        APPCalc = Application.Calculation
        AppEE = Application.EnableEvents
 
    'turn off some Excel functionality so your code runs faster
       Application.ScreenUpdating = False
       Application.DisplayStatusBar = False
       Application.Calculation = xlCalculationManual
       Application.EnableEvents = False
    
    ElseIf UCase(OnOff) = "ON" Then
    'Restore original values
        Application.ScreenUpdating = AppSU
        Application.DisplayStatusBar = AppDSB
        Application.Calculation = APPCalc
        Application.EnableEvents = AppEE
    End If
    On Error GoTo 0
End Sub
