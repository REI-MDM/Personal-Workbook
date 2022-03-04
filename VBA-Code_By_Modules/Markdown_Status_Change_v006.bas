Attribute VB_Name = "Markdown_Status_Change_v006"
Option Explicit

Sub MarkdownStatusChange()
Dim lastRow As Long
Dim statusChangeSheet As Worksheet
Dim batchCount As Long
Dim totalBatches As Long
Dim totalArts As Long
Dim i As Long

    LudicrousMode True
    On Error Resume Next
    Set statusChangeSheet = ActiveWorkbook.Worksheets("StatusChange")
    On Error GoTo 0
    If statusChangeSheet Is Nothing Then
        MsgBox ("We could not find the ""StatusChange"" Sheet in the active workbook." & _
            vbLf & "That is what I was expecting to find.  Aborting the Macro.")
            Exit Sub
    End If

'Find our last row
    lastRow = statusChangeSheet.Range("B" & Rows.Count).End(xlUp).row
    
'Sort by new status, then by variant
    statusChangeSheet.Sort.SortFields.Clear
    statusChangeSheet.Sort.SortFields.Add key:=Range("D5:D" & lastRow + 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    statusChangeSheet.Sort.SortFields.Add key:=Range("B5:B" & lastRow + 1), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    
    With statusChangeSheet.Sort
        .SetRange Range("B4:E" & lastRow + 1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'unmerge
    statusChangeSheet.Columns("B:E").MergeCells = False
'remove duplicates
    statusChangeSheet.Range("B4:E" & lastRow + 1).RemoveDuplicates _
        Columns:=Array(1, 2, 3, 4), Header:=xlNo
    
'Drop in H or D
    statusChangeSheet.Range("A5").Value = "H"
    batchCount = 1
    totalBatches = 1
    For i = 6 To lastRow
    'clean trim proper
        statusChangeSheet.Range("D" & i).Value = _
            ClTrProp(statusChangeSheet.Range("D" & i).Value)
        'statusChangeSheet.Range("E" & i).Value = _
            ClTrProp(statusChangeSheet.Range("E" & i).Value)
'determine if "H" or "D"
        If batchCount = 990 Or _
            statusChangeSheet.Range("D" & i).Value <> _
            statusChangeSheet.Range("D" & i - 1).Value Then
            'batch is max size or new status
            batchCount = 1 'reset batch count
            statusChangeSheet.Range("A" & i).Value = "H" 'new header
            totalBatches = totalBatches + 1
        Else
            'same batch
            statusChangeSheet.Range("A" & i).Value = "D"
            batchCount = batchCount + 1
        End If
    Next i
    
    totalArts = lastRow - 4
    
'Send the e-mail.  This should be a separate sub/function probably, but then
'I'd have to globally declare a couple variables.  So I will be lazy.
Dim OutlookApp As Object
Dim MailItem As Object
Dim messageTo As String
Dim messageCC As String
Dim subject As String
Dim messageBody As String
Dim fullName As String
 
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)
    fullName = GetDisplayName
    If fullName = "ERROR" Then
        fullName = InputBox(Prompt:="Your userid is not hardcoded into " & _
            "this macro.  Sorry.  How do you want your name to appear " & _
            "on the e-mail?", Title:="Who is you?", _
            Default:=Environ("username") & "@rei.com")
    End If


    messageTo = "ITSAPBasisAlerts@rei.com"
    messageCC = "bdidis@rei.com; masterdata@rei.com"
    subject = "Multiple background MASS MARC jobs up to " & totalBatches & _
        " background jobs - Status Changes for " & totalArts & " articles."
    
    If totalBatches < 20 Then
        messageBody = "Hello," & vbLf & vbLf & "Master Data is changing the " & _
            "status on " & totalArts & " articles as part of our monthly " & _
            "markdown status change process." & vbLf & vbLf & "We are going to " & _
            "enter up to " & totalBatches & " background MASS MARC batch jobs with between 1 and " & _
            "990 articles in each batch." & vbLf & vbLf & "Please let me know " & _
            "if you have any concerns. " & vbLf & vbLf & "Thanks," & vbLf & fullName
    Else
        messageBody = "Hello," & vbLf & vbLf & "Master Data would like to change " & _
            "the status on " & totalArts & " articles via MASS MARC background " & _
            "jobs.  However, due to the way our data is structured and the current " & _
            "limit of 990 articles per batch, this would require up to " & totalBatches & _
            " background jobs.  Will this be an issue?" & vbLf & vbLf & "If necessary we" & _
            " can process some in the foreground, or put off some of these updates until " & _
            "a later time." & vbLf & vbLf & "Thanks," & vbLf & fullName
    End If
            
    With MailItem
        .SentOnBehalfOfName = "masterdata@rei.com"
        .To = messageTo
        .CC = messageCC
        .subject = subject
        .BodyFormat = 1 ' Denotes olFormatHTML - HTML message formatting
        .Body = messageBody
        .HTMLBody = .HTMLBody
        'this .display could also be a .send to automatically send it.
        .Display
    End With
    
    If totalBatches > 20 Then
        MsgBox ("The SAP Basis Team has asked us to reach out to them if we " & _
            "are entering more than 20 batch jobs." & vbLf & "There should " & _
            "have been an email generated asking them for their opinion." & _
            vbLf & vbLf & "Please wait to see what they say before continuing.")
    Else
        MsgBox ("Ready to run ""MASS_Stat_Markdowns.TxR"" on this sheet." & vbLf & _
            vbLf & "It should take less than 5 minutes to run this script.")
    End If
    
    LudicrousMode False
    Set statusChangeSheet = Nothing
    Set MailItem = Nothing
    Set OutlookApp = Nothing
End Sub
        
Private Function ClTrProp(str As String) As String
    ClTrProp = WorksheetFunction.Clean(str)
    ClTrProp = WorksheetFunction.Trim(ClTrProp)
    ClTrProp = StrConv(ClTrProp, vbProperCase)
End Function

Private Function GetDisplayName() As String
'******************************************************************************
'Returns the active user's job title!  Does NOT currently work for anyone elses
'In the event of an error or no access to acrtive directory, this will return
'"ERROR"
'https://sdelisle.wordpress.com/2011/06/10/how-to-get-active-directory-accounts-information-using-vb-script/
'https://community.spiceworks.com/topic/361258-using-vba-to-report-user-s-full-name-maybe-from-ad
'******************************************************************************
Dim objAD As Object
Dim objUser As Object
Dim strJobTitle As String
    On Error Resume Next
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.UserName)
    GetDisplayName = objUser.DisplayName
    
    Set objAD = Nothing
    Set objUser = Nothing
    On Error GoTo 0
    If GetDisplayName = "" Then GetDisplayName = "ERROR"
End Function
'Adjusts Excel settings for faster VBA processing
Private Sub LudicrousMode(ByVal Toggle As Boolean)
    Application.ScreenUpdating = Not Toggle
    Application.EnableEvents = Not Toggle
    Application.DisplayAlerts = Not Toggle
    Application.EnableAnimations = Not Toggle
    Application.DisplayStatusBar = Not Toggle
    Application.PrintCommunication = Not Toggle
    Application.Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
End Sub

