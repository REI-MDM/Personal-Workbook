VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MDSErrorLog_v011 
   Caption         =   "MDS Error Log"
   ClientHeight    =   6480
   ClientLeft      =   15
   ClientTop       =   225
   ClientWidth     =   4635
   OleObjectBlob   =   "MDSErrorLog_v011.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MDSErrorLog_v011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'**********************************************************************
'**********************************************************************
'This Error log is to track errors sent to the Master Data team.
'On update the version name needs to be changed in the following locations
'Personal - Upload_File - 1 replacement
'Personal - Error log form - 11 replacements
'Mit - Upload_File - 1 replacement
'Assortment - SaveasCSV - 1 replacement


Private Sub Submit_Button_Click()

'forces task Number as required
If TaskNum.Value = "" Then
    TaskNum.SetFocus
    MsgBox "Dont Forget the Task Number"
    Exit Sub
Else
    If IsNumeric(Right(TaskNum.Value, 6)) = False Then
        TaskNum.SetFocus
        MsgBox "Make sure this is actually a Task number"
        Exit Sub
    End If
End If

'forces MDS user as required
If MDSRequest.Value = "" Then
    MDSRequest.SetFocus
    MsgBox "Dont Forget the MDS Requester"
    Exit Sub
End If
    
If T_F_Err.Value = "True" And ERRType.Value = "" And Notes.Value = "" Then
    ERRType.SetFocus
    MsgBox "Make sure to record what type of error is on this request!"
    Exit Sub
End If
 
    
    
'launches MDS_Error_Log sub
MDS_Error_Log

'unloadsform
Unload MDSErrorLog_v011

End Sub

Private Sub Submit_Second_Click()

If TaskNum.Value = "" Then
    TaskNum.SetFocus
    MsgBox "Dont Forget the Task Number"
    Exit Sub
Else
    If IsNumeric(Right(TaskNum.Value, 6)) = False Then
        TaskNum.SetFocus
        MsgBox "Make sure this is actually a Task number"
        Exit Sub
    End If
End If


'forces MDS user as required
If MDSRequest.Value = "" Then
    MDSRequest.SetFocus
    MsgBox "Dont Forget the MDS Requester"
    Exit Sub
End If

If T_F_Err.Value = "True" And ERRType.Value = "" And Notes.Value = "" Then
    ERRType.SetFocus
    MsgBox "Make sure to record what type of error is on this request!"
    Exit Sub
End If
 
    
'launches MDS_Error_Log sub
MDS_Error_Log

MsgBox "Error has been recorded."

End Sub

Public Sub UserForm_Initialize()

'This sub sets as much automaticly as possible, and populates our drop downs.
    
    REQType.AddItem "Article Create", 0
    REQType.AddItem "Article Maintain", 1
    REQType.AddItem "Maintain Promo", 2
    REQType.AddItem "Units of Measure", 3
    REQType.AddItem "Markdown", 4
    REQType.AddItem "Listings", 5
    REQType.AddItem "Assortment Group Create", 6
    REQType.AddItem "Assortment Group Maintain", 7
    REQType.AddItem "Create Promo", 8
    REQType.AddItem "Inspection Required", 9
    REQType.AddItem "Vendor Create", 10
    REQType.AddItem "Vendor Maintain", 11
    
    
    'this sets the TASK number automaticly on a create type. on error to leave blank if launched with no AWB
    On Error Resume Next
    REQType.Value = ActiveWorkbook.ContentTypeProperties("Request Type").Value
    'pulls TaskNum if on Article Create
    On Error Resume Next
    If ActiveWorkbook.ContentTypeProperties("Request Type").Value = "Article Create" Then
        TaskNum.Value = ActiveWorkbook.Worksheets("Article Create").Range("I8")
    End If
        
    On Error Resume Next
    
    If ActiveWorkbook.ContentTypeProperties("Request Type").Value = "Article Maintain" Then
        TaskNum.Value = ActiveWorkbook.Worksheets("Maintain Article").Range("CH1")
        
    End If
    'sets default to false, then populates the drop down to have both true and false.
    T_F_Err.Value = "False"
    T_F_Err.AddItem "True"
    T_F_Err.AddItem "False"
    
    'Sets our dropdown for error severity.
    ERRSev.AddItem "1"
    ERRSev.AddItem "2"
    ERRSev.AddItem "3"
      
    'sets our Error types - Based on Kyles work
    ERRType.AddItem "Add Variant - Already Exists/incorrect generic"
    ERRType.AddItem "Article # - Invalid/incorrect or missing"
    ERRType.AddItem "Assortment Create/Maintain"
    ERRType.AddItem "Characteristic profile"
    ERRType.AddItem "Color/Size"
    ERRType.AddItem "Data already exists in SAP"
    ERRType.AddItem "Drag/Drop incrementing"
    ERRType.AddItem "File Naming Convention"
    ERRType.AddItem "Hangtag Data"
    ERRType.AddItem "HTS"
    ERRType.AddItem "Incomplete Template"
    ERRType.AddItem "Incorrect weights"
    ERRType.AddItem "Merch Ops Approval"
    ERRType.AddItem "Missing Supporting Documents"
    ERRType.AddItem "Missing Template"
    ERRType.AddItem "Pricing issues"
    ERRType.AddItem "Promo Data"
    ERRType.AddItem "Size/Gender in description"
    ERRType.AddItem "UPC"
    ERRType.AddItem "Other"
       
    'populates the MD user name with whoever is running the REQ
    MDUser.Value = Application.UserName
    MDSRequest.Value = Name_Finder
    
    'This bit makes sure the error log pops up in the middle of the screen excel is running on
    'and not all over the place
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
        
End Sub
Public Function Name_Finder() As String

Dim BottomRow As Long
Dim SLSheet As Worksheet
Dim LastMDSUser As String
Dim MDUsers As String
Dim i As Integer

    MDUsers = "William Keister Nathanael Mokry Michael Hildrum Brad McKnight Kelly Brewer Ariela Humphries Kevin Hanson"
   
    'Select savelog sheet, then select Column F, then find the BottomRow of savelog,
    Set SLSheet = ActiveWorkbook.Worksheets("savelog")
    BottomRow = SLSheet.Range("F" & SLSheet.Rows.Count).End(xlUp).row
    'Sheets("Savelog").Select
    
    
    i = BottomRow
    LastMDSUser = ""
    Do
        If InStr(MDUsers, SLSheet.Range("F" & i).Value) = 0 Then
            LastMDSUser = SLSheet.Range("F" & i).Value
            End If
        i = i - 1
        If i = 0 Then
            Name_Finder = LastMDSUser
            Exit Function
        End If
                        
    Loop While LastMDSUser = ""
    
    Name_Finder = LastMDSUser
End Function
Public Function ResolveUserIDToJobTitle(sFromName) As String
    Dim OLApp As Object 'Outlook.Application
    Dim oRecip As Object 'Outlook.Recipient
    Dim oEU As Object 'Outlook.ExchangeUser
    Dim oEDL As Object 'Outlook.ExchangeDistributionList
    Dim oEUM As Object 'Exchange User Manager

    Set OLApp = CreateObject("Outlook.Application")
    Set oRecip = OLApp.session.CreateRecipient(sFromName)
    oRecip.Resolve
    If oRecip.Resolved Then
        Select Case oRecip.AddressEntry.AddressEntryUserType
            Case 0, 5 'olExchangeUserAddressEntry & olExchangeRemoteUserAddressEntry
                Set oEU = oRecip.AddressEntry.GetExchangeUser
                If Not (oEU Is Nothing) Then
                   ResolveUserIDToJobTitle = oEU.JobTitle
                   Set oEUM = oEU.GetExchangeUserManager
                End If
            Case 10, 30 'olOutlookContactAddressEntry & 'olSmtpAddressEntry
                    ResolveUserIDToJobTitle = oRecip.AddressEntry.Name
        End Select
    End If
    
    Set OLApp = Nothing
    Set oRecip = Nothing
    Set oEU = Nothing
    Set oEDL = Nothing
    Set oEUM = Nothing


End Function



Public Sub MDS_Error_Log()

Dim cn As ADODB.connection 'This is the direct DB connection
Dim strFile As String 'Filepath stored as string
Dim strCon As String 'our connection as a string
Dim strSQLInsert As String 'This is the SQL that inserts our errors

'These Variables Match Headers in our SQL DB Named the same for clarity
'used to pull data from a Form to the DB SQL string
Dim TaskNum As String
Dim MDUser As String
Dim MDSOpener As String
Dim REQType As String
Dim ErrorOnReq As Integer
Dim ErrorSeverity As Integer
Dim ErrorDate As Date
Dim ErrorType As String
Dim ErrorDetails As String
Dim OpenerTitle As String

'This is the base code to replace quotes to keep the DB clean and prevent dropping the tables in extreme cases
 ' Replace(OldString, Chr(34), "")


'this assigns variables to code to put in the DB insert
TaskNum = Replace(MDSErrorLog_v011.TaskNum.Value, Chr(34), "")
TaskNum = UCase(TaskNum)

MDUser = MDSErrorLog_v011.MDUser.Value
MDSOpener = Replace(MDSErrorLog_v011.MDSRequest.Value, Chr(34), "")
REQType = Replace(MDSErrorLog_v011.REQType.Value, Chr(34), "")
OpenerTitle = ResolveUserIDToJobTitle(MDSOpener)

'sets error to 1 or 0 for SQL insert
'SQL 1 is true 0 is false
If MDSErrorLog_v011.T_F_Err.Value = "True" Then
    ErrorOnReq = 1
    If MDSErrorLog_v011.ERRSev.Value = "" Then
        ErrorSeverity = 1
    Else
        ErrorSeverity = MDSErrorLog_v011.ERRSev.Value
    End If
Else
    ErrorOnReq = 0
    ErrorSeverity = 0
End If

ErrorDate = Date
ErrorType = MDSErrorLog_v011.ERRType.Value
ErrorDetails = Replace(MDSErrorLog_v011.Notes.Value, Chr(34), "")


'sets DB path
strFile = "G:\SC EVS\Master Data\ErrorLog\MDSErrorLog.accdb;"

'sets connection type and path
strCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFile
Set cn = CreateObject("ADODB.Connection")

'opens connection to file
cn.Open strCon

'string to insert data into DB
strSQLInsert = "('" & TaskNum & "','" & MDUser & "','" & MDSOpener & "','" & REQType & "','" & ErrorOnReq & "','" & ErrorSeverity & "','" & ErrorDate & "','" & ErrorType & "','" & ErrorDetails & "','" & OpenerTitle & "')"
'enters values from prompt into SQL DB
cn.Execute "INSERT INTO Errors (TaskNum, MDUser, MDSOpener, REQType, ErrorOnReq, ErrorSeverity, ErrorDate, ErrorType, ErrorDetails, OpenerTitle) VALUES" & strSQLInsert

'close connection, *and* dumps connection and unique table this may be overkill but should help to ensure we dont step on toes if checking at once
cn.Close
Set cn = Nothing

End Sub


