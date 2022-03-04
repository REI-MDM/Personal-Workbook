Attribute VB_Name = "Upload_File_v050"
Option Explicit

Dim ExcelDoc As Workbook
Const oldSP = "http://teamsites.rei.com/merchandising/Article and Vendor Master Data/"
Const newSP = "https://reiweb.sharepoint.com/sites/MDMresources/"
Dim WPCloseoutReady As Boolean

Sub Upload_Completed_File()
'******************************************************************************
' Upload_Completed_File Macro
' Originally written/adapted by NM
' Further edited by MH 01/04/2013
'
' Keyboard Shortcut: Ctrl+Shift+U
'
' 1. Get the current filename and file path, modify it to make it suitable for
' upload to our sharepoint site.
'
' 2. Save it to our sharepoint site. Once there grab the full name, and
' replace spaces, so that it will be a working hyperlink. Put that link in CC1
'
' 3. Re-save the file locally, so that you have a copy of the completed
' template and also so that you are no longer looking at / editing the server
' copy.
'
' 4. Copy the hyperlink so it is ready to pasted into the waypoint task
'
'******************************************************************************
Dim OldPath As String
Dim wbname As String
Dim LinkNoSpace As String
Dim Prop As Variant                 'He is not sure _exactly_ how this works.
Dim PropShouldBe As String          'The request type
Dim starttime As Date
Dim Diff As Variant
Dim response As String
Dim Retrysave As Boolean
Dim i As Long
Dim ftype As Long
Dim fpath As String
Dim Completefpath_old As String 'this is where we will be saving the the completed document on the "old" SharePoint site
Dim Completefpath As String 'this is where we will be saving the the completed document on the "new" SharePoint site

Dim ACSheet As Worksheet
Dim mySht As Worksheet


    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
    Set ExcelDoc = ActiveWorkbook
    If ExcelDoc.Name Like "TASK######*" Then
        MsgBox ("Why you uploading a MIT, BRO?  Aborting!")
        Set ExcelDoc = Nothing
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If

'Set mySht
    Set mySht = ActiveSheet
            
    Retrysave = False
    On Error Resume Next
    Set ACSheet = ExcelDoc.Worksheets("Article Create")
    On Error GoTo 0
    If Not ACSheet Is Nothing Then
    'Do some time stuff with the sheet
        'Set by the most recent run of Characteristic Lookup
        starttime = ACSheet.Range("W1").Value
        Diff = Now - starttime
        'Total Elapsed time
        ACSheet.Range("A4").NumberFormat = "hh:mm:ss"
        ACSheet.Range("A4").Value = Format(Diff, "hh:mm:ss")
        'Time as Waypoint likes it, numbers of hours.
        ACSheet.Range("A5").NumberFormat = "0.00"
        ACSheet.Range("A5").Value = Format(Diff * 24, "0.00")
    End If

   
    OldPath = ExcelDoc.fullName
    'find its current directory
    For i = Len(OldPath) To 1 Step -1
        Select Case Mid(OldPath, i, 1)
            Case Is = "/"
                fpath = Left(OldPath, i)
                Exit For
            Case Is = "\"
                fpath = Left(OldPath, i)
                Exit For
            Case Else
            'do nothing
        End Select
    Next i
    
    ftype = ExcelDoc.FileFormat
    wbname = ExcelDoc.Name
    
'Sharepoint does not allow the following characters in filenames
'  \ / : * ? " < > | # { } % ~ &
    wbname = Replace(wbname, "\", "_")
    wbname = Replace(wbname, "/", "_")
    wbname = Replace(wbname, ":", "_")
    wbname = Replace(wbname, "*", "")
    wbname = Replace(wbname, "?", "")
    wbname = Replace(wbname, """", "")
    wbname = Replace(wbname, "<", "(")
    wbname = Replace(wbname, ">", ")")
    wbname = Replace(wbname, "|", "_")
    wbname = Replace(wbname, "#", "num")
    wbname = Replace(wbname, "{", ")")
    wbname = Replace(wbname, "}", ")")
    wbname = Replace(wbname, "%", "")
    wbname = Replace(wbname, "~", "-")
    wbname = Replace(wbname, "&", "and")
    
    'Range("A11").Select

'Set the completed file path location - errors out if the document does not have the "Request Type" property.
'Updated to check for sheet names instead.  We "handle" properties later.
    If SheetExists("Z001 Main Vendor Record") Or SheetExists("Vendor Input") Then
        'Save to Vendor Ops Completed files
        Completefpath = newSP & "VC Completed Reqs " & Format(Now, "yy/") & wbname
        Completefpath_old = oldSP & "VC Completed Reqs " & Format(Now, "yy/") & wbname
        Else
        'Save to Master Data Completed Files
        Completefpath = newSP & "MD Complete Reqs " & Format(Now, "yy/") & wbname
        Completefpath_old = oldSP & "MD Complete Reqs " & Format(Now, "yy/") & wbname
    End If


'Save current file with new file name variable to Completed Request Folder
'if this directory is inaccessible due to connectivitiy issues, we get no
'error message.  Check for success later.
    
    'On Error Resume Next
    'No longer saving files to old SP site
    'ExcelDoc.SaveAs Filename:=Completefpath_old, FileFormat:=ftype, CreateBackup:=False
    ExcelDoc.SaveAs fileName:=Completefpath, FileFormat:=ftype, CreateBackup:=False
    'On Error GoTo 0

'A Fix for the "Properties not set" error
    For Each Prop In ExcelDoc.ContentTypeProperties
        If Prop.Name = "Request Type" Then 'And prop.Value = "" Then
            Retrysave = True
            If SheetExists("Article Create") Then
                PropShouldBe = "Article Create"
            ElseIf SheetExists("Maintain Article") Then
                PropShouldBe = "Article Maintain"
            ElseIf SheetExists("Inspection Required") Then
                PropShouldBe = "Inspection Required"
            ElseIf SheetExists("Initial & further MD") Then
                PropShouldBe = "Markdown"
            ElseIf SheetExists("PriceChange") Then
                PropShouldBe = "Markdown"
            ElseIf SheetExists("Maintain_Promo") Then
                PropShouldBe = "Maintain Promo"
            ElseIf SheetExists("Promotions") Then
                PropShouldBe = "Create Promo"
            ElseIf SheetExists("Unit of Measure") Then
                PropShouldBe = "Units of Measure"
            ElseIf SheetExists("assortment create") Then
                PropShouldBe = "Assortment Group Create"
            ElseIf SheetExists("assortment maintain") Then
                PropShouldBe = "Assortment Group Maintain"
            ElseIf SheetExists("Required") Then
                PropShouldBe = "Bonus Buy"
            ElseIf SheetExists("Temp Listings") Then
                PropShouldBe = "Listings"
            ElseIf SheetExists("Z001 Main Vendor Record") Then
                PropShouldBe = "Vendor Maintain"
            ElseIf SheetExists("Vendor Input") Then
                PropShouldBe = "Vendor Create"
            End If
            Prop.Value = PropShouldBe
            Exit For
        End If
    Next Prop

    If Retrysave Then
        'If properties were not initially set correctly, try and re-save now that it should be set
        ExcelDoc.Save
    End If

    'save to old sp as well - disable when newsp is "primary"
    'ExcelDoc.SaveAs Filename:=Completefpath, FileFormat:=ftype, CreateBackup:=False

'Obtain new file path , replace the spaces, and drop it into cell CI1
    
    LinkNoSpace = ExcelDoc.FullNameURLEncoded
    
'Check to see if we successfully saved to the network
    If Left(LinkNoSpace, 8) <> "https://" Then
        MsgBox ("It looks like we failed at uploading.  Please try to " & _
            "manually upload the file. Possible connectivity issue.")
        Exit Sub
    End If
    
    mySht.Columns("CI:CI").Hidden = False
    mySht.Range("CI1").Value = LinkNoSpace
        
'Save a local copy, to reset the default save path, otherwise if you keep it
'open, you are looking at the server copy.

    Application.DisplayAlerts = False
    ExcelDoc.SaveAs fileName:=fpath & wbname, _
        FileFormat:=ftype, CreateBackup:=False
    Application.DisplayAlerts = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

'Check for article maintain description/brand updates
    Call PI_Desc_Email_Checker
    
'Grab that cell CI1 and pop it into the clipboard
    ExcelDoc.activate
    mySht.activate
    mySht.Range("CI1:CI5").Copy
    
'Can this template run the Waypoint Closeout macro?
    'If the file is missing "Request Type" - No
    If ExcelDoc.ContentTypeProperties.Count > 0 Then
        'If the template is Article Create - Yes
        If ExcelDoc.ContentTypeProperties("Request Type") = "Article Create" Then
            WPCloseoutReady = True
        'If the template is Article Maintain - Yes
        ElseIf ExcelDoc.ContentTypeProperties("Request Type") = "Article Maintain" Then
            WPCloseoutReady = True
        'If the template is not AC or AM - No
        Else
            WPCloseoutReady = False
        End If
    Else
        WPCloseoutReady = False
    End If
     
'If the template can run the Waypoint Closeout macro, run it
    If WPCloseoutReady = True Then
    'Run Waypoint Closeout macro
        Call WaypointCloseout
        If WPCloseoutReady = True Then
        'Save and close template
            ExcelDoc.Save
            ExcelDoc.Close
        End If
    End If

'If the template can't run the Waypoint Closeout macro, the user must close it out manually
    If WPCloseoutReady = False Then
    'Close Task Manuallly
    response = MsgBox("File Uploaded, the link should be in your cursor. " & _
        "Close the template?", vbYesNo, "Close Template")
        If response = vbYes Then
        'Save and close template
            ExcelDoc.Save
            ExcelDoc.Close
        End If
    End If
    
'NOTE: Commented to allow for click-free WP Closeout
    'Find the version of the userform, and open it
    'open_MDSError_form
    'MDSErrorLog_v011.Show
       
On Error Resume Next
    Set ACSheet = Nothing
    Set ExcelDoc = Nothing
    Set mySht = Nothing
On Error GoTo 0

End Sub


Function SheetExists(strWSName As String) As Boolean
    Dim WS As Worksheet
    On Error Resume Next
    Set WS = ExcelDoc.Worksheets(strWSName)
    If Not WS Is Nothing Then SheetExists = True
    On Error Resume Next
    Set WS = Nothing
    On Error GoTo 0
End Function


Function MDS_Elfind() As String
'this function finds a specific VB component name
'in this case it finds the UF MDSErrorLog and returns it reguardless of version.
'this way i can stop annoying nathanael when i update the UF version.
'Can't seem to get the form to "unload" when called through this dynamic method.
'Bummer.

Dim VBComp As Object
Dim vbUF As String

For Each VBComp In ThisWorkbook.VBProject.VBComponents
    If InStr(VBComp.Name, "MDSErrorLog_") > 0 Then
       vbUF = VBComp.Name
        End If
        Next
             
MDS_Elfind = vbUF

End Function

Public Sub open_MDSError_form()
'this opens the userform
Dim uf As Object
    Set uf = UserForms.Add(MDS_Elfind)
    uf.Show
    Set uf = Nothing
End Sub
'******************************************************************************
'Changelog!
' -v044 - BC - Added WaypointCloseout sub. Added logic to Upload_Completed_File sub to determine if file is WPCloseoutReady (Lines 218 - 261). Disabled MDSErrorLog.
' -V020 - MH -
' -V019 - NM - After file is uploaded, grab a 4th cell of data for better closing notes (name)
' -V019 - NM - After file is uploaded, 3 Cells of data are copied/pasted for better closing notes
' -V018 - MH - Replace some characters which sharepoint does not allow in the filename
' -v017 - MH - Minor update to allow for uploading of some "Bad Templates"
' _v016 - MH - Possibly better support for .xlsx or .xls files, no forcing .xlsm
' _v015 - MH - Added support for new Markdown template properties.
'   Additionally always save as xlsm (hopefully)
' _v011 - MH - Check "http://" in LinkNoSpace as an indicator of successful
'   network save.
' _V011 - MH - Fixed (For good hopefully) the "Properties not set error"
' _v009 - MH - Adjusted "Closing Note Grabber" to work for everyone by default.
' _v008 - MH - Added Jenny to closing comments, removed Carson
' _v007 - MH - fix the bug I caused where it would do timely stuff on non AC
'    sheets
' _v006 - MH - Dropped elapsed time on the sheet in A4 and A5, error "ignoring"
'   for the "PropShouldBe" code.
' _v005 - CG - Added Nathanael to closing notes.
' _v004 - CG - Added closing notes copy option.
' _v003 - MH - Unhide the LinkNoSpace Destination - D97 had issues
' _v002 - NM - Old LinkNoSpace dest was hidden on creates, causing copy/paste
'   errors.  Changed location
'
    
Private Sub WaypointCloseout()

'This macro is designed to close out MDM Input Waypoint requests after completion.
'Written by Brian Combs - December 2020

'Requirements
    '1. Need to have SeleniumBasic installed and added as a reference library
    '2. Need to update chromedriver to latest driver.
    '3. Download latest driver for your version of Chrome here - "https://chromedriver.chromium.org/downloads".
    '4. Move chromedriver to this location "C:\Users\YOURUSERNAME\AppData\Local\SeleniumBasic"

'References
    'Selenium Type Library

'Variables
    Dim Bot As New WebDriver
    Dim AWB As Workbook
    Dim Task As String
    Dim req As String
    Dim CompletedURL As String
    Dim ClosureComments As String
    Dim ShortDesc As String
    Dim EnterTime As String
    Dim OpenedBy As String
    Dim AssignedTo As String
    Dim ErrorChoice As VbMsgBoxResult
    Dim TaskLink As Object
    Dim SLATask As Object
    Dim DynamicElement As Object
    Dim KeyObj As Selenium.keys

'Dev Variables
    Dim SaveChoice As VbMsgBoxResult

'Error handling
    On Error GoTo Retry
    
'Create Selenium Keys
    Set KeyObj = New Selenium.keys
       
'Set Workbook
    Set AWB = ActiveWorkbook
    
'Determine template type and properties
    'Check for content properties
    If AWB.ContentTypeProperties.Count > 0 Then
        'Article Create properties
        If AWB.ContentTypeProperties("Request Type") = "Article Create" Then
            Task = AWB.Worksheets("Article Create").Range("I8").Value
            req = AWB.ContentTypeProperties("Request Type")
            CompletedURL = AWB.Worksheets("Article Create").Range("CI1").Value
            ClosureComments = AWB.Worksheets("Article Create").Range("CI3").Value
            ShortDesc = " Auto"
            EnterTime = ".2"
        'Article Maintain properties
        ElseIf AWB.ContentTypeProperties("Request Type") = "Article Maintain" Then
            Task = AWB.Worksheets("Maintain Article").Range("CH1").Value
            'Currently only works on AM roll-up files unless the user manually enters Task in CH1 or G1
                If Task = "" Then
                    Task = AWB.Worksheets("Maintain Article").Range("G1").Value
                    If Task = "" Or Task = "MD Use Only" Then
                        WPCloseoutReady = False
                        Exit Sub
                    End If
                End If
            CompletedURL = AWB.Worksheets("Maintain Article").Range("CI1").Value
            ClosureComments = AWB.Worksheets("Maintain Article").Range("CI3").Value
            ShortDesc = AWB.Worksheets("Maintain Article").Range("CI5").Value & " Auto"
            EnterTime = ".1"
        Else
            MsgBox "The Waypoint Closeout macro only works with AC and AM templates right now."
            WPCloseoutReady = False
            Exit Sub
        End If
    Else
        WPCloseoutReady = False
        Exit Sub
    End If
'-------------------------------------------------------------------------------------------
'WAYPOINT HOME PAGE

'Open web browser and navigate to Waypoint
    Bot.Start "chrome", "https://rei2.service-now.com/sc_task_list.do?sysparm_query=assignment_group%3Df534206013b52a00b1d6b0322244b0e1&sysparm_first_row=1&sysparm_view=master_data"
    Bot.Window.Maximize
    Bot.Get "/"
       
'Enter task number in the "Number" search bar
    Bot.FindElementById("sc_task_table_header_search_control").SendKeys (Task)

'Hit the enter key
    Bot.SendKeys KeyObj.Enter

'-------------------------------------------------------------------------------------------
'SEARCH RESULTS PAGE

'Click on the task link
    Set TaskLink = Bot.FindElementsByLinkText(Task)
    TaskLink(1).Click

'-------------------------------------------------------------------------------------------
'TASK PAGE

'Determine "Opened By" and "Assigned To"
    OpenedBy = Bot.FindElementById("sc_task.request_item.opened_by_label").Value
    AssignedTo = Bot.FindElementById("sys_display.sc_task.assigned_to").Value

'Determine first names of "Opened By" and "Assigned To"
    OpenedBy = Split(OpenedBy, " ")(0)
    AssignedTo = Split(AssignedTo, " ")(0)

'Determine Closure Comments
    ClosureComments = "Hi " & OpenedBy & "," & vbCrLf & ClosureComments & vbCrLf & AssignedTo

'Paste Task "Enter Time"
    Bot.FindElementById("sc_task.u_time_worked").SendKeys (EnterTime)

'Paste Task summary in the "Short Description" title
   Bot.FindElementById("sc_task.short_description").SendKeys (ShortDesc)

'Change Task "State" to "Closed Complete"
    Bot.FindElementById("sc_task.state").AsSelect.SelectByValue ("3")

'Change Task "MD Was the SLA met?" to "Yes"
    Set DynamicElement = Bot.FindElementsByClass("cat_item_option")
    DynamicElement(11).AsSelect.SelectByValue ("Yes")

'Change Task "Was the template inaccurate or incomplete?" to "No"
    DynamicElement(16).AsSelect.SelectByValue ("No")

'Paste Task "Link to Completed Template"
    DynamicElement(22).SendKeys (CompletedURL)

'Paste Task "Other Closure Comments for Requestor"
    DynamicElement(24).SendKeys (ClosureComments)
    
''Click the "Update" button (for dev purposes only - can be deleted for prod)
'    SaveChoice = MsgBox("Do you want to update this task?", vbYesNo)
'        If SaveChoice = vbYes Then
'            Bot.FindElementById("sysverb_update").Click
'        Else
'            Exit Sub
'        End If
        
'Click the "Update" button
    Bot.FindElementById("sysverb_update").Click

'-------------------------------------------------------------------------------------------

'Close Chrome
    Bot.Quit

'The End
    Exit Sub
    
'-------------------------------------------------------------------------------------------
'This code pauses the macro for 5 seconds to allow the website to load. If the use chooses "No" it will exit the sub.
Retry:
    ErrorChoice = MsgBox("Selenium errored trying to find a web element. Would you like to wait 5 seconds and try again?", vbYesNo)
        If ErrorChoice = vbYes Then
            Bot.Wait (500)
            Resume
        Else
        MsgBox "Selenium errored trying to close this Task. Please close this Task manually." & vbCrLf & vbCrLf & "Please check that your Chrome driver matches your version of Chrome." & vbCrLf & vbCrLf & err.source & vbCrLf & vbCrLf & err.Number & vbCrLf & vbCrLf & err.Description & vbCrLf & vbCrLf & err.HelpContext
        WPCloseoutReady = False
        Exit Sub
        End If
   
End Sub

Private Sub PI_Desc_Email_Checker()
'Brian Combs - 2022

'This sub checks for description or brand updates on the Maintain Article sheet.
'If true, calls the PI_Description_Email sub.

'Variables
    Dim AWB As Workbook
    Dim WS As Worksheet
    Dim AMsheet As Worksheet
    Dim lastRow As Long
    Dim descRange As Range
    Dim brandRange As Range
    Dim singleCell As Range
    Dim maintFound As Boolean
    Dim emailNeeded As Boolean

'Set objects
    Set AWB = ActiveWorkbook

'Default booleans
    maintFound = False
    emailNeeded = False

'Loop through worksheets
    For Each WS In AWB.Worksheets
        If WS.Name = "Maintain Article" Then
            Set AMsheet = AWB.Worksheets("Maintain Article")
            maintFound = True
        End If
    Next

'If maintain aritcle sheet not found, exit sub
    If maintFound = False Then
        Exit Sub
    End If
    
'Count rows on maintSheet
    lastRow = 9
    Do While AMsheet.Range("A" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

'Set descRange
    Set descRange = AMsheet.Range("D9:D" & lastRow)
    Set brandRange = AMsheet.Range("X9:X" & lastRow)

'Loop through descRange
    For Each singleCell In descRange
        If singleCell.Value <> "" Then
            emailNeeded = True
        End If
    Next

'Loop through brandRange
    For Each singleCell In brandRange
        If singleCell.Value <> "" Then
            emailNeeded = True
        End If
    Next

'If emailNeeded = True call PI_Description_Email
    If emailNeeded = True Then
        Call PI_Description_Email
    End If

End Sub


Private Sub PI_Description_Email()
'Written by Kevin Hanson
'Updated by Brian Combs - 2022

Dim lastRow As Long
Dim AM As Worksheet                 'Article Maintain Worksheet
Dim reqSheet As Worksheet           'Article Maintain Request Worksheet
Dim dessheet As Worksheet           'Description Update Worksheet
Dim req As Workbook                 'Request Workbook
Dim arr As Variant                  'Array of Article, Blank, Current Descript, New Descript
Dim i As Long                       'Counter
Dim art As Long                     'Key for desdic
Dim tempdef(1 To 3) As String      'desdic Items 1-Current Description, 2-New Description, 3-Brand
Dim d As Integer                    'Counter
Dim desdic As Object                'Description dictionary
Dim fe As Variant                   'Object to loop through for each key in desdic
Dim rng As Range                    'Range of summary
Dim OutlookApp As Object            'Outlook
Dim MailItem As Object              'Email
Dim messageTo As String             'Email to
Dim messageCC As String             'Email CC
Dim subject As String               'Subject
Dim messageBody As String           'Message Body
Dim fullName As String              'Full Name of sender for signature
Dim Sig As String                   'Signature
Dim Articles As String              'Concatination of Articles for subject
Dim c As Long                       'Counter
Dim BrandWB As Workbook             'New Brand Request WB
Dim BrandRef As Worksheet           'Brand query
Dim BrandDic As Variant             'Brand Dictionary
Dim BrandArr() As Variant           'Brand Array

Dim Notes As String                 'Submitter Notes
Dim Reason As String                'Reason for Maintenance
Dim WS As Worksheet                 'Variable used to loop through all worksheets
Dim JobType As String               'Used to determine if AC + Maint or AM
Dim Submitter As String             'AM Submitter
Dim JobName As String               'File name
Dim ProjName As String              'Project Name
Dim Vendor As String                'Vendor name and Number
Dim Dept As String                  'Department name
Dim Season As String                'Season
Dim Priority As String              'Priority


Set req = ActiveWorkbook
Set AM = ActiveWorkbook.Worksheets("Maintain Article")
JobType = "AM"

'Determine if AC + Maintain or just Maintain
    For Each WS In req.Worksheets
        If WS.Name = "Article Create Request" Then
            JobType = "AC"
        End If
    Next
 
'Set worksheets based on jobtype
    If JobType = "AC" Then
        Set reqSheet = ActiveWorkbook.Worksheets("Article Create Request")
    ElseIf JobType = "AM" Then
        Set reqSheet = ActiveWorkbook.Worksheets("Article Maintain Request")
    End If

lastRow = AM.Range("A" & AM.Rows.Count).End(xlUp).row
arr = AM.Range("A9:X" & lastRow)
If desdic Is Nothing Then Set desdic = CreateObject("scripting.dictionary")

'Create Description dictionary
'If not Description Change present sub ends
d = 0
    For i = LBound(arr) To UBound(arr)
        If arr(i, 4) <> "" Or arr(i, 24) <> "" Then
            art = Left(arr(i, 1), 6)
            If Not desdic.Exists(art) Then
                tempdef(1) = arr(i, 3)
                tempdef(2) = arr(i, 4)
                tempdef(3) = arr(i, 24)
                desdic.Add art, tempdef
            End If
            d = d + 1
        End If

    Next i
    'Looks like there were descriptions or a brand update
    'lets add a new sheet to aggregate the 6 digit articles
    'the current and new descriptions and brand and do some formatting
    If d > 0 Then
        On Error Resume Next
        Set dessheet = req.Worksheets("DescriptionAdd")
        If dessheet.Name = "" Then
        Set dessheet = req.Worksheets.Add
        dessheet.Name = "DescriptionAdd"
        End If
        On Error GoTo 0
        dessheet.Range("A1") = "6 Digit Article #"
        dessheet.Range("A1").Interior.Color = 10079487
        dessheet.Range("B1") = "Current Description"
        dessheet.Range("B1").Interior.Color = 10079487
        dessheet.Range("C1") = "New Description"
        dessheet.Range("C1").Interior.Color = 10079487
        dessheet.Range("D1") = "New Brand"
        dessheet.Range("D1").Interior.Color = 10079487

        Workbooks.Open fileName:="\\reiweb.sharepoint.com@SSL\DavWWWRoot\sites\MasterDataManagement\Shared Documents\Daily Work Files\New_Brand_Request_Form.xlsm", UpdateLinks:=False, ReadOnly:=False
        Set BrandWB = Workbooks("New_Brand_Request_Form.xlsm")
        Set BrandRef = BrandWB.Worksheets("Brand Reference")
        lastRow = BrandRef.Range("A" & Rows.Count).End(xlUp).row
        BrandArr = BrandRef.Range("A2:B" & lastRow)

        Set BrandDic = CreateObject("scripting.dictionary")
        For i = LBound(BrandArr) To UBound(BrandArr)

            If Not BrandDic.Exists(WorksheetFunction.Text(str(BrandArr(i, 1)), "0000")) Then
               BrandDic.Add WorksheetFunction.Text(str(BrandArr(i, 1)), "0000"), BrandArr(i, 2)
            End If

        Next i

        BrandWB.Close Savechanges:=False


        For Each fe In desdic.keys
            lastRow = dessheet.Range("A" & Rows.Count).End(xlUp).row + 1
            'Drops article number
            dessheet.Range("A" & lastRow).Value = fe
            'Drops current description
            dessheet.Range("B" & lastRow).Value = desdic(fe)(1)
            'Drops new description
            dessheet.Range("C" & lastRow).Value = desdic(fe)(2)
            'Drops new Brand
            dessheet.Range("E" & lastRow).Value = desdic(fe)(3)
            If desdic(fe)(3) <> "" Then
                If BrandDic.Exists(WorksheetFunction.Text(desdic(fe)(3), "0000")) Then
                    dessheet.Range("D" & lastRow).Value = BrandDic(WorksheetFunction.Text(desdic(fe)(3), "0000"))
                Else
                    dessheet.Range("D" & lastRow).Value = InputBox("Looks like the brand on the request could not be found in the Brand Look up." & _
                    "Please input the Brand Name for " & desdic(fe)(3))
                    BrandDic.Add dessheet.Range("E" & lastRow).Value, dessheet.Range("D" & lastRow).Value
                End If
            End If
        Next



        dessheet.Columns("A:D").AutoFit
        lastRow = dessheet.Range("A" & Rows.Count).End(xlUp).row
        dessheet.Range("A1:D" & lastRow).Borders.LineStyle = xlContinuous


        'this creates part of subject line
        For c = 2 To lastRow
        If c = lastRow Then
            Articles = Articles & dessheet.Range("A" & c).Value
        Else
            Articles = Articles & dessheet.Range("A" & c).Value & ", "
        End If
        Next c

    Else
    'No descriptions or Brand updates found, GT#O
    Exit Sub
    End If


'Find last row of description update summary
lastRow = dessheet.Range("A" & Rows.Count).End(xlUp).row
Set rng = dessheet.Range("A1:D" & lastRow)

'initialize some objects

    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

'Set the variable FullName to the Result of the function GetUsername
'This is the signature on the email
fullName = GetUsername

'Check for data on reqSheet sheet. If blank, set the duplicate sheet as reqSheet (This is a temporary fix) - BC
    On Error Resume Next
    If reqSheet.Range("B1").Value = "" And JobType = "AM" Then
        Set reqSheet = ActiveWorkbook.Worksheets("Article Maintain Request (2)")
    End If
     
'Set some message defaults
    'Article Maintain
    If JobType = "AM" Then
        messageTo = "Productcopy@rei.com"
        JobName = req.Name
        ProjName = reqSheet.Range("B1").Value
        Notes = reqSheet.Range("G1").Value
        Reason = reqSheet.Range("D5").Value
        Submitter = reqSheet.Range("B15").Value
        Vendor = reqSheet.Range("B2").Value & " " & reqSheet.Range("B3").Value
        Dept = reqSheet.Range("B5").Value
        Season = reqSheet.Range("D2").Value
        Priority = reqSheet.Range("D3").Value
    'Article Create + Maintenance
    ElseIf JobType = "AC" Then
        messageTo = "Productcopy@rei.com"
        JobName = req.Name
        ProjName = reqSheet.Range("B1").Value
        Notes = reqSheet.Range("G1").Value
        Reason = "No Reason field on Article Create template."
        Submitter = reqSheet.Range("B15").Value
        Vendor = reqSheet.Range("B2").Value & " " & reqSheet.Range("B3").Value
        Dept = reqSheet.Range("B7").Value
        Season = reqSheet.Range("D2").Value
        Priority = reqSheet.Range("D3").Value
    End If
    On Error GoTo 0
'Create a subject line for the email
   subject = "SAP Description or Brand Update: " & Articles

'Create a message body
        messageBody = "Hello PI Team," & _
        vbLf & vbLf & _
        "Descriptions or Brand have been changed in SAP for the following Articles:" & _
        vbLf & vbLf & _
        "File name:" & JobName & vbLf & _
        "Project name:" & ProjName & vbLf & _
        "Submitter: " & Submitter & vbLf & _
        "Notes from merchant: " & Notes & vbLf & _
        "Reason for article maintenance: " & Reason & vbLf & _
        "Vendor: " & Vendor & vbLf & _
        "Department: " & Dept & vbLf & _
        "Season: " & Season & vbLf & _
        "Priority: " & Priority

'sign that message
        Sig = vbLf & vbLf & "Thanks,"

'Open up that e-mail
    With MailItem
        .SentOnBehalfOfName = "masterdata@rei.com"
        .To = messageTo
        .subject = subject
        .BodyFormat = 1 ' Denotes olFormatHTML - HTML message formatting
        .Body = messageBody
        .HTMLBody = .HTMLBody & RangetoHTML(rng)
        .HTMLBody = .HTMLBody & "<Br> <Br>" & Sig
        .HTMLBody = .HTMLBody & "<Br>" & fullName


        'If InStr(UCase("kehanso"), UCase(Environ("username"))) > 0 Or _
        'InStr(UCase("wikeist"), UCase(Environ("username"))) > 0 Then
        '.send
        'Else
        .Display
        '.End If

    End With

    'clear out our objects.  Not really necessary, but probably a good practice.
    Set MailItem = Nothing
    Set OutlookApp = Nothing

    'Hides newly created description summary sheet and reactives request sheet
    dessheet.Visible = xlSheetHidden
    req.activate

    'MsgBox "Send Description/Brand Update Email to the PI Team."

End Sub

Private Function RangetoHTML(rng As Range)
' By Ron de Bruin.
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to paste the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1, 1).PasteSpecial Paste:=8
        .Cells(1, 1).PasteSpecial xlPasteValues, , False, False
        .Cells(1, 1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1, 1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close Savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function
Function GetUsername() As String

Dim objAD As Object
Dim objUser As Object
Dim strDisplayName As String

Set objAD = CreateObject("ADSystemInfo")
Set objUser = GetObject("LDAP://" & objAD.UserName)
strDisplayName = objUser.DisplayName
GetUsername = strDisplayName
End Function










