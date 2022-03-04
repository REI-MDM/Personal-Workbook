Attribute VB_Name = "New_MIT_v023"
Option Explicit

Sub NewMIT()
Attribute NewMIT.VB_ProcData.VB_Invoke_Func = "q\n14"
'******************************************************************************
' Open_Master_Input Macro
' Macro recorded 5/4/2011 by jheller
' Last edited 0424/2013 by nmokry _v014
'
' Keyboard Shortcut: Ctrl+Q
'
' This function assumes you store a copy of the Master input template in a
' folder called "Master Input Template" on your desktop.
'
' If you use the function "MHNameFile" (which you do not have to), that sub
' assumes you have a folder called "Archived Winshuttles" within your
' "Master Input Template" folder on your desktop.
'
' In an effort to make things slightly more error resistant and possibly a tiny
' bit faster, more fully qualified names are used.  For example:
'
'                               Range("B3").Select
'                                   Becomes
' Workbooks(SrcBook).Sheets("Article Create").Range("B3").Select
'
' This in and of itself is not faster, but instead of selecting, copying,
' switching sheets and pasting, you can simply reference the other workbook
' directly.  For example instead of copying a range on the template and then
' later activating the MIT and pasting the info in, you can use:
'
'   MITACSheet.Range("C11", "BW" & RowsOfDataOnAC).Value = _
'        ACTemplate.Range("C11", "BW" & RowsOfDataOnAC).Value
'
' This is faster than copying and pasting.
'
' Need to update for new AM flow -
' Check if we are on an AM or AC template.
' If only AM, ignore AC Stuff
' if AC, with no AM, do only AC Stuff
' If AC with AC and AM, do both.
'
' Set Bools, wrap all the AC stuff in an if, and all the AC in an if?
' Leave everything in place, sheet switching, but each individual item in an if
' or an additional conditional - I like this better
'
' Hard reference the source AM sheet, WS Maintain Article - New Tab, WS AM_IN
' Hide unnecessary Worksheets based on what we are doing - Follow on?
'
' Resize AM tabs as well as AC tabs
'******************************************************************************

'Timer
Dim starttime As Date
starttime = Now

Application.ScreenUpdating = False

Dim SRCBook As String               'The Active workbook when the macro starts
Dim SrcFullPath As String           'The Full path of the source template
Dim MasInputVers As String          'MIT Filename
Dim CurrentVers As String           'The current version number
Dim CurrentMIT As String            'Our Current MIT Filename
Dim OldMIT As String                'Our OLD MIT Filename
Dim MasterInput As String           'MIT Filepath
Dim ACTemplate As Worksheet         'Explicit reference the AC template
Dim AMTemplate As Worksheet         'Explicit reference to the AM template
Dim LocalUsers As String            'A list of users who store things on C:\

Dim ThereIsAM As Boolean            'Article Maintenance template?
Dim ThereIsAC As Boolean            'Article Create present?

Dim Version As String               'Holds our template version

Dim BlankCounter As Long            'in case there are blank lines on AM
Dim RowsOfDataOnAM As Integer       'The number of rows on the AM Tab
Dim i As Long                       'Generic counter/increment

Dim AMPromo As Boolean              'Promos on the AM tab?
Dim RowsOfDataOnAC As Integer       'The number of rows on the AC Tab
Dim RowsOnMIT As Integer            'The number on your MIT
Dim MITACSheet As Object            'Explicit reference to the MIT AC sheet
Dim MITWSACInNew As Object          'Explicit reference to MIT WS_In - New
Dim MITAMSheet As Object            'Explicit reference to MIT AM sheet
Dim MITWSMTInNew As Object          'Explicit reference to MIT WS_MT - NEW
Dim MITMaster As Object             'Explicit reference to the MIT Master tab
Dim MDCheckin As Boolean            'Variable to check if you are checking in
Dim TaskNumber As String            'Task number of the request
Dim aps As String                   'Application path separator - / or \
Dim req As Worksheet                'MAC facing request sheet.


    SRCBook = ActiveWorkbook.Name
    SrcFullPath = ActiveWorkbook.fullName
    ActiveWorkbook.Save
    
'Set our SRC Sheets
    On Error Resume Next
    Set ACTemplate = Workbooks(SRCBook).Sheets("Article Create")
    Set AMTemplate = Workbooks(SRCBook).Sheets("Maintain Article")
    On Error GoTo 0

'If there is an AC tab, then we know that there is AC
    ThereIsAC = False
    If Not ACTemplate Is Nothing Then
        TaskNumber = ACTemplate.Range("I8").Value
        If Not IsNumeric(Right(TaskNumber, 6)) Then
               MDCheckin = True
        Else
               MDCheckin = False
        End If
        ThereIsAC = True
    End If
    
'If there are no article numbers on the AM tab, then ThereIsAM = False
'This is a little ugly, but it will allow for times when MAs leave blank lines
'on their template
'Keyed off column A - Article number
    i = 8       'The last known populated cell on every template
    BlankCounter = 0
    ThereIsAM = False
    If Not AMTemplate Is Nothing Then
        Do While BlankCounter < 300
            i = i + 1
            If AMTemplate.Range("A" & i).Value <> "" Then
                ThereIsAM = True
                BlankCounter = 0 'Reset the blank counter, because we found data
            Else
                BlankCounter = BlankCounter + 1
            End If
        Loop
    Else
        i = 300
    End If
    RowsOfDataOnAM = i - 300
    
    On Error Resume Next
    If ThereIsAC Then
        Set req = Workbooks(SRCBook).Seets("Article Create Request")
    Else
        Set req = Workbooks(SRCBook).Sheets("Article Maintain Request")
    End If
    On Error GoTo 0
'Version Data should be stored in cell "H1" on any create or maintain
'templates moving forward
    If ThereIsAC Then
    'Current create sheets have the version in the correct spot, grab that one
        Version = ACTemplate.Range("H1").Value
    ElseIf ThereIsAM Then
    'There is no AC, grab the AM template version, which should be in "H1"
        Version = AMTemplate.Range("H1").Value
        If Version = "" Then
        'This one did not have it in "H1"
            Version = AMTemplate.Range("G1").Value
        End If
    End If
     
'A version of LocalUsers lives in NewMIT, AutoOpenUpdater, and in the MIT in
'MHNameFile
'    LocalUsers = "lagapin"

'define MasInputVers as dependent on the Version of the AC template we recieve
    
    If Version = "V10.0" Then
        MasInputVers = "Article_Create_Master_Input_v10_0_BAPI.xlsb"
    Else
        MasInputVers = "Article_Create_Master_Input_v11_0_BAPI.xlsb"
    End If
    
    aps = Application.PathSeparator
    MasterInput = Replace(Environ("UserProfile") & "/Desktop/Master Input Template/", "/", aps)
    MasterInput = MasterInput & MasInputVers
    
    
'    If InStr(UCase(LocalUsers), UCase(Environ("username"))) > 0 Then
'        MasterInput = "C:\Users\" & Environ("username") & _
'            "\Desktop\Master Input Template\" & MasInputVers
'    Else
'        MasterInput = "\\ahqnas1.reicorpnet.com\users\" & _
'            Environ("username") & "\Profile\Desktop\Master Input Template\" & _
'            MasInputVers
'    End If

        

'Check if "Master Input Template" is open, and if so, close it
    If IsFileOpen(MasterInput) Then
        Workbooks(MasInputVers).Close False 'False indicates no popup dialog
    End If
    

'Find the number of rows on the AC sheet
'Keyed off column D - Article Category
'JUST KIDDING!  This is now hidden and unreliable  Key off Description - "G"
    If ThereIsAC Then
        'Go to the very bottom of the sheet.  If that is just formulas, step
        'up until we find actual data
        
        i = ACTemplate.Range("G" & Rows.Count).End(xlUp).row
        Do While ACTemplate.Range("G" & i).Value = ""
            i = i - 1
        Loop
        RowsOfDataOnAC = i

        'Make sure Article Category - "Generic" is filled in
        'Maybe a bad idea if we ever want to make singles with the AC template?
        ACTemplate.Range("D11:D" & RowsOfDataOnAC).Value = "Generic Article"
        

    'Unhides the "generic" and "article number" columns
        ACTemplate.Columns("A:B").EntireColumn.Hidden = False

    'Adds a formula into the vendor number column to populate it. This is
    'commonly overwritten/missing
        If Not (Range("J2") = 75 Or Range("J2") = 87 Or Range("J2") = _
            56 Or Range("J2") = 98) Then
            If Version = "V8.4" Then
                    Range("BO11:BO" & RowsOfDataOnAC).NumberFormat = "General"
                    Range("BO11:BO" & RowsOfDataOnAC).FormulaR1C1 = _
                        "=IF(RC7<>"""",R2C9,"""")"
                    Range("BO11:BO" & RowsOfDataOnAC).Calculate
                    Range("BO11:BO" & RowsOfDataOnAC).Value = _
                        Range("BO11:BO" & RowsOfDataOnAC).Value
                Else
                    Range("BR11:BR" & RowsOfDataOnAC).NumberFormat = "General"
                    Range("BR11:BR" & RowsOfDataOnAC).FormulaR1C1 = _
                        "=IF(RC7<>"""",R2C9,"""")"
                    Range("BR11:BR" & RowsOfDataOnAC).Calculate
                    Range("BR11:BR" & RowsOfDataOnAC).Value = _
                        Range("BR11:BR" & RowsOfDataOnAC).Value

            End If ' /End version check
        Else
            'Do nothing
        End If '/fill vendor code column
        
        'External references can get messed up sometimes.  Noticed an error
        'where crossdock values were flipping from C to W.  We'll paste values
        'in that column for templates earlier than V8.1
        If CDbl(Right(Version, 3)) < 8.1 Then
            Range("BJ11:BJ" & RowsOfDataOnAC).Value = _
                Range("BJ11:BJ" & RowsOfDataOnAC).Value
        End If
        
            
            
            
    End If  '/If ThereIsAC
    
'This opens the Master Input Template and finds your last row of formatting
    Workbooks.Open (MasterInput)
    
    On Error Resume Next
        Set MITACSheet = Workbooks(MasInputVers).Sheets("AC_Tmpt")
        Set MITWSACInNew = Workbooks(MasInputVers).Sheets("WS_AC")
        Set MITAMSheet = Workbooks(MasInputVers).Sheets("AM_Tmpt")
        Set MITWSMTInNew = Workbooks(MasInputVers).Sheets("WS_MT")
        Set MITMaster = Workbooks(MasInputVers).Sheets("Master")
    On Error GoTo 0

'Find the number of rows on our MIT

    RowsOnMIT = MITWSACInNew.Range("A65536").End(xlUp).row
    
'If your MIT is not big enough for the request, resize, with a small buffer
'This is kind of ugly and could be cleaned up probably, but it works.
    If RowsOfDataOnAC > RowsOnMIT Then
        RowsOfDataOnAC = RowsOfDataOnAC + 10
        Sheets("UPC Swap").Select
        With Rows(RowsOnMIT - 1 & ":" & RowsOnMIT)
            .AutoFill Destination:=.Resize(2 + RowsOfDataOnAC - RowsOnMIT)
        End With
        MITWSACInNew.Select
        With Rows(RowsOnMIT - 1 & ":" & RowsOnMIT)
            .AutoFill Destination:=.Resize(2 + RowsOfDataOnAC - RowsOnMIT)
        End With
    End If

'Same as above, but for the AM Tab.  Again, ugly, but it should work.
    If RowsOfDataOnAM > RowsOnMIT Then
        MITWSMTInNew.Select
        With Rows(RowsOnMIT - 1 & ":" & RowsOnMIT)
            .AutoFill Destination:=.Resize(2 + RowsOfDataOnAM - RowsOnMIT)
        End With
    End If

'Move the Article Create data into the Master Input Template
    If ThereIsAC Then
        ACTemplate.Range("A11:CF" & RowsOfDataOnAC).Copy
        MITACSheet.Range("A11:CF" & RowsOfDataOnAC).PasteSpecial _
            xlPasteValuesAndNumberFormats
        'MITACSheet.Cells.Font.Name = "Terminal"
    End If
'Move the AM Data
    If ThereIsAM Then
        MITAMSheet.Range("A9", "BK" & RowsOfDataOnAM).Value = _
        AMTemplate.Range("A9", "BK" & RowsOfDataOnAM).Value
    End If

'Populate the Master Tab with some data.  This is sorted in no particular
'order.  If we like, after this is more fully fleshed out, we can arrange these
'flags into some more sensical order.
    MITMaster.Range("B2").Value = SRCBook
    MITMaster.Range("B3").Value = SrcFullPath
    MITMaster.Range("B4").Value = ThereIsAM
    MITMaster.Range("B5").Value = AMPromo
    MITMaster.Range("B6").Value = ThereIsAC
    MITMaster.Range("B7").Value = RowsOfDataOnAC
    MITMaster.Range("B8").Value = RowsOfDataOnAM
    MITMaster.Range("B34").Value = TaskNumber
    MITMaster.Range("B39").Value = LastNonMDSaver(SRCBook)
'Hard coded "Output Created" to false on each open
    MITMaster.Range("B17").Value = False
    MITMaster.Range("B18").Value = Version
    MITMaster.Range("B27").Value = MDCheckin
    MITMaster.Range("T2").Value = Format(Now() - starttime, "hh:mm:ss")
'CharLookup Times
    If ThereIsAC Then
        MITMaster.Range("T9").Value = Format(ACTemplate.Range("W1").Value, "hh:mm:ss")
        MITMaster.Range("T10").Value = Format(ACTemplate.Range("W2").Value, "hh:mm:ss")
    End If
'Request info
    If Not req Is Nothing Then
        MITMaster.Range("B42").Value = req.Range("B1").Value    'Project Name:
        MITMaster.Range("B43").Value = req.Range("B2").Value    'Vendor #:
        MITMaster.Range("B44").Value = req.Range("B3").Value    'Vendor Name:
        MITMaster.Range("B45").Value = req.Range("B4").Value    'Brand #: (optional)
        MITMaster.Range("B48").Value = req.Range("B6").Value    'Additonal Vendor Contact Name (Optional):
        MITMaster.Range("B49").Value = req.Range("D6").Value    'Vendor Contact Email (Optional):
        MITMaster.Range("B51").Value = req.Range("D7").Value    'Vendor Catalog ?
        MITMaster.Range("B52").Value = req.Range("D2").Value    'Season :
        MITMaster.Range("B53").Value = req.Range("D3").Value    'Priority:
        MITMaster.Range("B54").Value = req.Range("G3").Value    'Reason for Critical Priority
        MITMaster.Range("B55").Value = req.Range("D4").Value    'Add to Promotion ?
        MITMaster.Range("B56").Value = req.Range("G1").Value    'Notes or special instructions:
        MITMaster.Range("B57").Value = req.Range("B15").Value   'Requested by:
        MITMaster.Range("B58").Value = req.Range("B16").Value   'Date Requested:
        If ThereIsAC Then
            MITMaster.Range("B50").Value = req.Range("B7").Value    'Department
            MITMaster.Range("B46").Value = req.Range("B5").Value    'Vendor Contact Name:
            MITMaster.Range("B47").Value = req.Range("D5").Value    'Vendor Contact Email:
        Else
        'article maintain mappings
            MITMaster.Range("B50").Value = req.Range("B5").Value    'Department
            MITMaster.Range("B59").Value = req.Range("D5").Value    'Reason for Article Maintenance
            MITMaster.Range("B60").Value = req.Range("G3").Value    'Crit price change approved by
        End If
    End If



'Kick off our "First Open" Macro in the MIT.  This runs various other macros.
    Application.Run "'" & MasterInput & "'" & "!FirstOpen"
    
Application.ScreenUpdating = True

'******************************************************************************
End Sub
Function IsFileOpen(Filename As String)
'******************************************************************************
' MH Pulled this function off the internet
' It checks to see if a file is open.  Imagine that.
'
' Previous versions would return errors if the filename you passed it did not
' exist
'
' Currently it is only in use for our Open Master Input Ctrl-Q macro,
' but I put it in a separate FunkShuns module.  That would be great if we built
' up a library of versatile useful functions. but so far we have not.  It is
' now included here for ease of use/transfer.
'******************************************************************************
    
    Dim iFileNum As Long
    Dim iErr As Long
     
    On Error Resume Next
    iFileNum = FreeFile()
    Open Filename For Input Lock Read As #iFileNum
    Close iFileNum
    iErr = err
    On Error GoTo 0
     
    Select Case iErr
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case Else: IsFileOpen = False
    End Select
'******************************************************************************
End Function
'******************************************************************************
' Change Log
'**_v013**
'   - Department numbers have changed.  D85 books used to exempt from the
'   vendor formula overwrite column BO.  That no longer exists.  D36 replaces.
'**_v012**
'   - Versioning is stupid.  I had hardcoded a check for "current Version"
'   which would need to be maintained.  This quick-n-dirty fix swaps behaviors
'   to check OLD versions (should not need to be supported, but left in) and
'   otherwise assumes the version is current.
'   - Crossdock References were getting borked.  Paste customer input values
'   before opening the MIT.
'**_v011**
'   - Updated MIT filetype to .xlsb.  Removed some "cleaning" code and moved
'   it into the MIT.  Added Task Number recording.  Removed unnecessary resize
'   of MIT "Tmpt" pages for large requests.  Ensured Calculation and values for
'   Vendor Number columns.
'**_v009 - v010**
'   - No Clue.  Ooops!
'**_v008**
'   - Forgot to update notes for 7, forget what exactly I did.  Added
'   localusers so that saves can happen for Lex and any other user who might
'   save to their C:\ drive.
'**_v006**
'   - Old Template version tracking was overwriting data for D97 creates, this
'   has been fixed.
'   - Fixed an error in the "Fill Vendor number formula" section that has
'   been around for a while but did not pop up often.
'**_v005**
'   - VPNs were getting improperly stripped of leading values in rare cases
'   with the ACSheet.values = template.values method of data transfer.  We
'   reverted back to a copy/paste method, which I like less, but should work
'    - Also the AM checker would fail if there were more than 5 blank lines at
'    the beginning, I have upped it to 300.  Which is not elegant, but oh well
'
'

Private Function LastNonMDSaver(bookname As String) As String
Dim SS As Worksheet
Dim lastRow As Long
Dim i As Long
    On Error Resume Next
    Set SS = Workbooks(bookname).Worksheets("Savelog")
    On Error Resume Next
    LastNonMDSaver = "Error"
    If SS Is Nothing Then Exit Function
    lastRow = SS.Range("B" & Rows.Count).End(xlUp).row
    For i = lastRow To 2 Step -1
        If SS.Range("C" & i).Value = False Then Exit For
    Next i
    If i = 2 Then Exit Function
    LastNonMDSaver = SS.Range("B" & i).Value
    Set SS = Nothing
End Function
