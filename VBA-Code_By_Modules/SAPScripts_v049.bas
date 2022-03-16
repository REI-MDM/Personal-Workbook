Attribute VB_Name = "SAPScripts_v049"
Option Explicit
Sub SAP_AssortDelete_WSOA6()

'**********************************************************************************************************************************************
'Built to run on Assortment Maintain when lots of deletions are needed. Can be used in other applications but with some tweeking.
'Ask Phil or Kevin about that
'-Kevin
'**********************************************************************************************************************************************

'Dim sapGuiAuto As Object
'Dim Application As Object
'Dim connection As Object
'Dim session As Object
Dim WB As Workbook
Dim WS As Worksheet
Dim i As Long
Dim Site As String


Set WB = ActiveWorkbook
Set WS = WB.Worksheets("WS Delete")


'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If
'start on row to because we have headers
i = 2

'launch WSOA6
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwsoa6"
session.FindById("wnd[0]").sendVKey 0
'enter assortment on first row in the assortment field
session.FindById("wnd[0]/usr/ssubSUB1:SAPLWSOWRSZ:1110/ctxtS_ASORT-LOW").Text = WS.Range("A" & i).Value
'click execute
session.FindById("wnd[0]/tbar[1]/btn[8]").press

'Do stuff until the last row
Do While WS.Range("A" & i) <> ""
    'Press Find button and drop in Site from row i
    session.FindById("wnd[0]/usr/cntlCUSTOM_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell[0]").PressButton "&FIND"
    Site = Format(WS.Range("C" & i), "0000000000")
    session.FindById("wnd[1]/usr/txtLVC_S_SEA-STRING").Text = Site
    session.FindById("wnd[1]").sendVKey 0
    
    'if site requested to be removed is not found in assortment get past pop up window and log that site was not in assortment
    On Error Resume Next
    session.FindById("wnd[1]/tbar[0]/btn[12]").press
    If err.Number <> 0 Then
        session.FindById("wnd[1]").sendVKey 0
        WS.Range("D" & i) = "Site not found in Assortment"
        err.Clear
    End If
    On Error GoTo 0
    
    'If the site was in Assortment, Delete Assortment and if there is no issue logged log that we removed site
    session.FindById("wnd[0]/usr/cntlCUSTOM_CONTAINER/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell[0]").PressButton "DELETE_ASSIGNMENT"
    If WS.Range("D" & i) = "" Then WS.Range("D" & i) = "Site removed from assortment Assortment"
    
    i = i + 1
    'save assortment and go onto next row
    'this currently loops poorly for same assortment removing multiple sites but I'll fix that later
    If WS.Range("A" & i) <> WS.Range("A" & i - 1) And WS.Range("A" & i) <> "" Then
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        'if nothing was saved this hits enter so we can keep going
        If session.FindById(session.ActiveWindow.Name & "/sbar").MessageType = "E" Then
            session.FindById("wnd[0]").sendVKey 0
        End If
        session.FindById("wnd[0]/usr/ssubSUB1:SAPLWSOWRSZ:1110/ctxtS_ASORT-LOW").Text = WS.Range("A" & i).Value
        session.FindById("wnd[0]/tbar[1]/btn[8]").press
        On Error Resume Next
        If session.FindById("wnd[1]").Text = "Assortment Assignment Tool" Then
            session.FindById("wnd[1]/usr/btnBUTTON_1").press
        End If
        On Error GoTo 0
    End If
    
    'if last line we're done and can exit WSOA6
    If WS.Range("A" & i) = "" Then
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        session.FindById("wnd[0]/tbar[0]/btn[15]").press
    End If
Loop

Call EndSAPCON

End Sub

Private Sub p_SAP_MBIMaintain_MASS_CHARVAL()
'Still in testing phase!
'KH

Dim WS As Workbook
Dim art As Long
Dim i As Integer

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

Set WS = ActiveWorkbook
i = 17
Do While WS.Worksheets("Maintain Article").Range("A" & i) <> ""
art = WS.Worksheets("Maintain Article").Range("A" & i)
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = art
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/btnOES_PDOWN").press
session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/btnOES_PDOWN").press
session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").Text = "NO"
session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").SetFocus
session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").CaretPosition = 2
session.FindById("wnd[0]/tbar[0]/btn[11]").press
i = i + 1

Loop

Call EndSAPCON

End Sub

Sub SAP_VendorContacts_XK02_MAP1_MAP2()

'*********************************************************************************************************************
'This is a work in progess to automate the input of contact data ACH + any Vendor Ops related work.
'Done - Step .5: error check a couple things before running anything into SAP (Does Vendor/Contact person Exist in SAP
'         are Emails formatted correctly with @ symbol) We can add more in but for now those are the major thinks that break.
'         Maybe do this while it runs but that will be harder.
'Done - Step 1: XK02 loops through all rows of vendor column and checks if ACH email infor exists.
'        If it does, we will want AP and VOps to submit what it should be (no ADD/REMOVES)
'        We then delete all emails addresses and then loop through the rows with the same vendor number and input them
'        in the order provided.
'Done - Step 2: XK02 loops through all rows of Contact Person column and checks if we are removing a vendor.
'        Message box prompts use to manually delete contact with specified First Name/Last Name/Department. Verify Contact Person #
'Step 3: MAP1 loops through all rows of Contact Person column and checks if we are adding a contact. Automated add of contact.
'        Possible issue with adding multiple emails when creating a new contact maybe we don't do this?
'Step 4: MAP2 loops through all rows of Contact Person column and checks if we are modifying a contact. Automated modification of contact.
'        similar looping to step one by deleting all emails and then adding them in the order provided.
'
'Ways to improve:
'When adding a new contact we can only enter one email. We may run into issues with this
'Currently all cell references are static which will cause issues if and when the vendor maintain template/Vendor
'MIT changes. Which it will.
'-Kevin Hanson
'*********************************************************************************************************************
Dim WB As Workbook
Dim WS As Worksheet
Dim i As Long
Dim ci As Integer
Dim ei As Integer
Dim ei1 As Integer
Dim Conti As Integer
Dim Vendor As Long
Dim Contact As Long
Dim responce As Integer
Dim startRow As Integer
Dim lastRow As Integer
Dim endRow As Integer
Dim Successcnt As Integer
Dim errorcnt As Integer
Dim Vendorcol As String
Dim ACHEmailcol As String
Dim ConActcol As String
Dim ConPersoncol As String
Dim Departcol As String
Dim FirstNamecol As String
Dim LastNamecol As String
Dim Emailcol As String
Dim Functioncol As String
Dim Telecol As String
Dim TeleExtcol As String
Dim Mobilecol As String
Dim Faxcol As String
Dim FaxExtcol As String
Dim RCEmailcol As String
Dim RCCOcol As String
Dim RCStreetcol As String
Dim RCStreet2col As String
Dim RCStreet3col As String
Dim RCStreet4col As String
Dim RCHousecol As String
Dim RCCitycol As String
Dim RCRegioncol As String
Dim RCPostalcol As String
Dim RCCountrycol As String
Dim RCTelecol As String
Dim RCTeleExtcol As String
Dim RCFaxcol As String
Dim RCFaxExtcol As String
Dim ACHLogcol As String
Dim RemovalLogcol As String
Dim AddsLogcol As String
Dim ModifyLogcol As String
Dim PopUpMessage As String
Dim StatusBar As String
Dim ErrorResponce As Integer
Dim ContactName As String
Dim Tableview As Integer
Dim HaveItBreak As Boolean


'Set some objects
Set WB = ActiveWorkbook
Set WS = WB.Worksheets("Maintain_WSData")
HaveItBreak = False

If UCase(Environ("username")) = "KEHANSO" Then
    If MsgBox("Have it Break?", vbYesNo) = vbYes Then
        HaveItBreak = True
    End If
End If
    




'Set Columns as variables so this is more dynamic
    i = 1
    Do While WS.Cells(6, i).Value <> ""
        
'        If i > 32 And i < 55 Then
        'Debug.Print ws.Cells(6, i).Value
'        End If
        Select Case Trim(WS.Cells(6, i).Value)
            Case Is = "Vendor Number"
                Vendorcol = columnletter(i)
                
            Case Is = "ACH E-mail" & Chr(10) & "(Enter in the order they appear in SAP)"
                ACHEmailcol = columnletter(i)
                
            Case Is = "Contact Action" & Chr(10) & "(Add, Modify, Remove)"
                ConActcol = columnletter(i)
                
            Case Is = "Contact Person #"
                ConPersoncol = columnletter(i)
                
            Case Is = "Department"
                Departcol = columnletter(i)
                
            Case Is = "First Name"
                FirstNamecol = columnletter(i)
                
            Case Is = "Name"
                LastNamecol = columnletter(i)
                
            Case Is = "E-mail" & Chr(10) & "(Entered in the order they should be in in SAP)"
                Emailcol = columnletter(i)
                
            Case Is = "Function"
                Functioncol = columnletter(i)
                
            Case Is = "Telephone"
                Telecol = columnletter(i)
                
            Case Is = "Telephone-Ext"
                TeleExtcol = columnletter(i)
                
            Case Is = "Mobile Phone"
                Mobilecol = columnletter(i)
                
            Case Is = "Fax"
                Faxcol = columnletter(i)
                
            Case Is = "Fax-Ext"
                FaxExtcol = columnletter(i)
                
            Case Is = "Returns Contact Email"
                RCEmailcol = columnletter(i)
                
            Case Is = "Returns Contact C/O"
                RCCOcol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "House Number"
                RCHousecol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Street"
                RCStreetcol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Street 2"
                RCStreet2col = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Street 3"
                RCStreet3col = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Street 4"
                RCStreet4col = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "City"
                RCCitycol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Region"
                RCRegioncol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Postal Code"
                RCPostalcol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Country"
                RCCountrycol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Telephone"
                RCTelecol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Telephone ext"
                RCTeleExtcol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Fax"
                RCFaxcol = columnletter(i)
                
            Case Is = "Returns Contact" & Chr(10) & "Fax ext"
                RCFaxExtcol = columnletter(i)
                
'            Case Is = "Dividend e"
'                ACHLogcol = ColumnLetter(i)
'
'            Case Is = "Dividend e"
'                RemovalLogcol = ColumnLetter(i)
'
'            Case Is = "Dividend e"
'                AddsLogcol = ColumnLetter(i)
'
'            Case Is = "Dividend e"
'                ModifyLogcol = ColumnLetter(i)
                


        End Select
        i = i + 1
    Loop
    
'Set Log columns based off the last column with data. WS log is 2 to the right these 4 are right after that
ACHLogcol = columnletter(i + 2)
RemovalLogcol = columnletter(i + 3)
AddsLogcol = columnletter(i + 4)
ModifyLogcol = columnletter(i + 5)


'Check to see if all columns got set, if someone updates the column header name this will break
'maybe set up a way to tell processor which column header didn't get set?
'Sorry but I didn't have time to set up a cool way to check which one broke. Set a break point at the if below and look
'in the immidiate window  to see which variable wasn't set. Ask Michael/Nathanael/Brian for help

'Add Street2 and Street3

If Vendorcol = "" Or _
    ACHEmailcol = "" Or _
    ConActcol = "" Or _
    ConPersoncol = "" Or _
    Departcol = "" Or _
    FirstNamecol = "" Or _
    LastNamecol = "" Or _
    Emailcol = "" Or _
    Functioncol = "" Or _
    Telecol = "" Or _
    TeleExtcol = "" Or _
    Mobilecol = "" Or _
    Faxcol = "" Or _
    FaxExtcol = "" Or _
    RCEmailcol = "" Or _
    RCCOcol = "" Or _
    RCHousecol = "" Or RCStreetcol = "" Or RCStreet2col = "" Or RCStreet3col = "" Or RCStreet4col = "" Or _
    RCCitycol = "" Or RCRegioncol = "" Or RCPostalcol = "" Or RCCountrycol = "" Or _
    RCTelecol = "" Or RCTeleExtcol = "" Or RCFaxcol = "" Or RCFaxExtcol = "" Then
        
    MsgBox "A header column text was changed and column variable did not get set. Exiting Sub. You'll have to step into this sub, " & _
    "figure out what column isn't coded correctly, and fix the live MIT. Reach out to Kevin or Michael for assistance."
    Exit Sub
End If




'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If


'***********************************************************************************************************
'Step .5 Check for things that will cause issues when inputting data into SAP

'Check vendor numbers and make sure they exist in SAP
'Lets only run these checks once
If WS.Range("CS1") = "" Then
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
    session.FindById("wnd[0]").sendVKey 0
    i = 7
    Do While WS.Range(Vendorcol & i) <> ""
        Vendor = WS.Range(Vendorcol & i)
        session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
        session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = ""
        session.FindById("wnd[0]/usr/ctxtRF02K-EKORG").Text = ""
        session.FindById("wnd[0]").sendVKey 8
        'session.findById("wnd[0]/usr/chkRF02K-D0110").Selected = True
        'session.findById("wnd[0]").sendVKey 0
        If session.FindById("wnd[0]/sbar").MessageType = "E" Then
            MsgBox "Looks like vendor #" & Vendor & " (line " & i & ") does not exist. Check in with submitter to figure out what vendor it should be."
            WS.Range(Vendorcol & i).Interior.ColorIndex = 6
            errorcnt = errorcnt + 1
        Else
    
        End If
        i = i + 1
    Loop
    
    'Check Contact Person Numbers and make sure they exist
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmap2"
    session.FindById("wnd[0]").sendVKey 0
    i = 7
    Do While WS.Range(Vendorcol & i) <> ""
    If UCase(WS.Range(ConActcol & i)) = "MODIFY" Or UCase(WS.Range(ConActcol & i)) = "REMOVE" Then
            Contact = WS.Range(ConPersoncol & i)
            Vendor = WS.Range(Vendorcol & i)
    
            session.FindById("wnd[0]/usr/ctxt*KNVK-PARNR").Text = Contact
            session.FindById("wnd[0]").sendVKey 0
        'Status bar message will be Error if the contact doesn't exist in SAP
        If session.FindById("wnd[0]/sbar").MessageType = "E" Then
            MsgBox "Looks like Contact Person #" & Contact & " (line " & i & ") does not exist. Check in with submitter to figure out what Contact it should be."
            WS.Range(ConPersoncol & i).Interior.ColorIndex = 6
            errorcnt = errorcnt + 1
            
        Else
            'in case submitter puts the wrong contact person number and it does exist we can check to confirm that contact belongs to the vendor specified
            If session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor Then
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                'SAP gives warning stat bar if multiple contacts exist with the same name click enter
                'so we can get back to MAP2 home screen
                If session.FindById("wnd[0]/sbar").MessageType = "W" Then
                    session.FindById("wnd[0]").sendVKey 0
                End If
            Else
                MsgBox "Contact " & Contact & " does not belong to Vendor " & Vendor & " Check in with submitter to confirm Contact Person number"
                WS.Range(ConPersoncol & i).Interior.ColorIndex = 6
                errorcnt = errorcnt + 1
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                'SAP gives warning stat bar if multiple contacts exist with the same name click enter
                'so we can get back to MAP2 home screen
                If session.FindById("wnd[0]/sbar").MessageType = "W" Then
                    session.FindById("wnd[0]").sendVKey 0
                End If
            End If
        End If
    End If
        i = i + 1
    Loop
    
    'Confirm Email Addresses (ACH, Contact, Returns Contact) have @ sign
    i = 7
    Do While WS.Range(Vendorcol & i) <> ""
        If WS.Range(ACHEmailcol & i) <> "" Then
            If InStr(WS.Range(ACHEmailcol & i).Text, "@") > 0 Or UCase(WS.Range(ACHEmailcol & i)) = "REMOVE" Then
            Else
            MsgBox "Looks like the ACH email address on line " & i & " (" & WS.Range(ACHEmailcol & i).Text & ") is missing the @ symbol. You will need to fix this before using this macro."
            WS.Range(ACHEmailcol & i).Interior.ColorIndex = 6
            errorcnt = errorcnt + 1
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    Loop
    
    i = 7
    Do While WS.Range(Vendorcol & i) <> ""
            If WS.Range(Emailcol & i) <> "" Or UCase(WS.Range(Emailcol & i)) = "REMOVE" Then
            If InStr(WS.Range(Emailcol & i).Text, "@") > 0 Then
            Else
            MsgBox "Looks like the Contact email address on line " & i & " (" & WS.Range(Emailcol & i).Text & ") is missing the @ symbol. You will need to fix this before using this macro."
            WS.Range(Emailcol & i).Interior.ColorIndex = 6
            errorcnt = errorcnt + 1
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    
    Loop
    
    i = 7
    Do While WS.Range(Vendorcol & i) <> ""
            If WS.Range(RCEmailcol & i) <> "" Or UCase(WS.Range(RCEmailcol & i)) = "REMOVE" Then
            If InStr(WS.Range(RCEmailcol & i).Text, "@") > 0 Then
            Else
            MsgBox "Looks like the Contact email address on line " & i & " (" & WS.Range(RCEmailcol & i).Text & ") is missing the @ symbol. You will need to fix this before using this macro."
            WS.Range(RCEmailcol & i).Interior.ColorIndex = 6
            errorcnt = errorcnt + 1
            End If
            i = i + 1
        Else
            i = i + 1
        End If
    
    Loop
    'If there were any issues, if issues don't run macro
    If errorcnt > 0 Then
        'ErrorResponce = MsgBox("Looks like there are some issues with Vendor/Contact numbers or with Email Address formating. Do you want to run on everything but those rows?", vbYesNo)
        MsgBox "Looks like there are some issues with Vendor/Contact numbers or with Email Address formating."
        MsgBox "If you can't reach the submittor and want to run this script, Delete the rows associated with the issue file on the maintain sheet and mark Needs Updates and add Notes for that file in summary when you close out."
        EndSAPCON
        Exit Sub
    End If
    WS.Range("CS1") = "Checked"
End If
'************************************************************************************************************
'Step 1 - ACH email on Address tab in XK02

If Not HaveItBreak Then
    On Error GoTo ACHEndLog
Else
    On Error GoTo 0
End If

errorcnt = 0
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
session.FindById("wnd[0]").sendVKey 0
i = 7

Do While WS.Range(Vendorcol & i) <> "" 'is there a vendor in vendor row
    If WS.Range(ACHLogcol & i).Interior.ColorIndex = 4 Then 'Skip logged success rows
    i = i + 1
    Else
        If WS.Range(ACHEmailcol & i) <> "" Then
            
            'XK02 enter vendor number, company code, purch org, unselect all, check Addess, Enter
            Vendor = WS.Range(Vendorcol & i)
            session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
            session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
            session.FindById("wnd[0]/usr/ctxtRF02K-EKORG").Text = "1000"
            session.FindById("wnd[0]").sendVKey 8
            session.FindById("wnd[0]/usr/chkRF02K-D0110").Selected = True
            session.FindById("wnd[0]").sendVKey 0
        
            'loop to delete all ACH email then add all emails from request
            ei = 0
            ci = 0
            Do While WS.Range(Vendorcol & i).Value = Vendor And WS.Range(ACHEmailcol & i).Text <> "" 'is vendor number the same as last
                session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA1:0300/subCOUNTRY_SCREEN:SAPLSZA1:0301/btnG_ICON_SMTP").press
                If ci = 0 Then 'have we deleted anything yet
                    'loop to delete all contacts
                    Do Until session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,0]").Text = "" 'Check this
                       session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6").GetAbsoluteRow(0).Selected = True
                       session.FindById("wnd[1]/tbar[0]/btn[14]").press 'delete line
                    Loop
                End If 'have we deleted anything yet
                'loop to add new email addresses until there are no more emails/vendor changes/No adds if Remove
                Do While WS.Range(ACHEmailcol & i).Text <> "" And WS.Range(Vendorcol & i).Value = Vendor
                    If UCase(WS.Range(ACHEmailcol & i).Text) = "REMOVE" Then
                    i = i + 1
                    WS.Range(ACHLogcol & i - 1) = "ACH Email Maintained for Vendor " & Vendor & " at " & Format(Now, "mm/dd/yyyy hh:nn")
                    WS.Range(ACHLogcol & i - 1).Interior.ColorIndex = 4
                    Else
                    session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0," & ci & "]").Text = UCase(WS.Range(ACHEmailcol & i).Text)
                    session.FindById("wnd[1]/tbar[0]/btn[13]").press
                If ci = 2 Then 'set position of next email after add button is hit
                    ci = 2
                    Else
                    ci = ci + 1
                End If 'set position of next email after add button is hit
                    ei = ei + 1
                    i = i + 1
                    End If
                Loop 'add new ACH email until vendor changes
                
            Loop 'is vendor number the same as last
            
            session.FindById("wnd[1]/tbar[0]/btn[0]").press ' ok code back to address page
            session.FindById("wnd[0]/tbar[0]/btn[11]").press 'lets end xk02 session to go on to finish all ACH address page stuff
        
        
        
            If session.ActiveWindow.Name = "wnd[0]" And session.FindById("wnd[0]/sbar").MessageType <> "E" Then ' No Error, not in a pop up window, and no error status
                WS.Range(ACHLogcol & i - ei & ":" & ACHLogcol & i - 1) = "Success: ACH Email Maintained for Vendor " & Vendor & " at " & Format(Now, "mm/dd/yyyy hh:nn")
                WS.Range(ACHLogcol & i - ei & ":" & ACHLogcol & i - 1).Interior.ColorIndex = 4
                ei = 0
            Else ' vba error, in a pop window, or error in status bar
                err.Raise (619)
ACHEndLog:
                lastRow = lastRow + 1
                'Err.Clear
                
                'set veriables for log
                If session.ActiveWindow.Name <> "wnd[0]" Then
                    PopUpMessage = session.ActiveWindow.PopupDialogText
                Else
                    StatusBar = session.FindById(session.ActiveWindow.Name & "/sbar").Text
                End If
                
                'close out of pop up windows to get to main window
                Do Until session.ActiveWindow.Name = "wnd[0]"
                    session.FindById(session.ActiveWindow.Name).Close
                Loop
                
                'exit out of main window without saving vendor
                session.FindById("wnd[0]/tbar[0]/btn[12]").press
                If session.ActiveWindow.Name = "wnd[1]" Then
                session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                End If
                session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
                session.FindById("wnd[0]").sendVKey 0
                
                'log error in correct rows of the vendor
                WS.Range(ACHLogcol & i - ei & ":" & ACHLogcol & i - 1) = "SAP SCRIPTING ERROR: " & PopUpMessage & " " & StatusBar
                WS.Range(ACHLogcol & i - ei & ":" & ACHLogcol & i - 1).Interior.ColorIndex = 3
                
                
                
                'increment errorcounter for msgbox
                errorcnt = errorcnt + ei
                PopUpMessage = ""
                StatusBar = ""
                Resume ACHENDLogExit
            End If
            
ACHENDLogExit:
        Else
            i = i + 1
            WS.Range(ACHLogcol & i - 1) = "No Changes Made to ACH Email"
        End If 'ACH update present
    End If ' Skip logged success rows
Loop ' is there a vendor in vendor row

If errorcnt = 0 Then
'Ws.Range("A2").ClearContents
Else
MsgBox errorcnt & " lines had errors for ACH Email requests. Check column " & ACHLogcol & " for logs to see what errors where."
End If
WS.Columns(ACHLogcol).AutoFit

'End Step 1
'****************************************************************************************************************
'Start Step 2 - Removal of contact in XK02

'list in XK02
errorcnt = 0
err.Clear
On Error GoTo 0
'On Error GoTo Removelog
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
session.FindById("wnd[0]").sendVKey 0
i = 7
Do While WS.Range(Vendorcol & i) <> "" 'is there a vendor in vendor row
    Vendor = WS.Range(Vendorcol & i)
    If WS.Range(RemovalLogcol & i).Interior.ColorIndex = 4 Then 'Skip logged success rows
        i = i + 1
    Else
        If WS.Range(ConActcol & i).Text = "Remove" Then
            If Left(WS.Range(Departcol & i).Text, 4) = "0007" Then
                MsgBox ("We should almost never delete a Returns contact. Skipping removal on row " & i & vbCrLf & _
                "Reach out to the submitter and ask them what's going on, then manually delete.")
                WS.Range(ConActcol & i & ":AQ" & i).Interior.ColorIndex = 6
                i = i + 1
            Else
                session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
                session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
                session.FindById("wnd[0]/usr/ctxtRF02K-EKORG").Text = "1000"
                session.FindById("wnd[0]").sendVKey 8
                session.FindById("wnd[0]/usr/chkWRF02K-D0380").Selected = True
                session.FindById("wnd[0]").sendVKey 0
'Lets automate this again...
                Tableview = session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER").VisibleRowCount
                Conti = 0
'loop through SAP contact table till the end page down when we reach the visable row count
                Do Until session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[1," & Conti & "]").Text = ""
'for each row in SAP contact table see if the contact matches
'First Name, Last name, and Dept for contact to be deleted on request
'if we found a match go into the partner details
                    If UCase(session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[1," & Conti & "]").Text) = UCase(WS.Range(FirstNamecol & i).Text) And _
                    UCase(session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAME1[2," & Conti & "]").Text) = UCase(WS.Range(LastNamecol & i).Text) And _
                    session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/ctxtKNVK-ABTNR[4," & Conti & "]").Text = Left(WS.Range(Departcol & i).Text, 4) Then
                        session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[1," & Conti & "]").SetFocus
                        session.FindById("wnd[0]/tbar[1]/btn[2]").press
'does contact person number match what is on request?
'if Contact number is correct go back check that SAP table
'is in the same position and if so delete contact
                        If session.FindById("wnd[0]/usr/txtKNVK-PARNR").Text = WS.Range(ConPersoncol & i).Text Then
                            session.FindById("wnd[0]/tbar[0]/btn[3]").press
'if Contact number is correct go back check that SAP table
'is in the same position and if so delete contact
                            
                    If UCase(session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[1," & Conti & "]").Text) = UCase(WS.Range(FirstNamecol & i).Text) And _
                    UCase(session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAME1[2," & Conti & "]").Text) = UCase(WS.Range(LastNamecol & i).Text) And _
                    session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/ctxtKNVK-ABTNR[4," & Conti & "]").Text = Left(WS.Range(Departcol & i).Text, 4) Then
                                session.FindById("wnd[0]/usr/tblSAPMF02KTCTRL_ANSPRECHPARTNER/txtKNVK-NAMEV[1," & Conti & "]").SetFocus
                                session.FindById("wnd[0]/tbar[1]/btn[14]").press
                                session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                                WS.Range(RemovalLogcol & i) = "Success: Contact " & WS.Range(ConPersoncol & i).Text & " was removed automatically at " & Format(Now, "mm/dd/yyyy hh:nn")
                                WS.Range(RemovalLogcol & i).Interior.ColorIndex = 4
                                Exit Do
'if SAP table moved make processor manually delete contact
                            Else
                                MsgBox "Looks like the SAP table moved... Please manually delete contact:" & vbCrLf _
                                & WS.Range(FirstNamecol & i).Text & " " & WS.Range(LastNamecol & i).Text & " Dept: " & Left(WS.Range(Departcol & i).Text, 4) & vbCrLf & _
                                "Check Contact Person matches (" & WS.Range(ConPersoncol & i).Text & ") before deleting. Delete contact, then click Ok."
                                WS.Range(RemovalLogcol & i) = "Success: Contact " & WS.Range(ConPersoncol & i).Text & " was removed manually at " & Format(Now, "mm/dd/yyyy hh:nn")
                                WS.Range(RemovalLogcol & i).Interior.ColorIndex = 4
                                
                                Exit Do
                            End If
                            
                        Else
'if contact number wasn't the correct don't delete and drop and error log
                            WS.Range(RemovalLogcol & i) = "Error: Contact number did not match with given First Name, Last Name, and Dept." & _
                                                            "Contact number was " & session.FindById("wnd[0]/usr/txtKNVK-PARNR").Text
                            WS.Range(RemovalLogcol & i).Interior.ColorIndex = 3
                            session.FindById("wnd[0]/tbar[0]/btn[3]").press
                            Conti = Conti + 1
'if we reached end of table view page down and set contact incremintor to 0 again
                            If Conti = Tableview Then
                            Conti = 0
                            session.FindById("wnd[0]").sendVKey 82
                            End If
                        End If 'does contact number match
                    Else
                        Conti = Conti + 1
'if we reached end of table view page down and set contact incremintor to 0 again
                        If Conti = Tableview Then
                        Conti = 0
                        session.FindById("wnd[0]").sendVKey 82
                        End If
                    
                    
                    
                    End If 'do FN, LN, dept match
                Loop
                
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                
                If WS.Range(RemovalLogcol & i) = "" Then
                    WS.Range(RemovalLogcol & i) = "Error: No Contact found matching:" & _
                        WS.Range(FirstNamecol & i).Text & " " & WS.Range(LastNamecol & i).Text & _
                        " Dept: " & Left(WS.Range(Departcol & i).Text, 4) & _
                        " Contact Person: " & WS.Range(ConPersoncol & i).Text
                    WS.Range(RemovalLogcol & i).Interior.ColorIndex = 3
                    
                End If
                
'                    MsgBox ("Please manually delete Contact:" & vbCrLf _
'                            & ws.Range(FirstNamecol & i).Text & " " & ws.Range(LastNamecol & i).Text & " Dept: " & Left(ws.Range(Departcol & i).Text, 4) & vbCrLf & _
'                            "Check Contact Person matches (" & ws.Range(ConPersoncol & i).Text & ") before deleting. Please delete contact, then click Ok.")
'                    session.findbyid("wnd[0]/tbar[0]/btn[11]").press
'                    ws.Range(RemovalLogcol & i) = "Success: Contact " & ws.Range(ConPersoncol & i).Text & " was removed manually at " & Format(Now, "mm/dd/yyyy hh:nn")
'                    ws.Range(RemovalLogcol & i).Interior.ColorIndex = 4
                i = i + 1
                
            End If
        Else
            WS.Range(RemovalLogcol & i) = "No Changes Made to Contacts for Vendor " & Vendor
            i = i + 1
        End If 'Removal to do
    End If 'skip logged success rows
    
Loop
WS.Columns(RemovalLogcol).AutoFit
'End Step 2
'*****************************************************************************************************************
'Start Step 3 - Creation of new contact with MAP1

errorcnt = 0
err.Clear

If Not HaveItBreak Then
    On Error GoTo AddConLog
Else
    On Error GoTo 0
End If

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmap1"
session.FindById("wnd[0]").sendVKey 0
i = 7
Do While WS.Range(Vendorcol & i) <> ""
    Vendor = WS.Range(Vendorcol & i)
    If WS.Range(AddsLogcol & i).Interior.ColorIndex = 4 Then 'skip logged success rows
        i = i + 1
    Else
        If UCase(WS.Range(ConActcol & i)) = "ADD" Then 'Add for this vendor?
            session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
            session.FindById("wnd[0]").sendVKey 0
    
            If WS.Range(Departcol & i) <> "" And Left(WS.Range(Departcol & i).Text, 4) <> "0000" Then 'Department
                session.FindById("wnd[0]/usr/ctxtKNVK-ABTNR").Text = Left(WS.Range(Departcol & i), 4)
            End If
    
            'Last Name
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_LAST").Text = WS.Range(LastNamecol & i).Text
            'First Name
            ContactName = WS.Range(FirstNamecol & i).Text & "|" & WS.Range(LastNamecol & i).Text & "|" & WS.Range(Departcol & i).Text
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_FIRST").Text = WS.Range(FirstNamecol & i).Text
            'Function
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-FUNCTION").Text = WS.Range(Functioncol & i).Text
            'Phone number
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_NUMBER").Text = WS.Range(Telecol & i).Text
            'Phone Ext
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_EXTENS").Text = WS.Range(TeleExtcol & i).Text
            'Mobile Phone
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-MOB_NUMBER").Text = WS.Range(Mobilecol & i).Text
            'Fax Number
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_NUMBER").Text = WS.Range(Faxcol & i).Text
            'Fax Ext
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_EXTENS").Text = WS.Range(FaxExtcol & i).Text
            
            'Email
            
            'add multiple emails for new contact
            ei = 0
            ci = 0
            session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/btnG_ICON_SMTP").press
            'loop to add new email addresses until there are no more emails/vendor changes/No adds if Remove
            Do While WS.Range(FirstNamecol & i).Text & "|" & WS.Range(LastNamecol & i).Text & "|" & WS.Range(Departcol & i).Text = ContactName And WS.Range(Emailcol & i).Text <> "" And WS.Range(Vendorcol & i).Value = Vendor
                session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0," & ci & "]").Text = UCase(WS.Range(Emailcol & i).Text)
                session.FindById("wnd[1]/tbar[0]/btn[13]").press
            If ci = 2 Then 'set position of next email after add button is hit
                ci = 2
                Else
                ci = ci + 1
            End If 'set position of next email after add button is hit
                ei = ei + 1
                i = i + 1
                
            Loop 'add new ACH email until vendor changes
                'reset i
                i = i - ei
            
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
            
    
            If Left(WS.Range(Departcol & i).Text, 4) = "0007" Then 'Is this a returns vendor?
                session.FindById("wnd[0]/tbar[1]/btn[18]").press
                'enter country code first cause it is a required field and SAP needs those filled out before an other button is clicked
                session.FindById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").Text = WS.Range(RCCountrycol & i)
                'open up all data fields in SAP view. We need this for CO and Street 2
                session.FindById("wnd[1]/tbar[0]/btn[6]").press
                session.FindById("wnd[1]/usr/txtADDR1_DATA-STREET").Text = WS.Range(RCStreetcol & i)
                session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL1").Text = WS.Range(RCStreet2col & i).Text
                session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL2").Text = WS.Range(RCStreet3col & i).Text
                session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL3").Text = WS.Range(RCStreet4col & i).Text
                If InStr(UCase(WS.Range(RCCOcol & i).Text), "C/O") > 0 Then
                    session.FindById("wnd[1]/usr/txtADDR1_DATA-NAME_CO").Text = WS.Range(RCCOcol & i).Text
                Else
                    session.FindById("wnd[1]/usr/txtADDR1_DATA-NAME_CO").Text = "C/O " & WS.Range(RCCOcol & i).Text
                End If
                session.FindById("wnd[1]/usr/txtADDR1_DATA-HOUSE_NUM1").Text = WS.Range(RCHousecol & i)
                session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").Text = WS.Range(RCPostalcol & i)
                session.FindById("wnd[1]/usr/txtADDR1_DATA-CITY1").Text = WS.Range(RCCitycol & i)
                session.FindById("wnd[1]/usr/ctxtADDR1_DATA-REGION").Text = WS.Range(RCRegioncol & i)
                'session.findById("wnd[1]/usr/txtADDR1_DATA-POST_CODE2").Text = Ws.Range(RCStreetcol & i)
                session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_NUMBER").Text = WS.Range(RCTelecol & i)
                session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_EXTENS").Text = WS.Range(RCTeleExtcol & i)
                session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_NUMBER").Text = WS.Range(RCFaxcol & i)
                session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_EXTENS").Text = WS.Range(RCFaxExtcol & i)
                session.FindById("wnd[1]/usr/btnG_ICON_SMTP").press
                session.FindById("wnd[2]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,0]").Text = WS.Range(RCEmailcol & i)
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/tbar[0]/btn[0]").press
            End If
    
            session.FindById("wnd[0]/tbar[0]/btn[11]").press
            
            
            'Log Some stuff
            If session.ActiveWindow.Name = "wnd[0]" And session.FindById("wnd[0]/sbar").MessageType <> "E" Then
                If ei <> 0 Then ' if emails were worked on we'll have to figure out how many rows go worked on
                    WS.Range(AddsLogcol & i & ":" & AddsLogcol & i + ei - 1) = "Success: Contact added to Vendor " & Vendor & " at " & Format(Now, "mm/dd/yyyy hh:nn")
                    WS.Range(AddsLogcol & i & ":" & AddsLogcol & i + ei - 1).Interior.ColorIndex = 4
                    i = i + ei
                    
                Else 'emails weren't worked on
                    WS.Range(ModifyLogcol & i) = "Success: Contact " & Contact & "maintained at " & Format(Now, "mm/dd/yyyy hh:nn")
                    WS.Range(ModifyLogcol & i).Interior.ColorIndex = 4
                    i = i + 1
                End If

            Else
                err.Raise (619)
AddConLog:
                'Err.Clear
                'set veriables for log
                If session.ActiveWindow.IsPopupDialog Then
                    PopUpMessage = session.ActiveWindow.PopupDialogText
                Else
                    StatusBar = session.FindById(session.ActiveWindow.Name & "/sbar").Text
                End If
                
                'close out of pop up windows to get to main window

                Do Until session.ActiveWindow.Name = "wnd[0]"
                    session.FindById(session.ActiveWindow.Name).Close
                    'if we errored in Business address popup we sometimes get a pop up confirming exiting
                    If session.FindById(session.ActiveWindow.Name).Text = "Cancel Address Editing" Then
                        session.FindById(session.ActiveWindow.Name & "/usr/btnSPOP-OPTION1").press
                    ElseIf session.FindById(session.ActiveWindow.Name).Text = "Error" Then
                        session.FindById(session.ActiveWindow.Name).Close
                    End If
                Loop
                
                'now we are at contact page and we need to exit out of main window without saving vendor
                session.FindById("wnd[0]/tbar[0]/btn[12]").press
                'sometimes we get popup confirming exit. if we get that click ok to exit without saving
                If session.FindById(session.ActiveWindow.Name).Text = "Cancel vendor" Then
                session.FindById(session.ActiveWindow.Name & "/usr/btnSPOP-OPTION1").press
                End If
                
                'log error in correct rows of the vendor
                If PopUpMessage & "" & StatusBar = "" Then
                    WS.Range(AddsLogcol & i) = "SAP SCRIPTING ERROR: VBA Error: " & err.Description & " Something went wrong and I couldn't figure out what. SAP VBA Scripting is hard..."
                    WS.Range(AddsLogcol & i).Interior.ColorIndex = 3
                    errorcnt = errorcnt + 1
                Else
                    WS.Range(AddsLogcol & i) = "SAP SCRIPTING ERROR: " & PopUpMessage & " " & StatusBar
                    WS.Range(AddsLogcol & i).Interior.ColorIndex = 3
                    errorcnt = errorcnt + 1
                End If
                
                'reset some log variables, get out of error handler, and go to next row
                i = i + 1
                
                PopUpMessage = ""
                StatusBar = ""
                Resume EndAddConLog
EndAddConLog:
            End If
        
        Else 'No Add Contact for this vendor go to next row
            WS.Range(AddsLogcol & i) = "No Changes Made to Contacts for Vendor " & Vendor
            i = i + 1
        End If 'End Adds for this vendor?
    End If 'skip logged success rows
Loop 'Vendor Column not blank

If errorcnt = 0 Then
'Ws.Range("A2").ClearContents
Else
MsgBox errorcnt & " lines had errors for Add Contact requests. Check column " & AddsLogcol & " for logs to see what errors where."
End If
'make log column fit
WS.Columns(AddsLogcol).AutoFit

'End Step 3
'*******************************************************************************************************
'Start Step 4 - Modify contact with MAP2

err.Clear
errorcnt = 0

If Not HaveItBreak Then
    On Error GoTo ModifyErrorLog
Else
    On Error GoTo 0
End If


session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmap2"
session.FindById("wnd[0]").sendVKey 0
i = 7
Do While WS.Range(Vendorcol & i) <> ""
    
    Vendor = WS.Range(Vendorcol & i)
    If WS.Range(ModifyLogcol & i).Interior.ColorIndex = 4 Then 'skip logged success row
        i = i + 1
    Else
        If UCase(WS.Range(ConActcol & i)) = "MODIFY" Then 'Modify on this row
            Contact = WS.Range(ConPersoncol & i).Value
            session.FindById("wnd[0]/usr/ctxt*KNVK-PARNR").Text = Contact
            session.FindById("wnd[0]").sendVKey 0
            
            'There should be no reason to modify the department. A new contact should be created
'            If WS.Range(Departcol & i) <> "" And Left(WS.Range(Departcol & i).Text, 4) <> "0000" Then 'Department
'                If UCase(WS.Range(Departcol & i)) = "REMOVE" Then
'                    session.FindById("wnd[0]/usr/ctxtKNVK-ABTNR").Text = ""
'                Else
'                    session.FindById("wnd[0]/usr/ctxtKNVK-ABTNR").Text = Left(WS.Range(Departcol & i), 4)
'                End If
'            End If
            
            If WS.Range(LastNamecol & i) <> "" Then 'Last Name
                If UCase(WS.Range(LastNamecol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_LAST").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_LAST").Text = WS.Range(LastNamecol & i).Text
                End If
            End If
            
            If WS.Range(FirstNamecol & i) <> "" Then 'First Name
                If UCase(WS.Range(FirstNamecol & i)) = "REMOVE" Then
                   session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_FIRST").Text = ""
                Else
                   session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-NAME_FIRST").Text = WS.Range(FirstNamecol & i).Text
                End If
            End If
            
            If WS.Range(Functioncol & i) <> "" Then 'Function
                If UCase(WS.Range(Functioncol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-FUNCTION").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtADDR3_DATA-FUNCTION").Text = Left(WS.Range(Functioncol & i).Text, 40)
                End If
            End If
            
            If WS.Range(Telecol & i) <> "" Then 'Phone number
                If UCase(WS.Range(Telecol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_NUMBER").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_NUMBER").Text = WS.Range(Telecol & i).Text
                End If
            End If
            
            If WS.Range(TeleExtcol & i) <> "" Then 'Phone Ext
                If UCase(WS.Range(TeleExtcol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_EXTENS").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-TEL_EXTENS").Text = WS.Range(TeleExtcol & i).Text
                End If
            End If
            
            If WS.Range(Mobilecol & i) <> "" Then 'Mobile Phone
                If UCase(WS.Range(Mobilecol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-MOB_NUMBER").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-MOB_NUMBER").Text = WS.Range(Mobilecol & i).Text
                End If
            End If
            
            If WS.Range(Faxcol & i) <> "" Then 'Fax Number
                If UCase(WS.Range(Faxcol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_NUMBER").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_NUMBER").Text = WS.Range(Faxcol & i).Text
                End If
            End If
            
            If WS.Range(FaxExtcol & i) <> "" Then 'Fax Ext
                If UCase(WS.Range(FaxExtcol & i)) = "REMOVE" Then
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_EXTENS").Text = ""
                Else
                    session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/txtSZA5_D0700-FAX_EXTENS").Text = WS.Range(FaxExtcol & i).Text
                End If
            End If
                      
            If Left(WS.Range(Departcol & i).Text, 4) = "0007" Then 'Is this a returns vendor?
                'go to business address screen
                session.FindById("wnd[0]/tbar[1]/btn[18]").press
                'sometimes a warning in status bar needs a second okay if there are multiple contacts with the same name
                If session.FindById("wnd[0]/sbar").MessageType = "W" Then
                    session.FindById("wnd[0]").sendVKey 0
                End If
                
                'Country Code is a required field and must be filled out before we do the next step
                If WS.Range(RCCountrycol & i) <> "" Then 'Returns contact Country
                    session.FindById("wnd[1]/usr/ctxtADDR1_DATA-COUNTRY").Text = WS.Range(RCCountrycol & i).Text
                End If
                
               'open up all data fields in SAP view. We need this for CO and Street 2
                session.FindById("wnd[1]/tbar[0]/btn[6]").press
                                   
                If WS.Range(RCStreetcol & i) <> "" Then 'Returns contact Street
                    If WS.Range(RCStreetcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STREET").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STREET").Text = WS.Range(RCStreetcol & i).Text
                    End If
                End If
                            
                If WS.Range(RCStreet2col & i) <> "" Then 'Returns contact Street 2
                    If WS.Range(RCStreet2col & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL1").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL1").Text = WS.Range(RCStreet2col & i).Text
                    End If
                End If
                            
                If WS.Range(RCStreet3col & i) <> "" Then 'Returns contact Street 3
                    If WS.Range(RCStreet3col & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL2").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL2").Text = WS.Range(RCStreet3col & i).Text
                    End If
                End If

                If WS.Range(RCStreet4col & i) <> "" Then 'Returns contact Street 4
                    If WS.Range(RCStreet4col & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL3").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-STR_SUPPL3").Text = WS.Range(RCStreet4col & i).Text
                    End If
                End If
                
                If WS.Range(RCCOcol & i) <> "" Then 'Returns contact C/O
                    If WS.Range(RCCOcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-NAME_CO").Text = ""
                    Else
                        If InStr(UCase(WS.Range(RCCOcol & i).Text), "C/O") > 0 Then
                           session.FindById("wnd[1]/usr/txtADDR1_DATA-NAME_CO").Text = WS.Range(RCCOcol & i).Text
                        Else
                          session.FindById("wnd[1]/usr/txtADDR1_DATA-NAME_CO").Text = "C/O " & WS.Range(RCCOcol & i).Text
                        End If
                    End If
                End If
                
                If WS.Range(RCHousecol & i) <> "" Then 'Returns contact House Number
                    If WS.Range(RCHousecol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-HOUSE_NUM1").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-HOUSE_NUM1").Text = WS.Range(RCHousecol & i).Text
                    End If
                End If
                
                If WS.Range(RCPostalcol & i) <> "" Then 'Returns contact Postal Code
                    If WS.Range(RCPostalcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").Text = WS.Range(RCPostalcol & i).Text
                    End If
                End If
                
                If WS.Range(RCCitycol & i) <> "" Then 'Returns contact City
                    If WS.Range(RCCitycol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-CITY1").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtADDR1_DATA-CITY1").Text = WS.Range(RCCitycol & i).Text
                    End If
                End If
                                
                If WS.Range(RCRegioncol & i) <> "" Then 'Returns contact Region
                    session.FindById("wnd[1]/usr/ctxtADDR1_DATA-REGION").Text = WS.Range(RCRegioncol & i).Text
                    session.FindById("wnd[1]/usr/ctxtADDR1_DATA-TIME_ZONE").Text = ""
                End If
                
                If WS.Range(RCTelecol & i) <> "" Then 'Returns Contact Tele
                    If WS.Range(RCTelecol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_NUMBER").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_NUMBER").Text = WS.Range(RCTelecol & i).Text
                    End If
                End If
                
                If WS.Range(RCTeleExtcol & i) <> "" Then 'Returns Contact Tele Ext
                    If WS.Range(RCTeleExtcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_EXTENS").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-TEL_EXTENS").Text = WS.Range(RCTeleExtcol & i).Text
                    End If
                End If
                
                If WS.Range(RCFaxcol & i) <> "" Then 'Returns Contact Fax
                    If WS.Range(RCFaxcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_NUMBER").Text = ""
                    Else
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_NUMBER").Text = WS.Range(RCFaxcol & i).Text
                    End If
                End If
                
                If WS.Range(RCFaxExtcol & i) <> "" Then 'Returns Contact Fax Ext
                    If WS.Range(RCFaxExtcol & i) = "REMOVE" Then
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_EXTENS").Text = WS.Range(RCFaxExtcol & i).Text
                    Else
                        session.FindById("wnd[1]/usr/txtSZA1_D0100-FAX_EXTENS").Text = WS.Range(RCFaxExtcol & i).Text
                    End If
                End If
                
                ei = 0
                ci = 0
                If WS.Range(RCEmailcol & i) <> "" Then 'RC email isn't blank
                    
                    Do While WS.Range(ConPersoncol & i).Value = Contact 'is contact number the same as last
                    
                        If WS.Range(RCEmailcol & i).Text <> "" Then 'RC contact email is updating
                            session.FindById("wnd[1]/usr/btnG_ICON_SMTP").press
                            If ci = 0 Then 'have we deleted anything yet
                                'loop to delete all contacts
                                Do Until session.FindById("wnd[2]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,0]").Text = "" 'Check this
                                   session.FindById("wnd[2]/usr/tblSAPLSZA6T_CONTROL6").GetAbsoluteRow(0).Selected = True
                                   session.FindById("wnd[2]/tbar[0]/btn[14]").press 'delete line
                                Loop
                            End If 'have we deleted anything yet
                            'loop to add new email addresses until there are no more emails/vendor changes/No adds if Remove
                            Do While WS.Range(RCEmailcol & i).Text <> "" And WS.Range(ConPersoncol & i).Value = Contact
                                session.FindById("wnd[2]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0," & ci & "]").Text = UCase(WS.Range(RCEmailcol & i).Text)
                                session.FindById("wnd[2]/tbar[0]/btn[13]").press
                            If ci = 2 Then 'set position of next email after add button is hit
                                ci = 2
                                Else
                                ci = ci + 1
                            End If 'set position of next email after add button is hit
                                ei = ei + 1
                                i = i + 1
                            Loop 'add new ACH email until vendor changes
                        End If 'End RC contact email updating
                                                    
                    Loop 'is Contact number the same as last
                    'reset i for contact email loop
                    i = i - ei
                    
            '        session.FindById("wnd[1]/usr/btnG_ICON_SMTP").press
            '        session.FindById("wnd[2]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,0]").Text = Ws.Range(RCEmailcol & i).Text
            '        session.FindById("wnd[2]/tbar[0]/btn[0]").press
                End If 'RCemail isn't blank end
                
                    If ci <> 0 Then
                        session.FindById("wnd[2]/tbar[0]/btn[0]").press ' ok code back to business address page cause email work was done
                    End If
                    
                    session.FindById("wnd[1]/tbar[0]/btn[0]").press 'ok code to save business address page work and go back to contact
                    
            End If 'Dept is 0007 end
            
            ci = 0
            ei1 = 0
            If WS.Range(Emailcol & i).Text <> "" Then 'Contact email is updating
                session.FindById("wnd[0]/usr/subADDRESS:SAPLSZA5:0900/btnG_ICON_SMTP").press
                If ci = 0 Then 'have we deleted anything yet
                    'loop to delete all contacts
                    Do Until session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0,0]").Text = "" 'Check this
                       session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6").GetAbsoluteRow(0).Selected = True
                       session.FindById("wnd[1]/tbar[0]/btn[14]").press 'delete line
                    Loop
                End If 'have we deleted anything yet
                'loop to add new email addresses until there are no more emails/vendor changes/No adds if Remove
                Do While WS.Range(Emailcol & i).Text <> "" And WS.Range(ConPersoncol & i).Value = Contact
                    session.FindById("wnd[1]/usr/tblSAPLSZA6T_CONTROL6/txtADSMTP-SMTP_ADDR[0," & ci & "]").Text = UCase(WS.Range(Emailcol & i).Text)
                    session.FindById("wnd[1]/tbar[0]/btn[13]").press
                If ci = 2 Then 'set position of next email after add button is hit
                    ci = 2
                    Else
                    ci = ci + 1
                End If 'set position of next email after add button is hit
                    ei1 = ei1 + 1
                    i = i + 1
                Loop 'add new ACH email until vendor changes
            End If 'Contact email updating
                            
            'reset i from email loop
            i = i - ei1
            
            If ci <> 0 Then
            session.FindById("wnd[1]/tbar[0]/btn[0]").press ' ok code back to address page cause email work was done
            End If
            session.FindById("wnd[0]/tbar[0]/btn[11]").press ' Save code to exit MAP2 session and work on next lines
            If session.FindById("wnd[0]/sbar").Text Like "*already exists*" Then
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
            End If
            
        
            'Log some stuff
            If session.ActiveWindow.Name = "wnd[0]" And session.FindById("wnd[0]/sbar").MessageType <> "E" Then
                If ei <> 0 Or ei1 <> 0 Then ' if emails were worked on we'll have to figure out how many rows go worked on
                    If ei >= ei1 Then
                        WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1) = "Success: Contact " & Contact & " modified at " & Format(Now, "mm/dd/yyyy hh:nn")
                        WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1).Interior.ColorIndex = 4
                        i = i + ei
                    Else
                        WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1) = "Success: Contact " & Contact & " modified at " & Format(Now, "mm/dd/yyyy hh:nn")
                        WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1).Interior.ColorIndex = 4
                        i = i + ei1
                    End If
                Else 'emails weren't worked on
                    WS.Range(ModifyLogcol & i) = "Success: Contact " & Contact & " maintained at " & Format(Now, "mm/dd/yyyy hh:nn")
                    WS.Range(ModifyLogcol & i).Interior.ColorIndex = 4
                    i = i + 1
                End If
            Else
            err.Raise (619)
ModifyErrorLog:
                'Err.Clear
                'set veriables for log
                If session.ActiveWindow.IsPopupDialog Then
                    PopUpMessage = session.ActiveWindow.PopupDialogText
                Else
                    StatusBar = session.FindById(session.ActiveWindow.Name & "/sbar").Text
                End If
                
                'close out of pop up windows to get to main window
                Do Until session.ActiveWindow.Name = "wnd[0]"
                    session.FindById(session.ActiveWindow.Name).Close
                    'if we errored in Business address popup we sometimes get a pop up confirming exiting
                    If session.FindById(session.ActiveWindow.Name).Text = "Cancel Address Editing" Then
                        session.FindById(session.ActiveWindow.Name & "/usr/btnSPOP-OPTION1").press
                    ElseIf session.FindById(session.ActiveWindow.Name).Text = "Error" Then
                        session.FindById(session.ActiveWindow.Name).Close
                    End If
                Loop
                
                'now we are at contact page and we need to exit out of main window without saving vendor
                session.FindById("wnd[0]/tbar[0]/btn[12]").press
                'sometimes we get popup confirming exit. if we get that click ok to exit without saving
                If session.FindById(session.ActiveWindow.Name).Text = "Cancel vendor" Then
                session.FindById(session.ActiveWindow.Name & "/usr/btnSPOP-OPTION1").press
                End If
                
                
                If ei <> 0 Or ei1 <> 0 Then ' if emails were worked on we'll have to figure out how many rows go worked on
                'log error in correct rows of the vendor
                    If ei >= ei1 Then
                        If PopUpMessage & StatusBar = "" Then
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1) = "SAP SCRIPTING ERROR: A Data field was too long. Not sure how to tell you which one yet. SAP VBA Scripting is hard..."
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1).Interior.ColorIndex = 3
                            errorcnt = errorcnt + ei
                            i = i + ei
                        Else
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1) = "SAP SCRIPTING ERROR: " & PopUpMessage & StatusBar
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei - 1).Interior.ColorIndex = 3
                            errorcnt = errorcnt + ei
                            i = i + ei
                        End If
                        
                    Else
                    
                        If PopUpMessage & StatusBar = "" Then
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1) = "SAP SCRIPTING ERROR: Unclear what error was. You're going to have to do some investigating. SAP VBA Scripting is hard..."
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1).Interior.ColorIndex = 3
                            errorcnt = errorcnt + ei1
                            i = i + ei1
                        Else
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1) = "SAP SCRIPTING ERROR: " & PopUpMessage & StatusBar
                            WS.Range(ModifyLogcol & i & ":" & ModifyLogcol & i + ei1 - 1).Interior.ColorIndex = 3
                            errorcnt = errorcnt + ei1
                            i = i + ei1
                        End If
                        
                    End If
                    
                Else 'no emails were worked on
                    If PopUpMessage & StatusBar = "" Then
                        WS.Range(ModifyLogcol & i) = "SAP SCRIPTING ERROR: VBA Error " & err.Description & "Unclear what error was. You're going to have to do some investigating. SAP VBA Scripting is hard..."
                        WS.Range(ModifyLogcol & i).Interior.ColorIndex = 3
                        errorcnt = errorcnt + 1
                        i = i + 1
                    Else
                        WS.Range(ModifyLogcol & i) = "SAP SCRIPTING ERROR: " & PopUpMessage & StatusBar
                        WS.Range(ModifyLogcol & i).Interior.ColorIndex = 3
                        errorcnt = errorcnt + 1
                        i = i + 1
                    End If
                End If
                
                PopUpMessage = ""
                StatusBar = ""
                Resume EndModifyErrorLog
            End If
        
            
EndModifyErrorLog:
    
        Else 'Modifying this row
            WS.Range(ModifyLogcol & i) = "No Changes Made to Contacts for Vendor " & Vendor
            i = i + 1
        End If 'Modifying this row end
    End If ' skip logged success row

    

Loop 'Vendor column not blank

If errorcnt = 0 Then
Else
MsgBox errorcnt & " lines had errors. Check column " & ModifyLogcol & " for logs to see what errors where."
End If
'make log column fit
WS.Columns(ModifyLogcol).AutoFit

'Ws.Range("A3").ClearContents
'Ws.Range("B2").ClearContents
'End Step 4
'**************************************************************************************************
'exit back to home screen and end connection with SAP
session.FindById("wnd[0]/tbar[0]/btn[12]").press
Call EndSAPCON
With WS.Range(ACHLogcol & "1:" & ModifyLogcol & "1")
    .Interior.Color = 52479
    .Font.Color = 16711680
End With
'if there were errors lets let the processor know
WS.Calculate
If Application.WorksheetFunction.Sum(WS.Range(ACHLogcol & "6:" & ModifyLogcol & "6")) > 0 Then
    MsgBox Application.WorksheetFunction.Sum(WS.Range(ACHLogcol & "6:" & ModifyLogcol & "6")) & " errors occured while processing contact request. Take a look at error logs and input manually"
    WS.Range(ACHLogcol & "1").activate
Else
    MsgBox "All contact requests were entered into SAP successfully. YAY! You Rock!"
End If
WS.Calculate
End Sub

Private Sub p_SAP_ActivatePromos_WAK16()
'KH
'this is the base code no work done
MsgBox ("not working yet. exiting sub")
Exit Sub
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwak16"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtRT_AKTNR").Text = "7965"
session.FindById("wnd[0]/usr/btn%_RT_MATNR_%_APP_%-VALU_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
session.FindById("wnd[1]/tbar[0]/btn[13]").press
session.FindById("wnd[1]/usr/btnSOFORT_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/usr/ctxtRT_AKTNR").Text = "7966"
session.FindById("wnd[0]/usr/ctxtRT_MATNR-LOW").SetFocus
session.FindById("wnd[0]/usr/ctxtRT_MATNR-LOW").CaretPosition = 10
session.FindById("wnd[0]/usr/btn%_RT_MATNR_%_APP_%-VALU_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[16]").press
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press
End Sub

Private Sub p_TestingErrors()
'KH

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "map2"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxt*KNVK-PARNR").Text = "2950"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[18]").press
Debug.Print session.FindById("wnd[0]/sbar[0]").Text
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").Text = "3291"
session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").SetFocus
session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").CaretPosition = 0
session.FindById("wnd[1]/usr/btnG_ICON_SMTP").press
session.FindById("wnd[2]/tbar[0]/btn[1]").press
session.FindById("wnd[0]/shellcont").Close
session.FindById("wnd[2]/tbar[0]/btn[0]").press
session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").Text = "93291"
session.FindById("wnd[1]/usr/txtADDR1_DATA-POST_CODE1").CaretPosition = 1
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/tbar[0]/btn[12]").press
session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
Call EndSAPCON

End Sub
Sub SAP_FinishBrand_MASS_ZSE16N_ZBD10()
'KH
Dim WB As Workbook
Dim NewBrandWS As Worksheet
Dim BrandUpdateWS As Worksheet
Dim ScratchWS As Worksheet
Dim ATPWS As Worksheet
Dim OldLR As Long
Dim NewLR As Long
Dim startRow As Long
Dim lastRow As Long
Dim Brand As Long


Set WB = Workbooks("New_Brand_Request_Form.xlsm")
Set NewBrandWS = WB.Worksheets("New Brands")
Set ScratchWS = WB.Worksheets("ScratchPad")
Set BrandUpdateWS = WB.Worksheets("Brand Updates")
Set ATPWS = WB.Worksheets("Articles to Push")

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'set some row variables
OldLR = NewBrandWS.Range("J10000").End(xlUp).row + 1
NewLR = OldLR
startRow = NewBrandWS.Range("J10000").End(xlUp).row + 1
lastRow = NewBrandWS.Range("I10000").End(xlUp).row

'Update Generics with Mass Tcode


session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmass"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtMASSSCREEN-OBJECT").Text = "bus1001001"
session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").Text = "brand"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press

'copy generics and set brand variable until the end of generics on sheet in row G
Do Until NewLR = lastRow
Brand = NewBrandWS.Range("F" & NewLR)

    'set new last row to know what generics to copy
    Do Until NewBrandWS.Cells(NewLR, 6).Value <> Brand And NewBrandWS.Cells(NewLR, 6).Value <> ""
        NewLR = NewLR + 1
        If NewBrandWS.Cells(NewLR, 7).Value = "" Or NewLR = lastRow + 1 Then
        Exit Do
        End If
    Loop
    'copy the generic
    NewLR = NewLR - 1
    NewBrandWS.Range("G" & OldLR & ":G" & NewLR).Copy
    
    session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press 'multiple selection button
    session.FindById("wnd[1]/tbar[0]/btn[16]").press 'trashcan button
    session.FindById("wnd[1]/tbar[0]/btn[24]").press 'paste button
    session.FindById("wnd[1]/tbar[0]/btn[8]").press 'excecute on generic selection window
    session.FindById("wnd[0]/tbar[1]/btn[8]").press 'excecute on mass window
    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD2-VALUE-LEFT[2,0]").Text = Brand 'put brand in dropdown field
    session.FindById("wnd[0]").sendVKey 0 'hit enter
    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press 'dropdown value to all in mass
    session.FindById("wnd[0]/tbar[0]/btn[11]").press 'save
    If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
    'great mass ran successfully
    Else
        NewBrandWS.Range("M" & OldLR) = "Run MASS again"
        NewBrandWS.Range("M" & OldLR).Interior.ColorIndex = 3
        MsgBox session.FindById("wnd[0]/usr/txtNR_E").Text & " Articles Errored on the Mass job for " & Brand & vbCrLf & vbCrLf & _
            "Error message dropped in Column M of the brands that need to be run again."
    End If
    session.FindById("wnd[0]/tbar[0]/btn[3]").press 'back button to mass
    session.FindById("wnd[0]/tbar[0]/btn[3]").press 'back button to brand mass page so we can do it again
    'reset NewLR variable after we copy pasted
    NewLR = NewLR + 1
    'exit loops if we're at bottom of articles
    If NewBrandWS.Cells(NewLR, 7).Value = "" Then
        Exit Do
    Else
    'reset OldLR to do next group of generics
        OldLR = NewLR
    End If

Loop
Application.CutCopyMode = False


'Get all variants for ZBD10 push
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "mara"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[18]").press
session.FindById("wnd[0]").sendVKey 71
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "matnr"
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").CaretPosition = 5
session.FindById("wnd[1]").sendVKey 0
session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").SetFocus
session.FindById("wnd[0]").sendVKey 71
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").Text = "satnr"
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").CaretPosition = 5
session.FindById("wnd[1]").sendVKey 0
session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/chkGS_SELFIELDS-MARK[5,0]").Selected = True
session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,0]").SetFocus
session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,0]").press
NewBrandWS.Range("G" & startRow & ":G" & lastRow).Copy
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
session.FindById("wnd[1]/tbar[0]/btn[0]").press

ATPWS.activate
ATPWS.Cells.ClearContents
ATPWS.Range("A1").PasteSpecial
ATPWS.Columns("A:A").Select
Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
    TrailingMinusNumbers:=True
ATPWS.Range("1:5").Delete
ATPWS.Range("B1:B" & ATPWS.Range("B1").End(xlDown).row).Copy


session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzbd10"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
ATPWS.Range("B1:B" & ATPWS.Range("B1").End(xlDown).row).Copy
session.FindById("wnd[1]/tbar[0]/btn[16]").press
session.FindById("wnd[1]/tbar[0]/btn[24]").press
ATPWS.Range("C1:C" & ATPWS.Range("C1").End(xlDown).row).Copy
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
session.FindById("wnd[1]/tbar[0]/btn[13]").press
session.FindById("wnd[1]/usr/btnSOFORT_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press

Call EndSAPCON

End Sub

Function columnletter(ColumnNumber As Long) As String
'This converts the column number to a the column string so Kevin doesn't have to rewrite the whole script to replace range with cell...


'Convert To Column Letter
  columnletter = Split(Cells(1, ColumnNumber).Address, "$")(1)
  
 
End Function
'**************************************************************************************************
'**************************************************************************************************
' Temp Listings Form
'**************************************************************************************************
'**************************************************************************************************
Sub SAP_Temp_Listing_WSM3_WSE6()
'******************************************************************************
' MH orig "wrote" in 10/2020 to process temp listings
' Assumes you are looking at a "Temp Listings" form and the auto-open macro
'   has dropped some data out on the "Sorted" sheet
' This is intended to account for the "bulk" case of new temp listings via
'   WSM3, but should also account for Ending VIF listings to site 2002 via
'   WSE6.
' As of 10/2020 utilizing Temp Listing form for VIF stuff is intended to be a
'   short term solution
'******************************************************************************
Dim sortwS As Worksheet
Dim lastcol As Long
Dim lastRow As Long
Dim i As Long
Dim msg As String

    'check if we are on the right template, if we have "sorted" data, etc.


    On Error Resume Next
    Set sortwS = ActiveWorkbook.Worksheets("Sorted")
    On Error GoTo 0
    
    If sortwS Is Nothing Then
        MsgBox "Did not find a 'Sorted' Worksheet?  Are you looking at a temp listing form?"
        Exit Sub
    End If
    sortwS.Visible = xlSheetVisible
    sortwS.activate
    
    lastcol = sortwS.Cells(2, Columns.Count).End(xlToLeft).Column
    
    If lastcol = 1 Then
        MsgBox "Did not find any data on the 'Sorted' Worksheet.  Maybe Auto-open didn't run?  Make sure that is populated and try again."
        Set sortwS = Nothing
        Exit Sub
    End If
    
    'MsgBox "Last column is " & lastcol
    
'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If
    
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwsm3"
    session.FindById("wnd[0]").sendVKey 0
    
    For i = 1 To lastcol Step 3
        'if we are working with a future end date, WSM3.
        'if we are working with VIF with a PAST end date, WSE6
        
        If sortwS.Cells(2, i + 1).Value >= Now Then    'end date in future
            '***** WSM3 *****
            'input starting screen stuff
            session.FindById("wnd[0]/usr/chkDATBE").Selected = False
            session.FindById("wnd[0]/usr/ctxtASORT-LOW").Text = sortwS.Cells(3, i + 1).Value
            session.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = sortwS.Cells(3, i).Value
            session.FindById("wnd[0]/usr/ctxtLSTFL").Text = "02"
            session.FindById("wnd[0]/usr/ctxtDATAB").Text = sortwS.Cells(2, i).Value
            session.FindById("wnd[0]/usr/ctxtDATBI").Text = sortwS.Cells(2, i + 1).Value
            
            'copy all sites into clipboard
            'sortwS.activate
            lastRow = sortwS.Cells(Rows.Count, i + 1).End(xlUp).row
            sortwS.Range(Cells(3, i + 1), Cells(lastRow, i + 1)).Select
            Selection.Copy
            
            'input all sites
            session.FindById("wnd[0]/usr/btn%_ASORT_%_APP_%-VALU_PUSH").press
            session.FindById("wnd[1]/tbar[0]/btn[16]").press        'delete previous entry
            session.FindById("wnd[1]/tbar[0]/btn[24]").press        'paste from clipboard
            session.FindById("wnd[1]/tbar[0]/btn[8]").press         'Execute
            
            'copy all articles into clipboard
            lastRow = sortwS.Cells(Rows.Count, i).End(xlUp).row
            sortwS.Range(Cells(3, i), Cells(lastRow, i)).Select
            Selection.Copy
            
            'input all articles
            session.FindById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
            session.FindById("wnd[1]/tbar[0]/btn[16]").press        'delete previous entry
            session.FindById("wnd[1]/tbar[0]/btn[24]").press        'paste from clipboard
            session.FindById("wnd[1]/tbar[0]/btn[8]").press         'Execute
            
            'execute (detect errors?)
            session.FindById("wnd[0]/tbar[1]/btn[8]").press
            
            'back out to "starting screen"
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        ElseIf sortwS.Cells(2, i + 1).Value < Now Then    'looks like we are ENDING VIF
            '***** WSE6 *****
                        
            'leave WSM3 for WSE6
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwse6"
            session.FindById("wnd[0]").sendVKey 0
            
            'enter starting site and article
            session.FindById("wnd[0]/usr/subSUB1:RWSORT48:1110/ctxtS_ASORT-LOW").Text = sortwS.Cells(3, i + 1).Value
            session.FindById("wnd[0]/usr/subSUB1:RWSORT48:1110/ctxtS_MATNR-LOW").Text = sortwS.Cells(3, i).Value
            
            'copy full site list
            'sortwS.activate
            lastRow = sortwS.Cells(Rows.Count, i + 1).End(xlUp).row
            sortwS.Range(Cells(3, i + 1), Cells(lastRow, i + 1)).Select
            Selection.Copy
            
            'paste in all sites
            session.FindById("wnd[0]/usr/subSUB1:RWSORT48:1110/btn%_S_ASORT_%_APP_%-VALU_PUSH").press
            session.FindById("wnd[1]/tbar[0]/btn[16]").press
            session.FindById("wnd[1]/tbar[0]/btn[24]").press
            session.FindById("wnd[1]/tbar[0]/btn[8]").press
            
            
            'copy full article list
            'sortwS.activate
            lastRow = sortwS.Cells(Rows.Count, i).End(xlUp).row
            sortwS.Range(Cells(3, i), Cells(lastRow, i)).Select
            Selection.Copy
            
            'Paste in all articles
            session.FindById("wnd[0]/usr/subSUB1:RWSORT48:1110/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
            session.FindById("wnd[1]/tbar[0]/btn[16]").press
            session.FindById("wnd[1]/tbar[0]/btn[24]").press
            session.FindById("wnd[1]/tbar[0]/btn[8]").press
            
            'execute
            session.FindById("wnd[0]/tbar[1]/btn[8]").press
            'select all
            session.FindById("wnd[0]/usr/cntlCONTAINER_0100/shellcont/shell").SelectAll
            'delete
            session.FindById("wnd[0]/usr/cntlCONTAINER_0100/shellcont/shell").PressToolbarButton "DELE"
            'back to wsm3
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwsm3"
            session.FindById("wnd[0]").sendVKey 0
        End If 'Checking for end date in current or future
        sortwS.Cells(1, i + 2).Value = "'<- Sent to SAP:"
        sortwS.Cells(2, i + 2).Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
        Application.Wait (Now + TimeValue("00:00:05"))
    Next i
    
    'back out of WSM3
    session.FindById("wnd[0]/tbar[0]/btn[15]").press
    'disconnect from SAP
    EndSAPCON
    
    Set sortwS = Nothing
End Sub
Sub SAP_PSI_MASS_CHARVAL()
'******************************************************************************
' MH orig "wrote" in 10/2020 to process VIF PSI updates
' Assumes you are looking at a "Temp Listings" form and the auto-open macro
'   has dropped some data out on the VIF PSI sheet
'******************************************************************************
Dim vifWS As Worksheet
Dim i As Long
Dim lastRow As Long
Dim startSelect As Long
Dim endSelect As Long
Dim chunkSize As Long   'number of articles we can process at a time
    
    
    On Error Resume Next
    Set vifWS = ActiveWorkbook.Worksheets("VIF PSI")
    On Error GoTo 0
    
    If vifWS Is Nothing Then
        MsgBox "Looks like there is no VIF PSI work here." & vbLf _
            & "If you believe this to be in error, validate " & _
            "'Create_Vif_Output' has run in your temp listings form."
        Exit Sub
    End If
    vifWS.Visible = xlSheetVisible
    vifWS.activate
    
    
    'Putting in too many articles will cause an SAP dump.
    '5000 - Too Many
    '4000 - too many
    '3000 - worked
    'chunkSize set to 2500
    chunkSize = 2500
    
    'connect to SAP - could be chained/combo'd with the above one
    SAPCON
    If session Is Nothing Then
        MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
        Exit Sub
    End If
    session.FindById("wnd[0]").Maximize
    For i = 1 To 3
        lastRow = vifWS.Cells(Rows.Count, i).End(xlUp).row
        If lastRow > 1 Then
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmass_charval"
            session.FindById("wnd[0]").sendVKey 0
            
            'load up our processing defaults
            session.FindById("wnd[0]/usr/radGV_RBBTH").Select
            session.FindById("wnd[0]/usr/chkGV_CLSFD").Selected = True
            session.FindById("wnd[0]/usr/ctxtGV_ATNAM").Text = "PRODUCTSOURCEINDICATOR"
            session.FindById("wnd[0]/usr/btnPB_EXP").press
            session.FindById("wnd[0]/usr/radGV_PRLL").Select                                'parallel
            session.FindById("wnd[0]/usr/radGV_ASYN").Select                                'Asynchronous
            session.FindById("wnd[0]/usr/txtGV_PRCNO").Text = "20"                          'num processes
            session.FindById("wnd[0]/usr/txtGV_RECNO").Text = "20"                          'records per process
            session.FindById("wnd[0]/usr/ctxtGV_UTLNO").Text = "20"                         'utilization percent
            session.FindById("wnd[0]/usr/ctxtGV_SVGRP").Text = "parallel_generators"        'server group
            
            'set our new value - same as column number on output sheet
            session.FindById("wnd[0]/usr/ctxtGV_NWVAL").Text = i
            

            
            'indicate the row we start our selection on
            startSelect = 2
            'default our endselect row
            endSelect = 0
            Do Until endSelect >= lastRow
                'grab 2500 article chunks until we get to the last row
                endSelect = WorksheetFunction.Min(startSelect + chunkSize, lastRow)
                
                'copy articles
                vifWS.activate
                vifWS.Range(Cells(startSelect, i), Cells(endSelect, i)).Select
                Selection.Copy
                
                'paste articles into our multi-select
                session.FindById("wnd[0]/usr/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press
                session.FindById("wnd[1]/tbar[0]/btn[16]").press
                session.FindById("wnd[1]/tbar[0]/btn[24]").press
                session.FindById("wnd[1]/tbar[0]/btn[8]").press
                
                'execute
                session.FindById("wnd[0]/tbar[1]/btn[8]").press
                '"Save" button
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                
                'unhandled error - no changes made.
                
                'Back out
                session.FindById("wnd[0]/tbar[0]/btn[3]").press
                startSelect = endSelect + 1
            Loop
            'back out
            'session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If  'lastrow > 1
    Next i
    EndSAPCON
End Sub

Sub SAP_RP_DC2St_MM42_Mass_ZBD10()
Dim WB As Workbook
Dim WS As Worksheet
Dim i As Integer
Dim j As Integer
Dim lastRow As Integer
Dim Article As Double
Dim Vendor As Double
Dim RP As Integer
Dim DCtoStore As Integer
Dim Validated As Integer



Set WB = ActiveWorkbook
Set WS = WB.Worksheets("Unit of Measure")
lastRow = WS.Range("A12").End(xlDown).row


Validated = MsgBox("Have you validated current CAR and PAC values in MARM table to make sure this req makes sense?", vbYesNo)

If Validated = vbNo Then
    MsgBox "Well do that first and the run again"
    Exit Sub
End If

If WS.Range("D12:D" & lastRow).HasFormula Then
Else
    WS.Range("D12").Formula = _
        "=IF(A12="""","""",IFERROR(INDEX($AA$10:$AB$1000,MATCH(A12,$AB$10:$AB$1000,0),1),""No PIR""))"
    WS.Range("D12").AutoFill Destination:=Range("D12:D1000")
End If

If WS.Range("AA10") = "" Then
    MsgBox "You need to drop PIRs/Articles in Column AA and AB before running this script"
    Exit Sub
End If

Application.Calculate

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

session.FindById("wnd[0]/tbar[0]/okcd").Text = "mm42"
session.FindById("wnd[0]").sendVKey 0

'Loop through all lines of UOM and adjust Car and PAC on Basic Data tab in MM42 as needed
For i = 12 To lastRow
    
    Article = WS.Range("A" & i).Value
    Vendor = WS.Range("C" & i).Value
    RP = WS.Range("O" & i).Value
    DCtoStore = WS.Range("Q" & i).Value
'If this line has a RP update and is a generic, do some stuff
    If RP > 0 Then
        If Article < 999999 Then
            session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Article
            session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = Vendor
            session.FindById("wnd[0]").sendVKey 0
            For j = 0 To 3
            'If RP is greater than 1 it is adding or modifying an existing Car row. Loop through the
            'UOM matrix on Basic Data tab to figure out the correct action
            If RP > 1 Then
                'If we hit a blank row then there was no Carton row before. Lets add one by adding CAR, EA, RP and log then exit for
                If session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text = "" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text = "car"
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2," & j & "]").Text = RP
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MESUB[3," & j & "]").Text = "ea"
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "Car UOM added to " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'If we find a Carton Row lets update the quantity to new RP
                ElseIf UCase(session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text) = "CAR" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2," & j & "]").Text = RP
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "Adjusted the existing Car UOM on " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'We should only ever have EA, PAC, and Car if we have more than that something is wrong...
                ElseIf j = 3 Then
                    MsgBox Article & " has more than 4 Alt units of measure and Kevin didn't code for that because he didn't think that could happen. Skipping this Generic. Add Car UOM in MM42 Manually..."
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "This Generic was skipped because it has a bunch of weird Alt UOMs..."
                    WS.Range("AC" & i).Interior.ColorIndex = 3
                End If
            Else 'RP = 1 so we want to delete the car row then exit for
                'Assuming that there is a car row if we are removing RP. Find Car row and delete it
                If UCase(session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text) = "CAR" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").SetFocus
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/btnME_DELETE").press
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "Removed existing Car UOM on " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'We didn't find Car row something when wrong...
                ElseIf j = 3 Then
                    MsgBox Article & " has more than 4 Alt units of measure and Kevin didn't code for that because he didn't think that could happen. Skipping this Generic. Add Car UOM in MM42 Manually..."
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "This Generic was skipped because it has a bunch of weird Alt UOMs..."
                    WS.Range("AC" & i).Interior.ColorIndex = 3
                End If
            End If 'RP > 1
            Next 'For J = 0 to 3 Looping through UOM matrix of Basic data tab
        End If 'Article < 999999
    End If 'RP > 0
    
'If this line has a DC to Store update and is a generic, do some stuff
    If DCtoStore > 0 Then
        If Article < 999999 Then
            session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Article
            session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = Vendor
            session.FindById("wnd[0]").sendVKey 0
            For j = 0 To 3
            'If DC to Store is greater than 1 it is adding or modifying an existing PAC row. Loop through the
            'UOM matrix on Basic Data tab to figure out the correct action
            If DCtoStore > 1 Then
                'If we hit a blank row then there was no PAC row before. Lets add one by adding CAR, EA, RP and log then exit for
                If session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text = "" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text = "PAC"
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2," & j & "]").Text = DCtoStore
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MESUB[3," & j & "]").Text = "EA"
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/radSMEINH-KZAUSME[6," & j & "]").Selected = True
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "PAC UOM added to " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'If we find a PAC Row lets update the quantity to new RP and exit for
                ElseIf UCase(session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text) = "PAC" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2," & j & "]").Text = DCtoStore
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/radSMEINH-KZAUSME[6," & j & "]").Selected = True
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "Adjusted the existing PAC UOM on " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'We should only ever have EA, PAC, and Car if we have more than that something is wrong...
                ElseIf j = 3 Then
                    MsgBox Article & " has more than 4 Alt units of measure and Kevin didn't code for that because he didn't think that could happen. Skipping this Generic. Add Car UOM in MM42 Manually..."
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "This Generic was skipped because it has a bunch of weird Alt UOMs..."
                    WS.Range("AC" & i).Interior.ColorIndex = 3
                End If
            Else 'Remove DC to Store by changing radio button back to EA
                If UCase(session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0," & j & "]").Text) = "PAC" Then
                    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/radSMEINH-KZAUSME[6,0]").Selected = True
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "Removed existing PAC UOM on " & Article & " at " & Format(Now(), "dd-mm-yyyy hh:nn:ss")
                    WS.Range("AC" & i).Interior.ColorIndex = 4
                    Exit For
                'We should only ever have EA, PAC, and Car if we have more than that something is wrong...
                ElseIf j = 3 Then
                    MsgBox Article & " has more than 4 Alt units of measure and Kevin didn't code for that because he didn't think that could happen. Skipping this Generic. Add Car UOM in MM42 Manually..."
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    WS.Range("AC" & i) = "This Generic was skipped because it has a bunch of weird Alt UOMs..."
                    WS.Range("AC" & i).Interior.ColorIndex = 3
                End If
            End If 'DCtoStore > 1
            Next 'j = 0 To 3 UOM matrix loop
        End If 'Article < 999999
    End If 'DCtoStore > 0
Next 'i = 12 To LastRow Loop through all rows and maybe do stuff

'Go back to session manager home screen
session.FindById("wnd[0]/mbar/menu[0]/menu[5]").Select

'Message box and drop in log column header
MsgBox "All done Maintaining Carton Qnt on Generics. Check Column AC for log"
WS.Range("AC11").Value = "MM42 Car qty updated for generics"
WS.Range("AC:AC").EntireColumn.AutoFit

Call SAP_Rounding_Profile_Mass

Call EndSAPCON



End Sub

Private Sub SAP_Rounding_Profile_Mass()
'Updates to Template Needed
'--Log Columns
'
'
'

Dim WB As Workbook
Dim WS As Worksheet
Dim i As Integer
Dim lastRow As Integer
Dim Article As Double
Dim Vendor As Double
Dim RP As Integer
Dim RPDic As Object
Dim key As Variant
Dim RemoveRP As Boolean
Dim AddRP As Boolean
Dim ZBD10Push As Integer
Dim ZBD10Cell As String
Dim StandAlone As Boolean
Dim MassErrors As Integer



'Set some objects and variables
Set WB = ActiveWorkbook
Set WS = WB.Worksheets("Unit of Measure")
Set RPDic = CreateObject("Scripting.Dictionary")
lastRow = WS.Range("A12").End(xlDown).row
MassErrors = 0
RemoveRP = False
AddRP = False

'Loop through all rows to get all RPs we are setting
'Used later to set Minstnqty
For i = 12 To lastRow
    RP = WS.Range("O" & i).Value
    If Not RPDic.Exists(RP) And RP > 0 Then
        RPDic.Add RP, True
        If RP = 1 Then
            RemoveRP = True
        ElseIf RP > 1 Then
            AddRP = True
        End If
        
    End If
Next

'Set filter on row 11 to be used later
WS.Rows("11:11").AutoFilter

'Drop headers for log columns Cause Kevin was too lazy to add to Template
WS.Range("AD11") = "Mass Add RP for Generics"
WS.Range("AE11") = "Mass Add RP for Variants"
WS.Range("AF11") = "Mass Remove RP for Generics"
WS.Range("AG11") = "Mass Remove RP for Variants"
WS.Range("AH11") = "Mass Minstdqty for Generics"
WS.Range("AI11") = "Mass Minstdqty for Generics"

'Standard setup for SAP script
If session Is Nothing Then
    SAPCON
    StandAlone = True
    'something went wrong setting up SAP connection Exit sub
    If session Is Nothing Then
        MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
        Exit Sub
    End If
End If

'Mass, Bus3003, RP, Execute, Press multiple selection button for PIR
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmass"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtMASSSCREEN-OBJECT").Text = "bus3003"
session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").Text = "rp"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press

'filter out No PIR rows (Field 4) and rows that don't have data in RP column (Field 15)
WS.Range("$A$11:$CI$1000").AutoFilter Field:=4, Criteria1:="<>No PIR"
WS.Range("$A$11:$CI$1000").AutoFilter Field:=15, Criteria1:="<>"

'Set D050 RP rule for new Rounding profiles for generics
If AddRP Then
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:="<999999"
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=15, Criteria1:=">1"
    lastRow = WS.Range("A1001").End(xlUp).row
    WS.Range("D12:D" & lastRow).Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
    If session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").Text <> "" Then
        session.FindById("wnd[0]/tbar[1]/btn[8]").press
        'If the number of entries is greater than blah. Click "display all records" button on popup
        If session.ActiveWindow.Name = "wnd[1]" Then
            session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End If
        'enter D050 in mass then drop it like its hot
        session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[4,0]").Text = "d050"
        session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        'Check for errors
        If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
            WS.Range("AD12:AD" & lastRow).Interior.ColorIndex = 4
            WS.Range("AD12:AD" & lastRow).Value = "Successful RP MASS RULE ADD"
        Else
            WS.Range("AD12:AD" & lastRow).Interior.ColorIndex = 3
            WS.Range("AD12:AD" & lastRow).Value = "ERROR RP MASS RULE ADD"
            MassErrors = MassErrors + 1
        End If
        'Back to RP Mass
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
    End If 'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
    'multiple selection, press delete to clear out for next set
    session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press
    session.FindById("wnd[1]/tbar[0]/btn[16]").press
    
    
    'set D050 RP rule for new Rounding Profiles for Variants
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:=">999999"
    lastRow = WS.Range("A1001").End(xlUp).row
    'if there are add RP for Variants copy PIRs, Paste in Multiple selection, no PIRs exist for variants skip
    If WS.Range("A1001").End(xlUp).row > 11 Then
        WS.Range("D12:D" & lastRow).Copy
        session.FindById("wnd[1]/tbar[0]/btn[24]").press
        session.FindById("wnd[1]/tbar[0]/btn[8]").press
        'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
        If session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").Text <> "" Then
            session.FindById("wnd[0]/tbar[1]/btn[8]").press
            'If the number of entries is greater than blah. Click "display all records" button on popup
            If session.ActiveWindow.Name = "wnd[1]" Then
                session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            End If
            'D050 in Mass and drop it like its hot
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[4,0]").Text = "d050"
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
            session.FindById("wnd[0]/tbar[0]/btn[11]").press
            'Check for errors and log accordingly
            If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
                WS.Range("AE12:AE" & lastRow).Interior.ColorIndex = 4
                WS.Range("AE12:AE" & lastRow).Value = "Successful RP MASS RULE ADD"
            Else
                WS.Range("AE12:AE" & lastRow).Interior.ColorIndex = 3
                WS.Range("AE12:AE" & lastRow).Value = "ERROR RP MASS RULE ADD"
                MassErrors = MassErrors + 1
            End If
            'back to RP mass
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If 'Don't run Mass if no PIRs are in low input
    End If 'Are there PIRs for Variants? Don't want to update all 2 Million PIRs...
End If 'are there add RP



'remove D050 RP rule for new Rounding profiles for Generics
If RemoveRP Then
    'Set filters for Generic and remove RP (<=1 in RP column), find last row, copy PIRs, press multiple selection, Delete, paste, execute
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:="<999999"
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=15, Criteria1:="<=1"
    lastRow = WS.Range("A1001").End(xlUp).row
    WS.Range("D12:D" & lastRow).Copy
    session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press
    session.FindById("wnd[1]/tbar[0]/btn[16]").press
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
    If session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").Text <> "" Then
        session.FindById("wnd[0]/tbar[1]/btn[8]").press
        'If the number of entries is greater than blah. Click "display all records" button on popup
        If session.ActiveWindow.Name = "wnd[1]" Then
            session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        End If
        'remove RP and drop it like its hot
        session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[4,0]").Text = ""
        session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        'Check for mass errors and log
        If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
            WS.Range("AF12:AF" & lastRow).Interior.ColorIndex = 4
            WS.Range("AF12:AF" & lastRow).Value = "Successful RP MASS RULE REMOVAL"
        Else
            WS.Range("AF12:AF" & lastRow).Interior.ColorIndex = 3
            WS.Range("AF12:AF" & lastRow).Value = "Error RP MASS RULE REMOVAL"
            MassErrors = MassErrors + 1
        End If
        'back to RP mass and delete out multiple selection to get ready for variants remove RP
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
    End If
    session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press
    session.FindById("wnd[1]/tbar[0]/btn[16]").press
    
    
    
    'remove D050 RP rule for new Rounding profiles for Variants
    WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:=">999999"
    lastRow = WS.Range("A1001").End(xlUp).row
    'in case no PIRs exist for variants, copy filtered data, paste, execute
    If WS.Range("A1001").End(xlUp).row > 11 Then
        WS.Range("D12:D" & lastRow).Copy
        session.FindById("wnd[1]/tbar[0]/btn[24]").press
        session.FindById("wnd[1]/tbar[0]/btn[8]").press
        'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
        If session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").Text <> "" Then
            session.FindById("wnd[0]/tbar[1]/btn[8]").press
            'If the number of entries is greater than blah. Click "display all records" button on popup
            If session.ActiveWindow.Name = "wnd[1]" Then
                session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
            End If
            'Remove D050 and drop it like its hot
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/ctxtHEADER_STRUC-FIELD5-VALUE-LEFT[4,0]").Text = "d050"
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
            session.FindById("wnd[0]/tbar[0]/btn[11]").press
            'Check for Mass errors
            If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
                WS.Range("AG12:AG" & lastRow).Interior.ColorIndex = 4
                WS.Range("AG12:AG" & lastRow).Value = "Successful RP MASS RULE REMOVAL"
            Else
                WS.Range("AG12:AG" & lastRow).Interior.ColorIndex = 3
                WS.Range("AG12:AG" & lastRow).Value = "ERROR RP MASS RULE REMOVAL"
                MassErrors = MassErrors + 1
            End If
            'Back to mass RP
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        End If 'Fail safe for to not mass 2 million PIRs if blank
    End If 'second fail safe if no PIRs for variants
End If 'Remove RP
'go back to MASS home page so we can do Minstdqty
session.FindById("wnd[0]/tbar[0]/btn[3]").press


'Set Minstdqty by filtering
session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").Text = "minstdqty"
session.FindById("wnd[0]/tbar[1]/btn[8]").press
'Add Minstdqty for Gen (i=1) and Vars (i=2)
For i = 1 To 2
'sort Gen or Var
    If i = 1 Then
        WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:="<999999"
    ElseIf i = 2 Then
        WS.Range("$A$11:$CI$1000").AutoFilter Field:=1, Criteria1:=">999999"
        
    End If
'sort by each RP on sheet and run mass
'in case no PIRs exist for variants
    If WS.Range("A1001").End(xlUp).row > 11 Then
        For Each key In RPDic.keys
            WS.Range("$A$11:$CI$1000").AutoFilter Field:=15, Criteria1:=key
            lastRow = WS.Range("A1001").End(xlUp).row
            session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press
            session.FindById("wnd[1]/tbar[0]/btn[16]").press
            WS.Range("D12:D" & lastRow).Copy
            session.FindById("wnd[1]/tbar[0]/btn[24]").press
            session.FindById("wnd[1]/tbar[0]/btn[8]").press
            'Fail safe to make sure we don't run mass on all PIRs if no PIR is entered
            If session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/ctxtMASSFREESEL-LOW[0,24]").Text <> "" Then
                session.FindById("wnd[0]/tbar[1]/btn[8]").press
                'If the number of entries is greater than blah. Click "display all records" button on popup
                    If session.ActiveWindow.Name = "wnd[1]" Then
                        session.FindById("wnd[1]/usr/btnSPOP-VAROPTION1").press
                    End If
                'remove or add RP?
                If key > 1 Then
                    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD5-VALUE-RIGHT[4,0]").Text = key
                    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD6-VALUE-RIGHT[5,0]").Text = key
                Else
                    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD5-VALUE-RIGHT[4,0]").Text = ""
                    session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD6-VALUE-RIGHT[5,0]").Text = 1
                End If
                session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                If i = 1 Then
                    If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
                        'ws.Cells(8, 33 + i).Value = ws.Cells(8, 33 + i).Value + 1
                        WS.Range("AH12:AH" & lastRow).Interior.ColorIndex = 4
                        WS.Range("AH12:AH" & lastRow).Value = "SUCCESSFUL MINSTDQTY SET"
                    Else
                        WS.Range("AH12:AH" & lastRow).Interior.ColorIndex = 3
                        WS.Range("AH12:AH" & lastRow).Value = "ERROR MINSTDQTY SET"
                        MassErrors = MassErrors + 1
                    End If
                ElseIf i = 2 Then
                    If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
                        WS.Range("AI12:AI" & lastRow).Interior.ColorIndex = 4
                        WS.Range("AI12:AI" & lastRow).Value = "SUCCESSFUL MINSTDQTY SET"
                    Else
                        WS.Range("AI12:AI" & lastRow).Interior.ColorIndex = 3
                        WS.Range("AI12:AI" & lastRow).Value = "ERROR MINSTDQTY SET"
                        MassErrors = MassErrors + 1
                    End If
                
                End If
                session.FindById("wnd[0]/tbar[0]/btn[3]").press
                session.FindById("wnd[0]/tbar[0]/btn[3]").press
            End If
        Next
    End If
Next

'remove filters
Selection.AutoFilter
WS.Rows("11:11").AutoFilter

'Go back to session manager home screen
session.FindById("wnd[0]/mbar/menu[0]/menu[3]").Select

'Add msgbox for success and failures
If MassErrors > 0 Then
    MsgBox "All done. All Mass runs were successful! You rock!"
Else
    MsgBox "All done. There were " & MassErrors & " MASS runs that had erros." & vbCrLf & vbCrLf & _
    "Check Columns AD:AI for errors and do those grouping manually."
End If

WS.Range("AD:AE").EntireColumn.AutoFit
WS.Range("AF:AF").EntireColumn.ColumnWidth = 30
WS.Range("AG:AI").EntireColumn.AutoFit
ZBD10Push = MsgBox("Would you like this macro to do ZBD10 push on all articles in Column A?", vbYesNo)
If ZBD10Push = vbYes Then
    ZBD10Cell = WS.Range("A12").Address
    SAP_Article_Push_ZBD10_Start_Cell (ZBD10Cell)
Else
    MsgBox "Make sure you push these articles before you close the waypoint task..."
End If

If StandAlone = True Then Call EndSAPCON

End Sub
Sub SAP_Article_Push_ZBD10_Start_Cell(Optional StartCell As String)


Dim StandAlone As Boolean
Dim Article As Range
Dim lastRow As Integer


If StartCell = "" Then
    On Error Resume Next
    Set Article = Application.InputBox(Prompt:="Please select First Article in a column of articles " & _
        "you would like to push with ZBD10, then hit OK.", _
        Title:="SPECIFY ARTICLES", Type:=8)
    On Error GoTo 0
    If Article Is Nothing Then
        MsgBox ("No range/cell selected.  Aborting Macro.")
        Exit Sub
    End If
Else
    Set Article = Range(StartCell)
End If

If session Is Nothing Then
'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If
StandAlone = True
End If

Range(Article.Address & ":" & Article.End(xlDown).Address).Copy

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzbd10"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[16]").press
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
session.FindById("wnd[1]/tbar[0]/btn[13]").press
session.FindById("wnd[1]/usr/btnSOFORT_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/mbar/menu[0]/menu[4]").Select

If StandAlone Then
Call EndSAPCON
End If

MsgBox "ZBD10 push completed! Yay!"

End Sub
Sub SAP_Article_Push_VLD_Start_Cell(Optional StartCell As String)

Dim StandAlone As Boolean
Dim Article As Range
Dim lastRow As Integer
Dim CondType As String
Dim Yesterday As String


If StartCell = "" Then
    On Error Resume Next
    Set Article = Application.InputBox(Prompt:="Please select First Article in a column of articles " & _
        "you would like to push with ZBD10, then hit OK.", _
        Title:="SPECIFY ARTICLES", Type:=8)
    On Error GoTo 0
    If Article Is Nothing Then
        MsgBox ("No range/cell selected.  Aborting Macro.")
        Exit Sub
    End If
Else
    Set Article = Range(StartCell)
End If

If session Is Nothing Then
'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If
StandAlone = True
End If

Range(Article.Address & ":" & Article.End(xlDown).Address).Copy
Yesterday = Format(Now() - 1, "mm/dd/yyyy")
Do Until CondType = "VKA0" Or CondType = "VKP0" Or CondType = "ZADP"
    CondType = UCase(InputBox("What condition type do you want to push Retail (VKP0), Promo (VKA0), or MAP (ZADP)? Enter Condition type code.", "Condition Type"))
    If CondType = "VKA0" Or CondType = "VKP0" Or CondType = "ZADP" Then
        Else: MsgBox "Enter correct condition type please."
    End If
Loop

session.FindById("wnd[0]").Maximize

session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nv/ld"
session.FindById("wnd[0]").sendVKey 0
If CondType = "ZADP" Then
    session.FindById("wnd[0]/usr/ctxtRV14A-KONLI").Text = "z2"
    session.FindById("wnd[0]").sendVKey 0
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[0]/usr/ctxtP_1-LOW").Text = "10"
    session.FindById("wnd[0]/usr/ctxtL_1-LOW").Text = ""
    session.FindById("wnd[0]/usr/ctxtKSCHL-LOW").Text = CondType
    session.FindById("wnd[0]").sendVKey 0
    session.FindById("wnd[0]/usr/btn%_L_2_%_APP_%-VALU_PUSH").press
    session.FindById("wnd[1]/tbar[0]/btn[16]").press
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    session.FindById("wnd[0]/usr/chkPAR_L").Selected = True
    session.FindById("wnd[0]/usr/ctxtDATUM-LOW").Text = Yesterday
    session.FindById("wnd[0]/usr/txtMAX_LINE").Text = ""
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[0]/mbar/menu[1]/menu[5]").Select
    session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = "pipjavals"
    session.FindById("wnd[1]").sendVKey 0
Else
    session.FindById("wnd[0]/usr/ctxtRV14A-KONLI").Text = "z1"
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[0]/usr/ctxtP_1-LOW").Text = "1000"
    session.FindById("wnd[0]/usr/ctxtP_2-LOW").Text = "10"
    session.FindById("wnd[0]/usr/ctxtKSCHL-LOW").Text = CondType
    session.FindById("wnd[0]/usr/btn%_L_1_%_APP_%-VALU_PUSH").press
    session.FindById("wnd[1]/tbar[0]/btn[16]").press
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    session.FindById("wnd[0]/usr/chkPAR_L").Selected = True
    session.FindById("wnd[0]/usr/ctxtDATUM-LOW").Text = Yesterday
    session.FindById("wnd[0]/usr/txtMAX_LINE").Text = ""
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[0]/mbar/menu[1]/menu[5]").Select
    session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = "pipjavals"
    session.FindById("wnd[1]").sendVKey 0
End If

session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press

If StandAlone Then
Call EndSAPCON
End If

MsgBox "V/LD push completed! Yay!"

End Sub
Sub SAP_Article_Push_WPMA_Start_Cell(Optional StartCell As String)

Dim StandAlone As Boolean
Dim Article As Range
Dim lastRow As Integer


If StartCell = "" Then
    On Error Resume Next
    Set Article = Application.InputBox(Prompt:="Please select First Article in a column of articles " & _
        "you would like to push with ZBD10, then hit OK.", _
        Title:="SPECIFY ARTICLES", Type:=8)
    On Error GoTo 0
    If Article Is Nothing Then
        MsgBox ("No range/cell selected.  Aborting Macro.")
        Exit Sub
    End If
Else
    Set Article = Range(StartCell)
End If

If session Is Nothing Then
'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If
StandAlone = True
End If

Range(Article.Address & ":" & Article.End(xlDown).Address).Copy

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwpma"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/chkPA_ART").Selected = True
session.FindById("wnd[0]/usr/ctxtPA_VKORG").Text = "1000"
session.FindById("wnd[0]/usr/ctxtPA_VTWEG").Text = "10"
session.FindById("wnd[0]/usr/ctxtSO_FISEL-LOW").Text = "rpos"
session.FindById("wnd[0]/usr/chkPA_ART").SetFocus
session.FindById("wnd[0]/usr/btn%_SO_MATAR_%_APP_%-VALU_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[16]").press
session.FindById("wnd[1]/tbar[0]/btn[24]").press
session.FindById("wnd[1]/tbar[0]/btn[8]").press
session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
session.FindById("wnd[1]/tbar[0]/btn[13]").press
session.FindById("wnd[1]/usr/btnSOFORT_PUSH").press
session.FindById("wnd[1]/tbar[0]/btn[11]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press

If StandAlone Then
Call EndSAPCON
End If

MsgBox "WPMA push completed! Yay!"

End Sub
Sub SAP_Article_Push_ZBD10()
    SAP_Article_Push_ZBD10_Start_Cell
End Sub
Sub SAP_Article_Push_VLD()
    SAP_Article_Push_VLD_Start_Cell
End Sub
Sub SAP_Article_Push_WPMA()
    SAP_Article_Push_WPMA_Start_Cell
End Sub


Sub SAP_UPCSwap_MM42()

Dim WB As Workbook
Dim WS As Worksheet
Dim TempWs As Worksheet
Dim lastRow As Long
Dim Article As String
Dim UPC As String
Dim UPCtoReplace As String
Dim i As Long
Dim j As Integer

Set WB = ActiveWorkbook
Set WS = WB.Worksheets("WS_MT")
Set TempWs = WB.Worksheets("AM_Tmpt")

lastRow = TempWs.Range("A10000").End(xlUp).row

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'type /nmm42 into command bar and hit enter
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmm42"
session.FindById("wnd[0]").sendVKey 0

'Step through WS_AM sheet and perform UPC swap
For i = 9 To lastRow
'If no swap article on row skip
    If WS.Range("Q" & i) <> "" Then
        Article = WS.Range("Q" & i)
        UPC = WS.Range("R" & i)
        UPCtoReplace = WS.Range("O" & i)
'enter into article and hit F7 to look at UPC data
        session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Article
        session.FindById("wnd[0]/tbar[1]/btn[19]").press
        session.FindById("wnd[0]/usr/tblSAPLMGMWTAB_CONT_0100").GetAbsoluteRow(0).Selected = True
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]").sendVKey 7
'look for the upc that should be swapped and swap it with the internal for that article
'Green success row
'Red Error row if we couldn't find the upc to swap
        j = 0
        Do Until session.FindById("wnd[0]/usr/tabsTABSPR1/tabpZU03/ssubTABFRA1:SAPLMGMW:2110/subSUB2:SAPLMGD2:8025/tblSAPLMGD2TC_EAN/txtMEAN-EAN11[3," & j & "]").Text = ""
            If session.FindById("wnd[0]/usr/tabsTABSPR1/tabpZU03/ssubTABFRA1:SAPLMGMW:2110/subSUB2:SAPLMGD2:8025/tblSAPLMGD2TC_EAN/txtMEAN-EAN11[3," & j & "]").Text = UPCtoReplace Then
                session.FindById("wnd[0]/usr/tabsTABSPR1/tabpZU03/ssubTABFRA1:SAPLMGMW:2110/subSUB2:SAPLMGD2:8025/tblSAPLMGD2TC_EAN/txtMEAN-EAN11[3," & j & "]").Text = UPC
                WS.Range("O" & i & ":R" & i).Interior.ColorIndex = 4
                Exit Do
            End If
                j = j + 1
        Loop
        If WS.Range("O" & i & ":R" & i).Interior.ColorIndex <> 4 Then
            WS.Range("O" & i & ":R" & i).Interior.ColorIndex = 3
        End If
'hit enter a bunch of times to get through yellow warnings about upc type changings and save variant
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
    End If
Next

'back to home screen
session.FindById("wnd[0]").sendVKey 3

EndSAPCON

MsgBox "UPC Swap finished. Check for any errored red rows. Row is red because the UPC to replace (Column O)" & _
        "was not on Article listed (Column Q). No data was changed for that article."
End Sub

Sub SAP_Add_Var_AxisFlipFixer_MM42()
Dim WB As Workbook
Dim VarOutputWS As Worksheet
Dim Vendor As String
Dim Generic As String
Dim UPC As String
Dim Cost As String
Dim VariantNum As String
Dim i As Integer
Dim lastRow As Integer
Dim Color_Frame_Flavor_Desc As String
Dim Size_Lens_Desc As String
Dim LoopInd As String
Dim UPCType As String
Dim GenStartRow As Range
Dim GenEndRow As Range

'This will run step two of the add variant script on either error rows or unprocessed rows
'if an axis flip occures it will fix it for you yay
'it will also log over the errored winshuttle
'it does not yet fix the error counter on the log sheet

'Set some objects
Set WB = ActiveWorkbook
Set VarOutputWS = WB.Worksheets("Output_Variants")

'Set some variables
lastRow = VarOutputWS.Range("A2").End(xlDown).row + 1
i = 2
Vendor = VarOutputWS.Range("K" & i)
LoopInd = VarOutputWS.Range("A" & i)

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'Launch MM42
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmm42"
session.FindById("wnd[0]").sendVKey 0

'loop through each generic that had errors
Do Until i >= lastRow
    If (UCase(VarOutputWS.Range("DE" & i)) Like "*NOT FOUND*" Or _
    UCase(VarOutputWS.Range("DE" & i)) Like "*TSV_TNEW_PAGE_ALLOC_FAILED*" Or _
    VarOutputWS.Range("DE" & i) = "") And _
    LoopInd = "H" Then
'Set Vars for the row
        Generic = VarOutputWS.Range("G" & i)
        Size_Lens_Desc = VarOutputWS.Range("Q" & i)
        Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
        LoopInd = VarOutputWS.Range("A" & i)
        UPC = VarOutputWS.Range("BA" & i)
        UPCType = VarOutputWS.Range("BB" & i)
        Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        Set GenStartRow = VarOutputWS.Range("DE" & i)
        
        session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Generic
        session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = Vendor
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB2:SAPLMGD2:1030/btnMATRIX_PUSH").press
        session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:1001/btnPOSITION").press
        session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Size_Lens_Desc
        session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Color_Frame_Flavor_Desc
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
'insert error handle for axis flip
        session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:1001/tblSAPLWMMBTC_SEL/chkSL01[1,0]").Selected = True
        session.FindById("wnd[0]").sendVKey 0
        i = i + 1
        
'Set Vars for the row
        Size_Lens_Desc = VarOutputWS.Range("Q" & i)
        Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
        LoopInd = VarOutputWS.Range("A" & i)
        UPC = VarOutputWS.Range("BA" & i)
        UPCType = VarOutputWS.Range("BB" & i)
        Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        
        Do While LoopInd = "D"
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:1001/btnPOSITION").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Size_Lens_Desc
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Color_Frame_Flavor_Desc
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Color_Frame_Flavor_Desc
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Size_Lens_Desc
                session.FindById("wnd[1]").sendVKey 0
            End If
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:1001/tblSAPLWMMBTC_SEL/chkSL01[1,0]").Selected = True
            session.FindById("wnd[0]").sendVKey 0
            i = i + 1
'Set Vars for the row
            Size_Lens_Desc = VarOutputWS.Range("Q" & i)
            Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
            LoopInd = VarOutputWS.Range("A" & i)
            UPC = VarOutputWS.Range("BA" & i)
            UPCType = VarOutputWS.Range("BB" & i)
            Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        Loop
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB5:SAPLMGD2:1040/ctxtRMMWZ-MEINH").Text = "ea"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB5:SAPLMGD2:1040/ctxtRMMWZ-NUMTP").Text = UPCType
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB5:SAPLMGD2:1040/btnPUSH_VAR_EAN").press
        
        Do While LoopInd = "D1"
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/btnPOSITION").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Size_Lens_Desc
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Color_Frame_Flavor_Desc
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Color_Frame_Flavor_Desc
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Size_Lens_Desc
                session.FindById("wnd[1]").sendVKey 0
            End If
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text = UPC
            session.FindById("wnd[0]").sendVKey 0
'Set Vars for the row
            i = i + 1
            Size_Lens_Desc = VarOutputWS.Range("Q" & i)
            Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
            LoopInd = VarOutputWS.Range("A" & i)
            UPC = VarOutputWS.Range("BA" & i)
            UPCType = VarOutputWS.Range("BB" & i)
            Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        Loop
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP02").Select
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03").Select
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=pb52"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=#sn2"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/tbar[0]/btn[7]").press
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        Do While LoopInd = "D2"
            session.FindById("wnd[0]/usr/subSUB3:SAPLWMMB:4000/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/btnPOSITION").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Size_Lens_Desc
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Color_Frame_Flavor_Desc
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Color_Frame_Flavor_Desc
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Size_Lens_Desc
                session.FindById("wnd[1]").sendVKey 0
            End If
            session.FindById("wnd[0]/usr/subSUB3:SAPLWMMB:4000/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text = Cost
            session.FindById("wnd[0]").sendVKey 0
'Set Vars for the row
            i = i + 1
            Size_Lens_Desc = VarOutputWS.Range("Q" & i)
            Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
            LoopInd = VarOutputWS.Range("A" & i)
            UPC = VarOutputWS.Range("BA" & i)
            UPCType = VarOutputWS.Range("BB" & i)
            Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        Loop
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP04").Select
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ensch"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/usr/ctxtRMMW1-VKORG").Text = "1000"
        session.FindById("wnd[1]/usr/ctxtRMMW1-VTWEG").Text = "10"
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05").Select
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06").Select
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB7:SAPLWRF_ARTICLE_SCREENS:2704/ctxtMARC-MTVFP").Text = "01"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07").Select
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=pb14"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD03").Select
        
        Do While LoopInd = "D3"
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD03/ssubMTX_SUBSC:SAPLWMMB:2101/btnPOSITION").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Size_Lens_Desc
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Color_Frame_Flavor_Desc
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
        If session.ActiveWindow.Name = "wnd[2]" Then
            session.FindById("wnd[2]/tbar[0]/btn[0]").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = Color_Frame_Flavor_Desc
            session.FindById("wnd[1]").sendVKey 0
            session.FindById("wnd[2]").sendVKey 0
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = Size_Lens_Desc
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
        End If
            VariantNum = session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD03/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text
            VarOutputWS.Range("D" & i) = VariantNum
        'Set Vars for the row
            i = i + 1
            Size_Lens_Desc = VarOutputWS.Range("Q" & i)
            Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
            LoopInd = VarOutputWS.Range("A" & i)
            UPC = VarOutputWS.Range("BA" & i)
            UPCType = VarOutputWS.Range("BB" & i)
            Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
        Loop
        Set GenEndRow = VarOutputWS.Range("DE" & i - 1)
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        
        If session.FindById("wnd[0]/sbar").MessageType = "S" Then
            VarOutputWS.Range(GenStartRow.Address, GenEndRow.Address).Value = session.FindById("wnd[0]/sbar").Text
            VarOutputWS.Range(GenStartRow.Address, GenEndRow.Address).Interior.Color = 10213059
        Else
            VarOutputWS.Range(GenStartRow.Address, GenEndRow.Address).Value = session.FindById("wnd[0]/sbar").Text
        End If
        
    Else 'not a H row or this gen group didn't have axis flip
    'Set Vars for the row
        i = i + 1
        Generic = VarOutputWS.Range("G" & i)
        Size_Lens_Desc = VarOutputWS.Range("Q" & i)
        Color_Frame_Flavor_Desc = VarOutputWS.Range("N" & i)
        LoopInd = VarOutputWS.Range("A" & i)
        UPC = VarOutputWS.Range("BA" & i)
        UPCType = VarOutputWS.Range("BB" & i)
        Cost = CStr(Round(VarOutputWS.Range("BQ" & i) + 0.000001, 2))
'log it over errors and maybe change color?
        
    End If 'not a H row or this gen group didn't have axis flip

Loop

EndSAPCON

MsgBox "You did an axis flip and only clicked like 4 times how cool is that!"
MsgBox "If you want to submit to automation again do the following:" & vbCrLf & _
        "-Confirm there were no odd hidden errors from this script (its new...)" & vbCrLf & _
        "-Run Add Variant script starting on script 3 run on unprocessed rows (in the forground)" & vbCrLf & _
        "-Fix Automation matrix on Log sheet to show now errors on Add Variant steps" & vbCrLf & _
        "-Submit to Automation if there are Generics to create."

End Sub

Sub SAP_EDI_IBM_XK02_MRM2_MN05_Z_Ariba()
Dim Vendor As Long
Dim AltVendor As Long
Dim OutPutArr(4) As String
Dim i As Integer
Dim VendorSave As Boolean
Dim AltVendorSave As Boolean
Dim OutputStatus(5) As Boolean
Dim ZASave As Boolean
Dim VendorSaveMsg As String
Dim AltVendorSaveMsg As String
Dim OutputStatusMsg As String
Dim ZASaveMsg As String

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'Set up some Arrays we will use later
OutPutArr(0) = "ZNB"
OutPutArr(1) = "ZNG"
OutPutArr(2) = "ZNW"
OutPutArr(3) = "ZNS"
OutPutArr(4) = "ZBO"


'Input Vendor Number we are working on
Vendor = InputBox("What is the Vendor Number?", "Ariba to IBM update")
AltVendor = 0

'Uncheck Payment adv by EDI for Vendor
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/chkRF02K-D0215").Selected = True
session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/chkLFB1-XEDIP").Selected = False
If session.FindById("wnd[0]/usr/ctxtLFB1-LNRZB").Text <> "" Then
    AltVendor = session.FindById("wnd[0]/usr/ctxtLFB1-LNRZB").Text
End If
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").Text Like "*not yet been agreed*" Then
    session.FindById("wnd[0]").sendVKey 0
End If

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    VendorSave = True
Else
    VendorSave = False
End If




'Uncheck Payment adv by EDI for Alt Payee if needed
If AltVendor <> 0 Then
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/chkRF02K-D0215").Selected = True
session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = AltVendor
session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
session.FindById("wnd[0]/usr/chkRF02K-D0215").SetFocus
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/chkLFB1-XEDIP").Selected = False
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").Text Like "*not yet been agreed*" Then
    session.FindById("wnd[0]").sendVKey 0
End If

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    AltVendorSave = True
Else
    AltVendorSave = False
End If

End If
'deleting all the Output conditions with MN05
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmn05"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "ZARB"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").press

For i = LBound(OutPutArr) To UBound(OutPutArr)
session.FindById("wnd[0]/usr/ctxtF001").Text = OutPutArr(i)
session.FindById("wnd[0]/usr/ctxtF002").Text = "1000"
session.FindById("wnd[0]/usr/ctxtF003-LOW").Text = Vendor
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY").GetAbsoluteRow(0).Selected = True
session.FindById("wnd[0]/tbar[1]/btn[14]").press
session.FindById("wnd[0]/tbar[0]/btn[11]").press
'if the condition doesn't exist go back and go onto next condition type
If session.FindById("wnd[0]/sbar").Text = "Saving not necessary. No changes were made" Then
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
    session.FindById("wnd[0]").sendVKey 0
Else
    If session.FindById("wnd[0]/sbar").MessageType = "S" Then
        OutputStatus(i) = True
    Else
        OutputStatus(i) = False
    End If
End If

'If session.FindById("wnd[0]/sbar").messagetype = "S" Then
'    OutputStatus(i) = True
'Else
'    OutputStatus(i) = False
'End If

Next

'Remove ERS1 and ERS6 Invoicing Outputs with MRM2
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmrm2"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "ERS1"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtF001-LOW").Text = Vendor
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY").GetAbsoluteRow(0).Selected = True
session.FindById("wnd[0]/tbar[1]/btn[14]").press
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    OutputStatus(4) = True
Else
    OutputStatus(4) = False
End If


session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "ERS6"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[0,0]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/usr/ctxtF001").Text = "1000"
session.FindById("wnd[0]/usr/ctxtF002-LOW").Text = Vendor
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY").GetAbsoluteRow(0).Selected = True
session.FindById("wnd[0]/tbar[1]/btn[14]").press
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    OutputStatus(5) = True
Else
    OutputStatus(5) = False
End If



'Remove vendor from Z_Ariba_Vendor Table
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nz_ariba_vendor"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/btnVIM_POSI_PUSH").press
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "1000"
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = Vendor
session.FindById("wnd[1]/tbar[0]/btn[0]").press
If session.FindById("wnd[0]/usr/tbl/ARBA/SAPLTABLE_MAINTCTRL_/ARBA/AN_VENDOR/ctxt/ARBA/AN_VENDOR-LIFNR[1,0]").Text = Vendor Then
    session.FindById("wnd[0]/usr/tbl/ARBA/SAPLTABLE_MAINTCTRL_/ARBA/AN_VENDOR").Rows(0).Selected = True
    session.FindById("wnd[0]/tbar[1]/btn[14]").press
    session.FindById("wnd[0]/tbar[0]/btn[11]").press
Else
    MsgBox "Vendor is not in Z_Ariba_Vendor Table. Consult with Vendor Ops."
End If
session.FindById("wnd[0]/tbar[0]/btn[11]").press
If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    ZASave = True
Else
    ZASave = False
End If

session.FindById("wnd[0]/tbar[0]/btn[3]").press

If AltVendor <> 0 Then
    If VendorSave And AltVendorSave And OutputStatus(0) And OutputStatus(1) And OutputStatus(2) _
    And OutputStatus(3) And OutputStatus(4) And OutputStatus(5) And ZASave Then
        MsgBox Vendor & " is now set up to transact with IBM in SAP system " & session.Info.SystemName
    Else
        MsgBox "Something went wrong... Manually make this vendor IBM EDI in system " & session.Info.SystemName, vbCritical
    End If
Else
    If VendorSave And OutputStatus(0) And OutputStatus(1) And OutputStatus(2) _
    And OutputStatus(3) And OutputStatus(4) And OutputStatus(5) And ZASave Then
        MsgBox Vendor & " is now set up to transact with IBM in SAP system " & session.Info.SystemName
    Else
        MsgBox "Something went wrong... Manually make this vendor IBM EDI in system " & session.Info.SystemName, vbCritical
    End If
End If

Call EndSAPCON


End Sub

Sub SAP_EDI_Ariba_XK02_MRM2_MN05_Z_Ariba()
Dim Vendor As Long
Dim AltVendor As Long
Dim OutPutArr(3) As String
Dim i As Integer
Dim VendorSave As Boolean
Dim AltVendorSave As Boolean
Dim OutputStatus(5) As Boolean
Dim ZASave As Boolean
Dim VendorSaveMsg As String
Dim AltVendorSaveMsg As String
Dim OutputStatusMsg As String
Dim ZASaveMsg As String

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'Set up some Arrays we will use later
OutPutArr(0) = "ZNB"
OutPutArr(1) = "ZNG"
OutPutArr(2) = "ZNW"
OutPutArr(3) = "ZNS"


'Input Vendor Number we are working on

Vendor = InputBox("What is the Vendor Number?", "IBM to Ariba update")

AltVendor = 0

'Check Payment adv by EDI for Vendor
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[8]").press
session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = Vendor
session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
session.FindById("wnd[0]/usr/chkRF02K-D0215").Selected = True
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/chkLFB1-XEDIP").Selected = True

If session.FindById("wnd[0]/usr/ctxtLFB1-LNRZB").Text <> "" Then
    AltVendor = session.FindById("wnd[0]/usr/ctxtLFB1-LNRZB").Text
End If
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").Text Like "*not yet been agreed*" Then
    session.FindById("wnd[0]").sendVKey 0
End If

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    VendorSave = True
Else
    VendorSave = False
End If

'Check Payment adv by EDI for Alt Payee if needed
If AltVendor <> 0 Then
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nxk02"
    session.FindById("wnd[0]").sendVKey 0
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    session.FindById("wnd[0]/usr/chkRF02K-D0215").Selected = True
    session.FindById("wnd[0]/usr/ctxtRF02K-LIFNR").Text = AltVendor
    session.FindById("wnd[0]/usr/ctxtRF02K-BUKRS").Text = "1000"
    session.FindById("wnd[0]").sendVKey 0
    session.FindById("wnd[0]/usr/chkLFB1-XEDIP").Selected = True
    session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").Text Like "*not yet been agreed*" Then
    session.FindById("wnd[0]").sendVKey 0
End If

    If session.FindById("wnd[0]/sbar").MessageType = "S" Then
        AltVendorSave = True
    Else
        AltVendorSave = False
    End If

End If

'Set up Output conditions with MN04
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmn04"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "zarb"
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").CaretPosition = 4
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[1,0]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").press

For i = LBound(OutPutArr) To UBound(OutPutArr)
    session.FindById("wnd[0]/usr/ctxtKOMB-BSART").Text = OutPutArr(i)
    session.FindById("wnd[0]/usr/ctxtKOMB-EKORG").Text = "1000"
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtKOMB-LIFNR[0,0]").Text = Vendor
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-PARVW[2,0]").Text = "LS"
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtRV13B-PARNR[3,0]").Text = "PIPJAVALS"
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-NACHA[4,0]").Text = "6"
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-VSZTP[5,0]").Text = "1"
    session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-SPRAS[6,0]").Text = "EN"
    session.FindById("wnd[0]/tbar[0]/btn[11]").press
    If session.FindById("wnd[0]/sbar").MessageType = "S" Then
        OutputStatus(i) = True
    Else
        OutputStatus(i) = False
    End If
Next

'Set up Invoicing Output conditions ERS1 and ERS6 with MRM1
session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMRM1"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "ERS1"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/txtKOMB-LIFRE[0,0]").Text = Vendor
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-PARVW[2,0]").Text = "LS"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtRV13B-PARNR[3,0]").Text = "PIPJAVALS"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-NACHA[4,0]").Text = "6"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-VSZTP[5,0]").Text = "3"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-SPRAS[6,0]").Text = "EN"
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    OutputStatus(4) = True
Else
    OutputStatus(4) = False
End If

session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/usr/ctxtRV13B-KSCHL").Text = "ERS6"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[0,0]").Select
session.FindById("wnd[1]/tbar[0]/btn[0]").press
session.FindById("wnd[0]/usr/ctxtKOMB-BUKRS").Text = "1000"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/txtKOMB-LIFRE[0,0]").Text = Vendor
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-PARVW[2,0]").Text = "LS"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtRV13B-PARNR[3,0]").Text = "PIPJAVALS"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-NACHA[4,0]").Text = "6"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-VSZTP[5,0]").Text = "4"
session.FindById("wnd[0]/usr/tblSAPMV13BTCTRL_FAST_ENTRY/ctxtNACH-SPRAS[6,0]").Text = "EN"
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    OutputStatus(5) = True
Else
    OutputStatus(5) = False
End If

session.FindById("wnd[0]/tbar[0]/btn[3]").press

'Add Vendor to Z_Ariba_Vendor Table
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nz_ariba_vendor"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/tbar[1]/btn[5]").press
session.FindById("wnd[0]/usr/ctxt/ARBA/AN_VENDOR-BUKRS").Text = "1000"
session.FindById("wnd[0]/usr/ctxt/ARBA/AN_VENDOR-LIFNR").Text = Vendor
session.FindById("wnd[0]/tbar[0]/btn[11]").press

If session.FindById("wnd[0]/sbar").MessageType = "S" Then
    ZASave = True
Else
    ZASave = False
End If

session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press

If AltVendor <> 0 Then
    If VendorSave And AltVendorSave And OutputStatus(0) And OutputStatus(1) And OutputStatus(2) _
    And OutputStatus(3) And OutputStatus(4) And OutputStatus(5) And ZASave Then
        MsgBox Vendor & " is now set up to transact with Ariba in SAP system " & session.Info.SystemName
    Else
        MsgBox "Something went wrong... Manually make this vendor Ariba EDI in system " & session.Info.SystemName, vbCritical
    End If
Else
    If VendorSave And OutputStatus(0) And OutputStatus(1) And OutputStatus(2) _
    And OutputStatus(3) And OutputStatus(4) And OutputStatus(5) And ZASave Then
        MsgBox Vendor & " is now set up to transact with Ariba in SAP system " & session.Info.SystemName
    Else
        MsgBox "Something went wrong... Manually make this vendor Ariba EDI in system " & session.Info.SystemName, vbCritical
    End If
End If


End Sub

Private Sub SAP_New_Generic_VOL_MM42_VK11()
'**********************************************************************************************************************
'This is intended to be a "short term" fix to help with Variant Overload issue for Generic Creates
'Step 1 and 2 of Bapi_Article_Create_Generic_v9.11 must have run prior to this
'
'This will do everything that steps 3-9 for Bapi_Article_Create_Generic_v9.11 does
'
'Loop through Outputs sheet
'If line is blank or memory error and is an H row
'Generic stuff
'-Basic data tab
'-ARN if needed (we can comment out if we want to have WS do this)
'-loop through H/D to create Variants
'-Listing Tab (Minus setting listings - we need to do this as a follow on cause site group data is blocked for a
'   while as SAP saves)
'-Purch Tab (may want to add cost loop if we want to comment out variant level stuff to have WS run 4-9)
'-Sales Tab (loop through variants to create retails - Look into using rule to copy all retails if there is no change
'   with in generic. This could save a ton of time)
'-Logistics DC Tab
'-Logistics Store Tab
'-POS Tab
'-Back to Basic data
'-Loop through variants (D4) to add upc/type to variants
'-Save Generic so PIR stuff works
'-Loop through Variants (D4) (we can comment it all out if we want WS to do scripts 4-9)
'
'


'Notes
'-Does not handle UPC Swap due to formulas in UPC Swap sheet.
'-Deal with .value everywhere to fix stupid random .text error
'-drop logs/update logsheet to show that all scrips ran?
'-Fixing the automation matrix
'-Comment better. this one is commented ok at this point but could be better
'-think about how else this could break and if it breaks is it easy to fix
'-size range flag?
'-Shits being locked out for a long time while SAP saves stuff. Deal with that somehow...
'**********************************************************************************************************************

Dim WB As Workbook
Dim WS As Worksheet
Dim LogSheet As Worksheet
Dim Generic As Long
Dim i As Integer
Dim j As Integer
Dim lastRow As Integer
Dim Article As Double
Dim Vendor As Double
Dim XYLooper As Integer
Dim VarCounter As Integer
Dim HLine As Integer
Dim D3Start As Integer
Dim D3End As Integer


Set WB = ActiveWorkbook
Set WS = WB.Worksheets("Output")
Set LogSheet = WB.Worksheets("LogSheet")
lastRow = WS.Range("A2").End(xlDown).row
i = 2


'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'check that Steps 1 and 2 have already be run

session.FindById("wnd[0]").Maximize


'
'for i = 2 to lastrow
'if logrow is blank or errored
'set generic variable
'if H do gen stuff plus setting variant retails
'end if gen stuf
'if D3 do all the Var specific stuff
'end if var stuff
'VK11 for generic
'i=i+1
'next
'



Do Until i > lastRow
    If (WS.Range("DI" & i) = "" Or WS.Range("DI" & i) Like "*TSV_TNEW_PAGE_ALLOC_FAILED*") And WS.Range("A" & i) = "H" Then
        'Generic stuff
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmm42"
        session.FindById("wnd[0]").sendVKey 0
        HLine = i
        Generic = WS.Range("C" & i)
        session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Generic
        session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = WS.Range("K" & i).Value
        'session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").SetFocus
        'session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").caretPosition = 5
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB1:SAPLMGD2:1010/txtMAKT-MAKTX").Text = WS.Range("E" & i).Value
        
        'If PAC set PAC
        If WS.Range("W" & i) > 1 And WS.Range("W" & i) <> "" Then
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/radSMEINH-KZAUSME[6,1]").Selected = True
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0,1]").Text = "PAC"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2,1]").Text = WS.Range("W" & i).Value
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MESUB[3,1]").Text = "EA"
        End If
        
        'If CAR set CAR
        If WS.Range("T" & i) = "CAR" Then
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/radSMEINH-KZBSTME[5,2]").Selected = True
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0,2]").Text = "CAR"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2,2]").Text = WS.Range("U" & i).Value
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MESUB[3,2]").Text = "EA"
        End If
        
        'If Rounding Profile set RP
        If WS.Range("CJ" & i) = True And WS.Range("V" & i) > 1 Then
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEINH[0,2]").Text = "CAR"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-AZSUB[2,2]").Text = WS.Range("V" & i).Value
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MESUB[3,2]").Text = "EA"
        End If
        
        'Drop Weights and Dims
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-BRGEW[12,0]").Text = Round(WS.Range("Z" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-NTGEW[13,0]").Text = Round(WS.Range("AA" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-GEWEI[14,0]").Text = WS.Range("AB" & i).Value
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-LAENG[15,0]").Text = Round(WS.Range("AC" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-BREIT[16,0]").Text = Round(WS.Range("AD" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-HOEHE[17,0]").Text = Round(WS.Range("AE" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-MEABM[18,0]").Text = WS.Range("AF" & i).Value
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-VOLUM[19,0]").Text = Round(WS.Range("AG" & i).Value + 0.000001, 2)
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-VOLEH[20,0]").Text = WS.Range("AH" & i).Value
        
        'Brand
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB8:SAPLWRF_ARTICLE_SCREENS:2020/ctxtMARA-BRAND_ID").Text = WS.Range("AJ" & i).Value
        'Country of Origin
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB12:SAPLZWRF_ARTICLE_SCREENS:2002/ctxtMAW1-WHERL").Text = WS.Range("AK" & i).Value
        'HTS Code
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB12:SAPLZWRF_ARTICLE_SCREENS:2002/ctxtMAW1-WSTAW").Text = WS.Range("AL" & i).Value
        'Tax PCode from script 9 Size range
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB11:SAPLWRF_ARTICLE_SCREENS:2001/ctxtMARA-EXTWG").Text = WS.Range("AM" & i).Value
        'Aptos PlaceHolder from New Info Chars script
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB12:SAPLZWRF_ARTICLE_SCREENS:2002/txtMARA-BISMT").Text = WS.Range("AZ" & i).Value
        
        If WS.Range("DB" & i) <> "" Then
            'Season Validity
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB14:SAPLWRF_ARTICLE_SCREENS:2004/ctxtMARA-SAISO").Text = WS.Range("DB" & i).Value
            'Season Year Validity
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB14:SAPLWRF_ARTICLE_SCREENS:2004/txtMARA-SAISJ").Text = WS.Range("FV" & i).Value
        End If
    
    'Info Chars
        'Gender
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").Text = WS.Range("AO" & i).Value
        'Model Year
        If WS.Range("AP" & i) <> "N/A" Then
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[1,32]").Text = WS.Range("AP" & i).Value
        End If
        'Dona
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[2,32]").Text = WS.Range("AQ" & i).Value
        'Online Only
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[3,32]").Text = WS.Range("AR" & i).Value
        'Dividend Eligible
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[4,32]").Text = WS.Range("AS" & i).Value
        'Employee Discount percent
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[5,32]").Text = WS.Range("AT" & i).Value
        'page down on info chars
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/btnOES_PDOWN").press
        'Validation Required
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").Text = WS.Range("AU" & i).Value
        'Inspection Required
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[5,32]").Text = WS.Range("AV" & i).Value
        'Page down on info chars
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/btnOES_PDOWN").press
        'MBI
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[0,32]").Text = WS.Range("CY" & i).Value
        'Liability Waiver
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[1,32]").Text = WS.Range("AX" & i).Value
        'Size Range from script 9 Size Range
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[3,32]").Text = WS.Range("DD" & i).Value
        
        
        
        'ARN Stuff in Alt Data from script 9 Size Range
        If WS.Range("FH" & i) = True Then
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=zu06"
            session.FindById("wnd[0]").sendVKey 0
            'ARN and Defaulted 0004
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpZU06/ssubTABFRA1:SAPLMGMW:2110/subSUB2:SAPLAEMM:2686/tblSAPLAEMMTCTRL_ADDI/ctxtWTADDI_EDIT-ADDIMAT[7,0]").Text = WS.Range("DA" & i).Value
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpZU06/ssubTABFRA1:SAPLMGMW:2110/subSUB2:SAPLAEMM:2686/tblSAPLAEMMTCTRL_ADDI/ctxtWTADDI_EDIT-ADDIFM[9,0]").Text = "0004"
            session.FindById("wnd[0]").sendVKey 0
            'Back to Basic Data for size range
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=Baba"
            session.FindById("wnd[0]").sendVKey 0
        End If
        
        
        'Click Variants button and create all the Variants by checking them
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=pb14"
        session.FindById("wnd[0]").sendVKey 0
        XYLooper = i
        
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=xych"
        session.FindById("wnd[0]").sendVKey 0
        'Click Position button, drop color and size in. axis flip fix if needed. Check box on article to create Variant
        Do While WS.Range("A" & XYLooper) = "H" Or WS.Range("A" & XYLooper) = "D"
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=posi"
            session.FindById("wnd[0]").sendVKey 0
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
            'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
            End If
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:1001/tblSAPLWMMBTC_SEL/chkSL01[1,0]").Selected = True
            XYLooper = XYLooper + 1
        Loop
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=back"
        session.FindById("wnd[0]").sendVKey 0
        
        
    
        'Listings stuff - We'll set some radio buttons but skip listing to 6 basic sites - known memory issue
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp02"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLWRF_ARTICLE_SCREENS:2214/chkRMMWZ-MPFLB").Selected = True
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLWRF_ARTICLE_SCREENS:2214/chkRMMWZ-NLIPR").Selected = True
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP02/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLWRF_ARTICLE_SCREENS:2214/chkRMMWZ-LILI").Selected = False
        
    
        'Purch tab stuff Add logic
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp03"
        session.FindById("wnd[0]").sendVKey 0
        'Regular Vendor Check
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/chkEINA-RELIF").Selected = True
        'Unlimited Check
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB4:SAPLWRF_ARTICLE_SCREENS:2222/chkEINE-UEBTK").Selected = True
        'Var Order Unit set to blank
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-VABME").Text = ""
        'VPN
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/txtEINA-IDNLF").Text = WS.Range("BK" & i).Value
        'Handling Type
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-ZZHANDTYP").Text = WS.Range("BM" & i).Value
        'RA Required
        If WS.Range("BY" & i) > 0 Then
            'session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-ZZRAREQ").Text = WS.Range("BY" & i).Value
        End If
        'Stock Type
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-ZZSTKTYP").Text = WS.Range("BN" & i).Value
        'Return Policy Code
        If WS.Range("BZ" & i) > 0 Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-ZZREPOL").Text = WS.Range("BZ" & i).Value
        End If
        'Minimun Standard Quantity & Rounding Profile Rule
        If WS.Range("CJ" & i) <> "" And WS.Range("CJ" & i) > 1 Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB4:SAPLWRF_ARTICLE_SCREENS:2222/txtEINE-MINBM").Text = WS.Range("V" & i).Value
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB4:SAPLWRF_ARTICLE_SCREENS:2222/ctxtEINE-RDPRF").Text = "D050"
        End If
        'Stage Time
        If WS.Range("BX" & i) > 0 Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB4:SAPLWRF_ARTICLE_SCREENS:2222/txtEINE-STAGING_TIME").Text = WS.Range("BX" & i).Value
        End If
        'Cost and per for Ea
        If WS.Range("T" & i) = "EA" Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-NETPR").Text = Round(WS.Range("BQ" & i).Value + 0.000001, 2)
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-PEINH").Text = "1"
        Else
            'Cost and per for Car
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-NETPR").Text = Round(WS.Range("BR" & i).Value + 0.000001, 2)
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-PEINH").Text = WS.Range("U" & i).Value
        End If
        'Defaulted EA value for Cost
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/ctxtEINE-BPRME").Text = "EA"
        'PrDateCntr
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/ctxtEINE-MEPRF").Text = "2"
        'Coop indicator
        If WS.Range("BW" & i) = "COOP" Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-EKKOL").Text = "COOP"
        End If
        
        
        
        'Sales tab to set Company Code 1000 Distribution Channels 10 and 30 and set retails of all variants
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp04"
        session.FindById("wnd[0]").sendVKey 0
        If session.FindById("wnd[0]/sbar").MessageType = "W" Then
            session.FindById("wnd[0]").sendVKey 0
        End If
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ensch"
        session.FindById("wnd[0]").sendVKey 0
        If session.FindById("wnd[0]/sbar").MessageType = "W" Then
            session.FindById("wnd[0]").sendVKey 0
        End If
        session.FindById("wnd[1]/usr/ctxtRMMW1-VKORG").Text = "1000"
        session.FindById("wnd[1]/usr/ctxtRMMW1-VTWEG").Text = "10"
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        'Set some defaults for Purch Price Determination and Sales Price Determination then loop through variants
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMW:2030/subSUB4:SAPLMGD2:2233/ctxtCALP-EKERV").Text = "01"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMW:2030/subSUB4:SAPLMGD2:2233/ctxtCALP-VKERV").Text = "01"
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=pb53"
        session.FindById("wnd[0]").sendVKey 0
        
        
        Do Until WS.Range("A" & XYLooper) = "D2"
            XYLooper = XYLooper + 1
        Loop
            'XYLooper = XYLooper - 1
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=xych"
        session.FindById("wnd[0]").sendVKey 0
        Do While WS.Range("A" & XYLooper) = "D2"
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=posi"
            session.FindById("wnd[0]").sendVKey 0
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
            'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
            End If
            If Right(Round(WS.Range("CE" & XYLooper).Value + 0.000001, 2) * 100, 2) = "73" Or Right(Round(WS.Range("CE" & XYLooper).Value + 0.000001, 2) * 100, 2) = "95" Or _
                Right(Round(WS.Range("CE" & XYLooper).Value + 0.000001, 2) * 100, 2) = "00" Then
                session.FindById("wnd[0]/usr/subSUB3:SAPLWMMB:4000/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text = WS.Range("CE" & XYLooper).Value
            Else
                session.FindById("wnd[0]/usr/subSUB3:SAPLWMMB:4000/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text = InputBox("Retail Price Looks Off... Price on line is " & WS.Range("CE" & XYLooper), "Price Looks Off", WS.Range("CE" & XYLooper))
            End If
                
            XYLooper = XYLooper + 1
        Loop
        
        session.FindById("wnd[0]/usr/subSUB2:SAPLMGW_SALES_PRICE_MATRIX:1220/chkCALP-VKABS").Selected = True
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=back"
        session.FindById("wnd[0]").sendVKey 0
        
        'Link to Comp 1000 Dist 30
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ensch"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/usr/ctxtRMMW1-VTWEG").Text = "30"
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        'Back to Dist 10
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/ensch"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/usr/ctxtRMMW1-VTWEG").Text = "10"
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        
        'Logistics: DC Tab
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp05"
        session.FindById("wnd[0]").sendVKey 0
        'Rounding Profile Blank it out
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB3:SAPLWRF_ARTICLE_SCREENS:2242/ctxtMARC-RDPRF").Text = ""
        'Purch Group
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB6:SAPLZWRF_ARTICLE_SCREENS:2244/ctxtMARC-EKGRP").Text = WS.Range("AN" & i).Value
        'Article Status
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB6:SAPLZWRF_ARTICLE_SCREENS:2244/ctxtMARC-MMSTA").Text = WS.Range("CI" & i).Value
        'Avail. Check
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB6:SAPLZWRF_ARTICLE_SCREENS:2244/ctxtMARC-MTVFP").Text = "01"
        'Crossdoc
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB6:SAPLZWRF_ARTICLE_SCREENS:2244/ctxtMARC-ZZCROSSD").Text = WS.Range("CK" & i).Value
        'Replenishment Type
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB3:SAPLWRF_ARTICLE_SCREENS:2242/ctxtMARC-DISMM").Text = WS.Range("CZ" & i).Value
    
        
        'Logistics: Store
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp06"
        session.FindById("wnd[0]").sendVKey 0
        If session.FindById("wnd[0]/sbar").MessageType = "W" Then
            session.FindById("wnd[0]").sendVKey 0
        End If
        'Purch Group
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB7:SAPLWRF_ARTICLE_SCREENS:2704/ctxtMARC-EKGRP").Text = WS.Range("AN" & i).Value
        'Article Status
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB7:SAPLWRF_ARTICLE_SCREENS:2704/ctxtMARC-MMSTA").Text = WS.Range("CI" & i).Value
        'Avail Check
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB7:SAPLWRF_ARTICLE_SCREENS:2704/ctxtMARC-MTVFP").Text = "01"
        'Replenishment Type
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB2:SAPLWRF_ARTICLE_SCREENS:2242/ctxtMARC-DISMM").Text = WS.Range("CZ" & i).Value
    
        
        'POS Tab
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp07"
        session.FindById("wnd[0]").sendVKey 0
        If session.FindById("wnd[0]/sbar").MessageType = "W" Then
        session.FindById("wnd[0]").sendVKey 0
        End If
        'Disc. Allowed
        If WS.Range("CL" & i) = "X" Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2251/chkWLK2-RBZUL").Selected = True
        End If
        'Price Required
        If WS.Range("CM" & i) = "X" Then
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2251/chkWLK2-PRERF").Selected = True
        End If
        'Memo Price
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2251/txtWLK2-ZZMEMO").Text = WS.Range("CN" & i).Value
        'Till Text stuff
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2273/tblSAPLMGD2TC_BON/ctxtMAMT-SPRAS[0,0]").Text = "EN"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2273/tblSAPLMGD2TC_BON/ctxtMAMT-MEINH[1,0]").Text = "EA"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2273/tblSAPLMGD2TC_BON/txtMAMT-MAKTM[4,0]").Text = WS.Range("CO" & i).Value

        'back to Basic data tab
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp01"
        session.FindById("wnd[0]").sendVKey 0

        Do Until WS.Range("A" & XYLooper) = "D4"
            XYLooper = XYLooper + 1
        Loop
        
        'Loop through D4 to set Variant level things
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB5:SAPLMGD2:1040/ctxtRMMWZ-MEINH").Text = "EA"
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB5:SAPLMGD2:1040/ctxtRMMWZ-NUMTP").Text = WS.Range("BB" & XYLooper).Value
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=pb30"
        session.FindById("wnd[0]").sendVKey 0
        
        Do While WS.Range("A" & XYLooper) = "D4"
            'Press position button, put in x char and y char
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/btnPOSITION").press
            session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
            session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
            'error handle for axis flip
            If session.ActiveWindow.Name = "wnd[2]" Then
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtX_CHAR_VAL").Text = WS.Range("Q" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
                session.FindById("wnd[2]/tbar[0]/btn[0]").press
                session.FindById("wnd[1]/usr/ctxtY_CHAR_VAL").Text = WS.Range("N" & XYLooper).Value
                session.FindById("wnd[1]").sendVKey 0
            End If
            'Variant we care about is now highlighted put in UPC and hit enter
            session.FindById("wnd[0]/usr/tabsMAIN_TS/tabpMD01/ssubMTX_SUBSC:SAPLWMMB:2101/tblSAPLWMMBTC_VAL/txtVL01[1,0]").Text = WS.Range("BA" & XYLooper).Value
            'session.FindById("wnd[0]").sendVKey 0
            'increment row
            XYLooper = XYLooper + 1
        Loop
        
        'back to basic data tab
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        
        'end generic stuff, Save the Generic then go back into Generic then go do variant specific stuff next
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Generic
        session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = WS.Range("K" & i).Value
        
        'sometimes you lock yourself out as it is trying to save. Keep clicking enter till youre not locked out
        Do Until session.ActiveWindow.Text Like "*Basic Data*"
            session.FindById("wnd[0]").sendVKey 0
        Loop
        
        
    'start Variant stuff 48:00
    'Skip over H, D, D1, D2, D3
        XYLooper = i
        Do Until WS.Range("A" & XYLooper) = "D4"
            XYLooper = XYLooper + 1
        Loop
        
        'Loop through D4 to set Variant level things
        VarCounter = 1
        Do While WS.Range("A" & XYLooper) = "D4"
            session.FindById("wnd[0]/tbar[1]/btn[13]").press
            'Variant number and drop in D3 column D
            session.FindById("wnd[1]/usr/ctxtRMMW1-VARNR").Text = Generic & Format(VarCounter, "0000")
            'WS.Range("D" & D4Line) = Generic & Format(VarCounter, "0000")
            'session.FindById("wnd[1]/usr/ctxtRMMW1-VARNR").caretPosition = 10
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
'            'UPC on basic data Tab
'            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/txtSMEINH-EAN11[8,0]").Text = WS.Range("BA" & XYLooper)
'            'UPC Catagory
'            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB4:SAPLMGD2:8022/tblSAPLMGD2TC_ME_8022/ctxtSMEINH-NUMTP[9,0]").Text = WS.Range("BB" & XYLooper)
            'Online Only
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB16:SAPLCTMS:4000/sub:SAPLCTMS:4000/ctxtRCTMS-MWERT[3,32]").Text = WS.Range("AR" & XYLooper).Value
            If WS.Range("DB" & XYLooper) <> "" Then
                'Season Validity
                session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB14:SAPLWRF_ARTICLE_SCREENS:2004/ctxtMARA-SAISO").Text = WS.Range("DB" & XYLooper).Value
                'Season Year Validity
                session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB14:SAPLWRF_ARTICLE_SCREENS:2004/txtMARA-SAISJ").Text = WS.Range("FV" & XYLooper).Value
            End If
            
            'Touch the Purch Tab to Create PIR
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp03"
            session.FindById("wnd[0]").sendVKey 0
            'VPN
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/txtEINA-IDNLF").Text = WS.Range("BK" & XYLooper).Value
            'Cost Update
            'Cost and per for Ea
            If WS.Range("T" & XYLooper) = "EA" Then
                session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-NETPR").Text = Round(WS.Range("BQ" & XYLooper).Value + 0.000001, 2)
            Else
                'Cost and per for Car
                session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB5:SAPLMGD2:2223/txtEINE-NETPR").Text = Round(WS.Range("BR" & XYLooper).Value + 0.000001, 2)
            End If
            'Var order Clear out
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP03/ssubTABFRA1:SAPLMGMW:2000/subSUB2:SAPLZMGD2:2221/ctxtEINA-VABME").Text = ""
            
            'Logistics: DC Tab Set Replenishment type
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp05"
            session.FindById("wnd[0]").sendVKey 0
            If session.FindById("wnd[0]/sbar").MessageType = "W" Then
                session.FindById("wnd[0]").sendVKey 0
            End If
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMW:2004/subSUB3:SAPLWRF_ARTICLE_SCREENS:2242/ctxtMARC-DISMM").Text = WS.Range("CZ" & XYLooper).Value
            'Logistics: Store Tab Set Replenishment type
            session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06").Select
        session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMW:2008/subSUB2:SAPLWRF_ARTICLE_SCREENS:2242/ctxtMARC-DISMM").Text = WS.Range("CZ" & XYLooper).Value
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "=sp01"
        session.FindById("wnd[0]").sendVKey 0
        VarCounter = VarCounter + 1
        XYLooper = XYLooper + 1
        Loop
        
        'After all variants have been touched lets go back to Gen and hit save button
        session.FindById("wnd[0]/tbar[1]/btn[13]").press
        session.FindById("wnd[1]/usr/ctxtRMMW1-VARNR").Text = ""
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        
'        'end generic stuff, Save the Generic then go back into Generic then go do variant specific stuff next
'        session.FindById("wnd[0]/tbar[0]/btn[11]").press
'        session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Generic
'        session.FindById("wnd[0]/usr/ctxtRMMW1-LIFNR").Text = WS.Range("K" & i).Value
'
'        'sometimes you lock yourself out as it is trying to save. Keep clicking enter till youre not locked out
'        Do Until session.ActiveWindow.Text Like "*Basic Data*"
'            session.FindById("wnd[0]").sendVKey 0
'        Loop
'        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        
        
        'Create Retail Condtion for Generic using Generic Variable set at beginning of if
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nvk11"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/ctxtRV13A-KSCHL").Text = "VKP0"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[1]/usr/sub:SAPLV14A:0100/radRV130-SELKZ[5,0]").Select
        session.FindById("wnd[1]/tbar[0]/btn[0]").press
        session.FindById("wnd[0]/usr/ctxtKOMG-VKORG").Text = "1000"
        session.FindById("wnd[0]/usr/ctxtKOMG-VTWEG").Text = "10"
        session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-MATNR[0,0]").Text = Generic
        session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/ctxtKOMG-VRKME[1,0]").Text = "EA"
        session.FindById("wnd[0]").sendVKey 0
        session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[3,0]").Text = WS.Range("CE" & HLine).Value
        'session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[3,1]").SetFocus
        'session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKONP-KBETR[3,1]").caretPosition = 16
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        session.FindById("wnd[0]/tbar[0]/btn[3]").press
        
''Insert WSM3 job
'    D3Start = i
'    Do Until WS.Range("A" & D3Start) = "D3"
'        D3Start = D3Start + 1
'    Loop
'    D3End = D3Start
'    Do Until WS.Range("A" & D3End) <> "D3"
'        D3End = D3End + 1
'    Loop
'        D3End = D3End - 1
'
'    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwsm3"
'    session.FindById("wnd[0]").sendVKey 0
'    session.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "1"
'    session.FindById("wnd[0]/usr/ctxtLSTFL").Text = "02"
'    session.FindById("wnd[0]/usr/ctxtASORT-LOW").Text = "1"
'    session.FindById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
'    session.FindById("wnd[1]/tbar[0]/btn[16]").press
'    WS.Range("D" & D3Start & ":D" & D3End).Copy
'    session.FindById("wnd[1]/tbar[0]/btn[24]").press
'    Application.CutCopyMode = False
'    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'    session.FindById("wnd[0]/usr/btn%_ASORT_%_APP_%-VALU_PUSH").press
'    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "2"
'    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "3"
'    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "3pl_all"
'    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "950"
'    session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "Rpos"
'    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'    session.FindById("wnd[0]/usr/chkDATBE").selected = False
'    session.FindById("wnd[0]/tbar[1]/btn[8]").press
'    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'    session.FindById("wnd[0]/tbar[0]/btn[3]").press
    

        i = i + 1
    Else
        i = i + 1
    End If 'H and log is error or blank
    
Loop 'Do until i > Lastrow

EndSAPCON

End Sub

Sub SAP_Modify_Int_Transit_MEK2()


Dim WS As Worksheet
Dim i As Integer
Dim j As Integer
Dim Portcode As String
Dim lastRow As Integer
Dim CreateNewTT As Integer

Set WS = ActiveWorkbook.Worksheets("Transit Time")

'Set lastrow variable
i = 5
Do Until WS.Range("E" & i) = " " Or WS.Range("E" & i) = ""
lastRow = i
i = i + 1
Loop
lastRow = lastRow

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

'Go to MEK2 Tcode
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmek2"
session.FindById("wnd[0]").sendVKey 0

'Enter Condition type, choose international, hit check box/enter
session.FindById("wnd[0]/usr/ctxtRV13A-KSCHL").Text = "zpdt"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[1]/tbar[0]/btn[0]").press

'Loop through all rows and for each row loop through the three DC columns
For i = 5 To lastRow
    If Not WS.Cells(i, 1) Like "EXAMPLE*" Then
        For j = 7 To 9
            'if there is nothing in a DC TT column skip to next column
            If WS.Cells(i, j) <> "" Then
                'enter port code(INCO 2), enter DC formatted 000#, hit excecute
                session.FindById("wnd[0]/usr/txtF001").Text = WS.Cells(i, 5).Value
                session.FindById("wnd[0]/usr/ctxtF002-LOW").Text = Format(j - 6, "0000")
                session.FindById("wnd[0]/tbar[1]/btn[8]").press
                
                'if no condition exists for PC/DC pair log that it needs to be manually created
                'and get back to previous page for next pair
                If session.FindById("wnd[0]/sbar").Text Like "*No condition*" Then
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    session.FindById("wnd[0]/usr/ctxtRV13A-KSCHL").Text = "zpdt"
                    session.FindById("wnd[0]").sendVKey 0
                    
                    'Looks like you don't always get the popup asking for Int or Dom...
                    'if we get it we'll choose Int
                    If session.ActiveWindow.Name = "wnd[1]" Then
                        session.FindById("wnd[1]/tbar[0]/btn[0]").press
                    End If
                    WS.Cells(i, j).Interior.ColorIndex = 3
                    WS.Cells(i, j + 16).Interior.ColorIndex = 3
                    WS.Cells(i, j + 16) = "Transit Time does not exist for site " & Format(j - 6, "0000") & " Please manually create."
                
                'There is an existing TT condition, lets modify the TT and save
                Else
                
                    session.FindById("wnd[0]/usr/tblSAPMV13ATCTRL_FAST_ENTRY/txtKOMG-ZZPLIFZ[2,0]").Text = WS.Cells(i, j).Value
                    session.FindById("wnd[0]/tbar[0]/btn[11]").press
                    
                    'if we get a success message then we all good, log that and move on.
                    'if we don't have user manually look into that...
                    If session.FindById("wnd[0]/sbar").MessageType = "S" Then
                        WS.Cells(i, j + 16).Interior.ColorIndex = 4
                        WS.Cells(i, j).Interior.ColorIndex = 4
                        WS.Cells(i, j + 16) = "Transit Time updated updated via SAP script - " & Format(Now(), "dd/mm/yyyy hh:nn")
                    Else
                        WS.Cells(i, j + 16).Interior.ColorIndex = 3
                        WS.Cells(i, j).Interior.ColorIndex = 3
                        WS.Cells(i, j + 16) = "Something Errored... Manually update this one - " & Format(Now(), "dd/mm/yyyy hh:nn")
                    End If
                End If
            End If
        Next
    End If
Next

'Back buttons to home screen
session.FindById("wnd[0]/tbar[0]/btn[3]").press
session.FindById("wnd[0]/tbar[0]/btn[3]").press

'Log Headers cause i don't want to update the blank template
WS.Range("W4") = "Sumner - 0001 Log"
WS.Range("X4") = "Bedford - 0002 Log"
WS.Range("Y4") = "Goodyear - 0003 Log"

'end message so you know you're done with script
MsgBox "All done. Take a look at logs in columns W-Y and Highlighting in columns G-I."

End Sub
Sub SAP_UPCSwap_MASS()
'Brian Combs - 2021

'This macro performs a MASS UPC swap on the Article MIT. This should be executed AFTER Article Create is complete and
'only if a category UPC swap is necessary.

'Variables
    Dim AWB As Workbook
    Dim UPCSheet As Worksheet
    Dim lastRow As Long
    Dim Article As Range
    Dim ArtRange As Range
    Dim NewUPC As String

'Set workbooks and sheets
    Set AWB = ActiveWorkbook
    Set UPCSheet = AWB.Worksheets("UPC Swap")

'Count rows on UPC Sheet
    lastRow = 8
    Do While UPCSheet.Range("D" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

'Set Article Range
    Set ArtRange = UPCSheet.Range("D8:D" & lastRow)
    
'SAP STUFF--------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON
    
'Navigate to T-code MASS
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nmass"
    session.FindById("wnd[0]").sendVKey 0
'Object Type - BUS1001001
    session.FindById("wnd[0]/usr/ctxtMASSSCREEN-OBJECT").Text = "bus1001001"
'Variant Name - UPC_CHG
    session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").Text = "UPC_CHG"
'Press Execute
    session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").SetFocus
    session.FindById("wnd[0]/usr/ctxtMASSSCREEN-VARNAME").CaretPosition = 7
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
    
'Loop through ArtRange
    For Each Article In ArtRange
    'Check if Article is a "Category Swap" (True) and not "IE" (Internal)
        If UPCSheet.Range("H" & Article.row).Value = "True" And UPCSheet.Range("P" & Article.row).Value <> "IE" Then
            NewUPC = UPCSheet.Range("O" & Article.row).Value
            Article.Copy
        'Press Multiple Selection (Arrow - Article)
            session.FindById("wnd[0]/usr/tabsTAB/tabpCHAN/ssubSUB_ALL:SAPLMASS_SEL_DIALOG:0200/ssubSUB_SEL:SAPLMASSFREESELECTIONS:1000/sub:SAPLMASSFREESELECTIONS:1000/btnMASSFREESEL-MORE[0,69]").press
        'Press Delete Entire Selection (Trash can)
            session.FindById("wnd[1]/tbar[0]/btn[16]").press
        'Press Upload from Clipboard (Clipboard)
            session.FindById("wnd[1]/tbar[0]/btn[24]").press
        'Press Copy (Execute)
            session.FindById("wnd[1]/tbar[0]/btn[8]").press
        'Press Execute (Execute)
            session.FindById("wnd[0]/tbar[1]/btn[8]").press
        'Enter New UPC in the "GTIN" field
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/ssubSUB_HEAD:SAPLMASSINTERFACE:0210/tblSAPLMASSINTERFACETCTRL_HEADER/txtHEADER_STRUC-FIELD3-VALUE-LEFT[3,0]").Text = NewUPC
        'Press Perform Mass Change (Down Arrows)
            session.FindById("wnd[0]/usr/tabsTBSTRP_TABLES/tabpTAB1/ssubFIELDS:SAPLMASSINTERFACE:0202/btnFDAE").press
        'Press Save
            session.FindById("wnd[0]/tbar[0]/btn[11]").press
            
        'Check for "Green Success bubbles"
            If session.FindById("wnd[0]/usr/txtNR_E").Text = "0" Then
            'If green, fill Article and UPC green
                Article.Interior.ColorIndex = 4
                UPCSheet.Range("O" & Article.row).Interior.ColorIndex = 4
            Else
            'If not green, fill Article and UPC red
                Article.Interior.ColorIndex = 3
                UPCSheet.Range("O" & Article.row).Interior.ColorIndex = 3
            End If
            
        'Press Back (Green Arrow)
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        'Press Back again (Green Arrow)
            session.FindById("wnd[0]/tbar[0]/btn[3]").press
        Else
        End If 'If Artcle is a "Category Swap"
    Next 'For Each Article In ArtRange

'Press Back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'Press Back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press

'Call ENDSAPCON
    EndSAPCON
    
'END SAP STUFF--------------------------------------------------------------

'Loop through ArtRange and look for red errors
    For Each Article In ArtRange
    'If the article is red
        If Article.Interior.ColorIndex = 3 Then
            MsgBox "Errors found! Check columns D and O of the UPC Swap worksheet for any red errors."
            Exit Sub
        Else
        End If
    Next

'Success
    MsgBox "UPC swap complete! No errors detected. Dang, you're good!"
    
End Sub


Sub SAP_HTS_Additional_Fix_SM30()
'Used as a part of HTS code updates. For some unknown reason when HTS codes are created they
'are set up incorrectly
'This script loops through column A of a sheet (HTS codes in column A that were set up incorrect
'and corrects the Additional info (T604 - NIHON) data element to ""


Dim i As Integer
Dim errored As Integer



'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

If session.Info.SystemName <> "ECD" Then
    MsgBox "This script can only be run in ECD. Exiting Sub."
    EndSAPCON
    Exit Sub
End If

session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsm30"
session.FindById("wnd[0]").sendVKey 0
session.FindById("wnd[0]/usr/ctxtVIEWNAME").Text = "V_T604"
session.FindById("wnd[0]/usr/btnUPDATE_PUSH").press

i = 1
Do While ActiveSheet.Range("A" & i) <> ""

session.FindById("wnd[0]/usr/btnVIM_POSI_PUSH").press
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = "US"
session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").Text = ActiveSheet.Range("A" & i).Text 'inputting code to find
session.FindById("wnd[1]").sendVKey 0
If session.FindById("wnd[0]/usr/tblSAPL080ETCTRL_V_T604/txtV_T604-STAWN[1,0]").Text = ActiveSheet.Range("A" & i).Text Then
    session.FindById("wnd[0]").sendVKey 2
    session.FindById("wnd[0]/usr/ctxtV_T604-NIHON").Text = "" 'blanking out
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
Else
    MsgBox "HTS code on your sheet does not exist for country code US. Highlighting yellow for you to look into..."
    ActiveSheet.Range("A" & i).Interior.ColorIndex = 6
    errored = errored + 1
End If

i = i + 1
Loop


EndSAPCON

MsgBox "Done updating " & i - 1 & " HTS codes' Additional data to Null. Continue working through documentation to set up transport."

If errored > 0 Then
MsgBox "Looks like you have " & errored & " HTS codes to look into. The script did not find them in the table..."
End If

End Sub

Sub SAP_Run_Errored_Listings_WSM3()
Dim GenWS As Worksheet
Dim VarWS As Worksheet
Dim LogWS As Worksheet
Dim lastRow As Integer
Dim WorkNeeded(2) As Boolean
Dim ListingColumn As Integer
Dim i As Integer

'Set WS variable and make sure we are running this on an AC Template Exit if not
On Error Resume Next
Set GenWS = ActiveWorkbook.Worksheets("Output")
Set VarWS = ActiveWorkbook.Worksheets("Output_Variants")
Set LogWS = ActiveWorkbook.Worksheets("LogSheet")
On Error GoTo 0
If err.Description <> "" Then
    MsgBox "This can only be run on a MIT brah... Exiting Sub."
    Exit Sub
End If

WorkNeeded(0) = False
WorkNeeded(1) = False
WorkNeeded(2) = False


If LogWS.Range("E5") > 0 And LogWS.Range("E5") <> "" Then WorkNeeded(0) = True
If LogWS.Range("E6") > 0 And LogWS.Range("E6") <> "" Then WorkNeeded(1) = True
If LogWS.Range("E23") > 0 And LogWS.Range("E23") <> "" Then WorkNeeded(2) = True

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "WSM3"
session.FindById("wnd[0]").sendVKey 0


For i = LBound(WorkNeeded) To UBound(WorkNeeded)

    If WorkNeeded(i) Then
    
        Select Case i
        Case Is = 0
        VarWS.activate
        ListingColumn = 110
        Case Is = 1
        VarWS.activate
        ListingColumn = 111
        Case Is = 2
        GenWS.activate
        ListingColumn = 112
        End Select
        
        
        
        lastRow = Range("D10000").End(xlUp).row
        Rows("1:1").AutoFilter Field:=ListingColumn, Criteria1:=RGB(247 _
                , 150, 70), Operator:=xlFilterCellColor

        session.FindById("wnd[0]/usr/chkDATBE").Selected = False
        session.FindById("wnd[0]/usr/ctxtASORT-LOW").Text = "1"
        session.FindById("wnd[0]/usr/ctxtMATNR-LOW").Text = "1565160001"
        session.FindById("wnd[0]/usr/ctxtLSTFL").Text = "02"
        session.FindById("wnd[0]/usr/btn%_ASORT_%_APP_%-VALU_PUSH").press
        session.FindById("wnd[1]/tbar[0]/btn[16]").press
        If i = 0 Or i = 2 Then
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "1"
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "2"
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "3"
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").Text = "950"
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").Text = "rpos"
        End If
        If i = 1 Or i = 2 Then
            session.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,5]").Text = "3pl_all"
        End If
        session.FindById("wnd[1]").sendVKey 0
        session.FindById("wnd[1]/tbar[0]/btn[8]").press
        session.FindById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
        Range("D2:D" & lastRow).SpecialCells(xlCellTypeVisible).Copy
        session.FindById("wnd[1]/tbar[0]/btn[16]").press
        session.FindById("wnd[1]/tbar[0]/btn[24]").press
        Application.CutCopyMode = False
        session.FindById("wnd[1]/tbar[0]/btn[8]").press
        session.FindById("wnd[0]/tbar[1]/btn[8]").press
        If session.FindById("wnd[0]/sbar").Text = "No article exists for these selection criteria" Then
            MsgBox "An article you tried to list didn't get created. Figure out what that is all about... Exiting sub"
            Exit Sub
        End If
        On Error Resume Next
        session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell").SetFocus
        If err.Description <> "" Then 'vba error meaning we couldn't find SAP error about locked site all is good log some stuff and say hurray
            Select Case i
            Case Is = 0
                LogWS.Range("E5").Value = 0
                WorkNeeded(0) = False
                MsgBox "Var 1 listing script errors have been fixed."
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Value = "Listings set by Kevin's cool listing VBA fixer @ " & Format(Now(), "HH:MM:SS")
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Interior.Color = 10213059
            Case Is = 1
                LogWS.Range("E6").Value = 0
                WorkNeeded(1) = False
                MsgBox "Var 2 listing script errors have been fixed."
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Value = "Listings set by Kevin's cool listing VBA fixer @ " & Format(Now(), "HH:MM:SS")
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Interior.Color = 10213059
            Case Is = 2
                LogWS.Range("E23").Value = 0
                WorkNeeded(2) = False
                MsgBox "Gen listing script errors have been fixed."
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Value = "Listings set by Kevin's cool listing VBA fixer @ " & Format(Now(), "HH:MM:SS")
                Range(Cells(2, ListingColumn), Cells(lastRow, ListingColumn)).Interior.Color = 10213059
            End Select
        Else 'No VBA error meaning we found the SAP error message about sites being locked
            Select Case i
            Case Is = 0
                MsgBox "Var 1 listing fix failed because someone was blocking you. You will need to run this macro again.", vbCritical
            Case Is = 1
                MsgBox "Var 2 listing fix failed because someone was blocking you. You will need to run this macro again.", vbCritical
            Case Is = 2
                MsgBox "Gen listing fix failed because someone was blocking you. You will need to run this macro again.", vbCritical
            End Select
            session.FindById("wnd[0]/tbar[0]/btn[3]").press 'gets us back to successful listings screen
            If session.ActiveWindow.Name = "wnd[1]" Then
                session.FindById("wnd[1]/tbar[0]/btn[0]").press 'if all sites were locked we get a pop up we need to get past
            End If
        End If
        On Error GoTo 0
        
        
        session.FindById("wnd[0]/tbar[0]/btn[3]").press 'gets us back to WSM3 input
        
    End If
    Rows("1:1").AutoFilter
Next

session.FindById("wnd[0]/tbar[0]/btn[3]").press 'gets us back to session manager
Application.Calculate
ActiveWorkbook.Worksheets("WS_AC").activate
ActiveSheet.Range("F4").activate

If Not WorkNeeded(0) And WorkNeeded(1) And WorkNeeded(2) Then
    MsgBox "Someone blocked you AGAIN. How silly. Run again to list", vbCritical
Else
    MsgBox "All listing errors present have been fixed. If listins were your only errors you should be good to submit to automation to finish up"
End If

EndSAPCON

End Sub

Sub SAP_New_Size_LC_MM42()

Dim WS As Worksheet
Dim i As Integer
Dim Generic As String
Dim Lenscoloradd As Integer


'Set WS variable and make sure we are running this on an AC Template Exit if not
On Error Resume Next
Set WS = ActiveWorkbook.Worksheets("Add Variant Check")
On Error GoTo 0
If err.Description <> "" Then
    MsgBox "This can only be run on an AC with Add Variant Check Sheet. Exiting Sub."
    Exit Sub
End If

'Connect to SAP with SAPCON Sub
SAPCON
If session Is Nothing Then
    MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
    Exit Sub
End If

Lenscoloradd = MsgBox("Are you adding new lens colors? If you select no that means you are adding new sizes.", vbYesNo)

'Start on row 2, Lauch MM42
i = 2
session.FindById("wnd[0]").Maximize
session.FindById("wnd[0]/tbar[0]/okcd").Text = "mm42"
session.FindById("wnd[0]").sendVKey 0


'Loop through rows 2 till the end of column and enter new sizes
Do Until WS.Range("M" & i) = ""
    Generic = WS.Range("M" & i).Text
    session.FindById("wnd[0]/usr/ctxtRMMW1-MATNR").Text = Generic
    session.FindById("wnd[0]/tbar[1]/btn[19]").press
    session.FindById("wnd[0]/usr/tblSAPLMGMWTAB_CONT_0100").GetAbsoluteRow(0).Selected = True
    session.FindById("wnd[0]").sendVKey 0
    session.FindById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMW:2008/subSUB2:SAPLMGD2:1030/btnSA_MERKMALE").press
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "=aufs"
    session.FindById("wnd[0]").sendVKey 0
    If Lenscoloradd = vbYes Then
        session.FindById("wnd[1]/usr/txtCLHP-CR_STATUS_TEXT").Text = "Lens"
    Else
        session.FindById("wnd[1]/usr/txtCLHP-CR_STATUS_TEXT").Text = "Size"
    End If
    session.FindById("wnd[1]").sendVKey 0
    session.FindById("wnd[2]/usr/cntlGRID1/shellcont/shell").ClickCurrentCell
    Do Until WS.Range("M" & i).Text <> Generic
        If WS.Range("N" & i).Text <> "" Then
            session.FindById("wnd[0]/tbar[1]/btn[16]").press
            session.FindById("wnd[0]/usr/sub:SAPLCTMS:0100/ctxtRCTMS-MWERT[0,35]").Text = Left(WS.Range("N" & i).Text, 6)
            session.FindById("wnd[0]").sendVKey 0
        End If
        i = i + 1
    Loop
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
    session.FindById("wnd[0]/tbar[0]/btn[11]").press
        
Loop

'Exit MM42 to get back to Session manager home screen
session.FindById("wnd[0]/tbar[0]/btn[3]").press


If Lenscoloradd = vbYes Then
    MsgBox "New Lens Colors added to generic. You rock!!"
Else
    MsgBox "New Sizes added to generic. You rock!!"
End If

EndSAPCON

End Sub

