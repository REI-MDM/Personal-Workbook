Attribute VB_Name = "SAPValidation_v007"
Option Explicit

Sub SAP_Validation_ZSE16N()
'Brian Combs - 2021
'Brian Combs & Carly Kane - 2022 repaired Assortment Maintain Removal Validation

'This sub will call the appropriate sub for different templates. Currently works for:
    'Maintain Promos
    'Temp Listings
    'VIF Temp Listings PSI Configuration
    'Assortment Creates
    'Assortment Maintains

'Variables
    Dim Template As Workbook
    Dim TemplateWorks As Boolean
    Dim TemplateChoice As String
    
'Error handling start
    On Error Resume Next

'Set Template
    Set Template = ActiveWorkbook
    TemplateWorks = False
    
'Auto call subs---------------------------------------
    
'Determine "Request Type" and call the appropriate sub
    If Template.ContentTypeProperties.Count > 0 Then
        'If the request type is Maintain Promo
        If Template.ContentTypeProperties("Request Type").Value = "Maintain Promo" Then
            TemplateWorks = True
            Call Validate_Maintain_Promo
        'If the request type is Temp Listings
        ElseIf Template.ContentTypeProperties("Request Type").Value = "Temp Listings" Then
            TemplateWorks = True
            Call Validate_Temp_Listings
        'If the request type is Assortment Create
        ElseIf Template.ContentTypeProperties("Request Type").Value = "Assortment Group Create" Then
            TemplateWorks = True
            Call Validate_Asst_Create
        'If the request type is Assortment Maintain
        ElseIf Template.ContentTypeProperties("Request Type").Value = "Assortment Group Maintain" Then
            TemplateWorks = True
            Call Validate_Asst_Maint
        Else
        End If
    End If

'Manual call subs---------------------------------------

'If there are no "Request Type" properties, give the user a selection
    If TemplateWorks = False Then
            TemplateChoice = Application.InputBox(Title:="Template Choice", Prompt:="Please enter the NUMBER of the template you want to validate." & vbNewLine & _
            "1 - Maintain Promo" & vbNewLine & _
            "2 - Temp Listings" & vbNewLine & _
            "3 - Assortment Create" & vbNewLine & _
            "4 - Assortment Maintain")
        'If user does not enter value
            If TemplateChoice = "" Then
                MsgBox "You did not enter a value. Please try again."
                Exit Sub
        'If user clicks cancel
            ElseIf TemplateChoice = "False" Then
                MsgBox "You selected cancel. Please try again."
                Exit Sub
        'If user selects AC
            ElseIf TemplateChoice = "1" Then
                TemplateWorks = True
                Call Validate_Maintain_Promo
        'If user selects Temp Listings
            ElseIf TemplateChoice = "2" Then
                TemplateWorks = True
                Call Validate_Temp_Listings
        'If user selects Assortment Create
            ElseIf TemplateChoice = "3" Then
                TemplateWorks = True
                Call Validate_Asst_Create
        'If user selects Assortment Maintain
            ElseIf TemplateChoice = "4" Then
                TemplateWorks = True
                Call Validate_Asst_Maint
        'If user selects the top secret code
            ElseIf TemplateChoice = "9" Then
                ActiveWorkbook.FollowHyperlink Address:="https://www.youtube.com/watch?v=6iFbuIpe68k"
        'If the user enters anything else
             Else
                MsgBox "Invalid entry. Please try again."
                Exit Sub
            End If
    End If
    
'Error handling end
    On Error GoTo 0
    
'If everything fails
    If TemplateWorks = False Then
        MsgBox "Sorry bud. Something went wrong. This macro only works with Maintain Promos, Temp Listings, and Asst Creates currently. Maybe the file is missing 'Request Type'?"
        Exit Sub
    End If
    
End Sub

Private Sub Validate_Maintain_Promo()
'Brian Combs - 2021

'Timer
    Debug.Print Now

'Variables
    Dim MPWB As Workbook
    Dim singleCell As Range
    
'Maintain Promo Worksheet Variables
    Dim MPSheet As Worksheet
    Dim MPLastRow As Long
    Dim PromoRange As Range
    Dim ActionRange As Range
    Dim VarRange As Range
    Dim PriceRange As Range
    Dim StatRange As Range
    Dim Action As String
    Dim MPKeyRange As Range
    Dim VLookRange As Range
    Dim MatchRange As Range

'Validation Worksheet Variables
    Dim ValidSheet As Worksheet
    Dim ValidLastRow As Long
    Dim ValidKeyRange As Range

'Turn off display alerts
    Application.DisplayAlerts = False
    
'Set Maintain Promo Workbook and Worksheet
    Set MPWB = ActiveWorkbook
    Set MPSheet = MPWB.Worksheets("Maintain_Promo")
    
'Delete the Validation sheet if it already exists
    On Error Resume Next
    Set ValidSheet = ActiveWorkbook.Worksheets("Validation")
    If Not ValidSheet Is Nothing Then
        ValidSheet.Delete
    Else
    End If
    On Error GoTo 0
    
'Create a Validation Worksheet
    MPWB.Sheets.Add(After:=MPSheet).Name = "Validation"
    Set ValidSheet = MPWB.Worksheets("Validation")

'Count rows on MP Sheet
    MPSheet.Select
    MPLastRow = 6
    Do While Range("A" & MPLastRow).Value <> ""
        MPLastRow = MPLastRow + 1
    Loop
    MPLastRow = MPLastRow - 1

'Set ranges
    Set PromoRange = MPSheet.Range("C6:C" & MPLastRow)
    Set VarRange = MPSheet.Range("I6:I" & MPLastRow)
    Set ActionRange = MPSheet.Range("A6:A" & MPLastRow)
    
'SAP STUFF--------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON

'Navigate to ZSE16N
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/NZSE16N"
    session.FindById("wnd[0]").sendVKey 0
    
'Enter WAKP in the Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WAKP"
    session.FindById("wnd[0]/usr/ctxtGD-TAB").CaretPosition = 4
    session.FindById("wnd[0]").sendVKey 0
    
'GoTo > Variants > Get > Variant PULL_PROMO_DATA
    session.FindById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select
    session.FindById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").Text = "PULL_PROMO_DATA"
    session.FindById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").SetFocus
    session.FindById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME").CaretPosition = 15
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
   
'Press the More button(Arrow) - Promotion
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").press

'Press the Clipboard button
    PromoRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press

'Press the Transfer Data (Execute) button
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
    
'Press the More button(Arrow) - Article
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").press

'Press the Clipboard button
    VarRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    
'Press the Transfer Data (Execute) button
    session.FindById("wnd[1]/tbar[0]/btn[8]").press

'Press the Online (Execute) button
    session.FindById("wnd[0]/tbar[1]/btn[8]").press

'Press Export button
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"

'Press Local File
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"

'Select "In the Clipboard"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus

'Press Continue (Green Check)
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

'Press Exit
    session.FindById("wnd[0]/tbar[0]/btn[15]").press

'Press Exit
    session.FindById("wnd[0]/tbar[0]/btn[15]").press

'Call ENDSAPCON
    EndSAPCON
    
'END SAP STUFF--------------------------------------------------------------

'Paste WAKP table into Validation sheet
    ValidSheet.Select
    Range("A1").Select
    Range("A1").PasteSpecial
    
'Run Text_To_Columns on Validation table
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    Resume Next

'Count rows on ValidSheet
    ValidSheet.Select
    ValidLastRow = 6
    Do While Range("B" & ValidLastRow).Value <> ""
        ValidLastRow = ValidLastRow + 1
    Loop
    ValidLastRow = ValidLastRow - 1

'Set Validation Key Range
    Set ValidKeyRange = ValidSheet.Range("A6:A" & ValidLastRow)
    ValidSheet.Range("A5") = "Promo|Article|Price|Status" & Now
    
'Drop in some formulas in Column A to create a Validation sheet key
    For Each singleCell In ValidKeyRange
    'If stat B (Promo & Article & Promo Price & Status)
        If Range("F" & singleCell.row).Value = "B        " Then
        singleCell.Value = "=TEXTJOIN(""|"",TRUE,B" & singleCell.row & _
                           ",C" & singleCell.row & _
                           ",D" & singleCell.row & _
                           ",TRIM(F" & singleCell.row & "))"

    'If stat F (Promo & Article & Status)
        ElseIf Range("F" & singleCell.row).Value = "F        " Then
        singleCell.Value = "=TEXTJOIN(""|"",TRUE,B" & singleCell.row & _
                           ",C" & singleCell.row & _
                           ",TRIM(F" & singleCell.row & "))"
        Else
    'If stat is blank
        singleCell.Value = "BLANK STATUS"
        End If
    Next

'Autofit ValidSheet
    ValidSheet.Columns("A:L").AutoFit

'Select MPSheet and set StatRange
    MPSheet.Select
    MPSheet.Range("DA5") = "Requested Status"
    Set StatRange = MPSheet.Range("DA6:DA" & MPLastRow)
    
'Loop through all variants and drop a "B" or "F" in column DA
    For Each singleCell In StatRange
        Action = Range("A" & singleCell.row).Value
        If Action = "Add Item and Price" Then
            singleCell.Value = "B"
        ElseIf Action = "Update Price" Then
            singleCell.Value = "B"
        ElseIf Action = "Remove (Deactivate)" Then
            singleCell.Value = "F"
        End If
    Next
    
'Drop in some formulas in Column DB (Promo & Article & Promo Price & Status) to create an MP sheet key
    Set MPKeyRange = MPSheet.Range("DB6:DB" & MPLastRow)
    MPSheet.Range("DB5") = "MPSheet Promo|Article|Price|Status"
    For Each singleCell In MPKeyRange
        'If status is "B" use Promo & Article & Promo Price & Status
        If MPSheet.Range("DA" & singleCell.row).Value = "B" Then
        singleCell.Value = "=TEXTJOIN(""|"",TRUE,C" & singleCell.row & _
                           ",I" & singleCell.row & _
                           ",P" & singleCell.row & _
                           ",DA" & singleCell.row & ")"
        'If status is "F" use Promo & Article & Status
        ElseIf MPSheet.Range("DA" & singleCell.row).Value = "F" Then
        singleCell.Value = "=TEXTJOIN(""|"",TRUE,C" & singleCell.row & _
                           ",I" & singleCell.row & _
                           ",DA" & singleCell.row & ")"
        End If
    Next

'Drop in some formulas in Column DC to find a match between Validation sheet key and MP sheet key
    Set VLookRange = MPSheet.Range("DC6:DC" & MPLastRow)
    MPSheet.Range("DC5") = "ValidSheet Promo|Article|Price|Status"
    For Each singleCell In VLookRange
        singleCell.Value = "=VLOOKUP(DB" & singleCell.row & ",Validation!$A$6" & ":$A$" & ValidLastRow & ",1,0)"
    Next

'Drop in match boolean
    Set MatchRange = MPSheet.Range("DD6:DD" & MPLastRow)
    MPSheet.Range("DD5") = "Match?"
    For Each singleCell In MatchRange
        singleCell.Value = "=IF(MATCH(DB" & singleCell.row & ",DC" & singleCell.row & ",0),TRUE,FALSE)"
    Next

'Autofit MPSheet
    MPSheet.Columns("DA:DD").AutoFit

'Loop through MatchRange to find errors. Prompt a Message Box if you find a false value
    For Each singleCell In MatchRange
        If singleCell.Value2 <> "True" Then
            MsgBox "Errors found! Please check column DD for FALSE or #N/A values"
            Exit Sub
        Else
        End If
    Next

'Success message!
    MsgBox "Sweet! Everything looks good! Check out column DD for confirmation."
    
'Turn off display alerts
    Application.DisplayAlerts = False

'Timer
    Debug.Print Now
   
End Sub

Private Sub Validate_Temp_Listings()
'Brian Combs - 2021

'Variables
    Dim TLWB As Workbook
    Dim singleCell As Range
    Dim WS As Worksheet
    
'Temp Listing Worksheet Variables
    Dim TLSheet As Worksheet
    Dim TLLastRow As Long
    Dim VarRange As Range
    Dim SiteRange As Range
    Dim TLKeyRange As Range
    Dim VLookRange As Range
    Dim MatchRange As Range

'Validation Worksheet Variables
    Dim ValidSheet As Worksheet
    Dim ValidLastRow As Long
    Dim ValidKeyRange As Range
    
'PSI Variables
    Dim UserChoice As VbMsgBoxResult

'Turn off display alerts
    Application.DisplayAlerts = False

'Set Temp Listing Workbook and Worksheet
    Set TLWB = ActiveWorkbook
    Set TLSheet = TLWB.Worksheets("Sheet1")

'Delete the Validation sheet if it already exists
    On Error Resume Next
    Set ValidSheet = ActiveWorkbook.Worksheets("Validation")
    If Not ValidSheet Is Nothing Then
        ValidSheet.Delete
    Else
    End If
    On Error GoTo 0

'Create a Validation Worksheet
    TLWB.Sheets.Add(After:=TLSheet).Name = "Validation"
    Set ValidSheet = TLWB.Worksheets("Validation")

'Count rows on TL Sheet
    TLSheet.Select
    TLLastRow = 2
    Do While Range("B" & TLLastRow).Value <> ""
        TLLastRow = TLLastRow + 1
    Loop
    TLLastRow = TLLastRow - 1

'Set ranges
    Set VarRange = TLSheet.Range("B2:B" & TLLastRow)
    Set SiteRange = TLSheet.Range("A2:A" & TLLastRow)

'SAP STUFF--------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON

'Navigate to ZSE16N
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
    session.FindById("wnd[0]").sendVKey 0

'Enter WLK1 in the Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WLK1"
    session.FindById("wnd[0]/usr/ctxtGD-TAB").CaretPosition = 4
    session.FindById("wnd[0]").sendVKey 0

'Press the More button(Arrow) - Assortment
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").press

'Press the Clipboard button
    SiteRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press

'Press the More button(Arrow) - Article
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").press

'Press the Clipboard button
    VarRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
    session.FindById("wnd[1]/tbar[0]/btn[8]").press

'Clear max number of hits
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").SetFocus
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").CaretPosition = 0

'Press the Transfer Data (Execute) button
    session.FindById("wnd[0]/tbar[1]/btn[8]").press

'Press the Export button
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"

'Press Local File
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"

'Select "In the Clipboard"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus

'Press Continue (Green Check)
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

'Press Exit
    session.FindById("wnd[0]/tbar[0]/btn[15]").press
    session.FindById("wnd[0]/tbar[0]/btn[15]").press

'Call ENDSAPCON
    EndSAPCON

'END SAP STUFF--------------------------------------------------------------

'Paste WLK1 table into Validation sheet
    ValidSheet.Select
    Range("A1").Select
    Range("A1").PasteSpecial

'Run Text_To_Columns on Validation table
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    Resume Next

'Count rows on ValidSheet
    ValidSheet.Select
    ValidLastRow = 6
    Do While ValidSheet.Range("B" & ValidLastRow).Value <> ""
        ValidLastRow = ValidLastRow + 1
    Loop
    ValidLastRow = ValidLastRow - 1

'Set Validation Key Range
    Set ValidKeyRange = ValidSheet.Range("A6:A" & ValidLastRow)
    ValidSheet.Range("A5") = "Site|Article|EndDate|StartDate" & Now

'Drop in some formulas in Column A to create a Validation sheet key
    For Each singleCell In ValidKeyRange
        singleCell.Value = "=B" & singleCell.row & "&C" & singleCell.row & "&E" & singleCell.row & "&G" & singleCell.row
    Next

'Autofit ValidSheet
    ValidSheet.Columns("A:L").AutoFit

'Select TLSheet and set TLKeyRange
    TLSheet.Select
    TLSheet.Range("E1") = "TLSheet Site|Article|EndDate|StartDate"
    Set TLKeyRange = TLSheet.Range("E2:E" & TLLastRow)

'Drop in some formulas in Column E to create a Temp Listings sheet key
    For Each singleCell In TLKeyRange
        singleCell.Value = "=A" & singleCell.row & "&B" & singleCell.row & "&VALUE(D" & singleCell.row & ")&VALUE(C" & singleCell.row & ")"
    Next

'Drop in some formulas in Column I to find a match between Validation sheet key and TL sheet key
    TLSheet.Range("F1") = "ValidSheet Site|Article|EndDate|StartDate"
    Set VLookRange = TLSheet.Range("F2:F" & TLLastRow)
    For Each singleCell In VLookRange
        singleCell.Value = "=VLOOKUP(E" & singleCell.row & ",Validation!$A$6" & ":$A$" & ValidLastRow & ",1,0)"
    Next

'Drop in match boolean
    TLSheet.Range("G1") = "Match?"
    Set MatchRange = TLSheet.Range("G2:G" & TLLastRow)
    For Each singleCell In MatchRange
        singleCell.Value = "=IF(MATCH(E" & singleCell.row & ",F" & singleCell.row & ",0),TRUE,FALSE)"
    Next

'Autofit TLSheet
    TLSheet.Columns("A:G").AutoFit

'Loop through MatchRange to find errors. Prompt a Message Box if you find a false value
    For Each singleCell In MatchRange
        If singleCell.Value2 <> "True" Then
            MsgBox "Errors found! Please check column G of Sheet1 for FALSE or #N/A values"
            Exit Sub
        Else
        End If
    Next

'Success message!
    MsgBox "Sweet! Everything looks good! Check out column G of Sheet1 for confirmation."


'Check if PSI work is needed
    For Each WS In TLWB.Worksheets
        If WS.Name = "VIF PSI" Then
            UserChoice = MsgBox("Looks like there is VIF PSI work. Would you like to validate that? Note: This query could take a few minutes to run.", vbYesNo)
        'If the user selects yes
            If UserChoice = vbYes Then
                Call Validate_PSI_VIF_Temp_Listings
        'If the user selects no
            Else
                MsgBox "Please validate VIF PSI work manually."
            End If
        Else
        End If
    Next

End Sub

Private Sub Validate_PSI_VIF_Temp_Listings()
'Brian Combs - 2021

'Variables
    Dim TLWB As Workbook
    Dim ValidSheet As Worksheet
    Dim PSISheet As Worksheet
    Dim PSILastRow As Long
    Dim AllPSILastRow As Long
    Dim PSIVars As Range
    Dim singleCell As Range
    Dim MatchRange As Range
    Dim VLookRange As Range
    
'Set sheets
    Set TLWB = ActiveWorkbook
    Set ValidSheet = TLWB.Worksheets("Validation")
    Set PSISheet = TLWB.Worksheets("VIF PSI")
    
'Check for articles in the PSI 1 column
    If PSISheet.Range("A2") = "" Then
        MsgBox "There are no PSI 1 articles in column A. This macro only checks for PSI 1 configuration. Please validate PSI 2 and PSI 3 articles manually."
        Exit Sub
    Else
    End If
    
'SAP STUFF--------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON

'Navigate to SQ01
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nsq01"
    session.FindById("wnd[0]").sendVKey 0

'Environment
    session.FindById("wnd[0]/mbar/menu[5]/menu[0]").Select

'Query areas
    session.FindById("wnd[1]/usr/radRAD1").Select

'Standard area
    session.FindById("wnd[1]/tbar[0]/btn[2]").press

'Other user group
    session.FindById("wnd[0]/tbar[1]/btn[19]").press

'ZMD > Choose
    session.FindById("wnd[1]/usr/cntlGRID1/shellcont/shell").SelectedRows = "0"
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

'ARTICLECHARVAL > Execute
    session.FindById("wnd[0]/usr/cntlGRID_CONT0050/shellcont/shell").SelectedRows = "0"
    session.FindById("wnd[0]/tbar[1]/btn[8]").press

'Enter Characteristic -PRODUCTSOURCEINDICATOR, Class Type - 026, Characteristic Value - 1
    session.FindById("wnd[0]/usr/txtSP$00003-LOW").Text = "PRODUCTSOURCEINDICATOR"
    session.FindById("wnd[0]/usr/ctxtSP$00004-LOW").Text = "026"
    session.FindById("wnd[0]/usr/txtSP$00005-LOW").Text = "1"
    session.FindById("wnd[0]/usr/txtSP$00005-LOW").SetFocus
    session.FindById("wnd[0]/usr/txtSP$00005-LOW").CaretPosition = 1
    session.FindById("wnd[0]/tbar[1]/btn[8]").press

'Press the Export button
    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
    session.FindById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").SelectContextMenuItem "&PC"

'In the clipboard
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus

'Press green check
    session.FindById("wnd[1]/tbar[0]/btn[0]").press

'Paste into valid sheet
    ValidSheet.Select
    ValidSheet.Range("AA1").PasteSpecial

'Press exit
    session.FindById("wnd[0]/tbar[0]/btn[15]").press
    session.FindById("wnd[0]/tbar[0]/btn[15]").press
    session.FindById("wnd[0]/tbar[0]/btn[15]").press

'Call ENDSAPCON
    EndSAPCON

'END SAP STUFF--------------------------------------------------------------

'Autofit ValidSheet
    ValidSheet.Columns("AA:AF").AutoFit
    
'Count rows in PSI table
    ValidSheet.Select
    AllPSILastRow = 9
    Do While ValidSheet.Range("AB" & AllPSILastRow).Value <> ""
        AllPSILastRow = AllPSILastRow + 1
    Loop
    AllPSILastRow = AllPSILastRow - 1
        
'Count variants on VIF PSI sheet
    PSISheet.Select
    PSILastRow = 2
    Do While PSISheet.Range("A" & PSILastRow).Value <> ""
        PSILastRow = PSILastRow + 1
    Loop
    PSILastRow = PSILastRow - 1

'Drop in some formulas in Column I to find a match between Validation sheet key and TL sheet key
    PSISheet.Range("D1") = "Articles with PSI 1"
    Set VLookRange = PSISheet.Range("D2:D" & PSILastRow)
    For Each singleCell In VLookRange
        singleCell.Value = "=VLOOKUP(A" & singleCell.row & ",Validation!$AB$9" & ":$AB$" & AllPSILastRow & ",1,0)"
    Next

'Drop in match boolean
    PSISheet.Range("E1") = "Match?"
    Set MatchRange = PSISheet.Range("E2:E" & PSILastRow)
    For Each singleCell In MatchRange
        singleCell.Value = "=IF(MATCH(D" & singleCell.row & ",A" & singleCell.row & ",0),TRUE,FALSE)"
    Next

'Loop through MatchRange to find errors. Prompt a Message Box if you find a false value
    On Error Resume Next
    For Each singleCell In MatchRange
        If singleCell.Value2 <> "True" Then
            MsgBox "Errors found! Please check column E of VIF PSI worksheet for FALSE or #N/A values"
            Exit Sub
        Else
        End If
    Next
    On Error GoTo 0

'Autofit ValidSheet
    PSISheet.Columns("D:E").AutoFit

'Success message!
    MsgBox "Sweet! Everything looks good! Check out column E of VIF PSI worksheet for confirmation."
    
'Check if PSI 2 or PSI 3 articles are on the request. As of July 2021, we are keeping all VIF articles at PSI 1.
    If PSISheet.Range("B2") <> "" Or PSISheet.Range("C2") <> "" Then
        MsgBox "This macro only checks articles with PSI 1 configuration. Please manually validate PSI 2 and PSI 3 articles."
    Else
    End If

End Sub

Private Sub Validate_Asst_Create()
'Brian Combs - 2021

'Note about what macro does - coming soon!

'Variables
    Dim AWB As Workbook
    Dim singleCell As Range
    Dim WS As Worksheet
    
'ValidSheet variables
    Dim ValidSheet As Worksheet
    Dim AsstWildRange As Range
    Dim BigListCell As Range
    Dim BigListRange As Range
    Dim BigListLastRow As Long
    Dim WRSZRange As Range
    Dim WRSZLastRow As Long
    Dim AsstNumRange As Range
    Dim AsstNumLastRow As Long
    Dim AsstNum As String
    Dim AsstName As String
    Dim VLookRange As Range
    Dim MatchRange As Range
    
'MDSheet variables
    Dim MDSheet As Worksheet
    Dim URL As String
    Dim EachCompletedNote As String
    Dim MDCurrentRow As Long

'HCSplitSheet variables
    Dim HCSplitSheet As Worksheet
    Dim HCSplitName As String
    Dim HCLastRow As Long
    Dim HCAsstIdRange As Range
    Dim AsstNameRange As Range
    Dim AllStoreCell As String
    Dim Counter As Integer
    Dim StoreSplitter As Variant
    Dim CurrentRow As Long
    Dim NumOfStores As Long

'ACSheet variables
    Dim ACSheet As Worksheet
    Dim ACLastRow As Long
    Dim ACAsstNameRange As Range
    Dim ACAsstCell As Range
    Dim ACAsstName As String
    Dim ACAsstNum As String
    Dim ACNumOfStores As String
    
'Turn off display alerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
'Set workbooks and worksheets
    Set AWB = ActiveWorkbook
    Set ACSheet = AWB.Worksheets("assortment create")
    Set MDSheet = AWB.Worksheets("MD")
    
'Find the Hot/Cold split worksheet (I don't love it... but this worksheet name is hard to find)
    For Each WS In AWB.Worksheets
        If WS.Range("A1").Value = "Assortment Create Template Naming Convention: DXX_ASSTC_SYY_Version Example: D36_ASSTC_S14_2" Then
            HCSplitName = WS.Name
        End If
    Next
    Set HCSplitSheet = AWB.Worksheets(HCSplitName)
  
'Count rows on HCSplitSheet
    HCLastRow = 3
    Do While HCSplitSheet.Range("D" & HCLastRow).Value <> ""
        HCLastRow = HCLastRow + 1
    Loop
    HCLastRow = HCLastRow - 1

'Set Assortment Name and HCAsstId Range
    Set AsstNameRange = HCSplitSheet.Range("D3:D" & HCLastRow)
    Set HCAsstIdRange = HCSplitSheet.Range("C3:C" & HCLastRow)
    
'Delete the Valid sheet if it already exists
    On Error Resume Next
    Set ValidSheet = AWB.Worksheets("Validation")
    If Not ValidSheet Is Nothing Then
        ValidSheet.Delete
    Else
    End If
    On Error GoTo 0
    
'Create a Valid Worksheet
    AWB.Sheets.Add.Name = "Validation"
    Set ValidSheet = AWB.Worksheets("Validation")

'Column headers on ValidSheet
    ValidSheet.Range("G1").Value = "Big List"
    
    ValidSheet.Range("E4").Value = "AsstWildcard"
    ValidSheet.Range("F4").Value = "Assortment"
    ValidSheet.Range("G4").Value = "Description"
    ValidSheet.Range("H4").Value = "Site"
    ValidSheet.Range("I4").Value = "Request Assortment|Site"
    ValidSheet.Range("J4").Value = "SAP Assortment|Site"
    ValidSheet.Range("K4").Value = "Match?"
    
    ValidSheet.Range("F5").Value = "----------"
    ValidSheet.Range("G5").Value = "----------"
    ValidSheet.Range("H5").Value = "----------"
    ValidSheet.Range("I5").Value = "----------"
    ValidSheet.Range("J5").Value = "----------"
    ValidSheet.Range("K5").Value = "----------"

'Set Assortment wildcard range
    For Each singleCell In HCAsstIdRange
        ValidSheet.Range("E" & singleCell.row + 2).Value = singleCell.Value & "*"
    Next
    
    Set AsstWildRange = ValidSheet.Range("E5:E" & HCLastRow + 2)

'SAP STUFF-----------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON

'ZSE16N > WRST Query Stuff
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
    session.FindById("wnd[0]").sendVKey 0
'Enter WRST into Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WRST"
    session.FindById("wnd[0]/usr/ctxtGD-TAB").CaretPosition = 4
    session.FindById("wnd[0]").sendVKey 0
'Press Assortment - More (Arros)
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").press
'Press Delete All Entries
    session.FindById("wnd[1]/tbar[0]/btn[34]").press
'Press Upload from Clipboard
    AsstWildRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
 'Press Transfer Data (Execute)
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'Press Description - More (Arrows)
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,3]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,3]").press
'Press Delete All Entries
    session.FindById("wnd[1]/tbar[0]/btn[34]").press
'Press Upload from Clipboard
    AsstNameRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
'Press Transfer Data (Execute)
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'Press Online (Execute)
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
'Press Export
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
'In the clipboard
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
'Press back (Green arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'Press back (Green arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press

'END SAP STUFF--------------------------------------------------------------

'Paste into Validation sheet
    ValidSheet.Select
    ValidSheet.Range("A1").Select
    ValidSheet.Range("A1").PasteSpecial
    
'Run "Text to Columns" on ValidSheet
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    On Error GoTo 0
    
'Count rows of Assortment Numbers
    AsstNumLastRow = 6
    Do While ValidSheet.Range("C" & AsstNumLastRow).Value <> ""
        AsstNumLastRow = AsstNumLastRow + 1
    Loop
    AsstNumLastRow = AsstNumLastRow - 1

'Set Assortment Number Range and paste into Sorted sheet
    Set AsstNumRange = ValidSheet.Range("C6:C" & AsstNumLastRow)
    
'Loop through Assortment Name Range and assign variables
    CurrentRow = 6
    For Each singleCell In AsstNameRange
        AsstName = singleCell.Value
        AllStoreCell = HCSplitSheet.Range("F" & singleCell.row)
        NumOfStores = HCSplitSheet.Range("E" & singleCell.row).Value
        ValidSheet.Range("B" & singleCell.row + 3).Value = NumOfStores
        
    'Make a big list of all Assortment Names and Stores (by splitting AllStoreCell at the " ")
        StoreSplitter = Split(AllStoreCell, " ")
        For Counter = 0 To NumOfStores - 1
            ValidSheet.Range("G" & CurrentRow).Value = AsstName
            ValidSheet.Range("H" & CurrentRow).Value = StoreSplitter(Counter)
            CurrentRow = CurrentRow + 1
        Next
    Next

'Count rows of BigList
    BigListLastRow = 6
    Do While ValidSheet.Range("G" & BigListLastRow).Value <> ""
        BigListLastRow = BigListLastRow + 1
    Loop
    BigListLastRow = BigListLastRow - 1

'Set BigList
    Set BigListRange = ValidSheet.Range("G6:G" & BigListLastRow)

'Loop through Assortment Number Range
    For Each singleCell In AsstNumRange
        AsstNum = singleCell.Value
        AsstName = Trim(ValidSheet.Range("D" & singleCell.row))
        'Loop through BigList
            For Each BigListCell In BigListRange
            'If the Asst Name in the Big List matches current Asst Name, assign Asst Number to it
                If BigListCell.Value = AsstName Then
                    ValidSheet.Range("F" & BigListCell.row).Value = AsstNum
                End If
            Next
    Next
    
'Loop through BigListRange and create an Assortment/Site key
    For Each BigListCell In BigListRange
        ValidSheet.Range("I" & BigListCell.row).Value = "=TEXTJOIN(""|"",TRUE,F" & BigListCell.row & ",H" & BigListCell.row & ")"
    Next

'SAP STUFF-----------------------------------------------------------------
    
'Navigate to ZSE16N
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
    session.FindById("wnd[0]").sendVKey 0
'Enter WRSZ into Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WRSZ"
'Clear Maximum no. of hits
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").SetFocus
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").CaretPosition = 0
'Hit Enter
    session.FindById("wnd[0]").sendVKey 0
'Press Assortment - More (Arrows)
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").press
'Press (Delete) All Entries
    session.FindById("wnd[1]/tbar[0]/btn[34]").press
'Press Upload From Clipboard
    AsstNumRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
'Press Transfer Data (Execute)
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'Press Online (Execute)
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
'Export
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
'In the clipboard
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
'Green check
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
'Press back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'Press back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press

'Call ENDSAPCON
    EndSAPCON
    
'END SAP STUFF--------------------------------------------------------------

'Paste into ValidSheet
    ValidSheet.Range("M1").PasteSpecial
    
'Run "Text to Columns" on ValidSheet
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    On Error GoTo 0

'Count rows on WRSZ table
    WRSZLastRow = 6
    Do While ValidSheet.Range("N" & WRSZLastRow).Value <> ""
        WRSZLastRow = WRSZLastRow + 1
    Loop
    WRSZLastRow = WRSZLastRow - 1

'Set WRSZ Range
    Set WRSZRange = ValidSheet.Range("N6:N" & WRSZLastRow)

'Loop through WRSZ Range and create assortment/site key
    For Each singleCell In WRSZRange
        ValidSheet.Range("AE" & singleCell.row).Value = "=TEXTJOIN(""|"",TRUE,N" & singleCell.row & ",P" & singleCell.row & ")"
    Next

'Drop in some formulas in Column J to find a match between Requst key and SAP key
    Set VLookRange = ValidSheet.Range("J6:J" & BigListLastRow)
    For Each singleCell In VLookRange
        singleCell.Value = "=VLOOKUP(I" & singleCell.row & ",Validation!$AE$6" & ":$AE$" & WRSZLastRow & ",1,0)"
    Next

'Drop in match boolean
    Set MatchRange = ValidSheet.Range("K6:K" & BigListLastRow)
    For Each singleCell In MatchRange
        singleCell.Value = "=IF(MATCH(I" & singleCell.row & ",J" & singleCell.row & ",0),TRUE,FALSE)"
    Next

'Auto Fit Columns
    ValidSheet.Columns("A:AE").AutoFit

'Count rows on Asst Create sheet
    ACLastRow = 3
    Do While ACSheet.Range("D" & ACLastRow).Value <> ""
        ACLastRow = ACLastRow + 1
    Loop
    ACLastRow = ACLastRow - 1

'Set Assortment Name Range
    Set ACAsstNameRange = ACSheet.Range("D3:D" & ACLastRow)

'Loop through AC Asst Name Range and create closure notes for the MD Sheet
    MDCurrentRow = 5
    For Each ACAsstCell In ACAsstNameRange
        ACAsstName = ACAsstCell.Value
        ACNumOfStores = ACSheet.Range("E" & ACAsstCell.row).Value
    'Find the name and number on the valid sheet
        For Each singleCell In AsstNumRange
            If Trim(ValidSheet.Range("D" & singleCell.row).Value) = Trim(ACAsstName) Then
                ACAsstNum = singleCell.Value
            End If
        Next
    'Create closure notes
        EachCompletedNote = ACAsstNum & " - " & ACAsstName & " with " & ACNumOfStores & " sites/stores"
        MDSheet.Range("C" & MDCurrentRow).Value = EachCompletedNote
        MDCurrentRow = MDCurrentRow + 1
    Next

'Loop through MatchRange to find errors. Prompt a Message Box if you find a false value
    For Each singleCell In MatchRange
        If singleCell.Value2 <> "True" Then
            MsgBox "Errors found! Please check column K of Validation worksheet for FALSE or #N/A values"
            Exit Sub
        Else
        End If
    Next

'Success message!
    MsgBox "Sweet! Everything looks good! Check out column K of Validation worksheet for confirmation."

'Turn on display alerts
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

Sub Validate_Asst_Maint()
'Brian Combs - 2021

'Note about what macro does

'Variables
    Dim AWB As Workbook
    Dim MDSheet As Worksheet
    Dim ValidSheet As Worksheet
    Dim singleCell As Range
    Dim WS As Worksheet
    Dim lastRow As Long
    Dim CurrentRow As Long
    
'Other variables
    Dim AMsheet As Worksheet
    Dim AddStoreCell As Range
    Dim RemStoreCell As Range

'Hot/Cold split variables
    Dim AsstNum As String
    Dim HCSingleCell As Range
    Dim HCSplitName As String
    Dim HCSplitSheet As Worksheet
    Dim HCAsstNumRange As Range
    Dim HCAsstNameRange As Range

'Valid sheet variables
    Dim StoreSplitter As Variant
    Dim AsstNumRange As Range
    Dim AsstNameRange As Range
    Dim Counter As Integer
    Dim AddListRange As Range
    Dim RemListRange As Range
    Dim WRSZRange As Range
    Dim AddVLookRange As Range
    Dim RemVLookRange As Range
    Dim AddMatchRange As Range
    Dim RemMatchRange As Range
    Dim ErrorsFound As Boolean
    Dim RemDateRange As Range
    Dim AddDateRange As Range
    
'Turn off display alerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
'Set workbooks and worksheets
    Set AWB = ActiveWorkbook
    Set AMsheet = AWB.Worksheets("assortment maintain")
    Set MDSheet = AWB.Worksheets("MD")
    
'Find the Hot/Cold split worksheet (I don't love it... but this worksheet name is hard to find)
    For Each WS In AWB.Worksheets
        If WS.Range("C1").Value = "Assortment Maintain Template Naming Convention: DXX_ASSTM_SYY_Version Example: D36_ASSTM_S14_2" Then
            HCSplitName = WS.Name
        End If
    Next
    Set HCSplitSheet = AWB.Worksheets(HCSplitName)
  
'Count rows on HCSplitSheet
    lastRow = 3
    Do While HCSplitSheet.Range("C" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

'Set HCAsstNumRange and HCAsstNameRange
    Set HCAsstNumRange = HCSplitSheet.Range("C3:C" & lastRow)
    Set HCAsstNameRange = HCSplitSheet.Range("D3:D" & lastRow)
    
'Delete the Valid sheet if it already exists
    On Error Resume Next
    Set ValidSheet = AWB.Worksheets("Validation")
    If Not ValidSheet Is Nothing Then
        ValidSheet.Delete
    Else
    End If
    On Error GoTo 0
    
'Create a Valid Worksheet
    AWB.Sheets.Add.Name = "Validation"
    Set ValidSheet = AWB.Worksheets("Validation")

'Column headers on ValidSheet
    ValidSheet.Range("G1").Value = "Add List - Note: #N/A values are BAD for adds"
    ValidSheet.Range("M1").Value = "Remove List - Note: #N/A values are GOOD for removes"
    
    ValidSheet.Range("G4").Value = "Assortment"
    ValidSheet.Range("H4").Value = "Site"
    ValidSheet.Range("I4").Value = "Request Add Assortment|Site"
    ValidSheet.Range("J4").Value = "SAP Assortment|Site"
    ValidSheet.Range("K4").Value = "Match?"
    
    ValidSheet.Range("M4").Value = "Assortment"
    ValidSheet.Range("N4").Value = "Site"
    ValidSheet.Range("O4").Value = "Request Remove Assortment|Site"
    ValidSheet.Range("P4").Value = "SAP Assortment|Site"
    ValidSheet.Range("Q4").Value = "Match?"
    
    ValidSheet.Range("G5").Value = "----------"
    ValidSheet.Range("H5").Value = "----------"
    ValidSheet.Range("I5").Value = "----------"
    ValidSheet.Range("J5").Value = "----------"
    ValidSheet.Range("K5").Value = "----------"
    ValidSheet.Range("M5").Value = "----------"
    ValidSheet.Range("N5").Value = "----------"
    ValidSheet.Range("O5").Value = "----------"
    ValidSheet.Range("P5").Value = "----------"
    ValidSheet.Range("Q5").Value = "----------"

'SAP STUFF-----------------------------------------------------------------

'Call SAPCON (from SAPConnector_v012 Module)
    SAPCON

'ZSE16N > WRST Query Stuff
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
    session.FindById("wnd[0]").sendVKey 0
'Enter WRST into Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WRST"
    session.FindById("wnd[0]/usr/ctxtGD-TAB").CaretPosition = 4
    session.FindById("wnd[0]").sendVKey 0
'Press Assortment - More (Arros)
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,2]").press
'Press Delete All Entries
    session.FindById("wnd[1]/tbar[0]/btn[34]").press
'Press Upload from Clipboard
    HCAsstNumRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
 'Press Transfer Data (Execute)
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'Press Online (Execute)
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
'Press Export
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
'In the clipboard
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
'Press back (Green arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'Press back (Green arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press

'END SAP STUFF--------------------------------------------------------------

'Paste into Validation sheet
    ValidSheet.Select
    ValidSheet.Range("A1").Select
    ValidSheet.Range("A1").PasteSpecial
    
'Run "Text to Columns" on ValidSheet
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    On Error GoTo 0

'Count rows of Assortment Numbers
    lastRow = 6
    Do While ValidSheet.Range("C" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

'Set Assortment Number Range
    Set AsstNumRange = ValidSheet.Range("C6:C" & lastRow)
    Set AsstNameRange = ValidSheet.Range("D6:D" & lastRow)
    
'Create add list--------------------

'Loop through Assortment Name Range and assign variables
    CurrentRow = 6
    For Each singleCell In HCAsstNumRange
        AsstNum = singleCell.Value
        Set AddStoreCell = HCSplitSheet.Range("E" & singleCell.row)
    'Make a big list of all Assortment Names and Stores (by splitting AddStoreCell at the " ")
        StoreSplitter = Split(AddStoreCell, " ")
        On Error Resume Next
        For Counter = 0 To UBound(StoreSplitter)
            ValidSheet.Range("G" & CurrentRow).Value = AsstNum
            ValidSheet.Range("H" & CurrentRow).Value = StoreSplitter(Counter)
            CurrentRow = CurrentRow + 1
        Next
        On Error GoTo 0
    Next

'Count rows of AddList
    lastRow = 6
    Do While ValidSheet.Range("G" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
   
'Set LastRow to 6 if Addlist is blank
    If lastRow <= 6 Then
        lastRow = 6
    Else
        lastRow = lastRow - 1
    End If

'Set AddList, AddVlook, and AddMatch Ranges
    Set AddListRange = ValidSheet.Range("G6:G" & lastRow)
    Set AddVLookRange = ValidSheet.Range("J6:J" & lastRow)
    Set AddMatchRange = ValidSheet.Range("K6:K" & lastRow)
    Set AddDateRange = ValidSheet.Range("L6:L" & lastRow)

'Loop through AddDateRange and print END OF TIME
    For Each singleCell In AddDateRange
        singleCell.Value = "END OF TIME"
    Next

'Loop through BigListRange and create an Assortment/Site key
    For Each singleCell In AddListRange
        ValidSheet.Range("I" & singleCell.row).Value = "=TEXTJOIN(""|"",TRUE,G" & singleCell.row & ",H" & singleCell.row & ",L" & singleCell.row & ")"
    Next

'Create remove list--------------------

'Loop through Assortment Name Range and assign variables
    CurrentRow = 6
    For Each singleCell In HCAsstNumRange
        AsstNum = singleCell.Value
        Set RemStoreCell = HCSplitSheet.Range("F" & singleCell.row)
    'Make a big list of all Assortment Names and Stores (by splitting AddStoreCell at the " ")
        StoreSplitter = Split(RemStoreCell, " ")
        On Error Resume Next
        For Counter = 0 To UBound(StoreSplitter)
            ValidSheet.Range("M" & CurrentRow).Value = AsstNum
            ValidSheet.Range("N" & CurrentRow).Value = StoreSplitter(Counter)
            CurrentRow = CurrentRow + 1
        Next
        On Error GoTo 0
    Next

'Count rows of RemList
    lastRow = 6
    Do While ValidSheet.Range("M" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    
'Set LastRow to 6 if Remlist is blank
    If lastRow <= 6 Then
        lastRow = 6
    Else
        lastRow = lastRow - 1
    End If

'Set RemList, RemVlook, and RemMatch ranges
    Set RemListRange = ValidSheet.Range("M6:M" & lastRow)
    Set RemVLookRange = ValidSheet.Range("P6:P" & lastRow)
    Set RemMatchRange = ValidSheet.Range("Q6:Q" & lastRow)
    Set RemDateRange = ValidSheet.Range("R6:R" & lastRow)
    
'Loop through RemDateRange and print REMOVED
    For Each singleCell In RemDateRange
       singleCell.Value = "REMOVED"
    Next
    
'Loop through RemListRange and create an Assortment/Site key
    For Each singleCell In RemListRange
        ValidSheet.Range("O" & singleCell.row).Value = "=TEXTJOIN(""|"",TRUE,M" & singleCell.row & ",N" & singleCell.row & ",R" & singleCell.row & ")"
    Next
    

'SAP STUFF-----------------------------------------------------------------
    
'Navigate to ZSE16N
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nzse16n"
    session.FindById("wnd[0]").sendVKey 0
'Enter WRSZ into Table field
    session.FindById("wnd[0]/usr/ctxtGD-TAB").Text = "WRSZ"
'Clear Maximum no. of hits
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").Text = ""
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").SetFocus
    session.FindById("wnd[0]/usr/txtGD-MAX_LINES").CaretPosition = 0
'Hit Enter
    session.FindById("wnd[0]").sendVKey 0
'Press Assortment - More (Arrows)
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").SetFocus
    session.FindById("wnd[0]/usr/tblSAPLZBC_210_E_SE16NSELFIELDS_TC/btnPUSH[4,1]").press
'Press (Delete) All Entries
    session.FindById("wnd[1]/tbar[0]/btn[34]").press
'Press Upload From Clipboard
    AsstNumRange.Copy
    session.FindById("wnd[1]/tbar[0]/btn[24]").press
'Press Transfer Data (Execute)
    session.FindById("wnd[1]/tbar[0]/btn[8]").press
'Press Online (Execute)
    session.FindById("wnd[0]/tbar[1]/btn[8]").press
'Export
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").PressToolbarContextButton "&MB_EXPORT"
'In the clipboard
    session.FindById("wnd[0]/usr/cntlRESULT_LIST/shellcont/shell").SelectContextMenuItem "&PC"
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
'Green check
    session.FindById("wnd[1]/tbar[0]/btn[0]").press
'Press back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
'Press back (Green Arrow)
    session.FindById("wnd[0]/tbar[0]/btn[3]").press

'Call ENDSAPCON
    EndSAPCON
    
'END SAP STUFF--------------------------------------------------------------

'Paste into ValidSheet
    ValidSheet.Range("S1").PasteSpecial
    
'Run "Text to Columns" on ValidSheet
    On Error Resume Next
        Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1)), TrailingMinusNumbers:=True
    On Error GoTo 0

'Count rows on WRSZ table
    lastRow = 6
    Do While ValidSheet.Range("T" & lastRow).Value <> ""
        lastRow = lastRow + 1
    Loop
    lastRow = lastRow - 1

'Set WRSZ Range
    Set WRSZRange = ValidSheet.Range("T6:T" & lastRow)

'Loop through WRSZ Range and create assortment/site key
    For Each singleCell In WRSZRange
        ValidSheet.Range("AK" & singleCell.row).Value = "=TEXTJOIN(""|"",TRUE,T" & singleCell.row & ",V" & singleCell.row & ",IF(AC" & singleCell.row & "=2958465,""END OF TIME"", ""REMOVED"")" & ")"
    Next

'Drop in some formulas in Column J to find a match between Requst key and WRSZ key
    For Each singleCell In AddVLookRange
        singleCell.Value = "=VLOOKUP(I" & singleCell.row & ",Validation!$AK$6" & ":$AK$" & lastRow & ",1,0)"
    Next

'Drop in match boolean
    For Each singleCell In AddMatchRange
        singleCell.Value = "=IF(MATCH(I" & singleCell.row & ",J" & singleCell.row & ",0),TRUE,FALSE)"
    Next

'Drop in some formulas in Column P to find a match between Requst key and WRSZ key
    For Each singleCell In RemVLookRange
        singleCell.Value = "=VLOOKUP(O" & singleCell.row & ",Validation!$AK$6" & ":$AK$" & lastRow & ",1,0)"
    Next

'Drop in match boolean
    For Each singleCell In RemMatchRange
        singleCell.Value = "=IF(MATCH(O" & singleCell.row & ",P" & singleCell.row & ",0),TRUE,FALSE)"
    Next
    
'Loop through AddMatchRange to find errors. Prompt a Message Box if you find a false value
    ErrorsFound = False 'Default to false
    If ValidSheet.Range("G6") <> "" Then
        For Each singleCell In AddMatchRange
            If ErrorsFound = False Then
                If singleCell.Text <> "TRUE" Then
                'Fill the cell red
                    singleCell.Interior.ColorIndex = 3
                    ErrorsFound = True
                Else
                'Fill the cell green
                    singleCell.Interior.ColorIndex = 4
                End If
            Else
            End If
        Next
    End If
    
'Loop through RemMatchRange to find errors. Prompt a Message Box if you find a false value
    If ValidSheet.Range("M6") <> "" Then
        For Each singleCell In RemMatchRange
            If ErrorsFound = False Then
                If singleCell.Text <> "TRUE" Then
                'Fill the cell red
                    singleCell.Interior.ColorIndex = 3
                    ErrorsFound = True
                Else
                'Fill the cell green
                    singleCell.Interior.ColorIndex = 4
                End If
            Else
            End If
        Next
    End If

'Loop through AsstNameRange and look for a match on HCAsstNameRange to see if descriptions match. Using cell colors as a boolean
    For Each singleCell In AsstNameRange
        For Each HCSingleCell In HCAsstNameRange
            If Trim(singleCell.Text) <> Trim(HCSingleCell.Text) And singleCell.Interior.ColorIndex <> 4 Then
            'If no match found, fill the cell red
                singleCell.Interior.ColorIndex = 3
            Else
            'If match found, fill the cell green
                singleCell.Interior.ColorIndex = 4
            End If
        Next
    Next

'Loop through AsstNameRange again and look for red cells
    For Each singleCell In AsstNameRange
        If ErrorsFound = False Then
            If singleCell.Interior.ColorIndex = 3 Then
                ErrorsFound = True
            End If
        End If
    Next
    
'Auto Fit Columns
    ValidSheet.Columns("A:AK").AutoFit

'Success or fail message
    If ErrorsFound = True Then
        MsgBox "Errors found! Please check columns D, K and Q of Validation worksheet for FALSE or #N/A values"
    ElseIf ErrorsFound = False Then
        MsgBox "Sweet! Everything looks good! Check out columns D, K and Q of Validation worksheet for confirmation."
    End If

'Turn on display alerts
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

'Version notes--------------------------------------------------------------
'v001 - Validates Maintain Promos
'v002 - Validates Temp Listings
'v003 - Validates VIF PSI Indicator (Temp Listings)
'v004 - Validates Assortment Creates
'v005 - Validate Assortment Create bug fix
'v006 - Validates Assortment Maintains
'v007 - Validates Assortment Maintain site removals







