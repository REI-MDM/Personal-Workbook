Attribute VB_Name = "Winshuttle_Launcher_v002"
Option Explicit
'Verison 0!
' Initial intention is that we use constants for script paths.
' Then have template specific subs for launching scripts with whatever
' checks are needed for them.
' data file should be saved and closed prior to launching the script.
' Note that there is currently no notification system (other than email)
' for when a script finishes.  IT would be tricky to build one in but
' still allow users to interact with excel.
'
' Could eventually be extended to "pick your script for you" based on template and work requested


'**** Change this for testing! ****
Const env = "ECP"
'**** **** **** **** **** **** ****


'Promo
Const Promo_Remove_Update = "http://winshuttle.rei.com/MultipleShuttleFiles/21834138-6931-4282-b852-7e9eb901ae57/V1.0/Promo_FFD_Push_Deactivate_Updates_Removes.Txr"
Const Promo_DOTD = "http://winshuttle.rei.com/MultipleShuttleFiles/33048436-e214-4469-bdc5-248a9bc8431b/V1.0/WAK1_DOTX_Promotion.TxR"

'AM
Const VK11_Retail = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Article_Maintain/VK11_Retail_Change_v11.Txr"
Const Create_Cost = "http://winshuttle.rei.com/MultipleShuttleFiles/c9092174-e8e9-47d8-ba62-52ab5ee09743/V1.0/MEK1_Create_Cost_AM_v9.11.Txr"
Const ZADP_MAP = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Article_Maintain/VK11_ZADP_MIN_ADV_PRC_2.Txr"
Const VPN = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Article_Maintain/MM42_VPN_AM_v9.11.Txr"

'AC
Const Add_Size = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Article_Create/1-MM42_Add_New_Size_11.Txr"

'Vendor
Const Create_Vendor = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Vendor/XK01_Create_Vendor_V1.8.TxR"
Const Maintain_Vendor = "http://winshuttle.rei.com/ShuttleFiles/MasterData/Vendor/XK02_VendorMaintenance_V6.3.TxR"

Sub WS_Pick_Single_Script()
'******************************************************************************
' Allow user to pick any single script in the library
'
'******************************************************************************
Dim AWB As Workbook
Dim WS As Worksheet
Dim sheetname As String
Dim theScript As String
Dim scriptname As String
Dim awbpath As String
Dim response As Long

'modify this for various scripts
'    sheetname = "Maintain_Promo"        'Worksheet to check for
'    theScript = Promo_Remove_Update     'Set to correct full http script path or constant

    '"simple" way of seeing if we are looking at the right template.
    Set AWB = ActiveWorkbook
    
'    On Error Resume Next
'    Set ws = awb.Worksheets(sheetname)
'    On Error GoTo 0
'
'    If ws Is Nothing Then
'        MsgBox ("Aborting:" & vbLf & "Please ensure the template you are " & _
'            "trying to run on has a sheet named: " & vbLf & sheetname)
'        GoTo garbage
'    End If



    'could look at the active workbook to suggest an "initialFileName" to start looking in.
    'but not doing that here.
rePick:
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .InitialFileName = "http://winshuttle.rei.com/ShuttleFiles/MasterData/"
        .Filters.Clear
        'don't allow selection of linked scripts - *.TxR(Link)
        .Filters.Add Description:="Single Scripts", Extensions:="*.TxR"
        
        'if user selected something, capture that
        'when '.show' ends, it returns a 0 for cancel/x'd out, or -1 for something chosen
        If .Show = -1 Then
            theScript = .SelectedItems(1)
        End If
    End With
    
    If theScript = "" Then
        MsgBox ("No script selected.  Aborting")
        GoTo garbage
    End If
    
' Not needed, lock down the filter in the dialog guy
    
'    If InStr(theScript, "(Link") > 0 Then
'        MsgBox "Apologies.  Linked scripts are not currently supported with this method." & _
'             vbLf & vbLf & "You Selected:" & vbLf & theScript
'        response = MsgBox("Would you like to pick a different script?", vbYesNo)
'        If response = vbYes Then
'            GoTo rePick
'        Else
'            GoTo garbage
'        End If
'    End If


    'get a friendly script name
    scriptname = Right(theScript, Len(theScript) - InStrRev(theScript, "/"))
    
    'give the user a chance to review and cancel if desired
    response = MsgBox("Going to launch the script:" & vbLf & scriptname & vbLf & vbLf & "Against the file: " & vbLf & AWB.Name & vbLf & vbLf & "(We'll save and close the file first)", vbOKCancel)
    If response <> vbOK Then GoTo garbage
    
    'once we close awb, we can't reference its fullname anymore, so lets store that.
    awbpath = AWB.fullName
    AWB.Save
    AWB.Close
    
    'Just Launch A Script
    JLAScript script:=theScript, fpath:=awbpath
    
    'Cleanup
garbage:
    On Error Resume Next
    Set AWB = Nothing
    Set WS = Nothing
    On Error GoTo 0
End Sub


Sub WS_Promo_Remove_Update()
'******************************************************************************
' Initial "Template" for script launching. Intended to be copy-pasteable to use
'   for other scripts with minor modifications
' validate we are looking at the right template
' if so, save it, close it, launch a script
'******************************************************************************
Dim AWB As Workbook
Dim WS As Worksheet
Dim sheetname As String
Dim theScript As String
Dim scriptname As String
Dim awbpath As String
Dim response As Long

'modify this for various scripts
    sheetname = "Maintain_Promo"        'Worksheet to check for
    theScript = Promo_Remove_Update     'Set to correct full http script path or constant
    
    '"simple" way of seeing if we are looking at the right template.
    Set AWB = ActiveWorkbook
    On Error Resume Next
    Set WS = AWB.Worksheets(sheetname)
    On Error GoTo 0
    
    If WS Is Nothing Then
        MsgBox ("Aborting:" & vbLf & "Please ensure the template you are " & _
            "trying to run on has a sheet named: " & vbLf & sheetname)
        GoTo garbage
    End If
    
    'get a friendly script name
    scriptname = Right(theScript, Len(theScript) - InStrRev(theScript, "/"))
    
    'give the user a chance to review and cancel if desired
    response = MsgBox("Going to launch the script:" & vbLf & scriptname & vbLf & vbLf & "Against the file: " & vbLf & AWB.Name, vbOKCancel)
    If response <> vbOK Then GoTo garbage
    
    'once we close awb, we can't reference its fullname anymore, so lets store that.
    awbpath = AWB.fullName
    AWB.Save
    AWB.Close
    
    'Just Launch A Script
    JLAScript script:=theScript, fpath:=awbpath
    
    'Cleanup
garbage:
    On Error Resume Next
    Set AWB = Nothing
    Set WS = Nothing
    On Error GoTo 0
End Sub


Private Sub JLAScript(script As String, fpath As String)
'******************************************************************************
' Just launch a script.
' No other "Fanciness"
' needs the dblQT function
'******************************************************************************
Dim AppPath As String
Dim client As String
Dim system As String
Dim eml As String
Dim fullcommand As String
   
    AppPath = "C:\Program Files (x86)\Winshuttle\Studio\Winshuttle.Studio.Console.exe"

    client = "100"
    
    'Application
    AppPath = dblQt(AppPath)
    'Script
    script = " -SapTransaction -run" & dblQt(script)
    'Result File Name/File Path
    fpath = " -rfn" & dblQt(fpath)
    'System - Hard Set currently
    system = " -sys" & dblQt(env)
    'Client - maybe need to adjust for ECD?
    client = " -cid" & dblQt(client)
    'email!
    eml = " -eml" & dblQt(LCase(Environ("Username")) & "@rei.com")
    
    fullcommand = AppPath & script & fpath & system & client & eml & " -dsw" & dblQt("True") 'dsw = "Disable production server warning"
    Debug.Print fullcommand
    Shell (fullcommand)

End Sub

Private Function dblQt(sInput As String) As String
    dblQt = Chr(34) & sInput & Chr(34)
End Function
