Attribute VB_Name = "SAPConnector_v012"
Option Explicit

'Below are the variables needed for VBA SAP Scriping. Hierarchy is in the order below.
Public sapGuiAuto As Object 'SAP Program application (think login screen)
Public SAP As Object        'GUI Aplication highest hierarchy
Public connection As Object 'SAP client connection (ECP, ECQ, EQ2, etc)
Public session As Object    'Active session(open window) of connection.


Sub SAPCON()


Dim openSAP As Long         'used to open SAP with command prompt line
Dim myCMD As String         'command prompt line to open SAP
Dim Tcode As String         'used to set session. We want the session to be an unused SAP session so no unsaved data is lost
Dim il As Integer           'counter to help set connection
Dim it As Integer           'counter to help set session
Dim KillNewSess As Boolean  'Not currently functioning - if we open an new session lets close it after so number of SAP sessions is the after script as it is before
Dim response As Integer     'if user has 6 sessions open that are not at manager session tcod, this give them the chance to end sub or to use one of the open sessions
Dim systemID As String
Dim ECPyesno As VbMsgBoxResult
Dim client As String
Dim RanCMD As Boolean
Dim user As String
Dim OpenCons As Object


Set OpenCons = CreateObject("scripting.dictionary")
user = Environ("username")
RanCMD = False
Const SAPpath = "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\"

'Set desired SAP Client.
ECPyesno = MsgBox("Do you want to connect to the ECP client?", vbYesNo)
    If ECPyesno = vbYes Then
        systemID = "ECP"
    Else
        Do Until systemID <> ""
        systemID = UCase(InputBox("What Client do you want to link to?" & vbCrLf & _
                                    "ECP" & vbCrLf & _
                                    "ECQ" & vbCrLf & _
                                    "EQ2" & vbCrLf & _
                                    "ECD", "SAP Client"))
        If systemID = "ECP" Or systemID = "ECQ" Or systemID = "EQ2" Or systemID = "ECD" Then
        Else
        MsgBox "Please enter a valid SAP client from the list provided"
        systemID = ""
        End If
        Loop
    End If
    
'string for shortcut to execute
myCMD = "sapshcut -system=" & systemID & " -CLIENT=100 -LANGUAGE=EN -COMMAND=s000 -TYPE=TRANSACTION -snc_name=p:CN=SID, O=SAP-AG, C=DE-reuse, -t=ZUPC"
KillNewSess = False

'Lets makes things go faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    
    
'Lets set SAP as an Object.
    On Error Resume Next
    Set sapGuiAuto = GetObject("SAPGUI")        'If this errors login screen not open
    Set SAP = sapGuiAuto.GetScriptingEngine     'This sets up ability to run GuiScripting
    
'If SAP is not open Lets launch it.
'Depending on users security settings in SAP they may need to have them click allow for security question.
'If they click allow they shouldn't see that pop up again for furture runs

    If sapGuiAuto Is Nothing Then
        openSAP = Shell(SAPpath & myCMD)
        RanCMD = True
        'activate msgbox for non mdm users so they know what is happening/sap security question
        'MsgBox "We just tried to launch SAP. Click OK once the SAP home screen opens successfully. If SAP asks for permission to log on, check Remember my Decision and click Allow."
        Do Until Not sapGuiAuto Is Nothing
        Set sapGuiAuto = GetObject("SAPGUI")
        Loop
    
        Set SAP = sapGuiAuto.GetScriptingEngine     'This sets up ability to run GuiScripting
    
'If user took a really long time or hit deny for security question
        If sapGuiAuto Is Nothing Then
    
            Application.ScreenUpdating = True
            Application.DisplayStatusBar = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            MsgBox "Sorry, we could not connect to SAP for you. Did you hit Deny on the security question? Please manually log onto SAP and try again."
            Exit Sub
        End If
    
'Loop through open clients and set connection to match desired Client
        
        Do Until Not session Is Nothing
            Set connection = SAP.Children(0)
            Set session = connection.Children(0)
        Loop
    
    End If
On Error GoTo 0


'SAP application variables have been set.
'Lets cycle through the open SAP GUI connections and sessions to set the correct session




'If desired client is not open, open desired client ECP works with cmd line prompt and auto
'log in for primary account
'updated sap won't log into other clients if ecp is open with cmd prompt so we'll use
'openconnection method of SAP
If Not RanCMD Then
    il = 0
    
    Do Until il = SAP.Children.Count
        If UCase(SAP.Children(il + 0).ConnectionString) Like "*REICORPNET.COM*" Then ' looks like a client we care about
            Set connection = SAP.Children(il + 0)
            Set session = connection.Children(0)
            client = session.Info.SystemName
            If Not OpenCons.Exists(client) Then
                OpenCons.Add client, True
            End If
        End If
        il = il + 1
    Loop
    
    If Not OpenCons.Exists(systemID) Then
        
        Select Case systemID
            
                Case Is = "ECP"
                    If Not session Is Nothing Then
                        SAP.OpenConnection ("ECP - Production System")
                    Else
                        openSAP = Shell(SAPpath & myCMD)
                    End If
                    
                Case Is = "ECQ"
                    If Not session Is Nothing Then
                        SAP.OpenConnection ("ECQ - Quality System")
                    Else
                        openSAP = Shell(SAPpath & myCMD)
                    End If
                    
                Case Is = "EQ2"
                    If Not session Is Nothing Then
                        SAP.OpenConnection ("EQ2 - Project System")
                    Else
                        openSAP = Shell(SAPpath & myCMD)
                    End If
                    
                Case Is = "ECD"
                    If Not session Is Nothing Then
                        SAP.OpenConnection ("ECD - Development System")
                    Else
                        openSAP = Shell(SAPpath & myCMD)
                    End If
                    
            
        End Select
        
        'activeate for non mdm users
        'MsgBox "We just tried to launch SAP. Click OK once the SAP home screen opens successfully. If SAP asks for permission to log on, check Remember my Decision and click Allow."
    
'        If OpenCons.Count = 0 Then
'            On Error Resume Next
'            Do Until Not connection Is Nothing
'                Set connection = SAP.Children(0)
'            Loop
'            On Error GoTo 0
'        End If

        On Error Resume Next
        il = 0
        Do Until client = systemID Or il > SAP.Children.Count
            If UCase(SAP.Children(il + 0).ConnectionString) Like "*REICORPNET.COM*" Then ' looks like a client we care about
                Set connection = SAP.Children(il + 0)
                Set session = connection.Children(0)
                client = session.Info.SystemName
            End If
            il = il + 1
            If il > SAP.Children.Count Then il = 0
        Loop
        
        'if we are logging on using openconnection method this will help skip the log in screen.
        'we could make this better for loggin onto different clients or tertiary account in the future.
        If session.Info.Transaction = "S000" Then
            session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = "100"
            session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = user
            session.FindById("wnd[0]").sendVKey 0
        End If
    
        On Error GoTo 0
    Else ' Looks like desired client is already open Loop through all connections to find desired client
        il = 0
        Do Until client = systemID
            If UCase(SAP.Children(il + 0).ConnectionString) Like "*REICORPNET.COM*" Then ' looks like a client we care about
                Set connection = SAP.Children(il + 0)
                Set session = connection.Children(0)
                client = session.Info.SystemName
                il = il + 1
            End If
        Loop
    End If 'is the desired client already open?
End If 'did we run CMD prompt?

'If we are logged onto desired client on another machine log that machine out
If session.ActiveWindow.Text = "License Information for Multiple Logons" Then
    session.FindById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select
    session.FindById("wnd[1]").sendVKey 0
End If
'Loop though open sessions of desired client to either create new session to set to existing home transation

    For it = 0 To connection.Children.Count - 1
        Set session = connection.Children(it + 0)
        Tcode = session.Info.Transaction
        If Tcode <> "SESSION_MANAGER" And it = connection.Children.Count - 1 Then
            If connection.Children.Count <> 6 Then
                session.CreateSession
                Application.Wait (Now + TimeValue("0:00:03"))
                KillNewSess = True
            Else
                response = MsgBox("Looks like you have 6 sessions of ECP open." & vbCrLf & vbCrLf & _
                                    "Click ok if you want to play roulette and lose one of your sessions." & vbCrLf & _
                                    "Click Cancel to chose which session to close, so you don't lose any work.", vbOKCancel)
                If response = 2 Then
                    
                
                    Set sapGuiAuto = Nothing
                    Set SAP = Nothing
                    Set connection = Nothing
                    Set session = Nothing
                    
                    Application.ScreenUpdating = True
                    Application.DisplayStatusBar = True
                    Application.Calculation = xlCalculationAutomatic
                    Application.EnableEvents = True
                    
                    Exit Sub
                Else
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                End If
                
            End If
        Else
        If Tcode = "SESSION_MANAGER" Then
            Exit For
        End If
        End If
    Next


'We now know there is a Session Manager session open
'lets loop through one more time to set session object
    For it = 0 To connection.Children.Count - 1
        Set session = connection.Children(it + 0)
        Tcode = session.Info.Transaction
        If Tcode = "SESSION_MANAGER" Then
            Set session = connection.Children(it + 0)
            Exit For
        End If
    Next
'EndSAPCON
End Sub

Sub EndSAPCON()

    Set sapGuiAuto = Nothing
    Set SAP = Nothing
    Set connection = Nothing
    Set session = Nothing
End Sub
