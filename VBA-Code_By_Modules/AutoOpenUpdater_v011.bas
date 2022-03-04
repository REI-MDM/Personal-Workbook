Attribute VB_Name = "AutoOpenUpdater_v011"
Option Explicit
Sub Auto_Open()
'******************************************************************************
' This is intended to simplify the updating of teammates macros and functions
' in use on their personal macro workbook.
'
' MH 03/13/2013
'
' By default it is named "Auto_Open" so that it will run whenever your personal
' macro workbook is opened.  You could easily change this sub name to make it
' into an "on demand" updater.
'
' Here is an outline of the steps this macro takes:
'
' Step 1: Available Modules
' It looks for any files in the "MacroRepository" folder and capture their
'   filename into the "AvailableModules" array. It expects filenames to be
'   in the format of "***_v0##.b" where *** is the module name, and ## is the
'   version number.  It does not do any error checking, and files in the
'   MacroRepository currently must be manually maintained and kept clean.
'--Potential Update--
'   Error check filenames, or only capture filenames that match our desired
'       format
'
' Step 2: Installed Modules
' Look through all installed modules and capture their data in the "Installed
'   Modules" array.  If the module doesn't look "versioned" then it will
'   version it with _v000.
'
' Step 3: Compare Versions, decide if we want to update.
' Next we nest loop through all our installed and available modules.  If we
'   find a pair where the Base Module matches, we compare versions. If the
'   "Available" version is greater than the "Installed" version, we give the
'   User an option to auto-install all updates. If they select no, we ask if
'   they would like to individually choose each update, or make no updates.
'   **Auto Install** - Installs all newer mods, and outputs a list of updated
'       Modules when done.
'   **Choose Each** - Gives the user a dialog box for each updateable module
'       And asks if we should install or not.  Gives a list of updated mods at
'       The finish.
'   **Do Not Update** - Gives the user no more message boxes.
'   **Update Available to this module** - Displays a message box that
'       upgrading will need to be done via the "ImportModule" mod. This always
'       displays if there is an update available.
'
' Step 4: Remove Old Versions
' Remove Old Versions - Look through "InstalledModules" for anything marked
'       "YES" for uninstall. Remove this.
'--Potential Update-- could export a module as a backup if needed
'
' Step 5: Install Updates
' Loop through "AvailableModules" looking for anything marked "YES" for install
'   And then install them, and add them to the list of "UpdatedModules"
'
' Step 6: Re-Check for versioning.
' In case someone made an update, but stripped the versioning off, this will
'       Reinstate the version number.
'--Potential Update-- This will only catch things that are not versioned at
'   all!  It will not catch "Old" Versioning.  This will need to be fixed.
'
' Most of this code is very redundant with "ImportModule" and that is because
' they rely on eachother for updates. Neither can update itself.
' btw this will break at 100 you know...fyi...probably
'******************************************************************************
'This is pointed at our macro folder
Dim file As String
Dim i As Long
Dim NumAvail As Long
Dim AvailableModules() As Variant
Dim MacroRepository As String
Dim TWB As Workbook
'Dim Initialize As Object
Dim user As String


    Set TWB = Workbooks("PERSONAL.XLSB")
    
    user = LCase(Environ("username"))
    
MacroRepository = "C:\Users\" & user & "\RECREATIONAL EQUIPMENT INC\Master Data Management - Shared Documents\Projects\Macro files\Macros\"

'******************** Attempt to initialize network location *****************'
'    Set Initialize = CreateObject("WScript.Network")
'    On Error Resume Next
'    Initialize.MapNetworkDrive "M:", MacroRepository
'    DoEvents
'    On Error GoTo 0
'Really we should check to see if the drive exists now.  Chdir?  If it exists
'Either because we just mapped it or it already existed
'    If Err.Number <> 0 Then
'        'We could pop up a message box if we need to here
'    Else
'        'MsgBox ("Drive Mapped OK")
'    End If
        
'********************* START Capture Available Modules ************************'
    i = 0
    'Simply counting available Modules, perhaps redundant with filling the array
    'On Error GoTo TheEnd:
    file = Dir(MacroRepository)
    'On Error GoTo 0
    If file <> "" Then
        i = 1
    End If
    Do While file <> ""
        file = Dir
        i = i + 1
    Loop
    NumAvail = i - 1
    
    'Size the array

ReDim AvailableModules(1 To NumAvail, 1 To 4)

'               AvailableModules Array looks like this:
'AvailableModules(i, 1) AvailableModules(i,2)   AvailableModules(i,3)   AvailableModules(i,4)'
'Filename               Base Module             Version                 MarkForInstall
'NewMIT_v001.b          NewMIT                  001                     "Yes"/"No"   '

'Now we will fill the array'

    file = Dir(MacroRepository)
    For i = 1 To NumAvail '
    'Filename
        AvailableModules(i, 1) = file
    'Base Macro Name (Everything before the _v###.b)
        AvailableModules(i, 2) = Left(AvailableModules(i, 1), _
            Len(AvailableModules(i, 1)) - 7)
    'Version number (After "_v" but before the ".b"
        AvailableModules(i, 3) = Mid(AvailableModules(i, 1), _
            Len(AvailableModules(i, 2)) + 3, 3)
    'Default to "NO" for mark for Install
        AvailableModules(i, 4) = "NO"
    'Get the next file
        file = Dir
    Next i
    
'********************** END Capture Available Modules ************************'

'********************** START Capture Installed Modules **********************'

Dim InstalledModules() As String
Dim numitems As Long
Dim nummodules As Long
Dim j As Long
Dim CurMod As String
Dim AllVersioned As Boolean

    AllVersioned = True
    
'This count will count sheets and "ThiWorkbook" and possibly userforms
    numitems = TWB.VBProject.VBComponents.Count
    
'Modules return type "1" here.  We will count the Modules
    For i = 1 To numitems
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Or TWB.VBProject.VBComponents.Item(i).Type = 3 Then
        nummodules = nummodules + 1
        End If
    Next i

'Resize the Array (for the first time, so really just size it)

ReDim InstalledModules(1 To nummodules, 1 To 4)
    
'               InstalledModules Array looks like this:
'InstalledModules(i, 1)     InstalledModules(i,2)       InstalledModules(i,3)       InstalledModules(i,4)'
'Module Name                Base Module                 Version                     MarkForRemoval
'NewMIT_v001                NewMIT                      001                         "Yes"/"No"   '
    
'Fill in just ModuleName
    j = 1
    For i = 1 To numitems
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Or TWB.VBProject.VBComponents.Item(i).Type = 3 Then
            CurMod = TWB.VBProject.VBComponents.Item(i).Name
        'Module Name
            InstalledModules(j, 1) = CurMod
        'Default to "NO" for removal
            InstalledModules(j, 4) = "NO"
        'Try to grab BaseName and Version
            If InStr(CurMod, "_v0") <> 0 Then
            'This guy looks versioned
                InstalledModules(j, 2) = Left(CurMod, InStr(CurMod, "_v0") - 1)
                InstalledModules(j, 3) = Right(CurMod, 3)
            Else
            'Not Versioned.  Lets try and ignore it.
                InstalledModules(j, 2) = "N/A"
                AllVersioned = False
            'Not Versioned
            End If
            
'Previously this was versioning unversioned Modules.  But it was alittle buggy
'And would error out when there was already a Module1_v000 and there was a new
'Module1 it was trying to version.  I have left it here in case I want to fix
            
'            If Mid(CurMod, Len(CurMod) - 4, 3) <> "_v0" Then
'            'This item does not look versioned. Let's call it version 000
'            MsgBox (CurMod & " is the current modules name")
'                CurMod = CurMod & "_v000"
'                TWB.VBProject.VBComponents.Item(i).Name = _
'                    CurMod
'            End If
            'Full Mod name
'            InstalledModules(j, 1) = TWB.VBProject. _
'                VBComponents.Item(i).Name
'            'Base Mod name ( - "_v0xx")
'            InstalledModules(j, 2) = Left(InstalledModules(j, 1), _
'                (Len(InstalledModules(j, 1)) - 5))
'            'Version number (last 3 - "0xx")
'            InstalledModules(j, 3) = Right(InstalledModules(j, 1), 3)
'            'default to "NO" for removal

            j = j + 1
        End If
    Next i
    
'Find Unversioned mods and see if we can version them. Rememeber to check for
'Duplicate Base Mods (Module1_v000 and Module1, etc) Oh wait, I have to step
'through numitems to update module names.  Ooops.  Maybe ignore non-versioned
'items?
'    If Not AllVersioned Then
'        For i = 1 To UBound(InstalledModules, 1)
'            If InstalledModules(i, 3) = "" Then
'            CurMod = InstalledModules(i, 1)
'                For j = 1 To UBound(InstalledModules, 1)
'                    If i <> j And InstalledModules(i, 2) = _
'                        InstalledModules(j, 2) Then
'                    'i is the one we are changing
'
'    End If
        
        

'************************ END Capture Installed Modules **********************'

'************************ Start Comparing Versions ***************************'
'************************ Start Mark for Upgrade/Removal**********************'

Dim UpgradeAllAuto As Boolean
Dim ChooseEach As Boolean
Dim SkipUpdates As Boolean
Dim MadeAChoice As Boolean
Dim response As String
Dim Automsg As String
Dim Individmsg As String
Dim ErrorUpdateCurrent As String

UpgradeAllAuto = False
ChooseEach = False
SkipUpdates = False
MadeAChoice = False

Automsg = "Updates have been found!  Would you like to automatically update" & _
    " all possible installed modules?"
Individmsg = "Choosing yes will prompt you at each upgradeable macro for " & _
    "installation.  Choosing no will skip updates."
ErrorUpdateCurrent = "It looks like there is a newer version of " & _
    "AutoOpenUpdater - the currently running mod.  We cannot update the " & _
    "module while it is running.  You should be able to use the " & _
    "ImportModule macro to update it, however."
    


'Step through all installed modules
    For i = 1 To UBound(InstalledModules)
    'Check to see if we have a matching Module "Available"
        For j = 1 To UBound(AvailableModules)
            If AvailableModules(j, 2) = InstalledModules(i, 2) Then
            'We found a matching module "Available"
                If AvailableModules(j, 3) > InstalledModules(i, 3) Then
                'We found an upgrade!
                    If Not MadeAChoice Then
                    'Prompt the user for desired action
                        response = MsgBox(Automsg, vbYesNo)
                        If response = vbYes Then
                        'Auto upgrade all
                            MadeAChoice = True
                            UpgradeAllAuto = True
                        ElseIf response = vbNo Then
                        'Don't upgrade all. Ask for upgrade selective, or no
                        'upgrades at all.
                            response = MsgBox(Individmsg, vbYesNo)
                            If response = vbYes Then
                            'Upgrade selectively
                                ChooseEach = True
                                MadeAChoice = True
                            ElseIf response = vbNo Then
                            'No upgrades - exit the sub
                                MadeAChoice = True
                                SkipUpdates = True
                                Exit Sub
                            End If
                        End If
                    End If
                    If MadeAChoice And UpgradeAllAuto Then
                    'We're upgrading automatically.  Mark the intalled one for
                    'removal, and the upgradeably one for installation
                        InstalledModules(i, 4) = "YES"
                        AvailableModules(j, 4) = "YES"
                    End If 'Ending the "mark for upgrade - Auto"
                    If Not UpgradeAllAuto And ChooseEach And MadeAChoice Then
                    'Were choosing for each.
                        response = MsgBox("Would you like to upgrade " & _
                            InstalledModules(i, 2) & " from version " & _
                            InstalledModules(i, 3) & " to version " & _
                            AvailableModules(j, 3) & " ?", vbYesNo)
                        If response = vbYes Then
                        'Yes means install. Mark as such
                            InstalledModules(i, 4) = "YES"
                            AvailableModules(j, 4) = "YES"
                        Else
                        'No means do not install.  This is superfluous as
                        'we currently don't check for "NO", only "YES"
                            InstalledModules(i, 4) = "NO"
                            AvailableModules(j, 4) = "NO"
                        End If ' Ending the dialog for "upgrade this one?"
                    End If ' "Ending the Choosing Each" If
                    If AvailableModules(j, 2) = "AutoOpenUpdater" Then
                    'We cannot update the currently running macro. Mark to no
                    'for install new and uninstall current.  Pop a message box
                        InstalledModules(i, 4) = "NO"
                        AvailableModules(j, 4) = "NO"
                        MsgBox (ErrorUpdateCurrent)
                    End If ' Ending current mod is running mod.
                End If 'Ending the "Available Version > Installed Version"
            End If  'Ending the "If we found a match"
        Next j  'Looping through AvailableModules looking for matches
    Next i  'Looping through InstalledModules
    
'************************ End Comparing Versions *****************************'
'************************ End Mark for Upgrade/Removal ***********************'


'************************ Start Remove Installed Modules *********************'
Dim RemovedModules As String

    For i = 1 To UBound(InstalledModules)
        If InstalledModules(i, 4) = "YES" Then
        'This module is marked for removal
            TWB.VBProject.VBComponents.Remove _
            TWB.VBProject.VBComponents.Item(InstalledModules(i, 1))
            RemovedModules = RemovedModules & " " & InstalledModules(i, 2)
        End If
    Next i
'************************ End Remove Installed Modules ***********************'

'************************ Start Install Updates ******************************'
Dim UpdatedModules As String

    For j = 1 To UBound(AvailableModules)
        If AvailableModules(j, 4) = "YES" Then
        'This module is marked for installation
            TWB.VBProject.VBComponents.Import _
                (MacroRepository & AvailableModules(j, 1))
            UpdatedModules = UpdatedModules & " " & AvailableModules(j, 2)
        End If
    Next j
    
'************************ End Install Updates ********************************'
            
'************* Verify that newly installed modules are versioned *************'

    'Step through all items again
    For i = 1 To numitems
        'We only want to look at "Modules"
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Then
            CurMod = TWB.VBProject.VBComponents.Item(i).Name
            'Parse this for the "BaseMod" Name if it is versioned
            If InStr(1, CurMod, "_v0") > 0 Then
                'looks versioned, strip the versioning
                CurMod = Left(CurMod, InStr(1, CurMod, "_v0") - 1)
            End If
            For j = 1 To UBound(AvailableModules)
            'Lets step through our available array looking for this module
                If AvailableModules(j, 2) = CurMod And _
                    AvailableModules(j, 4) = "YES" Then
                    CurMod = CurMod & "_v" & AvailableModules(j, 3)
                    TWB.VBProject.VBComponents.Item(i) _
                        .Name = CurMod
                    'I should probably put escape conditions here, but I am
                    'feeling a little lazy, so just loop through everything
                End If
            Next j 'Looping through AvailableArray
        End If  'Ending "If Module" if
    Next i ' Stepping through next NumItem
            
'****************** Everything should be versioned again! ********************'
    If UpdatedModules <> "" Then
        MsgBox ("We updated the following modules: " & UpdatedModules)
        TWB.Save
    End If

TheEnd:
'******************** UnMap Network drive ************************************'
    'On Error Resume Next
    'Initialize.RemoveNetworkDrive "M:"
    'On Error GoTo 0

'Check our MIT
    Call MitChecker
    
'Do some Michael Stuff
    If LCase(Environ("Username")) = "mhildru" _
        Or LCase(Environ("Username")) = "johubba" Then
        'Sync O Common
            'Call OCommonSync
        'Refresh the MAP template
            Call MAPRefresh
        'get current templates
            'Call GetCurrentTemplates
        'Clean some files out
            'Call cleanDirs
    End If
End Sub
Sub MitChecker()
'******************************************************************************
' Intended to compare your "local" MITs save date with the latest one in our
' Mit Repository.
'
'******************************************************************************
Dim MITRepo As String
Dim LocalRepo As String
Dim CurFile As String
Dim CurMITDate As Date
Dim CurMITName As String
Dim RepoMITDate As Date
Dim RepoMITName As String
Dim response As String
Dim MitType As String
Dim i As Integer
Dim aps As String

'This points to the teamsite location our MIT lives on.

    
    'New Team site sycned to onedrive
    MITRepo = "C:\Users\" & LCase(Environ("username")) & "\RECREATIONAL EQUIPMENT INC\Master Data Management - Shared Documents\Projects\Master Input Templates\"
    'MITRepo = "\\teamsites.rei.com\DavWWWRoot\sites\FinanceDept\mdm\Shared Documents\Projects\Master Input Templates\"
        
'We'll define where to look for our MIT based on if the user typically saves
'things locally (laptoppers) or to a networked drive (default).

'    If InStr(LocalUsers, Environ("Username")) > 0 Then
'    'This is Lex's preference.
'        LocalRepo = "C:\Users\" & Environ("username") & "\Desktop\Master Input Template\"
'    Else
'        LocalRepo = "\\ahqnas1.reicorpnet.com\users\" & Environ("username") & _
'        "\Profile\Desktop\Master Input Template\"
'    End If

'Change it to be based on environment.
    aps = Application.PathSeparator
    LocalRepo = Replace(Environ("UserProfile") & "\Desktop\Master Input Template\", "\", aps)


'Make sure the directory exists
    RecursDir (LocalRepo)
  

'Loop through the Article MIT first then the Vendor MIT

'set counter for loop
i = 1
Do While i < 3

'Set the variable that determines which MIT to check for updates (Article MIT, or Vendor MIT)
    If i = 1 Then
        MitType = "Article_Create_Master_Input_"
    Else
        MitType = "Vendor_MIT_"
    End If
  
            'Check the save date time on our currently saved MIT
            CurMITDate = Now - 1000
           
            CurFile = Dir(LocalRepo)
            Do While CurFile <> ""
                If InStr(CurFile, MitType) >= 1 Then ' Or InStr(CurFile, "Vendor_MIT_") >= 1 Then
                    If FileDateTime(LocalRepo & CurFile) > CurMITDate Then
                        CurMITDate = FileDateTime(LocalRepo & CurFile)
                        CurMITName = CurFile
                    End If
                End If
                CurFile = Dir
            Loop
                        
         'Check the save date time on the MIT in the repository
            RepoMITDate = Now - 1000
            
            CurFile = Dir(MITRepo)
            Do While CurFile <> ""
                If InStr(CurFile, MitType) >= 1 Then
                    If FileDateTime(MITRepo & CurFile) > RepoMITDate Then
                        RepoMITDate = FileDateTime(MITRepo & CurFile)
                        RepoMITName = CurFile
                    End If
                End If
                CurFile = Dir
            Loop
                   
        'Check to see if the Repo one is newer.  If so, offer to copy it over.
                   
            If RepoMITDate > CurMITDate Then
                response = MsgBox("I have found a newer MIT - " & RepoMITName & _
                    " saved on " & RepoMITDate & " - versus yours, - " & CurMITName & _
                    " saved on " & CurMITDate & " - Would you like to update?", vbYesNo)
                If response = vbYes Then
                    FileCopy source:=MITRepo & RepoMITName, _
                    Destination:=LocalRepo & RepoMITName
                Else
                    MsgBox ("Okay.  No worries.  Have a good day.")
                End If
            End If
 'increment counter
 i = i + 1
 Loop
    
End Sub
Sub RecursDir(filePath As String)
'******************************************************************************
' Intended to be used/inlcuded before we try to save a file to a location
' to ensure that the directory exists.
'
' It should start from the relative "root" of the file path and ensure each
' directory out from there to the final directory
'
'******************************************************************************
Dim CurPath As String
Dim EachDir As Variant
Dim Start As Long
Dim i As Long

    If Right(filePath, 1) = "\" Then
        filePath = Left(filePath, Len(filePath) - 1)
    End If
'Grab each directory name by splitting the path on "\"
    EachDir = Split(filePath, "\")
    
'Set our current path - Network locations have "\\" at the start, which will
'cause EachDir(0) and EachDir(1) to be null strings.
    CurPath = EachDir(0)
    If CurPath = vbNullString Then
        If EachDir(0) = vbNullString And _
            EachDir(1) = vbNullString And _
            EachDir(2) = "ahqnas1.reicorpnet.com" And _
            EachDir(3) = "users" Then
        'All these conditions should be true for a network drive path
                CurPath = "\\ahqnas1.reicorpnet.com\users"
                Start = 4
        End If
    Else
        Start = 1
    End If
    
    'Step through and check/create the filepath
    For i = Start To UBound(EachDir)
        CurPath = CurPath & "\" & EachDir(i)
        If Dir(CurPath & "\", vbDirectory) = vbNullString Then
            MkDir CurPath
        End If
    Next i
End Sub
Private Sub OCommonSync()
'******************************************************************************
' Intended to sync up the below listed OCommon and Local directories with the
' most recently save versions of your files.
'
' It does not currently support any folders within the above listed
' directories, only files within the specified ones.
'
' Potentially useful for project work where you may be working with other REI
' people in O:/Common and want to sync up files in the event of O:/Common
' wipes.
'******************************************************************************
Dim fn As String
Dim OCommonPath As String
Dim LocalPath As String
Dim NoOCommon As Boolean
Dim NoLocal As Boolean
Dim OCommonDate As Date
Dim LocalDate As Date
Dim AlreadyChecked As String
Dim FilesSynced As Long
Dim starttime As Date
Dim aps As String
starttime = Now

'Change these directories as you see fit!
    OCommonPath = "O:\Common\" & LCase(Environ("username")) & "\"
    
    aps = Application.PathSeparator
    
    LocalPath = Environ("userprofile") & aps & "Desktop" & aps & LCase(Environ("username")) & aps


'Check if O:\Common path exists
    If Dir(OCommonPath, vbDirectory) = vbNullString Then
        RecursDir OCommonPath
        OCommonPath = OCommonPath & "\"
        NoOCommon = True
    End If
    
'Check if "Local" path exists
    If Dir(LocalPath, vbDirectory) = vbNullString Then
        RecursDir LocalPath
        NoLocal = True
    End If

    If NoLocal And NoOCommon Then
        'Do nothing, no files in either location
    ElseIf NoLocal And Not NoOCommon Then
        'O:\Common existed, but not local, copy files from OCommon to Local
        fn = Dir(OCommonPath)
        Do While fn <> ""
            FileCopy source:=OCommonPath & fn, Destination:=LocalPath & fn
            FilesSynced = FilesSynced + 1
            fn = Dir
        Loop
    ElseIf Not NoLocal And NoOCommon Then
        'Local exists, OCommon did not.  Copy all from Local to OCommon
        fn = Dir(LocalPath)
        Do While fn <> ""
            FileCopy source:=LocalPath & fn, Destination:=OCommonPath & fn
            FilesSynced = FilesSynced + 1
            fn = Dir
        Loop
    ElseIf Not NoLocal And Not NoOCommon Then
        'Both locations existed, compare save dates for all files
    'Check Local for Newer versions
        fn = Dir(LocalPath)
        Do While fn <> ""
            LocalDate = FileDateTime(LocalPath & fn)
            
            On Error Resume Next
            'if the file exists in Ocommon, get that date, otherwise this
            'errors and OCommonDate stays at #00:00:00#
                OCommonDate = FileDateTime(OCommonPath & fn)
            On Error GoTo 0
            
            If LocalDate > OCommonDate Then
                FileCopy source:=LocalPath & fn, Destination:=OCommonPath & fn
                FilesSynced = FilesSynced + 1
            End If
            
            'reset OCommonDate in the event of an "error"
            OCommonDate = #12:00:00 AM#
            fn = Dir
        Loop
    'Check O:Common for newer versions
        fn = Dir(OCommonPath)
        Do While fn <> ""
            OCommonDate = FileDateTime(OCommonPath & fn)
            
            On Error Resume Next
            'if the file exists in Local, get that date, otherwise this
            'errors and LocalDate stays at #00:00:00#
                LocalDate = FileDateTime(LocalPath & fn)
            On Error GoTo 0
            
            If OCommonDate > LocalDate Then
                FileCopy source:=OCommonPath & fn, Destination:=LocalPath & fn
                FilesSynced = FilesSynced + 1
            End If
            
            'reset LocalDate in the event of an "error"
            LocalDate = #12:00:00 AM#
            fn = Dir()
        Loop
    End If
    If FilesSynced > 0 Then
        MsgBox ("It took " & Format(Now - starttime, "hh:mm:ss") & _
            " to sync " & FilesSynced & " files between " & OCommonPath _
            & " and " & LocalPath & ".")
    End If
End Sub
Private Sub MAPRefresh()
Dim repo As String
Dim repofn As String
Dim sp As String
Dim spfn As String

    repo = "\\teamsites.rei.com\DavWWWRoot\merchandising\MSR\Winshuttle Tools\"
    repofn = "MDS_Request_MAP.xlsm"
    sp = "\\teamsites.rei.com\DavWWWRoot\merchandising\Article and Vendor Master Data\Templates for Waypoint\"
    spfn = "MAP_Request_Form.xlsm"
    'MsgBox FileDateTime(repo & repofn)
    'MsgBox FileDateTime(sp & spfn)
    If FileDateTime(repo & repofn) > FileDateTime(sp & spfn) Then
        FileCopy repo & repofn, sp & spfn
    End If

End Sub
Private Sub cleanDirs()
Dim fpArray(1 To 3) As String 'Adjust as necessary
Dim deleteTime As Long
Dim fn As Variant
Dim i As Long
Dim aps As String
Dim cleanCount As Long
Dim starttime As Date
    
    starttime = Now
    
    Toggle "OFF"

    deleteTime = 90 'stuff older than 90 days we'll delete
    aps = Application.PathSeparator
    fpArray(1) = Environ("userprofile") & aps & "Desktop" & aps & "\Temp SAP Work In Progress\" 'Temp SAP
    fpArray(2) = Environ("userprofile") & aps & "Desktop" & aps & "\Master Input Template\Archived Winshuttles\" 'Archived MITs
    fpArray(3) = "C:\Users\" & LCase(Environ("username")) & "\CMITS"
    For i = LBound(fpArray) To UBound(fpArray)
        fn = Dir(fpArray(i))
        Do Until fn = ""
            If FileDateTime(fpArray(i) & fn) < Now - deleteTime Then
                cleanCount = cleanCount + 1
                'MsgBox fn & " is from " & FileDateTime(fpArray(i) & fn) & ". Killing it."
                Kill fpArray(i) & fn
            End If
            fn = Dir
        Loop
    Next i
    
    Toggle "ON"
    'MsgBox ("Killed " & cleanCount & " in " & Format(Now - starttime, "hh:mm:ss"))

    

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


