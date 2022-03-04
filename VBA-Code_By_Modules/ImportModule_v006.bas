Attribute VB_Name = "ImportModule_v006"
Sub ImportModule()
'******************************************************************************
' Select Module files to import.
'
'
'******************************************************************************
Dim InstalledModules() As String          'Info about currently installed mods
Dim numitems As Long                    'Numner of total "items" attached
Dim nummodules As Long                  'Number of modules
Dim i As Long                           'Counter for stepping through items
Dim j As Long                           'Counter - incrementing InstalledModules
Dim CurMod As String                    'helper variable for mod names
Dim RemoveAMod As Boolean               'Boolean to see if we need to remove

Dim SelModules() As String              'Array for holding items to import
Dim Module As Variant                   'Helps in populating above array
Dim Filechosen As Long                  'Ensure selectino for install
Dim MacroRepository As String           'Where the macros be stored at
Dim LMR As Long                         'Helper variable - Len(MacroRepo)

Dim ImportCurrentError As String        'The message we display if you selected
                                        'this mod for import/update
Dim WebModName As Variant
Dim TWB As Workbook                     'This workbook

ImportCurrentError = "Sorry, you cannot install/update the currently " & _
    "running module.  You should be able to install updates to Import " & _
    "Module with the Update module."

'*********************** Gather Info about currently installed mods ***********
    
    
    MacroRepository = "C:\Users\" & LCase(Environ("username")) & "\RECREATIONAL EQUIPMENT INC\Master Data Management - Shared Documents\Projects\Macro files\Macros\"

    'MacroRepository = "\\teamsites.rei.com\DavWWWRoot\sites\FinanceDept\mdm\Shared Documents\Projects\Macro files\Macros\"
    LMR = Len(MacroRepository)
    RemoveAMod = False
    Set TWB = ThisWorkbook

    'This count will include "ThisWorkbook" and any worksheets
    numitems = TWB.VBProject.VBComponents.Count
    For i = 1 To numitems
    'item.type will return "1" for modules (unsure about userforms or classes)
    'and "100" for Worksheets and "ThisWorkbook" I believe.
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Or TWB.VBProject.VBComponents.Item(i).Type = 3 Then
            nummodules = nummodules + 1
        End If
    Next i
    
ReDim InstalledModules(1 To nummodules, 1 To 4)

    j = 1
    For i = 1 To numitems
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Or TWB.VBProject.VBComponents.Item(i).Type = 3 Then
            CurMod = TWB.VBProject.VBComponents.Item(i).Name
            'Module Name
            InstalledModules(j, 1) = CurMod
            'default to "NO" for removal
            InstalledModules(j, 4) = "NO"
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
            
            
'            If Len(CurMod) < 5 Then
'                CurMod = CurMod & "_v000"
'                TWB.VBProject.VBComponents.Item(i).Name = _
'                    CurMod
'            End If
'            If Mid(CurMod, Len(CurMod) - 4, 3) <> "_v0" Then
'            'This item does not look versioned. Let's call it version 000
'                CurMod = CurMod & "_v000"
'                TWB.VBProject.VBComponents.Item(i).Name = _
'                    CurMod
'            End If
            'Full Mod name
'            InstalledModules(j, 1) = TWB.VBProject. _
'                VBComponents.Item(i).Name
            'Base Mod name ( - "_v0xx")
'            InstalledModules(j, 2) = Left(InstalledModules(j, 1), _
'                (Len(InstalledModules(j, 1)) - 5))
            'Version number (last 3 - "0xx")
'            InstalledModules(j, 3) = Right(InstalledModules(j, 1), 3)

'            InstalledModules(j, 4) = "NO"
            j = j + 1
        End If
    Next i
    
'*********************** Allow User to select mods to install *****************
    
    i = 1
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the module(s) to import"
        .InitialFileName = MacroRepository
        .InitialView = msoFileDialogViewSmallIcons
        .Filters.Clear
        .Filters.Add "Sharepoint .bas files", "*.b"
        .Filters.Add "Sharepoint .frm files", "*.f"
    'Excel will take files with any extension and try to add them.
    'However, My versioning/BaseMod parsing is not robust enough to take
    'anything other than a one character file extension.  I have commented
    'out the all files filter.
        '.Filters.Add "All Files", "*.*"
        .ButtonName = "Import"
        .AllowMultiSelect = True
        Filechosen = .Show
        If Filechosen = -1 Then
        'Hooray!  They selected something
            ReDim SelModules(1 To .SelectedItems.Count, 1 To 4)
            For Each Module In .SelectedItems
            'Module here returns a WEB address, not a network address. We will
            'Parse the web address for the text after the last "/"
                WebModName = Split(Module, "\")
                'Full path to the module, including name and extension
                SelModules(i, 1) = MacroRepository & _
                    WebModName(UBound(WebModName))
                'BaseModule - Take off full path, and remove the "_v0xx.b"
                SelModules(i, 2) = Mid(SelModules(i, 1), LMR + 1, _
                    Len(SelModules(i, 1)) - LMR - 7)
                'Version number - Find where "_v0" starts, add two to get past
                'the "_v" and grab the next three characters, the 0xx"
                SelModules(i, 3) = Mid(SelModules(i, 1), InStr(LMR, _
                    SelModules(i, 1), "_v0") + 2, 3)
                'default to "YES" for install
                SelModules(i, 4) = "YES"
                i = i + 1
            Next
        Else
        'Cancelled or closed the box.  Or no files were chosen.  There is
        'nothing more to do.
            MsgBox ("No Files Chosen")
            Exit Sub
        End If
    End With
    
'************* Check to see if we should install/clean anything ***************
    
    For i = 1 To UBound(SelModules)
        For j = 1 To UBound(InstalledModules)
            If SelModules(i, 2) = InstalledModules(j, 2) Then
            'One of the selected items is already installed
                If SelModules(i, 3) > InstalledModules(j, 3) Then
                'Selected item is newer version, remove previous version
                    RemoveAMod = True
                    InstalledModules(j, 4) = "YES"
                ElseIf SelModules(i, 3) = InstalledModules(j, 3) Then
                'Selected item is installed version, do not install
                    SelModules(i, 4) = "NO"
                ElseIf SelModules(i, 3) < InstalledModules(j, 3) Then
                'Installed version is more up to date.  Do nothing for now.
                'If "Export" macro gets moving, possibly suggest that here.
                End If
                'Check to see if we are trying to update "Import Module"
                'Also check if it has been marked for upgrade/removal -
                'only display a message if we actually thought we would install
                'it. Then reset the install/remove
                If SelModules(i, 2) = "ImportModule" And _
                    SelModules(i, 4) = "YES" Then
                    'show an error message
                    MsgBox (ImportCurrentError)
                    'unmark for install
                    SelModules(i, 4) = "NO"
                    'unmark for removal
                    InstalledModules(j, 4) = "NO"
                End If
            End If
          Next j
    Next i
                
'*********************** If we have mods to remove, let's do that. ************

    If RemoveAMod Then
        For j = 1 To UBound(InstalledModules)
            If InstalledModules(j, 4) = "YES" Then
            'we found one we need to remove
                TWB.VBProject.VBComponents.Remove _
                    TWB.VBProject.VBComponents.Item _
                        (InstalledModules(j, 1))
            End If
        Next j
    End If
    
'*********************** Now we can install our selected mods *****************

    For i = 1 To UBound(SelModules)
        If SelModules(i, 4) = "YES" Then
            TWB.VBProject.VBComponents.Import _
                (SelModules(i, 1))
        End If
    Next i
    
'******** Mods should be installed, let's ensure they are versioned ***********

    For i = 1 To numitems
        If TWB.VBProject.VBComponents.Item(i).Type = 1 Or TWB.VBProject.VBComponents.Item(i).Type = 3 Then
            CurMod = TWB.VBProject.VBComponents.Item(i).Name
            If InStr(1, CurMod, "_v0") > 0 Then
            'looks versioned, strip the versioning.
                CurMod = Left(CurMod, InStr(1, CurMod, "_v0") - 1)
            End If
            For j = 1 To UBound(SelModules)
                If CurMod = SelModules(j, 2) And SelModules(j, 4) = "YES" Then
                'Same basemod AND we installed it, - Let's append our filename
                ' version.
                    CurMod = CurMod & "_v" & SelModules(j, 3)
                    TWB.VBProject.VBComponents.Item(i). _
                        Name = CurMod
                End If
            Next j
        End If
    Next i

    MsgBox ("Done installing modules")
    TWB.Save
End Sub
