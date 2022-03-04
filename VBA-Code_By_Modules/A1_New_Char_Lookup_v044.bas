Attribute VB_Name = "A1_New_Char_Lookup_v044"
Option Explicit

'Customer Site
Const CSRefLoc = "C:\Temp\Color Size Flavor Lens Ref.xlsb"
Const GRefLoc = "G:\SC EVS\Master Data\Automation\Winshuttle Queries\Color Size Flavor Lens Ref.xlsb"
'Const CSRefLoc = "C:\Users\mhildru\Desktop\Temp SAP Work In Progress\Color Size Flavor Lens Ref.xlsb"
'Where Winshuttle drops raw data
'Const WSRefLoc = "G:\IT Master Data & Vendor Compliance\Master Data\Char Reference\Color Size Flavor Lens REF_WS_Blank.xlsm"
'A manually curated list of allowed abbreviations or "misspellings" for colors
Const AbbrevList = "G:\SC EVS\Master Data\Automation\Winshuttle Queries\ColorAbbrevs.txt"

Private Version As String
Private ACSheet As Worksheet
Private AWB As Workbook
Private CSWB As Workbook
Private AllGenDic As Object
Private allowedAbbrevsDic As Object
Private NetezzaSux As Boolean
Private Sub CheckRefDates()
'******************************************************************************
' 07/2019 - Move "lookup location" from G-drive to local computer.  Should be
' marginally better for checking in while teleworking.
' This little bit should check the G drive file and your local file, and copy
' the Gdrive one to your machine if it is newer.
'******************************************************************************
Dim localFile As Date
Dim gFile As Date
On Error Resume Next
    localFile = FileDateTime(CSRefLoc)
    gFile = FileDateTime(GRefLoc)
On Error GoTo 0
    If gFile > localFile Then
        FileCopy GRefLoc, CSRefLoc
    End If
End Sub

Private Sub fillvars()
'set some of our variables we're going to reuse
    
    Set AWB = ActiveWorkbook
    
    On Error Resume Next
    Set ACSheet = AWB.Worksheets("Article Create")
    On Error GoTo 0
    If Not ACSheet Is Nothing Then
    'we found an AC sheet.  Set some stuff..
    
        'validate if our "local" char reference is up to date.
        CheckRefDates
        
        'set our allowed color abbreviations
        If allowedAbbrevsDic Is Nothing Then
            Set allowedAbbrevsDic = CreateObject("scripting.dictionary")
            On Error Resume Next
            popDicFromText dic:=allowedAbbrevsDic, fileloc:=AbbrevList
            On Error GoTo 0
        End If
        
        'start off on our AC sheet
        ACSheet.Select
        Version = ACSheet.Range("H1").Value

    Else
        MsgBox ("Could not find ""Article Create"" sheet. Could be a " & _
        """Bad Template"" or perhaps you hit this accidentally.  Aborting!")

        garbage
        'Turn on some Excel functionality to it initial state
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True

        End
    End If
    
End Sub

Sub A1NewCharLookup()
'******************************************************************************
' Originally written by jheller
' Reworked by mhildru - 1/12/2013
'
' Look up Characteristic Values from our Color Size Reference Sheet.
'
' Initial reworking involves optimizing runtime by eliminating unnecessary
' selection.finds through the use of the UniqueChars array.
'
' Creates a very basic "Checkin" sheet with new colors/sizes and mismatches
'
' Greenifies characteristics that were "NEW" or "Unfound" but now have a code
' Orangifies characteristics where we looked up a different value than what was
' on the template
'
' Return colors/frames with the same description, but other color families on
' the summary sheet.
'
'
' The next steps:
' Overall logic review
' Check for New Values and Mismatches within inital code lookups - if all good
'   then we don't need to create/touch the "Checkin" sheet.
'
' Re-Write!  I realized on the ride in that there could be minor speed and
' major clarity/readability improvements if re-written as such
'
' Step 1: Step through CharArray looking for unique characteristics to fill
'   the UniqueChars array.
' Step 2: Fill codes for all characteristics in UniqueChars by looking up all
'   info on the master tab (to eliminate sheet switching).  Master tab may need
'   slightly different formatting in the format sort data macro to accomodate
'   this. --Master column "I" now has trimmed Characteristic type, name and
'   family all concat. I.E. COLORAPPLERED (Color, Apple, CF. Red) that should
'   be used for such a lookup.--
'   Step 2a: We can set NewPrimary and NewSecondary Bools here.
'   Step 2b: We can fill AltCodes for anything that is unfound in the sheet.
' Step 3: Fill in CharArray with the info we have by searching through data
'   gathered into UniqueChars.
' Step 4: Summary/MD specific stuff can be done.
'
' Even better for 90% of cases - Build a dictionary of all our lookup values
' and build a dictionary of our listed values on the sheet, then compare them.
'
' The limitation is that looking up alternate color families for "APPLE"
' would not work if we key off COLORAPPLERED in our dictionary.
'
'******************************************************************************
Dim lastRow As Long                     'Last row of data on the sheet
Dim CharArray As Variant                'Our main characteristic array
Dim UniqueChars(450, 7) As Variant      'An array containing only unique
                                        'characteristics. Arbitrarily sized.
                                        'Michael doesn't know enough yet to do
                                        'this better.
Dim i As Long                           'Steps through CharArray
Dim j As Long                           'secondary counter - spaces, output
Dim k As Long                           'Steps through UniqueChars,
Dim NumCode1s As Long                   'The number of unique Color,Flavor,
                                        'FrameColor codes
Dim NumCode2s As Long                   'The number of unique Size,Lens color
                                        'codes.
Dim NewValue As Boolean                 'Determines if we are looking up a new
                                        'unique characteristic
Dim NewPrimary As Boolean               'New Primary Characteristic
Dim NewSecondary As Boolean             'New Secondary Characteristic
Dim BadPrimary As Boolean               'Primary Code Mismatch
Dim BadSecondary As Boolean             'Secondary code Mismatch
Dim MDUsers As String                   'List of MD User Usernames
Dim IAmMD As Boolean                    'Are you Master Data?
Dim starttime As Date                   'A timer
Dim cl As Range                         'For use in cleaning and trimming data
Dim Dept As String                      'For storing the department number'
Dim TaskNumber As String                'For Checking for Checkin
Dim Message As String                   'Show a summary message

'Alternate Charactersitic code variables
Dim AltCodes(1 To 150, 1 To 5) As Variant
                                        'This array will hold the data for
                                        'Possible alternate codes. It too
                                        'should not be decalared statically
                                        'but I don't know how to do that.
Dim DescToSearch As String              'For finding alternate codes
Dim FoundCell As Range                  'To verify a find
Dim LastCell As Range                   'Where we found our last match
Dim FirstAddr As String                 'To know when to exit the loop
Dim n As Long                           'Counter for our alt codes

'Check Variables
Dim MDCheckin As Boolean                'If MD, and checking in
Dim UPCSheet As Worksheet               'For UPC checking
Dim FunkyCharSheet As Worksheet        'For FunkyChars
Dim BadUPCs As Boolean                  'For Bad UPCs
Dim FunkyChars As Boolean               'For FunkyChars
Dim BadHTSCodes As Boolean              'For BadHTS
Dim toCTRLQ As Boolean                  'Should we do ctrl-q macro?
'Other junk - Globalize all the things when I get some time to re-write
Dim pfusheet As Worksheet
Dim PFULast As Long
Dim AVSheet As Worksheet                'for Add Variant Checking



'turn off some Excel functionality so your code runs faster
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    starttime = Now
    
    fillvars
        
    MDUsers = Environ("username") '"nmokry phamilt mhildru avega wikeist kbrewer"
    
    MDCheckin = False
    
    If InStr(UCase(MDUsers), UCase(Environ("username"))) > 0 Then
        IAmMD = True
        toCTRLQ = True
        TaskNumber = ACSheet.Range("I8").Value
        If IsNumeric(Right(TaskNumber, 6)) Then
            MDCheckin = False
        Else
            MDCheckin = True
        End If
        
        starttime = Now
        Call CheckForFUs 'cannot check AVs here because codes are not present/"fix
        Set pfusheet = AWB.Worksheets("PFUs")
        ACSheet.activate
        ACSheet.Range("V6").Value = "CheckforFUs RunTime"
        ACSheet.Range("W6").Value = _
            Format(Now() - starttime, "hh:mm:ss")
        'reset starttime for charlookup
        starttime = Now
    Else
        IAmMD = False
        toCTRLQ = False
    End If
    
    NumCode1s = 1
    NumCode2s = 1
    NewPrimary = False
    NewSecondary = False
    BadPrimary = False
    BadPrimary = False
    
'Grab Department number - as of 2013-08-21 this is only used for D45 to not
'replace "LARGE" size with "L" for sleeping bags.
    Dept = Range("J2").Value
    
    
'Find the last row of data on the workbook - key off the "Generic Description" column
'which is what the ctrl-q uses
    lastRow = ACSheet.Range("G" & ACSheet.Rows.Count).End(xlUp).row
    Do Until ACSheet.Range("G" & lastRow).Value <> ""
        lastRow = lastRow - 1
    Loop
        

'Copy data to array and grab three extra columns for OrigChar1 and OrigChar2
'and oldart
    CharArray = ACSheet.Range("I11:Q" & lastRow).Value
 
'Open the Ref Sheet Read only and unhide columns
    If IAmMD Then
        'CheckForNewerRef
    End If
    
    'this is janky - sometimes excel 2013 doesn't want to open our csref
    On Error Resume Next
    Application.DisplayAlerts = False
    Workbooks.Open CSRefLoc, ReadOnly:=True
    Set CSWB = Workbooks("Color Size Flavor Lens Ref.xlsb")
    i = 1
    Do Until Not CSWB Is Nothing Or i = 4
        'Application.Wait Now + TimeValue("0:00:01") 'maybe wait a second before trying again?
        Workbooks.Open CSRefLoc, ReadOnly:=True
        Set CSWB = Workbooks("Color Size Flavor Lens Ref.xlsb")
        i = i + 1
    Loop
    On Error GoTo 0
    
    If CSWB Is Nothing Then
        MsgBox ("We had some error accessing the color size reference " & _
            "sheet - Please retry.  If you continue to get errors... uh, " & _
            "give up and go home early?")
        End
    End If
    Application.DisplayAlerts = True
    'end of jankiness
    
    
'Find old articles that were created with the old size table.  We could do this
'in our chararray, but maybe lets drop char profile onto our AC sheet why not?'
    
'fill our old generic dictionary with data from the color size reference sheet
    FillAllGen
    
    For i = 1 To (lastRow - 10)
        
    'Check if we are looking up SIZE or SIZE_SAP
        If AllGenDic.Exists(ACSheet.Range("E" & 10 + i).Value) Then
            'We have Data on this!  Validate Char profile
            'This doesn't work for Frame Lens though.  Oops.  Needs to be more better
            If UCase(AllGenDic(ACSheet.Range("E" & 10 + i).Value)(1)) <> _
                UCase(ACSheet.Range("I" & 10 + i).Value) Then
                'we have a char profile mismatch!
                'Color it red.  Log PFU? I forget how to log.  Color Red
                'Maybe PFU in the AV Checker? by comparing "CR" and "I"?
                ACSheet.Range("I" & 10 + i).Interior.Color = 255
            End If
            
            'plop the second characteristic into our lookup array
            CharArray(i, 9) = AllGenDic(ACSheet.Range("E" & 10 + i).Value)(3)
            ACSheet.Range("CS" & 10 + i).Value = AllGenDic(ACSheet.Range("E" & 10 + i).Value)(1)
        Else
            'Generic was not recorded in Early 2018, assume it is color/size_SAP for size lookup
            'and that the char profile on the sheet is correct.  But it could be wrong if it is Color Only
            CharArray(i, 9) = "SIZE_SAP"
        End If
    Next i

    
'erase this dictionary, we are done with it
    Set AllGenDic = Nothing
    
    CSWB.Sheets("Color").Select
    Range("A:F").EntireColumn.Hidden = False
    CSWB.Sheets("Size").Select
    Range("A:F").EntireColumn.Hidden = False
    CSWB.Sheets("Flavor").Select
    Range("A:F").EntireColumn.Hidden = False
    CSWB.Sheets("Lens Color").Select
    Range("A:F").EntireColumn.Hidden = False
    CSWB.Sheets("Frame").Select
    Range("A:F").EntireColumn.Hidden = False
    
On Error GoTo ErrHandler:

'CharArray looks like this, populated with data from template columns I:P originally
'
'(i, 1)         (i,2)       (i, 3)  (i, 4)  (i, 5)      (i, 6)  (i, 7)      (i, 8)      (i, 9)
'Char Profile   Color,F,F   Code1   CF      Size/Lens   Code2   OrigCode1   OrigCode2   SIZE or SIZE_SAP
'
'
'UniqueChars looks almost the same, but starts empty and there are two notable
'differences
'
'(k, 1) = Code1 type (FRAME_COLOR_C,FLAVOR,COLOR) & Descriptor & CF if
'applicable for example, "FLAVORTeriyaki" or "COLORStoneGray" or
'"FRAME_COLOR_CStoneGray"
'
'(k, 7) = Code2 type (LENS_COLOR_S, SIZE_SAP,SIZE) & Descriptor
' LENS_COLOR_CBlue or SIZEXL or SIZE_SAPXS
'

'** Start the i loop! **
    For i = 1 To (lastRow - 10)
        'Trim Characteristics
        CharArray(i, 1) = WorksheetFunction.Proper(CharArray(i, 1))
        CharArray(i, 2) = WorksheetFunction.Trim(UCase(CharArray(i, 2)))
        CharArray(i, 4) = WorksheetFunction.Trim(UCase(CharArray(i, 4)))
        CharArray(i, 5) = WorksheetFunction.Trim(UCase(CharArray(i, 5)))
        'Copy original values for mismatch checking
        CharArray(i, 7) = CharArray(i, 3)
        CharArray(i, 8) = CharArray(i, 6)
        
' Begin Searching for Charactersitic code 1 (Primary)

'*** FLAVOR ***
        If InStr(CharArray(i, 1), "Flavor") Then
        'Assume we will have to look it up on the reference sheet unless we
        'find it in UniqueChars
            NewValue = True
        'Search UniqueChars first
            For k = 1 To NumCode1s
                If "FLAVOR" & CharArray(i, 2) = UniqueChars(k, 1) Then
                    NewValue = False
                    CharArray(i, 3) = UniqueChars(k, 3)
                    Exit For
                End If
            Next k
            If NewValue Then
                NumCode1s = NumCode1s + 1 'Add another unique value
                UniqueChars(NumCode1s, 1) = "FLAVOR" & CharArray(i, 2)
                'Search on Flavor Tab
                Sheets("Flavor").Select
                Range("B:B").Select
                Selection.Find(What:=CharArray(i, 2), After:=ActiveCell, _
                    LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).activate
                CharArray(i, 3) = ActiveCell.Offset(0, 1).Value
                UniqueChars(NumCode1s, 2) = CharArray(i, 2) 'Flavor Descriptor
                UniqueChars(NumCode1s, 3) = CharArray(i, 3) 'Flavor Code
            End If
                            
'*** FRAME ***
        ElseIf InStr(CharArray(i, 1), "Frame") Then
        'pet peeve of mine - misspelled FUCHSIA
            CharArray(i, 2) = Replace(CharArray(i, 2), "FUSCHIA", "FUCHSIA", , , vbTextCompare)
        'Color Family
            If LCase(CharArray(i, 4)) = "neutral/kh" Then
                CharArray(i, 4) = "NEUTRAL/KHAKI"
            End If
            If LCase(CharArray(i, 4)) = "khaki/neutral" Then
                CharArray(i, 4) = "NEUTRAL/KHAKI"
            End If
            NewValue = True
            For k = 1 To NumCode1s
                If "FRAME_COLOR_C" & CharArray(i, 2) & CharArray(i, 4) = _
                    UniqueChars(k, 1) Then
                        NewValue = False
                        CharArray(i, 2) = UniqueChars(k, 2) 'Spaced descriptor
                        CharArray(i, 3) = UniqueChars(k, 3) 'Code
                        Exit For
                End If
            Next k
            If NewValue Then
                NumCode1s = NumCode1s + 1
                UniqueChars(NumCode1s, 1) = "FRAME_COLOR_C" & CharArray(i, 2) _
                    & CharArray(i, 4)
                'Search on Frame Tab
                Sheets("Frame").Select
                Range("B:B").Select
                Selection.Find(What:=CharArray(i, 2) & CharArray(i, 4), _
                    After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False).activate
                CharArray(i, 3) = ActiveCell.Offset(0, 2).Value
            'Add in leading spaces
                If ActiveCell.Offset(0, 6) <> 0 Then '5
                    For j = 1 To ActiveCell.Offset(0, 6)
                        CharArray(i, 2) = " " & CharArray(i, 2)
                    Next j
                End If
                UniqueChars(NumCode1s, 2) = CharArray(i, 2)
                UniqueChars(NumCode1s, 3) = CharArray(i, 3)
                UniqueChars(NumCode1s, 4) = CharArray(i, 4)
            End If
            
'*** COLORS ***
        ElseIf InStr(CharArray(i, 1), "Color") And InStr(CharArray(i, 1), _
            "Lens") = 0 Then
        'pet peeve of mine - misspelled FUCHSIA
            CharArray(i, 2) = Replace(CharArray(i, 2), "FUSCHIA", "FUCHSIA", , , vbTextCompare)
        'slash spacing
            CharArray(i, 2) = Replace(CharArray(i, 2), " /", "/", , , vbTextCompare)
            CharArray(i, 2) = Replace(CharArray(i, 2), "/ ", "/", , , vbTextCompare)
        'Color Family
            If LCase(CharArray(i, 4)) = "neutral/kh" Then
                CharArray(i, 4) = "NEUTRAL/KHAKI"
            End If
            If LCase(CharArray(i, 4)) = "khaki/neutral" Then
                CharArray(i, 4) = "NEUTRAL/KHAKI"
            End If
            NewValue = True
            For k = 1 To NumCode1s
                If "COLOR" & CharArray(i, 2) & CharArray(i, 4) = _
                    UniqueChars(k, 1) Then
                    NewValue = False
                    CharArray(i, 3) = UniqueChars(k, 3) ' code
                    CharArray(i, 2) = UniqueChars(k, 2) ' spaced desc.
                    Exit For
                    End If
            Next k
            If NewValue Then
                NumCode1s = NumCode1s + 1
                UniqueChars(NumCode1s, 1) = "COLOR" & CharArray(i, 2) & _
                    CharArray(i, 4)
            'Search on Color Tab
                Sheets("Color").Select
                Range("B:B").Select
                Selection.Find(What:=CharArray(i, 2) & CharArray(i, 4), _
                    After:=ActiveCell, LookIn:=xlValues, LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                    MatchCase:=False, SearchFormat:=False).activate
                CharArray(i, 3) = ActiveCell.Offset(0, 2).Value
                UniqueChars(NumCode1s, 3) = CharArray(i, 3)
                'Add in leading spaces
                If ActiveCell.Offset(0, 6) <> 0 Then
                    For j = 1 To ActiveCell.Offset(0, 6)
                        CharArray(i, 2) = " " & CharArray(i, 2)
                    Next j
                End If
                UniqueChars(NumCode1s, 2) = CharArray(i, 2)
                UniqueChars(NumCode1s, 4) = CharArray(i, 4)
            End If
        Else   'If there is no characteristic to be found fill in N/A
            CharArray(i, 3) = "N/A"
        End If
   
'Search for Size and Lens - Code2s

'*** LENS ***
        If InStr(CharArray(i, 1), "Lens") Then
            NewValue = True
            For k = 1 To NumCode2s
                If "LENS_COLOR_S" & CharArray(i, 5) = UniqueChars(k, 7) Then
                    NewValue = False
                    CharArray(i, 6) = UniqueChars(k, 6)
                    Exit For
                End If
            Next k
            If NewValue Then
                NumCode2s = NumCode2s + 1
                UniqueChars(NumCode2s, 7) = "LENS_COLOR_S" & CharArray(i, 5)
                UniqueChars(NumCode2s, 5) = CharArray(i, 5)
           'Search on Lens Tab
               Sheets("Lens Color").Select
               Range("B:B").Select
                Selection.Find(What:=CharArray(i, 5), After:=ActiveCell, _
                    LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).activate
                CharArray(i, 6) = ActiveCell.Offset(0, 1).Value
                UniqueChars(NumCode2s, 6) = CharArray(i, 6)
            End If
        
'*** SIZE ***
        ElseIf InStr(CharArray(i, 1), "Size") Then
        ' Replace common sizing differences
            If Dept <> "45" Then
            'Deptartment 45, Sleeping bags, wants "LARGE" sleeping bags, so we will leave
                '"LARGE"s as is.
                If CharArray(i, 5) = "XSmall" Then CharArray(i, 5) = "XS"
                If CharArray(i, 5) = "X Small" Then CharArray(i, 5) = "XS"
                If CharArray(i, 5) = "Small" Then CharArray(i, 5) = "S"
                If CharArray(i, 5) = "SML" Then CharArray(i, 5) = "S"
                If CharArray(i, 5) = "SM" Then CharArray(i, 5) = "S"
                If CharArray(i, 5) = "Medium" Then CharArray(i, 5) = "M"
                If CharArray(i, 5) = "Med" Then CharArray(i, 5) = "M"
                If CharArray(i, 5) = "Large" Then CharArray(i, 5) = "L"
                If CharArray(i, 5) = "LG" Then CharArray(i, 5) = "L"
                If CharArray(i, 5) = "LRG" Then CharArray(i, 5) = "L"
                If CharArray(i, 5) = "XLarge" Then CharArray(i, 5) = "XL"
                If CharArray(i, 5) = "X Large" Then CharArray(i, 5) = "XL"
            End If
            NewValue = True
            For k = 1 To NumCode2s
                If CharArray(i, 9) & CharArray(i, 5) = UniqueChars(k, 7) Then
                    NewValue = False
                    CharArray(i, 6) = UniqueChars(k, 6) 'Code
                    Exit For
                End If
            Next k
            If NewValue Then
                NumCode2s = NumCode2s + 1
                UniqueChars(NumCode2s, 5) = CharArray(i, 5) 'Size/Lens desc.
                UniqueChars(NumCode2s, 7) = CharArray(i, 9) & CharArray(i, 5)

        'Search on Size Tab
                If CharArray(i, 9) = "SIZE" Then
            'If it is an old variant search old sizes
                    Sheets("Master").Select
                    ActiveSheet.ListObjects("MasterTable").Range.AutoFilter Field:=1, Criteria1:="SIZE"
                Else
                    Sheets("Size").Select
                End If
                Range("B:B").Select
                Selection.Find(What:=CharArray(i, 5), After:=ActiveCell, _
                    LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, MatchCase:=False, _
                    SearchFormat:=False).activate
                If IsNumeric(ActiveCell.Offset(0, 1).Value) Then
                    CharArray(i, 6) = ActiveCell.Offset(0, 1).Value
                Else
                    CharArray(i, 6) = ""
                End If
                UniqueChars(NumCode2s, 6) = CharArray(i, 6) 'Size/Lens Code
            End If
        Else   'If there is no characteristic to be found fill in N/A
            CharArray(i, 6) = "N/A"
        End If
        'Check for a Primary mismatch on this iteration
        If CharArray(i, 3) <> CharArray(i, 7) And CharArray(i, 7) < _
            999999 And CharArray(i, 7) > 1 Then
            BadPrimary = True
        End If
        'Check for a Secondary mismatch on this iteration
        If CharArray(i, 6) <> CharArray(i, 8) And CharArray(i, 8) < _
            999999 And CharArray(i, 8) > 1 Then
            BadSecondary = True
        End If
    Next i
    
'Now lets check our UniqueChars array to see if we have any "new" items

    For k = 1 To WorksheetFunction.Max(NumCode1s, NumCode2s)
        If UniqueChars(k, 2) <> "" And UniqueChars(k, 3) = "" Then
            NewPrimary = True
        End If
        If UniqueChars(k, 7) <> "" And UniqueChars(k, 6) = "" Then
            NewSecondary = True
        End If
    Next k
    
'And NOW let's check to see if we have any "same name different color family"
'characteristics.
    
    If NewPrimary Then
        For k = LBound(UniqueChars, 1) To NumCode1s
        'check to see if we have found a code
            If UniqueChars(k, 3) = "" Then
            'Loop through the rest of the array searching for a duplicate
                For i = k + 1 To NumCode1s
                    If UniqueChars(k, 2) = UniqueChars(i, 2) And _
                        UniqueChars(i, 3) = "" Then
                    'prepend a space to the "second" one
                        UniqueChars(i, 2) = " " & UniqueChars(i, 2)
                    End If '/End descs match, i has no code
                Next i 'loop through the rest of the unique chars
            End If '/end checking if the code is still blank
        Next k 'Primary loop through unique chars
    End If '/End check if new primaries
    
'Now, if we have new charactersistics, lets look for alternates for Colors
'and Frames.  Maybe we have an "Unassigned" or a neighboring color family
'that would work.  We only need to do this for Primary Characteristics,
'as we should not have duplicates for secondaries.
    
    n = 1
    
'First we will step through UniqueChars looking for colors, then come back
'and look for Frames

'*** Alt Color Codes ***
If NewPrimary Then
    Sheets("Color").Select
    For k = 1 To WorksheetFunction.Max(NumCode1s, NumCode2s)
        If Left(UniqueChars(k, 1), 5) = "COLOR" And UniqueChars(k, 3) = "" Then
            DescToSearch = UniqueChars(k, 2)
            With Range("C:C")
                Set LastCell = .Cells(.Cells.Count)
            End With
            Set FoundCell = Range("C:C").Find(What:=DescToSearch, _
                LookIn:=xlValues, LookAt:=xlWhole, After:=LastCell)
    
            If Not FoundCell Is Nothing Then
                FirstAddr = FoundCell.Address
                AltCodes(n, 1) = "COLOR"
                AltCodes(n, 2) = FoundCell.Value
                AltCodes(n, 3) = FoundCell.Offset(0, 1).Value
                AltCodes(n, 4) = FoundCell.Offset(0, 4).Value
                AltCodes(n, 5) = FoundCell.Offset(0, 5).Value
                n = n + 1
            End If
            Do Until FoundCell Is Nothing
                Set FoundCell = Range("C:C").FindNext(After:=FoundCell)
                If FoundCell.Address = FirstAddr Then
                    Exit Do
                End If
                AltCodes(n, 1) = "COLOR"
                AltCodes(n, 2) = FoundCell.Value
                AltCodes(n, 3) = FoundCell.Offset(0, 1).Value
                AltCodes(n, 4) = FoundCell.Offset(0, 4).Value
                AltCodes(n, 5) = FoundCell.Offset(0, 5).Value
                n = n + 1
            Loop
        End If
    Next k

'*** Alt Frame codes ***
    Sheets("Frame").Select
    For k = 1 To NumCode1s
        If Left(UniqueChars(k, 1), 13) = "FRAME_COLOR_C" And _
            UniqueChars(k, 3) = "" Then
            DescToSearch = UniqueChars(k, 2)
            With Range("C:C")
                Set LastCell = .Cells(.Cells.Count)
            End With
            Set FoundCell = Range("C:C").Find(What:=DescToSearch, _
                    LookIn:=xlValues, LookAt:=xlWhole, After:=LastCell)

            If Not FoundCell Is Nothing Then
                FirstAddr = FoundCell.Address
                AltCodes(n, 1) = "FRAME_COLOR_C"
                AltCodes(n, 2) = FoundCell.Value
                AltCodes(n, 3) = FoundCell.Offset(0, 1).Value
                AltCodes(n, 4) = FoundCell.Offset(0, 4).Value
                AltCodes(n, 5) = FoundCell.Offset(0, 5).Value
                n = n + 1
            End If
            Do Until FoundCell Is Nothing
                Set FoundCell = Range("C:C").FindNext(After:=FoundCell)
                If FoundCell.Address = FirstAddr Then
                    Exit Do
                End If
                AltCodes(n, 1) = "FRAME_COLOR_C"
                AltCodes(n, 2) = FoundCell.Value
                AltCodes(n, 3) = FoundCell.Offset(0, 1).Value
                AltCodes(n, 4) = FoundCell.Offset(0, 4).Value
                AltCodes(n, 5) = FoundCell.Offset(0, 5).Value
                n = n + 1
            Loop
        End If
    Next k
End If

'error handling code
ErrHandler:
    Range("B1").Select
Resume Next


    'Close the Ref Sheet and activate the template
    Workbooks("Color Size Flavor Lens Ref.xlsb").Close Savechanges:=False
    
    'Write data back to template and highlight missing values
    
    'Make sure our size column will be text formatted, so "9/10" comes out as
    '"9/10", and not as date "10-Sep" or number/date "41527"
    ACSheet.Range("M11:M" & lastRow).NumberFormat = "@"
    ACSheet.Range("I11:N" & lastRow) = CharArray
    ACSheet.Range("K11:K" & lastRow & ",N11:N" & lastRow).SpecialCells _
        (xlCellTypeBlanks).Select
    With Selection
        .Value = "Unfound"
        .Interior.ColorIndex = 6
    End With
    
'Greenify New values
    For i = 1 To (lastRow - 10)
    'Greenify Primaries
        If CharArray(i, 3) > 1 And CharArray(i, 3) < 999999 And _
            (CharArray(i, 7) < 1 Or CharArray(i, 7) > 999999) Then
        'Hopefully this means that the original value was not numeric/between 1 and 999999,
        'And our lookup value is a number between 1 and 999999.
            AWB.Sheets("Article Create").Range("K" & 10 + i).Interior.Color = 65280
        End If
    'Greenify Secondaries
        If CharArray(i, 6) > 1 And CharArray(i, 6) < 999999 And _
            (CharArray(i, 8) < 1 Or CharArray(i, 8) > 999999) Then
            AWB.Sheets("Article Create").Range("N" & 10 + i).Interior.Color = 65280
        End If
    Next i

'Populate the fact that we checked colors on the PFUSheet
    PFULast = pfusheet.Range("a65536").End(xlUp).row + 1
    pfusheet.Range("A" & PFULast).Value = "ColorCheck"
    pfusheet.Range("B" & PFULast).Value = BadPrimary Or BadSecondary Or NewPrimary Or NewSecondary
    If pfusheet.Range("B" & PFULast) = True Then
        pfusheet.Range("C" & PFULast).Value = "Check the ColorCheck Sheet for new colors or mismatches"
    End If

    'Drop size sort data into column CK
    Call jamInSizeQuery

'****************************************************************************************
'Master Data Section - could be a separate Sub, but I am not smart enough yet to pass
'arrays around.
'****************************************************************************************
'
' If the user is Master Data, let's produce a summary sheet.
' If this isn't buggy, we could roll out a "Checkin Sheet" into the template
' and prettyify it, and leave it hidden except for us.
If BadPrimary Or BadSecondary Or NewPrimary Or NewSecondary Then
    If IAmMD Then
    
    'Declare some variables on MD Cares about
        Dim ColorSheet As Worksheet             'An Output sheet for MD users
        Dim ErrorThisIteration As Boolean       'To check if we should increment j
        Dim NewCharCounter As Long              'How many new characteristics do we have?
        
        
        toCTRLQ = False
        NewCharCounter = 0
        BadUPCs = False
        FunkyChars = False
        
        Message = "There is data on the ColorCheck sheet."
    'Select/create Checkin sheet
        On Error Resume Next
        Set ColorSheet = AWB.Sheets("ColorCheck")
        On Error GoTo 0
        If Not ColorSheet Is Nothing Then
            ColorSheet.Cells.ClearContents
            ColorSheet.Select
        Else
            AWB.Sheets.Add().Name = "ColorCheck"
            Set ColorSheet = AWB.Sheets("ColorCheck")
            ColorSheet.Select
        End If
        
    'Column Headers - These should be prettied and more descriptive if we roll
    'this into the MA's view
        Range("A1").Value = "CharType"
        Range("B1").Value = "Possible new Primary"
        Range("D1").Value = "Color Family"
        Range("E1").Value = "2nd CharType"
        Range("F1").Value = "Possible new Secondary"
        Range("H1").Value = "Line"
        Range("I1").Value = "Primary Descriptor"
        Range("J1").Value = "Lookup Code"
        Range("K1").Value = "Primary C.F"
        Range("L1").Value = "OrigCode"
        Range("M1").Value = "Secondary Descriptor"
        Range("N1").Value = "Lookup Code"
        Range("O1").Value = "Orig Code"
        Range("Q1").Value = "Quick_Color_Code"
        Range("Q2").Value = "First New Color -->"
        Range("V2").Value = "First New Lens -->"

    'Ensure formatting for size columns is "text"
        Range("F:F").NumberFormat = "@"
        Range("M:M").NumberFormat = "@"
    'New Primary Characteristics?
        i = 2
        For k = 1 To NumCode1s
            If UniqueChars(k, 1) <> "" And UniqueChars(k, 3) = "" Then
                If Left(UniqueChars(k, 1), 5) = "COLOR" Then
                    Range("A" & i).Value = "COLOR"
                ElseIf Left(UniqueChars(k, 1), 13) = "FRAME_COLOR_C" Then
                    Range("A" & i).Value = "FRAME_COLOR_C"
                ElseIf Left(UniqueChars(k, 1), 6) = "FLAVOR" Then
                    Range("A" & i).Value = "FLAVOR"
                End If
                Range("B" & i).Value = UniqueChars(k, 2)
                Range("C" & i).Value = UniqueChars(k, 3)
                Range("D" & i).Value = UniqueChars(k, 4)
                NewCharCounter = NewCharCounter + 1
                NewPrimary = True
                i = i + 1
            End If
        Next k
        'Call our spellchecker!
        SpellCheck rng:=Range("B2:B" & i)
        
    'New Secondary Characteristics?
        j = 2
        For k = 1 To NumCode2s
            If UniqueChars(k, 7) <> "" And UniqueChars(k, 6) = "" Then
                If Left(UniqueChars(k, 7), 8) = "SIZE_SAP" Then
                    Range("E" & j).Value = "SIZE_SAP"
                ElseIf Left(UniqueChars(k, 7), 12) = "LENS_COLOR_S" Then
                    Range("E" & j).Value = "LENS_COLOR_S"
                ElseIf Left(UniqueChars(k, 7), 4) = "SIZE" Then
                    Range("E" & j).Value = "SIZE (old)"
                End If
                Range("F" & j).Value = UniqueChars(k, 5)
                NewCharCounter = NewCharCounter + 1
                NewSecondary = True
                j = j + 1
            End If
        Next k
        
    'Now we can output some  alt codes if we found them
        If n > 1 Then
            j = WorksheetFunction.Max(i, j) + 2
            Range("A" & j).Value = "Alternate codes"
            j = j + 1
            Range("A" & j).Value = "CharType"
            Range("B" & j).Value = "Description"
            Range("C" & j).Value = "Code"
            Range("D" & j).Value = "C.F."
            Range("E" & j).Value = "NumSpaces"
            j = j + 1
            For i = 1 To n
                Range("A" & j).Value = AltCodes(i, 1)
                Range("B" & j).Value = AltCodes(i, 2)
                Range("C" & j).Value = AltCodes(i, 3)
                Range("D" & j).Value = AltCodes(i, 4)
                Range("E" & j).Value = AltCodes(i, 5)
                j = j + 1
            Next i
        End If
    
    'Lets now step through CharArray and do some mismatch callouts
        j = 2
        For i = 1 To (lastRow - 10)
            ErrorThisIteration = False
        'Check if there is a primary mismatch
            If CharArray(i, 3) <> CharArray(i, 7) And CharArray(i, 7) < _
                999999 And CharArray(i, 7) > 1 Then
                Range("H" & j).Value = i + 10
                Range("I" & j).Value = CharArray(i, 2)
                Range("J" & j).Value = CharArray(i, 3)
                Range("K" & j).Value = CharArray(i, 4)
                Range("L" & j).Value = CharArray(i, 7)
                BadPrimary = True
                ErrorThisIteration = True
            'Highlight the cell on the template
                AWB.Sheets("Article Create").Range("K" & 10 + i).Interior.Color = 49407
            End If
        'Check for a secondary mismatch
            If CharArray(i, 6) <> CharArray(i, 8) And CharArray(i, 8) < _
                999999 And CharArray(i, 8) > 1 Then
                Range("H" & j).Value = i + 10
                Range("M" & j).Value = CharArray(i, 5)
                Range("N" & j).Value = CharArray(i, 6)
                Range("O" & j).Value = CharArray(i, 8)
                BadSecondary = True
                ErrorThisIteration = True
            'Highlight the error on the original template
                AWB.Sheets("Article Create").Range("N" & 10 + i).Interior.Color = 49407
            End If
            If ErrorThisIteration Then
                j = j + 1
            End If
        Next i
        Cells.Select
        Cells.EntireColumn.AutoFit
        
    'Hide unused columns
        If Not NewPrimary And n = 1 Then
            Range("A:D").EntireColumn.Hidden = True
        End If
        If Not NewSecondary And n = 1 Then
            Range("E:F").EntireColumn.Hidden = True
        End If
        If Not BadPrimary And Not BadSecondary Then
            Range("H:O").EntireColumn.Hidden = True
        ElseIf Not BadPrimary Then
            Range("I:L").EntireColumn.Hidden = True
        ElseIf Not BadSecondary Then
            Range("M:O").EntireColumn.Hidden = True
        End If
    End If '/MD Users
End If '/booleans true This determined if we needed to make a "Color" sheet
    

'and once again, reselct AC, just for funsies
    ACSheet.activate
    ACSheet.Range("V1:W8").Font.ColorIndex = 2
    ACSheet.Range("W1:W2").NumberFormat = "hh:mm:ss;@"
    ACSheet.Range("V1").Value = "CharLookup StartTime"
    ACSheet.Range("W1").Value = starttime
    ACSheet.Range("V2").Value = "CharLookup RunTime"
    ACSheet.Range("W2").Value = _
        Format(Now() - starttime, "hh:mm:ss")
    
    starttime = Now
    'try to jam in the AV checker
    'If Not NetezzaSux Then
    Call AVChecker
    ACSheet.Range("V5").Value = "AVChecker RunTime"
    ACSheet.Range("W5").Value = _
        Format(Now() - starttime, "hh:mm:ss")
        
        
    On Error Resume Next
    Set UPCSheet = Sheets("UPC Check")
    Set FunkyCharSheet = Sheets("FunkyChars")
    Set AVSheet = AWB.Worksheets("Add Variant Check")
    On Error GoTo 0
    If Not UPCSheet Is Nothing Then
    'UPC Sheet did get set, we output a summary above
        BadUPCs = True
        toCTRLQ = False
        Message = Message & vbLf & "There is data on the UPC Check Sheet"
    End If
    If Not FunkyCharSheet Is Nothing Then
    'FunkyChars had an output
        FunkyChars = True
        toCTRLQ = False
        Message = Message & vbLf & "There is data on the FunkyChars Sheet"
    End If
    If BadHTSCodes Then
        toCTRLQ = False
        Message = Message & vbLf & "There are bad HTS codes.  Check " & _
            "column BN."
    End If
    If Not AVSheet Is Nothing Then
    'UPC Sheet did get set, we output a summary above
        toCTRLQ = False
        Message = Message & vbLf & "There is data on the Add Variant Check Sheet"
    End If
'Now we need to determine what sheets to show:
        On Error Resume Next
        Set ColorSheet = Sheets("ColorCheck")
        On Error GoTo 0

'check our PFU sheet
    toCTRLQ = True
    PFULast = pfusheet.Range("A" & Rows.Count).End(xlUp).row

    If IAmMD Then
        For i = 2 To PFULast
            If pfusheet.Range("B" & i) = True Then
                toCTRLQ = False
                pfusheet.activate
            End If
        Next i
    End If
        
        
  


'****************************************************************************************
'End MD section
'****************************************************************************************

'toCTRLQ starts out as TRUE if IAMMD, and False if not MD.  It will get changed
'to false if there is a color sheet, UPC sheet, FunkyChars sheet, or BadHTSCodes.
    If toCTRLQ Then
    'There was no reason to create a color sheet, and I am MD
        ACSheet.activate
        Call NewMIT
    End If
    
    'Turn on some Excel functionality to it initial state
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    garbage
    
End Sub

Private Sub CheckForFUs()
Dim pfusheet As Worksheet
Dim BadHTSCodes As Boolean

'Checks for R1C1 formatting, and corrects if set

    Call FormatCheck

'check for a PFU sheet and create if necessary
    On Error Resume Next
    Set pfusheet = ActiveWorkbook.Worksheets("PFUs")
    If pfusheet Is Nothing Then
        Sheets.Add().Name = "PFUs"
        Sheets("PFUs").activate
        Set pfusheet = ActiveWorkbook.Worksheets("PFUs")
    Else
        pfusheet.activate
    End If
    On Error GoTo 0

'Clear out/Format PFU
    pfusheet.Cells.ClearContents
    pfusheet.Range("A1").Value = "Checked Item"
    pfusheet.Range("B1").Value = "Looks Bad?"
    pfusheet.Range("C1").Value = "Instructions"

'Do some error checks.
    Call LookForBadChars(ActiveWorkbook)
'Re-selct our AC Sheet
    ACSheet.activate
'Check UPCs - living in GTINChecker module (set to DB)
    Call GTIN_Netezza(AWB)
'Check Vendors are associated with Generics on Add Variant
    Call Vendor_Snow
'Re-selct our AC Sheet
    ACSheet.activate
'Check HTS codes
    BadHTSCodes = BadHTS(ACSheet)
'Even though we should not have switched sheets at all, reselect ACSheet
    ACSheet.activate
'And check for unique articles
    Call UniqueIDer
'Check for Valid Vendors
    Call ValidVendorCheck
'again, ensure AC sheet is active
    ACSheet.activate



End Sub

Private Sub LookForBadChars(TWB As Workbook)
'******************************************************************************
' Intended to be used on customer templates. Can also be used to auto-clean
' common replacements.  It will also call out with red highlighting any funky
' characters it runs across.
'
' Here it is called from Characteristic Lookup
'******************************************************************************
'Dim TWB As Workbook
Dim StartWS As Worksheet
Dim pfusheet As Worksheet
Dim WS As Variant
Dim cl As Variant
Dim CurString As String
Dim CurChar As String
Dim newchar As String
Dim AscCurChar As Long
Dim i As Long
Dim ErrorCount As Long
Dim BadDic As Object
Dim FunkySheet As Worksheet
Dim itm As Variant
Dim ErrorThisIt As Boolean
Dim qt As String
Dim PFULast As Long
Dim starttime As Date

'Log the time.  For funsies
    starttime = Now

'Commented out because called from within char lookup where this is already off
'turn stuff off for faster processing
'    Application.ScreenUpdating = False
'    Application.DisplayStatusBar = False
'    Application.Calculation = xlCalculationManual
'    Application.EnableEvents = False

'Set some initial things
    ErrorCount = 0
    Set StartWS = ActiveSheet
    Set pfusheet = ActiveWorkbook.Worksheets("PFUs")
    qt = Chr(34)
'    Set TWB = ThisWorkbook
    Set BadDic = CreateObject("Scripting.dictionary")
    
'loop through each worksheet. Maybe overkill, but catches all style plan input
'type things, that we normally would not be able to get by selecting text cells
    For Each WS In TWB.Worksheets
        If WS.CodeName = "Article_Create" Or WS.CodeName = "Article_Maintain" _
            Or WS.CodeName = "Style_Plan_Report_Input" Or WS.CodeName = "Lookups" Then
            
            WS.activate
        'grab any text fields we find
            On Error Resume Next
            Selection.SpecialCells(xlCellTypeConstants, xlTextValues).Select
            For Each cl In Selection
        'and loop through each text cell!
            'clean and trim!
                CurString = WorksheetFunction.Clean(WorksheetFunction.Trim(cl.Value))
                If Len(CurString) > 0 Then
                    For i = 1 To Len(CurString)
                        CurChar = Mid(CurString, i, 1)
                        AscCurChar = Asc(CurChar)
                        If AscCurChar < 32 Or _
                            AscCurChar > 127 Then
                        'Keep track that we had an error in this cell
                            ErrorThisIt = True
                            ErrorCount = ErrorCount + 1
                        'color it red so we can easily find the field
                            cl.Interior.Color = 255
                        'get a replacement char
                            newchar = charclean(AscCurChar)
                        'add it to our output dictionary
                            If Not BadDic.Exists(WS.Name & " " & cl.Address) Then
                                BadDic.Add WS.Name & " " & cl.Address, "From " & qt & CurString & qt & ", Replaced: " & _
                                qt & CurChar & qt & " -> " & qt & newchar & qt
                            Else
                                BadDic(WS.Name & " " & cl.Address) = BadDic(WS.Name & " " & cl.Address) & _
                                ", " & Chr(34) & CurChar & Chr(34) & " -> " & qt & newchar & qt
                            End If '/Dictionary entry exists
                        'clean up our curstring with the replacement we got from charclean
                            CurString = Left(CurString, i - 1) & newchar & Right(CurString, Len(CurString) - i)
                        End If '/AscCurChar outside of range
                    Next i '/looping through characters
                End If '/len curstring > 0
            'if we found an error, replace the cell value, and finish off our output dic entry
                If ErrorThisIt Then
                're-trim in case we substituted lots of space characters
                    CurString = WorksheetFunction.Trim(CurString)
                    BadDic(WS.Name & " " & cl.Address) = BadDic(WS.Name & " " & cl.Address) & ", to " & qt & CurString & qt
                    ErrorThisIt = False 'we are moving to a new cell.  Reset this.
                End If '/end checking if there was an error this iteration
                cl.Value = CurString
            Next '/next cell in selected text cells
            On Error GoTo 0
            Cells(1, 1).Select
        End If
Nextws:
    Next WS
    
'If we ran into funkyness - output a summary on a new sheet
    If ErrorCount > 0 Then
        On Error Resume Next
        Set FunkySheet = TWB.Worksheets("FunkyChars")
        On Error GoTo 0
        If FunkySheet Is Nothing Then
            Sheets.Add().Name = "FunkyChars"
            Set FunkySheet = TWB.Worksheets("FunkyChars")
        End If
        FunkySheet.activate
        FunkySheet.Cells.ClearContents
        FunkySheet.Range("A1").Value = "Address"
        FunkySheet.Range("B1").Value = "Funkyness"
        i = 2
        For Each itm In BadDic.keys
            FunkySheet.Range("A" & i).Value = itm
            FunkySheet.Range("B" & i).Value = BadDic(itm)
            i = i + 1
        Next itm
        FunkySheet.Cells.EntireColumn.AutoFit
    Else
'if no funkyness, re-select whatever sheet was originally active
        StartWS.activate
    End If
'here we will output times onto the article create sheet.
On Error Resume Next
    ActiveWorkbook.Sheets("Article Create").Range("V1:W3").Font.ColorIndex = 2
    ActiveWorkbook.Sheets("Article Create").Range("W1:W3").NumberFormat = "hh:mm:ss;@"
    ActiveWorkbook.Sheets("Article Create").Range("V3").Value = "LookForBadChars Runtime"
    ActiveWorkbook.Sheets("Article Create").Range("W3").Value = Format(Now - starttime, "hh:mm:ss")
On Error GoTo 0

'output some data onto the PFUsheet
On Error Resume Next
    pfusheet.Range("B" & PFULast).Value = ErrorCount > 0
    If ErrorCount > 0 Then
        pfusheet.Range("C" & PFULast).Value = "If Funky Chars in customer submitted data, feedback to merch"
    End If
On Error GoTo 0

    StartWS.activate
'Commented out because exits to char lookup where we want this off
'    Application.ScreenUpdating = True
'    Application.DisplayStatusBar = True
'    Application.Calculation = xlCalculationAutomatic
'    Application.EnableEvents = True

    'MsgBox ("It took " & Format(Now - Starttime, "hh:mm:ss") & " to find " & errorcount & " funky characters.")
 
End Sub
Private Function charclean(ASCIICODE As Long) As String
'******************************************************************************
' Intended to function in conjunction with the LookForBadChars sub.
' takes as an input an ascii code (>128 is intended) and returns an ascii code
' for a character that is similar, but <= 127
'******************************************************************************
    Select Case ASCIICODE
        Case Is = 130
            charclean = Chr(44) ' Ç with ,
        Case Is = 131
            charclean = Chr(102) ' É with f
        Case Is = 132
            charclean = Chr(44) & Chr(44) ' Ñ with ,,
        Case Is = 133
            charclean = Chr(46) & Chr(46) & Chr(46) ' Ö with ...
        Case Is = 136
            charclean = Chr(94) ' à with ^
        Case Is = 137
            charclean = Chr(37) ' â with %
        Case Is = 138
            charclean = Chr(83) ' ä with S
        Case Is = 139
            charclean = Chr(60) ' ã with <
        Case Is = 141
            charclean = Chr(32) ' ç with
        Case Is = 142
            charclean = Chr(90) ' é with Z
        Case Is = 143
            charclean = Chr(32) ' è with
        Case Is = 144
            charclean = Chr(32) ' ê with
        Case Is = 145
            charclean = Chr(39) ' ë with '
        Case Is = 146
            charclean = Chr(39) ' í with '
        Case Is = 147
            charclean = Chr(34) ' ì with "
        Case Is = 148
            charclean = Chr(34) ' î with "
        Case Is = 150
            charclean = Chr(45) ' ñ with -
        Case Is = 151
            charclean = Chr(45) ' ó with -
        Case Is = 152
            charclean = Chr(126) ' ò with ~
        Case Is = 154
            charclean = Chr(115) ' ö with s
        Case Is = 155
            charclean = Chr(62) ' õ with >
        Case Is = 157
            charclean = Chr(32) ' ù with
        Case Is = 158
            charclean = Chr(122) ' û with z
        Case Is = 159
            charclean = Chr(89) ' ü with Y
        Case Is = 160
            charclean = Chr(32) ' † with
        Case Is = 161
            charclean = Chr(105) ' ° with i
        Case Is = 165
            charclean = Chr(89) ' • with Y
        Case Is = 166
            charclean = Chr(124) ' ¶ with
        Case Is = 171
            charclean = Chr(60) & Chr(60) ' ´ with <<
        Case Is = 173
            charclean = Chr(45) ' ≠ with -
        Case Is = 180
            charclean = Chr(39) ' ¥ with '
        Case Is = 181
            charclean = Chr(117) ' µ with u
        Case Is = 184
            charclean = Chr(44) ' ∏ with ,
        Case Is = 189
            charclean = "1/2" ' Ω with 1/2
        Case Is = 191
            charclean = Chr(63) ' ø with ?
        Case Is = 192
            charclean = Chr(65) ' ¿ with A
        Case Is = 193
            charclean = Chr(65) ' ¡ with A
        Case Is = 194
            charclean = Chr(65) ' ¬ with A
        Case Is = 195
            charclean = Chr(65) ' √ with A
        Case Is = 196
            charclean = Chr(65) ' ƒ with A
        Case Is = 197
            charclean = Chr(65) ' ≈ with A
        Case Is = 199
            charclean = Chr(67) ' « with C
        Case Is = 200
            charclean = Chr(69) ' » with E
        Case Is = 201
            charclean = Chr(69) ' … with E
        Case Is = 202
            charclean = Chr(69) '   with E
        Case Is = 203
            charclean = Chr(69) ' À with E
        Case Is = 204
            charclean = Chr(73) ' Ã with I
        Case Is = 205
            charclean = Chr(73) ' Õ with I
        Case Is = 206
            charclean = Chr(73) ' Œ with I
        Case Is = 207
            charclean = Chr(73) ' œ with I
        Case Is = 208
            charclean = Chr(68) ' – with D
        Case Is = 209
            charclean = Chr(78) ' — with N
        Case Is = 210
            charclean = Chr(79) ' “ with O
        Case Is = 211
            charclean = Chr(79) ' ” with O
        Case Is = 212
            charclean = Chr(79) ' ‘ with O
        Case Is = 213
            charclean = Chr(79) ' ’ with O
        Case Is = 214
            charclean = Chr(79) ' ÷ with O
        Case Is = 215
            charclean = Chr(120) ' ◊ with x
        Case Is = 216
            charclean = Chr(79) ' ÿ with O
        Case Is = 217
            charclean = Chr(85) ' Ÿ with U
        Case Is = 218
            charclean = Chr(85) ' ⁄ with U
        Case Is = 219
            charclean = Chr(85) ' € with U
        Case Is = 220
            charclean = Chr(85) ' ‹ with U
        Case Is = 221
            charclean = Chr(89) ' › with Y
        Case Is = 223
            charclean = Chr(66) ' ﬂ with B
        Case Is = 224
            charclean = Chr(97) ' ‡ with a
        Case Is = 225
            charclean = Chr(97) ' · with a
        Case Is = 226
            charclean = Chr(97) ' ‚ with a
        Case Is = 227
            charclean = Chr(97) ' „ with a
        Case Is = 228
            charclean = Chr(97) ' ‰ with a
        Case Is = 229
            charclean = Chr(97) ' Â with a
        Case Is = 231
            charclean = Chr(99) ' Á with c
        Case Is = 232
            charclean = Chr(101) ' Ë with e
        Case Is = 233
            charclean = Chr(101) ' È with e
        Case Is = 234
            charclean = Chr(101) ' Í with e
        Case Is = 235
            charclean = Chr(101) ' Î with e
        Case Is = 236
            charclean = Chr(105) ' Ï with i
        Case Is = 237
            charclean = Chr(105) ' Ì with i
        Case Is = 238
            charclean = Chr(105) ' Ó with i
        Case Is = 239
            charclean = Chr(105) ' Ô with i
        Case Is = 240
            charclean = Chr(111) '  with o
        Case Is = 241
            charclean = Chr(110) ' Ò with n
        Case Is = 242
            charclean = Chr(111) ' Ú with o
        Case Is = 243
            charclean = Chr(111) ' Û with o
        Case Is = 244
            charclean = Chr(111) ' Ù with o
        Case Is = 245
            charclean = Chr(111) ' ı with o
        Case Is = 246
            charclean = Chr(111) ' ˆ with o
        Case Is = 247
            charclean = Chr(47) ' ˜ with /
        Case Is = 248
            charclean = Chr(111) ' ¯ with o
        Case Is = 249
            charclean = Chr(117) ' ˘ with u
        Case Is = 250
            charclean = Chr(117) ' ˙ with u
        Case Is = 251
            charclean = Chr(117) ' ˚ with u
        Case Is = 252
            charclean = Chr(117) ' ¸ with u
        Case Is = 253
            charclean = Chr(121) ' ˝ with y
        Case Is = 255
            charclean = Chr(121) ' ˇ with y
        Case Else
            charclean = vbNullString
        End Select
      
'        curstring = Replace(curstring, Chr(133), Chr(46) & Chr(46) & Chr(46)) ' Ö with ...
'        curstring = Replace(curstring, Chr(136), Chr(94)) ' à with ^
'        curstring = Replace(curstring, Chr(137), Chr(37)) ' â with %
'        curstring = Replace(curstring, Chr(138), Chr(83)) ' ä with S
'        curstring = Replace(curstring, Chr(139), Chr(60)) ' ã with <
'        curstring = Replace(curstring, Chr(141), Chr(32))  ' ç with
'        curstring = Replace(curstring, Chr(142), Chr(90)) ' é with Z
'        curstring = Replace(curstring, Chr(143), Chr(32))  ' è with
'        curstring = Replace(curstring, Chr(144), Chr(32))  ' ê with
'        curstring = Replace(curstring, Chr(145), Chr(39)) ' ë with '
'        curstring = Replace(curstring, Chr(146), Chr(39)) ' í with '
'        curstring = Replace(curstring, Chr(147), Chr(34)) ' ì with "
'        curstring = Replace(curstring, Chr(148), Chr(34)) ' î with "
'        curstring = Replace(curstring, Chr(150), Chr(45)) ' ñ with -
'        curstring = Replace(curstring, Chr(151), Chr(45)) ' ó with -
'        curstring = Replace(curstring, Chr(152), Chr(126)) ' ò with ~
'        curstring = Replace(curstring, Chr(154), Chr(115)) ' ö with s
'        curstring = Replace(curstring, Chr(155), Chr(62)) ' õ with >
'        curstring = Replace(curstring, Chr(157), Chr(32)) ' ù with
'        curstring = Replace(curstring, Chr(158), Chr(122)) ' û with z
'        curstring = Replace(curstring, Chr(159), Chr(89)) ' ü with Y
'        curstring = Replace(curstring, Chr(160), Chr(32)) ' † with
'        curstring = Replace(curstring, Chr(161), Chr(105)) ' ° with i
'        curstring = Replace(curstring, Chr(165), Chr(89)) ' • with Y
'        curstring = Replace(curstring, Chr(166), Chr(124)) ' ¶ with |
'        curstring = Replace(curstring, Chr(171), Chr(60) & Chr(60)) ' ´ with <<
'        curstring = Replace(curstring, Chr(173), Chr(45)) ' ≠ with -
'        curstring = Replace(curstring, Chr(180), Chr(39)) ' ¥ with '
'        curstring = Replace(curstring, Chr(181), Chr(117)) ' µ with u
'        curstring = Replace(curstring, Chr(184), Chr(44)) ' ∏ with ,
'        curstring = Replace(curstring, Chr(191), Chr(63)) ' ø with ?
'        curstring = Replace(curstring, Chr(192), Chr(65)) ' ¿ with A
'        curstring = Replace(curstring, Chr(193), Chr(65)) ' ¡ with A
'        curstring = Replace(curstring, Chr(194), Chr(65)) ' ¬ with A
'        curstring = Replace(curstring, Chr(195), Chr(65)) ' √ with A
'        curstring = Replace(curstring, Chr(196), Chr(65)) ' ƒ with A
'        curstring = Replace(curstring, Chr(197), Chr(65)) ' ≈ with A
'        curstring = Replace(curstring, Chr(199), Chr(67)) ' « with C
'        curstring = Replace(curstring, Chr(200), Chr(69)) ' » with E
'        curstring = Replace(curstring, Chr(201), Chr(69)) ' … with E
'        curstring = Replace(curstring, Chr(202), Chr(69)) '   with E
'        curstring = Replace(curstring, Chr(203), Chr(69)) ' À with E
'        curstring = Replace(curstring, Chr(204), Chr(73)) ' Ã with I
'        curstring = Replace(curstring, Chr(205), Chr(73)) ' Õ with I
'        curstring = Replace(curstring, Chr(206), Chr(73)) ' Œ with I
'        curstring = Replace(curstring, Chr(207), Chr(73)) ' œ with I
'        curstring = Replace(curstring, Chr(208), Chr(68)) ' – with D
'        curstring = Replace(curstring, Chr(209), Chr(78)) ' — with N
'        curstring = Replace(curstring, Chr(210), Chr(79)) ' “ with O
'        curstring = Replace(curstring, Chr(211), Chr(79)) ' ” with O
'        curstring = Replace(curstring, Chr(212), Chr(79)) ' ‘ with O
'        curstring = Replace(curstring, Chr(213), Chr(79)) ' ’ with O
'        curstring = Replace(curstring, Chr(214), Chr(79)) ' ÷ with O
'        curstring = Replace(curstring, Chr(215), Chr(120)) ' ◊ with x
'        curstring = Replace(curstring, Chr(216), Chr(79)) ' ÿ with O
'        curstring = Replace(curstring, Chr(217), Chr(85)) ' Ÿ with U
'        curstring = Replace(curstring, Chr(218), Chr(85)) ' ⁄ with U
'        curstring = Replace(curstring, Chr(219), Chr(85)) ' € with U
'        curstring = Replace(curstring, Chr(220), Chr(85)) ' ‹ with U
'        curstring = Replace(curstring, Chr(221), Chr(89)) ' › with Y
'        curstring = Replace(curstring, Chr(223), Chr(66)) ' ﬂ with B
'        curstring = Replace(curstring, Chr(224), Chr(97)) ' ‡ with a
'        curstring = Replace(curstring, Chr(225), Chr(97)) ' · with a
'        curstring = Replace(curstring, Chr(226), Chr(97)) ' ‚ with a
'        curstring = Replace(curstring, Chr(227), Chr(97)) ' „ with a
'        curstring = Replace(curstring, Chr(228), Chr(97)) ' ‰ with a
'        curstring = Replace(curstring, Chr(229), Chr(97)) ' Â with a
'        curstring = Replace(curstring, Chr(231), Chr(99)) ' Á with c
'        curstring = Replace(curstring, Chr(232), Chr(101)) ' Ë with e
'        curstring = Replace(curstring, Chr(233), Chr(101)) ' È with e
'        curstring = Replace(curstring, Chr(234), Chr(101)) ' Í with e
'        curstring = Replace(curstring, Chr(235), Chr(101)) ' Î with e
'        curstring = Replace(curstring, Chr(236), Chr(105)) ' Ï with i
'        curstring = Replace(curstring, Chr(237), Chr(105)) ' Ì with i
'        curstring = Replace(curstring, Chr(238), Chr(105)) ' Ó with i
'        curstring = Replace(curstring, Chr(239), Chr(105)) ' Ô with i
'        curstring = Replace(curstring, Chr(240), Chr(111)) '  with o
'        curstring = Replace(curstring, Chr(241), Chr(110)) ' Ò with n
'        curstring = Replace(curstring, Chr(242), Chr(111)) ' Ú with o
'        curstring = Replace(curstring, Chr(243), Chr(111)) ' Û with o
'        curstring = Replace(curstring, Chr(244), Chr(111)) ' Ù with o
'        curstring = Replace(curstring, Chr(245), Chr(111)) ' ı with o
'        curstring = Replace(curstring, Chr(246), Chr(111)) ' ˆ with o
'        curstring = Replace(curstring, Chr(247), Chr(47)) ' ˜ with /
'        curstring = Replace(curstring, Chr(248), Chr(111)) ' ¯ with o
'        curstring = Replace(curstring, Chr(249), Chr(117)) ' ˘ with u
'        curstring = Replace(curstring, Chr(250), Chr(117)) ' ˙ with u
'        curstring = Replace(curstring, Chr(251), Chr(117)) ' ˚ with u
'        curstring = Replace(curstring, Chr(252), Chr(117)) ' ¸ with u
'        curstring = Replace(curstring, Chr(253), Chr(121)) ' ˝ with y
'        curstring = Replace(curstring, Chr(255), Chr(121)) ' ˇ with y

End Function


Private Function BadHTS(WSToCheck As Worksheet) As Boolean
'******************************************************************************
' Here integrated into Charlookup as a function.  Returns "TRUE" if there are
' Bad HTS codes.  Returns "FALSE" if there are no hts codes, or the HTScodes
' are good.
'******************************************************************************
Dim ValidHTSLoc As String
Dim HTSWB As Workbook
Dim GoodHTSArray() As Variant
Dim GoodDic As Object
Dim TWBHTSArray As Variant
Dim endRow As Long
Dim CountBad As Long
Dim HTStoCheck As Boolean
Dim starttime As Date
Dim i As Long
Dim pfusheet As Worksheet
Dim PFULast As Long
''Dim Version As String               'Holds our template version

'Current create sheets have the version in the correct spot, grab that one
'        Version = WSToCheck.Range("H1").Value


    starttime = Now
    
    Set pfusheet = ActiveWorkbook.Worksheets("PFUs")
    PFULast = pfusheet.Range("A65536").End(xlUp).row + 1
    
    
    ValidHTSLoc = "http://teamsites.rei.com/merchandising/Article" _
        & " and Vendor Master Data/Shared Documents/Valid_HTS.xlsb"
        
    Set GoodDic = CreateObject("scripting.dictionary")
    HTStoCheck = False
'Capture the HTS codes on this create sheet
'Need to conditionalize this based upon Version
    
    If Version = "V8.4" Then
        endRow = WSToCheck.Range("BN65536").End(xlUp).row
        TWBHTSArray = WSToCheck.Range("BN11:BO" & endRow).Value
    Else
        endRow = WSToCheck.Range("BQ65536").End(xlUp).row
        TWBHTSArray = WSToCheck.Range("BQ11:BR" & endRow).Value
    End If

'Clean our HTS array, and check to see if we have any actual HTS codes to check
    For i = LBound(TWBHTSArray, 1) To UBound(TWBHTSArray, 1)
        TWBHTSArray(i, 1) = Replace(TWBHTSArray(i, 1), "-", "")
        TWBHTSArray(i, 1) = Replace(TWBHTSArray(i, 1), ".", "")
        TWBHTSArray(i, 1) = Replace(TWBHTSArray(i, 1), " ", "")
        TWBHTSArray(i, 1) = WorksheetFunction.Clean(TWBHTSArray(i, 1))
        TWBHTSArray(i, 1) = WorksheetFunction.Trim(TWBHTSArray(i, 1))
        If IsNumeric(TWBHTSArray(i, 1)) Then
            HTStoCheck = True
        End If
    Next i

'Check to see if we have any actual HTS codes to check
    If HTStoCheck Then
    'Grab the GoodHTSArray
        Workbooks.Open Filename:=ValidHTSLoc, ReadOnly:=True
        Set HTSWB = Workbooks("Valid_HTS.xlsb")
        endRow = HTSWB.Worksheets(1).Range("A65536").End(xlUp).row
        GoodHTSArray = HTSWB.Worksheets(1).Range("A2:A" & endRow).Value
        Workbooks("Valid_HTS.xlsb").Close False
'        MsgBox (GoodDic.Count)
    'Build a dic for quicker lookups
        For i = LBound(GoodHTSArray) To UBound(GoodHTSArray, 1)
            If Not GoodDic.Exists(GoodHTSArray(i, 1)) Then
                GoodDic.Add CStr(GoodHTSArray(i, 1)), "Good!"
            End If
        Next i
'        MsgBox (GoodDic.Count)
        
    'now check TWBHTSArray
        For i = LBound(TWBHTSArray, 1) To UBound(TWBHTSArray, 1)
            If IsNumeric(TWBHTSArray(i, 1)) Then
                If Not GoodDic.Exists(TWBHTSArray(i, 1)) Then
                    TWBHTSArray(i, 2) = "BAD!"
                    If Version = "V8.4" Then
                        WSToCheck.Range("BN" & 10 + i).Interior.Color = 255
                    Else
                        WSToCheck.Range("BQ" & 10 + i).Interior.Color = 255
                    End If
                    CountBad = CountBad + 1
                Else
                    'WSToCheck.Range("BQ" & 10 + i).Value = TWBHTSArray(i, 1)
                    TWBHTSArray(i, 2) = "OK"
                End If '/end not in GoodDic
            End If '/end isnumeric check
        Next i
    End If
    If CountBad > 0 Then
        BadHTS = True
    Else
        BadHTS = False
    End If


    WSToCheck.Range("BN9").Font.ColorIndex = 2
    WSToCheck.Range("BN9").NumberFormat = "hh:mm:ss"
    WSToCheck.Range("BN9").Value = Format(Now - starttime, "hh:mm:ss")
    
    pfusheet.Range("A" & PFULast).Value = "HTS Codes"
    pfusheet.Range("B" & PFULast).Value = BadHTS
    If BadHTS Then
        pfusheet.Range("C" & PFULast).Value = "Check column ""BN"" on the AC Sheet for Red invalid HTS Codes"
    End If

    WSToCheck.activate

        
End Function

Private Sub UniqueIDer()
Dim i As Long
Dim j As Long
Dim IDs() As Variant ' this is an array
Dim curID As String
Dim ErrorFound As Boolean
Dim pfusheet As Worksheet
Dim PFULast As Long
Dim ACLast

    Set pfusheet = AWB.Worksheets("PFUs")
    PFULast = pfusheet.Range("A65536").End(xlUp).row + 1
    ErrorFound = False
    ACLast = 11
'Find last populated row, but ignore "0" from Style Plan Input, and formulas
    Do While Range("G" & ACLast + 1).Value <> "" And _
        Range("G" & ACLast + 1).Value <> 0
        ACLast = ACLast + 1
    Loop
'If there is more than one item, we need to check for uniques
    If ACLast > 11 Then
    
        ACSheet.Range("CI10").Value = "DESC|GENDER|COLOR|SIZE"
        ACSheet.Range("CI11").FormulaR1C1 = _
            "=UPPER(TRIM(RC7)&""|"" &TRIM(RC8)& ""|"" &TRIM(RC10)& ""|"" &TRIM(RC13))"
        ACSheet.Range("CI11").AutoFill Destination:=Range("CI11:CI" & ACLast)
        ACSheet.Range("CI11:CI" & ACLast).Select
        Selection.Calculate
        Selection.FormatConditions.AddUniqueValues
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        Selection.FormatConditions(1).DupeUnique = xlDuplicate
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        Selection.FormatConditions(1).StopIfTrue = False
        IDs = Range("CI11:CI" & ACLast).Value
    
        For i = 1 To UBound(IDs)
            curID = IDs(i, 1)
            For j = i + 1 To UBound(IDs)
                If IDs(j, 1) = curID Then
                    'MsgBox ("Check Column ""CI"" for possible duplicate articles")
                    ErrorFound = True
                    GoTo Endsub
                End If
            Next j
        Next i
    End If '/End more than one item check
Endsub:
    pfusheet.Range("A" & PFULast).Value = "Duplicate Articles"
    pfusheet.Range("B" & PFULast).Value = ErrorFound
    If ErrorFound Then
        pfusheet.Range("C" & PFULast).Value = "Check column ""CI"" on the AC Sheet for Red duplicate looking articles"
    End If

End Sub
Private Sub FillAllGen()
Dim myarray As Variant
Dim Lookups As Worksheet
Dim tempdef(1 To 3) As Variant 'AC Sheet description, Code, Second Char
Dim i As Long
    Set AllGenDic = CreateObject("Scripting.dictionary")
    Set Lookups = Workbooks("Color Size Flavor Lens Ref.xlsb").Worksheets("Lookup")
    myarray = Lookups.Range("AllGen")
    For i = LBound(myarray) To UBound(myarray)
        If Not AllGenDic.Exists(myarray(i, 1)) Then
            tempdef(1) = myarray(i, 2)
            tempdef(2) = myarray(i, 3)
            tempdef(3) = myarray(i, 4)
            AllGenDic.Add myarray(i, 1), tempdef
        End If
    Next i
    Erase myarray
    Set Lookups = Nothing

End Sub
Private Sub SpellCheck(rng As Range)
'******************************************************************************
' Check spelling in a range, highlight yellow entire cells where there is
' a potential misspelling
'
' allowedAbbrevsDic is set up in "fillvars"
'*******************************************************************************
Dim cl As Range
Dim tempsplit As Variant
Dim i As Long
    
    rng.Interior.ColorIndex = xlNone
    'disable all caps ignore
    Application.SpellingOptions.IgnoreCaps = False
    For Each cl In rng
        'split it up into individual words.  Might not be needed
        'but separating things with a slash in the middle is likely
        'needed.
        tempsplit = Split(Replace(cl.Text, "/", " "), " ")
        'go through each word.  Could maybe flag specific misspellings,
        'but not now.
        For i = LBound(tempsplit) To UBound(tempsplit)
            If Not Application.CheckSpelling(Word:=tempsplit(i)) Then
                If Not allowedAbbrevsDic.Exists(tempsplit(i)) Then
                    cl.Interior.Color = vbYellow
                End If
            End If
        Next i
    Next cl
    
End Sub

Private Sub FormatCheck()

    With Application
    If .ReferenceStyle = xlR1C1 Then
        .ReferenceStyle = xlA1
    Else
        .ReferenceStyle = xlA1
    End If
    End With

End Sub
Private Sub GetAbbrevs()
'******************************************************************************
' Just used this to test the popDicFromText sub
'******************************************************************************
Dim allowedAbbrevs As Object
Dim fn As String
    Set allowedAbbrevs = CreateObject("scripting.dictionary")
    fn = "G:\SC EVS\Master Data\Automation\Winshuttle Queries\ColorAbbrevs.txt"
    popDicFromText dic:=allowedAbbrevs, fileloc:=fn
    MsgBox allowedAbbrevs.Count
End Sub
Private Sub popDicFromText(ByRef dic As Object, fileloc As String)
'******************************************************************************
' Grab data from a text file and populate a dictionary with that data.
' Currently written to put the whole line as a "KEY" and the definition is just
' TRUE for all items.  Maybe later we want to adapt this or use it for
' something else?
'******************************************************************************
Dim line As String
    Open fileloc For Input As #1
    Do Until EOF(1)
        Line Input #1, line
        If Not dic.Exists(line) Then
            dic.Add line, True
        End If
    Loop
    Close #1
End Sub
Private Sub garbage()
'******************************************************************************
' erase or unset everything we might have set.  Probably not needed
'******************************************************************************
    On Error Resume Next
        Set ACSheet = Nothing
        Set allowedAbbrevsDic = Nothing
        Set AWB = Nothing
        Set CSWB = Nothing
        Set AllGenDic = Nothing

    On Error GoTo 0
End Sub
Private Sub ValidVendorCheck()

Dim pfusheet As Worksheet
Dim VV As Workbook
Dim PFULast As Long
Dim VVLast As Long
Dim VVArr() As Variant
Dim Vendor As Long
Dim ErrorFound As Boolean
Dim i As Long

    Vendor = ACSheet.Range("I2")
    ACSheet.Calculate
    If ACSheet.Range("J3") <> "" Then
        ErrorFound = Not ACSheet.Range("J3")
        If ErrorFound = False Then
            Exit Sub
        Else
            Set pfusheet = AWB.Worksheets("PFUs")
            PFULast = pfusheet.Range("A65536").End(xlUp).row + 1
            pfusheet.Range("A" & PFULast).Value = "Valid Vendor"
            pfusheet.Range("B" & PFULast).Value = ErrorFound
            pfusheet.Range("C" & PFULast).Value = Vendor & " is either not a valid vendor number or has a PO block on it. Check SAP."
        End If
    Else
    Exit Sub
    End If
    
    ErrorFound = False
End Sub
Private Sub jamInSizeQuery()
'******************************************************************************
' Let's look at size codes on an article create and bring back SAP sort order
' from MDC_PROD.  Might be broken.
'
' MH 09/21/2018
'******************************************************************************
Dim sizeDic As Object
Dim TWB As Workbook
Dim AC As Worksheet
Dim sdCol As Long
Dim lastRow As Long
Dim i As Long
Dim wherein As String
Dim rsArray As Variant
    'Set this workbook to a variable
    Set TWB = ActiveWorkbook
    'Set the article create sheet in this workbook to a variable
    Set AC = TWB.Sheets("Article Create")
    'Create a dictionary
    Set sizeDic = CreateObject("scripting.dictionary")
    
    'Search the Headers for the column used for Size Description
    sdCol = AC.Rows("10:10").Find(What:="Size or Lens Color", LookIn:=xlValues, LookAt:=xlWhole).Column
    lastRow = Cells(Rows.Count, sdCol).End(xlUp).row
    
    If lastRow < 11 Then
        Exit Sub
    End If
    
    'MsgBox "sdCol = " & sdCol & " and lastrow is " & lastrow
    'start our sql wherein statement
    wherein = "('"
    'go through and find all unique sizes, add to dictionary and wherein statement
    For i = 11 To lastRow
        If AC.Cells(i, sdCol).Value <> "" And AC.Cells(i, sdCol).Value <> "N/A" Then
            If Not sizeDic.Exists(AC.Cells(i, sdCol).Value) Then
                sizeDic.Add AC.Cells(i, sdCol).Value, 0
                wherein = wherein & AC.Cells(i, sdCol).Value & "', '"
            End If
        End If 'looks like size is populated.  may be broken
    Next i
    'finish off our wherein statement.
    wherein = Left(wherein, Len(wherein) - 3) & ")"
    
    'now lets query.  We could do that in a different sub.
    
    rsArray = sizeSort(wherein)
    
    'If rsArray is not populated, because no data was brought back, skip to the end.
    If Not IsArray(rsArray) Then GoTo TheEnd
    'rsarray(0, i) = text
    'rsarray(1, i) = sort
    For i = LBound(rsArray, 2) To UBound(rsArray, 2)
        sizeDic(rsArray(0, i)) = rsArray(1, i)
    Next i
    
    For i = 11 To lastRow
        If AC.Cells(i, sdCol).Value <> "" And AC.Cells(i, sdCol).Value <> "N/A" Then
            AC.Range("CK" & i).Value = sizeDic(AC.Cells(i, sdCol).Value)
        End If 'looks like size is populated.  may be broken
    Next i
TheEnd:
    On Error Resume Next
    Set AC = Nothing
    Set TWB = Nothing
    Set sizeDic = Nothing
    Erase rsArray
    On Error GoTo 0
End Sub
Private Function sizeSort(wherein As String) As Variant
'******************************************************************************
' wherein looks like ('S', 'M', 'L', 'XL') and is built in jamInSizeQuery
'
' should return an array to that same sub.
' Lifted from Wes's "OpenSizeRngRS" and adapted for no object dependencies
' and re-worked just a bit based on how I like things.
'******************************************************************************
Dim ObjMyConn As Object 'ADODB.Connection
Dim objMyRecordset As Object 'ADODB.Recordset
Dim objmycommand As Object 'ADODB.Command
Dim SQLStr As String
    
    ' define the connection, recordset, and sql command
    Set ObjMyConn = CreateObject("ADODB.Connection")
    Set objMyRecordset = CreateObject("ADODB.Recordset")
    Set objmycommand = CreateObject("ADODB.Command")
    

    'Define connection string to style plan database
    'Dev database string we will use for testing/dev - comment out before moving to production
    'ObjMyConn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Data Source=sharesqldevnm.reicorpnet.com;Initial Catalog=MDC_PROD_NTAP;"
    'Prod:
    ObjMyConn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Data Source=wkhqpsqldb01c;Initial Catalog=MDC_PROD;"
        
    ObjMyConn.Open
     
    'Open Command for stored procedure'
     SQLStr = "SELECT distinct SizeDesc, SortOrder FROM [MDC_PROD].[dbo].[app_sp_SizeRange] where sizeDesc in " & wherein
     'SQLStr = "SELECT distinct SizeDesc, SortOrder FROM [MDC_PROD].[dbo].[app_sp_SizeRange] where sizeDesc in ('S', 'M', 'L', 'XL')"
     With objmycommand
        .CommandText = SQLStr
        .ActiveConnection = ObjMyConn
     End With
         

         
    'Open Recordset'
    Set objMyRecordset.ActiveConnection = ObjMyConn
    Set objMyRecordset = objmycommand.Execute
    

    On Error Resume Next
    sizeSort = objMyRecordset.GetRows(objMyRecordset.RecordCount)        'This puts the RS into an array
    'if no results are found from the query, then sizesort will be empty. Could check for BOF or EOF
    'but I don't know about that.
    On Error GoTo 0
    
    Set ObjMyConn = Nothing
    Set objMyRecordset = Nothing
    Set objmycommand = Nothing
End Function


Sub Vendor_Snow()
'******************************************************************************
' Read generics for Add Variants off of the AC sheet
'   "Clean" them
'   Build SQL string
'   Allow for timeouts
'
'   Write results to sheet.
'
'******************************************************************************
'Database objects
Dim cn As Object  'ADODB.Connection 'This is the direct DB connection (late Bound)
Dim rs As Object 'ADODB.Recordset 'Gives us a record to return DB results
'set our datasource.  Preconfigured on users's system
Const strCon = "DSN=Snow"

'Dim strCon As String 'our connection as a string
Dim i As Long
Dim j As Long  'holds num of upc's for loop later
Dim DBResultsArray As Variant 'array to shift rows.
Dim PFULastRow As Long         'Last row on the PFU Sheet
Dim pfusheet As Worksheet

Dim AddVendorSheet As Worksheet
Dim AddVendorNeeded As Boolean
Dim UPCsDUPES As Boolean
Dim errMsg As String
Dim rowDic As Object
Dim GenDic As Object
Dim selSQL As String
Dim starttime As Date
Dim selectfromtimer As String
Dim SQLGenList As String
Dim SQLVendor As String
Dim generic As Variant




    
    Set pfusheet = AWB.Worksheets("PFUs")
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    Set rowDic = CreateObject("Scripting.dictionary")
    Set GenDic = CreateObject("Scripting.dictionary")
    
    errMsg = "Likely Timed out checking Generic PIRs.  Please validate them manually."
    
    'Fill Generic Dictionary GenDIc with all our generics we are going to check
    
    i = 11
            
    Do Until ACSheet.Range("G" & i) = ""
        If ACSheet.Range("E" & i) <> "" Then
            If Not GenDic.Exists(ACSheet.Range("E" & i).Value) Then
                GenDic.Add ACSheet.Range("E" & i).Value, ACSheet.Range("E" & i).Value
            End If
        End If
        i = i + 1
    Loop
    
    'skip query if GenDic.Count = 0
    'No add variants then no need to ping snowflake
    If GenDic.Count > 0 Then
    For Each generic In GenDic
        SQLGenList = SQLGenList & "'" & generic & "', "
    Next
    
    'finish SQLGenList and SQLVendor
    SQLGenList = Trim(SQLGenList)
    SQLGenList = Left(SQLGenList, Len(SQLGenList) - 1) & ")"
    SQLVendor = "'" & ACSheet.Range("I2").Value & "'"
    'start our "Select" sql
    selSQL = "select product from PRODUCT.PRODUCT.MERCH_PRODUCT_VENDOR_COST" & vbLf & _
                "where product in (" & SQLGenList & " and vendor_id=" & SQLVendor
    
    'sql should be ready.
    
        On Error Resume Next
        'opens connection to Netezza
        cn.Open strCon
        If Error <> "" Then GoTo Timeout
        On Error GoTo 0
    
        'start timer for select sql
        starttime = Now
    
        Set rs.ActiveConnection = cn
        rs.Open selSQL
        
    
        'convert record set to an array to limit touch time on DB
        If Not rs.EOF Then DBResultsArray = rs.GetRows
    
        'close connection
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing


        'Generic count on AC is different then Generics associated with Vendor in Snowflake
        'Ubound has +1 becuase it is 0 referenced array
        If UBound(DBResultsArray, 2) + 1 <> GenDic.Count Then
            AddVendorNeeded = True
            On Error Resume Next
            Set AddVendorSheet = Sheets("Add Vendor Needed")
            On Error GoTo 0
            If Not AddVendorSheet Is Nothing Then
                If Not AddVendorSheet.Visible = xlSheetVisible Then
                    AddVendorSheet.Visible = xlSheetVisible
                End If
                AddVendorSheet.Select
            Else
                AWB.Sheets.Add().Name = "Add Vendor Needed"
                Set AddVendorSheet = Sheets("Add Vendor Needed")
                AddVendorSheet.Select
            End If
         
            'Get rid of generics that are associated with vendor and drop the ones that aren't on PFU sheet
            For i = LBound(DBResultsArray, 2) To UBound(DBResultsArray, 2)
                If GenDic.Exists(CDbl(DBResultsArray(0, i))) Then
                    GenDic.Remove (CDbl(DBResultsArray(0, i)))
                End If
            Next
            'Clear contents just in case something got left there
            AddVendorSheet.Cells.ClearContents
            'Create some ghetto column headers
            AddVendorSheet.Range("A1").Value = "Generic Not Associated with Vendor"
            AddVendorSheet.Range("B1").Value = "Vendor"
            'Drop out anything where we found a match
            j = 2 'row to insert to in ++'d per iteration
            For Each generic In GenDic
                    AddVendorSheet.Range("A" & j).Value = generic
                    AddVendorSheet.Range("B" & j).Value = ACSheet.Range("I2").Value
                    j = j + 1
            Next
            AddVendorSheet.Cells.EntireColumn.AutoFit
            
        
        End If 'Generic count on AC is different then Generic count from snowflake
    
    End If '/Are there add variants to check
    
    'log things on PFU sheet
    PFULastRow = pfusheet.Range("A" & Rows.Count).End(xlUp).row + 1
    pfusheet.Range("A" & PFULastRow).Value = "Generics Associated with Vendor"
    pfusheet.Range("B" & PFULastRow).Value = AddVendorNeeded
    If AddVendorNeeded = True Then
        pfusheet.Range("C" & PFULastRow).Value = "See Generics that need add vendor first on Add Vendor Needed Sheet"
    End If
    pfusheet.Range("A:C").EntireColumn.AutoFit
Timeout:
    AWB.Worksheets("Article Create").activate
    If Error <> "" Then
        MsgBox Error & vbLf & errMsg
        On Error Resume Next
        rs.Close
        cn.Close
        Set cn = Nothing
        Set rs = Nothing
        On Error GoTo 0
    End If
    On Error Resume Next
    Erase DBResultsArray
    On Error GoTo 0
    Set rowDic = Nothing
    Set AddVendorSheet = Nothing
    Set pfusheet = Nothing
End Sub







