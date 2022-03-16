Attribute VB_Name = "AVChecker_v007"
Option Explicit
Option Base 0

Public bigArr As Variant        'Stores ALL data on the AC sheet
Public bigAVDic As Object       'stores AV Data from bigarr
Public AVDic As Object          'Stores JUST Generic number and description
Public AVArr() As Variant       'Stores JUST Generic number and description
Public Product_List As String   'for sql statement
'Const conStr = "DSN=Netezza"    'this is what I named my netezza connection in ODBC
Const conStr = "DSN=Snow"    'this is what I named my netezza connection in ODBC
Public rsArray As Variant       'The netezza records returned
Public existingSizes As Variant 'Array for sizes associated - from netezza
Public newSizes As Variant      'Array for new sizes on the AC sheet
Public existingVars As Variant  'array for all existing variants of a style
Public dupeVarsDic As Object    'Keyed off style|colorID|sizeID
Public sizeCheckDic As Object   'Keyed off style|sizeID
Public colorCheckDic As Object  'Keyed off style|colorDESC (trimmed)|CCode|CFAM?
Public charProfDic As Object    'keyed off generic|charProf
Public ACWB As Workbook         'Set this to our AC workbook somehow nicely.
Public issuesheet As Worksheet  'where we log errors
Public NetezzaTimeout As Boolean
Public CPDic As Object          'Translate Char Profile NUMBER from product table into AC Char Profile



Sub AVChecker()
'******************************************************************************
' Main controller - intended to be the "workhorse" for looking at add variants
' on an AC sheet, and querying netezza for existing sizes/colors on those
' generics.
'
' - I originally wanted to userform, but decided I did not want to figure that
' out, so now it just outputs to a worksheet and "errors"
'
' Should do a userform to show if a size-color combo already exists which is an
' "immediate" sendback to MDS, as well as a helper to show if any new sizes
' "match" current sizing.  I.E. Current size is "EU 42" - probably don't want
' to add a NEW "42"
'
' Currently does NOT account for new sizes or items which do not need both
' color and size.
'******************************************************************************
Dim ACSheet As Worksheet
Dim lastRow As Long
Dim fe As Variant
Dim i As Long, k As Long
Dim dvkey As String
Dim sckey As String
Dim cckey As String
Dim cpkey As String
Dim newCount As Long
Dim newDic As Object
Dim pfusheet As Worksheet
Dim tempkey As String


    Set ACWB = ActiveWorkbook
    Set ACSheet = ACWB.Worksheets("Article Create") 'ActiveSheet  'ThisWorkbook.Worksheets("ACSHEET")
    'set some dictionaries
    If AVDic Is Nothing Then Set AVDic = CreateObject("Scripting.dictionary")
    If bigAVDic Is Nothing Then Set bigAVDic = CreateObject("Scripting.dictionary")
    If dupeVarsDic Is Nothing Then Set dupeVarsDic = CreateObject("Scripting.dictionary")
    If sizeCheckDic Is Nothing Then Set sizeCheckDic = CreateObject("Scripting.dictionary")
    If newDic Is Nothing Then Set newDic = CreateObject("scripting.dictionary")
    If colorCheckDic Is Nothing Then Set colorCheckDic = CreateObject("scripting.dictionary")
    If charProfDic Is Nothing Then Set charProfDic = CreateObject("scripting.dictionary")
    
    'clear out old issuesheet (if there is data there)
    clearIssues
    
    'build our dictionary which translates Char profile "Code" into what we expect to see on an AC
    fillCPDic
    
    newCount = 1
    lastRow = ACSheet.Range("E" & Rows.Count).End(xlUp).row
    bigArr = ACSheet.Range("E11:N" & lastRow)
    
    Product_List = "("
    
    For i = LBound(bigArr, 1) To UBound(bigArr, 1)
        If IsNumeric(bigArr(i, 1)) And bigArr(i, 1) <> "" Then
        'Build a "duplicate Values" key based on generic and color and size
        'code if they exist.  This will still not catch "NEW" color or size codes
            tempkey = bigArr(i, 1) & "|"
            'if we have a color code, add that in, otherwise blank
            If IsNumeric(bigArr(i, 7)) Then
                tempkey = tempkey & Format(bigArr(i, 7), "000000")
            End If
            tempkey = tempkey & "|"
            'if we have a size code, add that in, otherwise blank
            If IsNumeric(bigArr(i, 10)) Then
                tempkey = tempkey & Format(bigArr(i, 10), "000000")
            End If
            dvkey = tempkey
            'dvkey = bigArr(i, 1) & "|" & Format(bigArr(i, 7), "000000") & _
                "|" & Format(bigArr(i, 10), "000000") '826025|000321|100008
            If Not dupeVarsDic.Exists(dvkey) Then
                dupeVarsDic.Add dvkey, bigArr(i, 1)
            End If
            
            'build a key got our size checker.  Combine generic and size codes
            'if size code exists.  Will not catch "NEW" size codes.
            tempkey = bigArr(i, 1) & "|"
            If IsNumeric(bigArr(i, 10)) Then
                tempkey = tempkey & Format(bigArr(i, 10), "000000")
            End If
            sckey = tempkey
            'sckey = bigArr(i, 1) & "|" & Format(bigArr(i, 10), "000000") '826025|100008
            If Not sizeCheckDic.Exists(sckey) Then
                sizeCheckDic.Add sckey, bigArr(i, 1)
            End If
            
            'build a key for our char profile checker
            cpkey = bigArr(i, 1) & "|" & bigArr(i, 5)
            If Not charProfDic.Exists(cpkey) Then
                charProfDic.Add cpkey, bigArr(i, 5)
            End If
            
            
            '##This doesn't work yet!  I need to figure it out! ##
            '##Just Kidding, I am pretty sure this works!##
            'Build a key for our color checker.  Combine generic and trimmed
            'color desc - intended to catch when we are trying to add an
            '"EARTH" in color family tan when there is already an "earth" in
            'color family brown there.
            
            'If we look through existing variants and find something where the
            'first two "fields" match, but the whole thing does not, we should
            'throw a warning.  Ideally, we should throw a warning if the
            'existing data has a similar issue.  Add to the dic as we go if
            'new?
            
            'I.E. On the ACSheet we have "112362|EARTH|006867|BROWN"
            'and already existing is "112362|EARTH|000779|NEUTRAL/KHAKI"
            'Then we should "error"
            'When going through our RSarray, we should re-create this key.
            'We should find that second "|" character and compare everything
            'up until it with everything pre-existing, if we find a match, we
            'further check to see if the whole thing is a match.  If we find
            'something where the first two match but not the whole thing,
            'we should output that item somehow
 
            If bigArr(i, 6) <> "" And bigArr(i, 6) <> "N/A" Then
                tempkey = bigArr(i, 1) & "|"
                tempkey = tempkey & UCase(Trim(bigArr(i, 6))) & "|"
                tempkey = tempkey & Format(bigArr(i, 7), "000000") & "|"
                tempkey = tempkey & UCase(Left(bigArr(i, 8), 10))
                cckey = tempkey
                If Not colorCheckDic.Exists(cckey) Then
                    colorCheckDic.Add cckey, bigArr(i, 1)
                End If
            End If
            '## END This doesn't work yet! Section 1##
            
            
            If Not bigAVDic.Exists(bigArr(i, 1)) Then
                bigAVDic.Add bigArr(i, 1), bigArr(i, 3)
                Product_List = Product_List & "'" & bigArr(i, 1) & "',"
            End If
            'Add Variant
        End If '/if isnumeric
    Next i '/loop through bigarr
    
    If Len(Product_List) = 1 Then GoTo cleanup
    '"Finish" Product list.  We want ('123456','456789','789123')
    'but we tacked on an extra , and never terminated our parens.
    Product_List = Left(Product_List, Len(Product_List) - 1) & ")"
    
    QueryNetezza 'This fills rsarray
    
    'if rsarray is empty it should be because we timed out.  Quit.
    If NetezzaTimeout Then GoTo cleanup
    'now we gots to loop through ALL our netezza RS to check for
    'styles with new sizes, or variants which already exist.

'rsarray looks like this - "Backwards" from the normal way I array, starts at 0
'(0, i) - Generic/style - 826025
'(1, i) - Article/SKU - 8260250001
'(2, i) - "color" code - 000321
'(3, i) - "color" description - Black
'(4, 1) - "color" family - Black
'(5, i) - "size" code - 100008
'(6, i) - "size" desc - XS
'(7, i) - "ACCHARPROF" from the master data repository Char_profs table - if an entry is included there
'(7, i) = "Characteristic_Profile_Code" from Merch sku - 23, or 1, or 11.

'Check for char profile issues.  go through rsarray and bigarr and highlight rows on the ac sheet which
'are mismatch

    For i = LBound(rsArray, 2) To UBound(rsArray, 2)
        'replace cp CODE with ACCHARPROF from my dictionary
        If CPDic.Exists(rsArray(7, i)) Then
            'if it does not exist, we want to get a mismatch I think
            rsArray(7, i) = CPDic(rsArray(7, i))
        End If
        dvkey = rsArray(0, i) & "|" & rsArray(2, i) & "|" & rsArray(5, i)
        sckey = rsArray(0, i) & "|" & rsArray(5, i)
        cpkey = rsArray(0, i) & "|" & rsArray(7, i)
        cckey = rsArray(0, i) & "|" & UCase(rsArray(3, i)) & "|" & Format(rsArray(2, i), "000000") & "|" & UCase(rsArray(4, i))
        If dupeVarsDic.Exists(dvkey) Then
            'this is a problem! They are trying to create something which already exists
            If Not AVDic.Exists(rsArray(0, i)) Then
                AVDic.Add rsArray(0, i), bigAVDic(rsArray(0, i))
            End If
            ReportIssue issue:="DUPE", Generic:=rsArray(0, i), desc1:=dvkey, desc2:=rsArray(1, i)
            'MsgBox dvkey & " already exists!  We cannot recreate it!"
        End If
        If sizeCheckDic.Exists(sckey) Then
            'this is "good" - it means there is already an association
            'between the size code and the style/generic.
            sizeCheckDic.Remove sckey
            'we remove it from the dictionary.  It is no longer a potential "new" size
        End If
        
        'Check Char Profile
        If Not charProfDic.Exists(cpkey) And Len(cpkey) > 7 Then
        'Error.  We should have matching data here. but let's just report once.
            For Each fe In charProfDic.keys
                If Left(fe, 6) = rsArray(0, i) Then
                    ReportIssue issue:="CHARPROF", Generic:=rsArray(0, i), desc1:=charProfDic(fe), desc2:=rsArray(7, i)
                    charProfDic.Remove fe
                    Exit For
                End If
            Next fe
            'charProfDic.Remove cpkey
        End If
        
        'Check color family
        If Not colorCheckDic.Exists(cckey) Then
            tempkey = Left(cckey, InStr(8, cckey, "|"))
        'we don't already have variants with this color.  Not really a problem, unless we have variants with the same color name and different color family!
            For Each fe In colorCheckDic.keys
            'loop through all the "requested" colors
                If Left(fe, InStr(8, fe, "|")) = tempkey Then
                'If it looks like the same color description and generic, but does not match exactly (checked above)
                'that is a problem
                    ReportIssue issue:="CFAM", Generic:=rsArray(0, i), desc1:=Right(fe, Len(fe) - 7), desc2:=Right(cckey, Len(cckey) - 7)
                    colorCheckDic.Remove fe 'remove this entry so we don't multi-report the same thing
                    Exit For
                End If
            Next fe
        End If
    Next i
    
    'see if anything remains in our sc dic
    If sizeCheckDic.Count > 0 Then
        For Each fe In sizeCheckDic.keys
            If Not AVDic.Exists(Format(sizeCheckDic(fe), "@")) Then
                AVDic.Add Format(sizeCheckDic(fe), "@"), bigAVDic(sizeCheckDic(fe))
            End If
        Next fe
    End If

    
    
    Set bigAVDic = Nothing
    If AVDic.Count > 0 Then
        
        ReDim AVArr(0 To AVDic.Count - 1, 0 To 1)
        i = LBound(AVArr, 1)
        For Each fe In AVDic.keys
            'populate an array to hopefully show in a list box
            AVArr(i, 0) = fe
            AVArr(i, 1) = AVDic(fe)
            i = i + 1
        Next fe


        For Each fe In AVDic.keys
            FillArrays (fe)
        Next fe
    End If 'AVDic has entries
    
    'Should check for variants which already exist as a "hard stop" and
    'a "soft stop" for verifying new sizes make sense.
    'which means checking this before "continuing on" with the rest of execution.

    'Mostly done here?  I think?  Need to build a function to fill some arrays.
    'based on a "selected" style
    'Existing Sizes - list all unique sizes associated with a style/generic
    'New Sizes - list any sizes on the sheet/in bigArr which are not in rsarray with the style
    'existing variants - list all variants which currently exist for a style
    'MsgBox "BAM!"
    
    
        Set pfusheet = ACWB.Worksheets("PFUs")
        lastRow = pfusheet.Range("A" & Rows.Count).End(xlUp).row + 1
        pfusheet.Range("A" & lastRow).Value = "Add Variant Check"
        pfusheet.Range("B" & lastRow).Value = False
        pfusheet.Range("C" & lastRow).Value = ""


    If Not issuesheet Is Nothing Then
        issuesheet.Columns("A:P").EntireColumn.AutoFit
        pfusheet.Range("B" & lastRow).Value = True
        pfusheet.Range("C" & lastRow).Value = "Check the ws for 'valid' new " & _
            "sizes and variants which already exist"
    End If
    
cleanup:
    Erase bigArr

    Set issuesheet = Nothing
    Set ACWB = Nothing
    Set ACSheet = Nothing
    Set AVDic = Nothing
    Set dupeVarsDic = Nothing 'added
    Set sizeCheckDic = Nothing 'added
    Set colorCheckDic = Nothing
    Set charProfDic = Nothing
    Set newDic = Nothing 'added
    Set pfusheet = Nothing
    Set CPDic = Nothing
    On Error Resume Next
    Erase rsArray
    Erase existingSizes
    Erase newSizes
    Erase existingVars
    Erase AVArr 'added
    On Error GoTo 0
End Sub
Private Sub QueryNetezza()
'******************************************************************************
' Queries Netezza for stuff and junk.
' ConStr needs to be set publicly - Const conStr = "DSN=Netezza"
' Product_List needs to be populated, (and global/public)
' rsarray will be filled (which should also be publicly declared)
'******************************************************************************
Dim strsql As String
Dim cn As Object
Dim rs As Object

    NetezzaTimeout = False
    'late bound - just because.
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
'   SELECT ms.PRODUCT, ms.SKU, ms.COLOR_ID, ms.COLOR_DESC, ms.COLOR_FAMILY, ms.SIZE_ID, ms.SIZE_DESC, cp.ACCHARPROF
'   FROM Product.Product.MERCH_SKU ms
'       left outer join MASTERDATA_REPOSITORY.DWADMIN.CHAR_PROFS cp
'       on ms.PRODUCT = cp.product
'    WHERE ms.Product in ('144593', '103391')
'    ORDER BY SKU
'
    
'Update this to pull char profile code from Merch Sku
'    strsql = "SELECT MS.PRODUCT, MS.SKU, MS.COLOR_ID, MS.COLOR_DESC, MS.COLOR_FAMILY, " & _
'        "MS.SIZE_ID, MS.SIZE_DESC, CP.ACCHARPROF " & _
'        "FROM Product.Product.MERCH_SKU MS " & _
'        "LEFT OUTER JOIN MASTERDATA_REPOSITORY.DWADMIN.CHAR_PROFS CP " & _
'        "ON MS.PRODUCT = CP.PRODUCT " & _
'        "WHERE MS.PRODUCT in " & Product_List & " ORDER BY MS.SKU;"
    
    strsql = "SELECT MS.PRODUCT, MS.SKU, MS.COLOR_ID, MS.COLOR_DESC, MS.COLOR_FAMILY, " & _
        "MS.SIZE_ID, MS.SIZE_DESC, MS.CHARACTERISTIC_PROFILE_CODE " & _
        "FROM Product.Product.MERCH_SKU MS " & _
        "WHERE MS.PRODUCT in " & Product_List & " ORDER BY MS.SKU;"
    
    
    'ActiveSheet.Range("S1").Value = strsql
    On Error Resume Next
    
    cn.Open conStr
    rs.ActiveConnection = cn
    rs.Open strsql
    If Error <> "" Then
        MsgBox Error & vbLf & "Add Variant Checker Timed out.  Please validate manually"
        NetezzaTimeout = True
    Else
        rsArray = rs.GetRows(rs.RecordCount)        'This puts the RS into an array
    End If
    

    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
On Error GoTo 0
End Sub
Private Function FillArrays(style As Long) As Boolean
'******************************************************************************
' Intended to take as an argument a style/generic and fill some "interesting"
' arrays as a result.
' Public existingSizes As Variant   -Sizes already associated with a style
' Public newSizes As Variant        -sizes on the AC sheet not associated
' Public existingVars As Variant    -data on variants associated with a style
'
' Should erase, redim, and populate.
' Existing and new should be "just" code/desc
' existingVars should have "all" data on existing variants
'
' should also be able to be extended to provide similar data for colors.
'
' Data source for existing should be rsarray
' data source for new should be the AC sheet/bigarr
'
' Would be called whenever top listbox selection changes to redraw dependant
' listbox arrays.
'******************************************************************************
Dim exSizeDic As Object
Dim newSizeDic As Object
Dim exVarCount As Long
Dim i As Long
Dim j As Long
Dim fe As Variant
Dim rep1 As String
Dim rep2 As String

    Set exSizeDic = CreateObject("scripting.dictionary")
    Set newSizeDic = CreateObject("scripting.dictionary")
    exVarCount = 0
'clear out existing data
    On Error Resume Next
    Erase existingSizes
    Erase newSizes
    Erase existingVars
    On Error GoTo 0
    
    
'First let's look through rsarray and gather data on existing vars and populate
'existing sizes.  We'll use a dictionary for existing sizes and codes.
'Oh DIP!  we have to "count" existing variants in our array before we can
'resize existingvars!  This is the worst.  Oh well.  We'll make a pass, count
'make another pass, fill.
    For i = LBound(rsArray, 2) To UBound(rsArray, 2)
        If rsArray(0, i) = style Then
            exVarCount = exVarCount + 1
            If Not exSizeDic.Exists(rsArray(5, i)) Then
                'this fills existing size dictionary
                exSizeDic.Add rsArray(5, i), rsArray(6, i)
            End If
        End If
    Next i
    
    ReDim existingVars(0 To exVarCount - 1, 0 To 7)
    exVarCount = 0

    For i = LBound(rsArray, 2) To UBound(rsArray, 2)
        If rsArray(0, i) = CStr(style) Then
            'same as above, but here we need to fill
            For j = 1 To UBound(rsArray, 1)
            'rsarray is "flipped" with generic as element 0
                existingVars(exVarCount, j - 1) = rsArray(j, i)
            Next j
            exVarCount = exVarCount + 1
        End If
    Next i
    'now we have existingvars filled.
    
    'fill existingsizes
    ReDim existingSizes(0 To exSizeDic.Count - 1, 0 To 1)
    i = 0
    For Each fe In exSizeDic.keys
        existingSizes(i, 0) = fe
        existingSizes(i, 1) = exSizeDic(fe)
        i = i + 1
    Next fe
    'now existingsizes array is filled with all existing sizes.
    
    'now we need to loop through bigarr to see if there are any new sizes!
    For i = LBound(bigArr, 1) To UBound(bigArr, 1)
        If bigArr(i, 1) = style Then
            If IsNumeric(bigArr(i, 10)) Or UCase(bigArr(i, 10)) = "NEW" Then 'if not numeric, likely "new" or "N/A"
                If Not exSizeDic.Exists(Format(bigArr(i, 10), "000000")) And Not _
                    newSizeDic.Exists(Format(bigArr(i, 10), "000000")) Then
                        newSizeDic.Add Format(bigArr(i, 10), "000000"), bigArr(i, 9)
                End If
            End If
        End If
    Next i
    
    'and populate our newsizearr if there are new sizes
    If newSizeDic.Count > 0 Then
        i = 0
        ReDim newSizes(0 To newSizeDic.Count - 1, 0 To 1)
        For Each fe In newSizeDic.keys
            newSizes(i, 0) = Format(fe, "000000")
            newSizes(i, 1) = newSizeDic(fe)
            i = i + 1
        Next fe
        For i = 0 To WorksheetFunction.Max(UBound(newSizes, 1), UBound(existingSizes, 1))
            If i <= UBound(newSizes, 1) Then
                rep1 = newSizes(i, 0) & " - " & newSizes(i, 1)
            Else
                rep1 = ""
            End If
            If i <= UBound(existingSizes, 1) Then
                rep2 = existingSizes(i, 0) & " - " & existingSizes(i, 1)
            Else
                rep2 = ""
            End If
            ReportIssue issue:="NEWSIZE", Generic:=style, desc1:=rep1, desc2:=rep2
        Next i
    End If

    'MsgBox exVarCount

Set exSizeDic = Nothing
Set newSizeDic = Nothing
End Function

Private Function ReportIssue(issue As String, Generic As Variant, desc1 As Variant, desc2 As Variant)
'******************************************************************************
' Ummm... take an issue and log it?
' Create an "error" sheet if it does not exist within wb
' determine what type of issue we are reporting:
' "Dupe" = add variant exists - output requested item key and existing variant
' "NewSize" = adding a size to an item, output item key, and list of sizes?
' other?
'******************************************************************************
Dim lastRow As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long

    On Error GoTo makesheet
    Set issuesheet = ACWB.Sheets("Add Variant Check")
    On Error GoTo 0
    If issuesheet Is Nothing Then
makesheet:
        ACWB.Sheets.Add().Name = "Add Variant Check"
        Set issuesheet = ACWB.Worksheets("Add Variant Check")
    End If
On Error GoTo 0
    'Drop Headers statically
    If issuesheet.Cells(1, 1).Value = "" Then
        With issuesheet
            'Dupes
            .Cells(1, 1).Value = "Generic"
            .Cells(1, 2).Value = "Duplicate Key"
            .Cells(1, 3).Value = "Existing Variant"
            'New Sizing
            .Cells(1, 13).Value = "Generic"
            .Cells(1, 14).Value = "New Sizes"
            .Cells(1, 15).Value = "Existing Sizes"
            
            'Kelly's Message
            .Cells(1, 16).Font.Bold = True
            .Cells(1, 16).Value = "Validate before moving forward."
            .Cells(2, 16).Font.Bold = True
            .Cells(2, 16).Value = "Do the new sizes being requested make sense to what is already on the article?"
            
            'conflicting color family
            .Cells(1, 9).Value = "Generic"
            .Cells(1, 10).Value = "New Color"
            .Cells(1, 11).Value = "Existing Colors"
            
            'CharProf
            .Cells(1, 5).Value = "Generic"
            .Cells(1, 6).Value = "This AC CharProf"
            .Cells(1, 7).Value = "Actual CharProf"
        End With
    End If
    
    'Determine where we're dropping data
    Select Case UCase(issue)
        Case "DUPE"
            c1 = 1
            c2 = 2
            c3 = 3
        Case "NEWSIZE"
            c1 = 13
            c2 = 14
            c3 = 15
        Case "CFAM"
            c1 = 9
            c2 = 10
            c3 = 11
        Case "CHARPROF"
            c1 = 5
            c2 = 6
            c3 = 7
    End Select
    'find our last reported issue of this type
    lastRow = issuesheet.Cells(Rows.Count, c1).End(xlUp).row + 1
    
    'log the issue
    issuesheet.Cells(lastRow, c1).Value = Generic
    issuesheet.Cells(lastRow, c2).Value = desc1
    issuesheet.Cells(lastRow, c3).Value = desc2

   
End Function
Private Function fillCPDic()
'HARDCODE some definitions here.

    If CPDic Is Nothing Then Set CPDic = CreateObject("scripting.dictionary")
    
    'only "good" ones are
    CPDic.Add 1, "Color-Size"
    CPDic.Add 2, "Color Only"
    CPDic.Add 3, "Size Only"
    CPDic.Add 4, "Flavor-Size"
    CPDic.Add 5, "Flavor Only"
    CPDic.Add 11, "Frame Color-Lens Color"
    CPDic.Add 21, "Size Only"
    CPDic.Add 22, "Flavor-Size"
    CPDic.Add 23, "Color-Size"
    'Any BAD char profiles cause issues
End Function
Private Sub clearIssues()
    On Error Resume Next
    Set issuesheet = ACWB.Sheets("Add Variant Check")
    On Error GoTo 0
    If Not issuesheet Is Nothing Then
        issuesheet.Cells.ClearContents
    End If
End Sub
