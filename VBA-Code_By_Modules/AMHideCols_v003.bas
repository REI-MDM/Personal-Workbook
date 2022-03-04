Attribute VB_Name = "AMHideCols_v003"
Sub AM_Hide_Cols()
Attribute AM_Hide_Cols.VB_ProcData.VB_Invoke_Func = "H\n14"
'
' AM_Hide_Cols Macro
'
' Keyboard Shortcut: Ctrl+Shift+H
' The following lines of code hide columns that don't have data in them
Dim WS As Worksheet
Dim i As Integer

On Error Resume Next
Set WS = ActiveWorkbook.Worksheets("Maintain Article")

If Not WS Is Nothing Then
    If Application.WorksheetFunction.CountBlank(Range("B9:B500")) = 492 Then
        Columns("B:B").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("C9:C500")) = 492 Then
        Columns("C:C").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("D9:D500")) = 492 Then
        Columns("D:D").ColumnWidth = 0.5
    End If
     If Application.WorksheetFunction.CountBlank(Range("E9:E500")) = 492 Then
        Columns("E:E").ColumnWidth = 0.5
    End If
     If Application.WorksheetFunction.CountBlank(Range("F9:F500")) = 492 Then
        Columns("F:F").ColumnWidth = 0.5
    End If
     If Application.WorksheetFunction.CountBlank(Range("G9:G500")) = 492 Then
        Columns("G:G").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("H9:H500")) = 492 Then
        Columns("H:H").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("I9:I500")) = 492 Then
        Columns("I:I").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("J9:J500")) = 492 Then
        Columns("J:J").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("K9:K500")) = 492 Then
        Columns("K:K").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("L9:L500")) = 492 Then
        Columns("L:L").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("M9:M500")) = 492 Then
        Columns("M:M").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("N9:N500")) = 492 Then
        Columns("N:N").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("O9:O500")) = 492 Then
        Columns("O:O").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("P9:P500")) = 492 Then
        Columns("P:P").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("Q9:Q500")) = 492 Then
        Columns("Q:Q").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("R9:R500")) = 492 Then
        Columns("R:R").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("S9:S500")) = 492 Then
        Columns("S:S").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("T9:T500")) = 492 Then
        Columns("T:T").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("U9:U500")) = 492 Then
        Columns("U:U").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("V9:V500")) = 492 Then
        Columns("V:V").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("W9:W500")) = 492 Then
        Columns("W:W").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("X9:X500")) = 492 Then
        Columns("X:X").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("y9:Y500")) = 492 Then
        Columns("Y:Y").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("Z9:Z500")) = 492 Then
        Columns("Z:Z").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AA9:AA500")) = 492 Then
        Columns("AA:AA").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AB9:AB500")) = 492 Then
        Columns("AB:AB").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("Ac9:Ac500")) = 492 Then
        Columns("Ac:Ac").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AD9:AD500")) = 492 Then
        Columns("AD:AD").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AE9:AE500")) = 492 Then
        Columns("Ae:Ae").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AF9:AF500")) = 492 Then
        Columns("AF:AF").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AG9:AG500")) = 492 Then
        Columns("AG:AG").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AH9:AH500")) = 492 Then
        Columns("AH:AH").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AI9:AI500")) = 492 Then
        Columns("AI:AI").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AJ9:AJ500")) = 492 Then
        Columns("AJ:AJ").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AK9:AK500")) = 492 Then
        Columns("AK:AK").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AL9:AL500")) = 492 Then
        Columns("AL:Al").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AM9:AM500")) = 492 Then
        Columns("AM:AM").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AN9:AN500")) = 492 Then
        Columns("AN:AN").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AO9:AO500")) = 492 Then
        Columns("AO:AO").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AO9:AO500")) = 492 Then
        Columns("AP:AP").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AQ9:AQ500")) = 492 Then
        Columns("AQ:AQ").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AR9:AR500")) = 492 Then
        Columns("AR:AR").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AS9:AS500")) = 492 Then
        Columns("AS:AS").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AT9:AT500")) = 492 Then
        Columns("AT:AT").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AU9:AU500")) = 492 Then
        Columns("AU:AU").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AV9:AV500")) = 492 Then
        Columns("AV:AV").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AW9:AW500")) = 492 Then
        Columns("AW:AW").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AX9:AX500")) = 492 Then
        Columns("AX:AX").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AY9:AY500")) = 492 Then
        Columns("AY:AY").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("AZ9:AZ500")) = 492 Then
        Columns("AZ:AZ").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BA9:BA500")) = 492 Then
        Columns("BA:BA").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BB9:BB500")) = 492 Then
        Columns("BB:BB").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BC9:BC500")) = 492 Then
        Columns("BC:BC").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BD9:BD500")) = 492 Then
        Columns("BD:BD").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BE9:BE500")) = 492 Then
        Columns("BE:BE").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BF9:BF500")) = 492 Then
        Columns("BF:BF").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BG9:BG500")) = 492 Then
        Columns("BG:BG").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BH9:BH500")) = 492 Then
        Columns("BH:BH").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BI9:BI500")) = 492 Then
        Columns("BI:BI").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BJ9:BJ500")) = 492 Then
        Columns("BJ:BJ").ColumnWidth = 0.5
    End If
    If Application.WorksheetFunction.CountBlank(Range("BK9:BK500")) = 492 Then
        Columns("BK:BK").ColumnWidth = 0.5
    End If
    
    

'Unhide Master Data Tools
    Columns("BL:BP").Select
    Selection.EntireColumn.Hidden = False

'Drag down MD tool formulas that were possibly messed up by merchant
    If Range("BC8").Value = "Staging Time" Then
        Range("BC9:BE9").Select
        Selection.AutoFill Destination:=Range("BC9:BE500"), Type:=xlFillDefault
    End If
    

    Range("A9").Select
Else
    Set WS = ActiveWorkbook.Worksheets("Maintain_WSData")
    If Not WS Is Nothing Then
        i = 4
        Do Until WS.Range(Cells(6, i), Cells(6, i)).Value = ""
            If Application.WorksheetFunction.CountBlank(Range(Cells(7, i), Cells(500, i))) = 494 Then
                Columns(i).ColumnWidth = 0.5
            End If
            i = i + 1
        Loop
    End If
End If

End Sub

