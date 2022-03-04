Attribute VB_Name = "Quick_Color_Code_v004"
Option Explicit

Sub Quick_Color_Code()

'This macro allows a user to add color codes to "unknown colors" quickly. The color and code table must be copied from a
'Waypoint request and pasted into cell R2 of the ColorCheck worksheet on an AC template.
'Written by Brian Combs - January 2021

'Variables
    Dim ACSheet As Worksheet
    Dim ColorSheet As Worksheet
    Dim Cell As Range
    Dim CodeAndColorTable As Range
    Dim ColorTable As Range
    Dim Code As String
    Dim Color As String
    Dim SpaceNumStr As String
    Dim SpaceNum As Integer
    Dim Count As Integer
    Dim ACColorTable As Range
    Dim FoundCell As Range
    Dim ACLastRow As Long
    Dim CCLastRow As Long
    
'Turn off display alerts
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False

'Set ACSheet and ACColorTable
    Set ACSheet = ActiveWorkbook.Worksheets("Article Create")
    ACSheet.Select
    ACLastRow = 11
    Do While Range("G" & ACLastRow).Value <> ""
        ACLastRow = ACLastRow + 1
    Loop
    ACLastRow = ACLastRow - 1
    Set ACColorTable = Worksheets("Article Create").Range("J11:J" & ACLastRow)

'Set ColorSheet and CodeAndColor Table
    Set ColorSheet = Worksheets("ColorCheck")
    ColorSheet.Select
    CCLastRow = 2
    Do While Range("R" & CCLastRow).Value <> ""
        CCLastRow = CCLastRow + 1
    Loop
    CCLastRow = CCLastRow - 1
    Set CodeAndColorTable = ColorSheet.Range("R2:R" & CCLastRow)
    ColorSheet.Columns("R").AutoFit

'Make sure the CodeAndColor Table is pasted in correct spot
    If ColorSheet.Range("R2").Value = "" Then
        MsgBox "Error. Please make sure that the first new color is pasted in cell R2 of the ColorCheck worksheet."
        Exit Sub
    End If

'Loop through the CodeAndColorTable
    For Each Cell In CodeAndColorTable
    'Find the Code
        Code = Left(Cell, 6)
        Range("S" & Cell.row).Value = Code
    'Find the Color
        Color = Right(Cell.Value, Len(Cell) - 7)
        Range("T" & Cell.row).Value = Color
        ColorSheet.Columns("S:T").AutoFit
    Next

'Loop through the ColorTable
    Set ColorTable = ColorSheet.Range("T2:T" & CCLastRow)
    For Each Cell In ColorTable
    'Find colors with spaces
        If InStr(Cell.Value, "(") Then
        'Extract the number from the string
            SpaceNumStr = Right(Cell.Value, 9)
            SpaceNumStr = Split(SpaceNumStr, " ")(0)
            SpaceNumStr = Right(SpaceNumStr, 1)
            SpaceNum = CInt(SpaceNumStr)
            Range("U" & Cell.row).Value = SpaceNum
        'Remove "(X Space)" from Color
            Cell = Split(Cell, " (")(0)
        End If
    Next

'Loop through the ColorTable again (because I suck at writing)
    For Each Cell In ColorTable
        Color = Cell.Value
        Code = Cell.Offset(0, -1)
        SpaceNum = Cell.Offset(0, 1)
    'Loop through the ACColorTable, enter codes, add spaces
        For Each FoundCell In ACColorTable
            'Enter color codes on AC sheet
            If FoundCell.Value = Color Then
                FoundCell.Offset(0, 1) = Code
                'Enter spaces on AC colors
                If SpaceNum > 0 Then
                    For Count = 1 To SpaceNum
                        FoundCell.Value = " " & FoundCell.Value
                    Next
                End If
            End If
        Next
    Next

'Success!
    MsgBox "Color codes added successfully. Rock on!"
    
'Quick lens codes is not ready yet. Share the news...
    If ColorSheet.Range("W2") <> "" Then
        MsgBox "Sorry, the Quick Lens Code section is not complete. Please add lens codes manually."
    End If
   
'Turn on display alerts
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    
End Sub


