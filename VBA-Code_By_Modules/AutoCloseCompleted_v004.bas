Attribute VB_Name = "AutoCloseCompleted_v004"
Option Explicit

Sub Auto_Close_Completed()

Dim fp As String
Dim fn As String
Dim WB As Workbook
Dim WS As Worksheet
Dim FollowOnCells As Range
Dim FollowOnCell As Range
Dim FollowOnWork As Boolean
Dim user As String



user = Environ("username")
fp = "G:\SC EVS\Master Data\Automation\Transaction\" & user & "\Completed\"

fn = Dir(fp)

Do Until fn = ""
        Set WB = Workbooks.Open(fp & fn)
        Set WS = WB.Worksheets("WS_AC")
        Set FollowOnCells = WS.Range("A1, A3:A4, C1, D1:D3, E2, E5, F1, F3, F5")
        'Follow on work cells
        'A1 - Promo work
        'A3 - UPC Swap
        'A4 - Prop 65
        'D1 -New Brand
        'D2 - Special Listing for rental/special buy
        'D3 - HTS code on add var
        'E2 - Mixed UPC types and ALT/CAR UPC follow on work
        'F1 - Hazmat
        'F3 - AM!
        'F5 -MBI
        If WS.Range("A6") <> "" Then WS.Range("A6") = ""
        If WS.Range("C4") <> "" Then WS.Range("C4") = ""
        If WS.Range("E1") <> "" Then WS.Range("E1") = ""
        
        FollowOnWork = False
    If UCase(Left(fn, 7)) = "WSMAINT" Then
        FollowOnWork = True
    Else
        For Each FollowOnCell In FollowOnCells
            If FollowOnCell.Value <> "" Then
                FollowOnWork = True
            End If
        Next FollowOnCell
    End If
    
        If FollowOnWork Then
        Else
            Application.Run ("'" & WB.Name & "'" & "!FillInTemplate")
            WB.Close True
        End If
        fn = Dir
    
Loop

End Sub
