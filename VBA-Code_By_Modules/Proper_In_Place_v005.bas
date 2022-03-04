Attribute VB_Name = "Proper_In_Place_v005"
Option Explicit

Sub ProperInPlace()
Dim Start As Range
Dim cl As Range
Dim row As Long
Dim Column As Long
Dim PropRange As Boolean
Dim response As Integer
Dim tempstring As String

        On Error Resume Next
        Set Start = Application.InputBox(Prompt:="Please select the first " & _
            "item in the column you would like to Properize, then hit OK.", _
            Title:="SPECIFY START", Type:=8)
        On Error GoTo 0
        If Start Is Nothing Then
            MsgBox ("No range/cell selected.  Aborting Macro.")
            Exit Sub
        End If
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
    
            
        row = Start.row
        Column = Start.Column
        With ActiveSheet
            Do While .Cells(row, Column).Value <> ""
                tempstring = CStr(.Cells(row, Column).Value)
            'clean/trim stuff
                tempstring = Replace(tempstring, Chr(160), " ")
                tempstring = WorksheetFunction.Clean(tempstring)
                tempstring = WorksheetFunction.Trim(tempstring)
                tempstring = WorksheetFunction.Proper(tempstring)
                    
            'Lets try only doing this if the whole thing is uppercase/lowercase
'                If tempstring = UCase(tempstring) Or _
'                    tempstring = LCase(tempstring) Then
            'The meat
            
            'exceptions - careful with letter combinations that might be in
                'regular words. whitespace can help select against these
                
                'Numbers
                    tempstring = Replace(tempstring, " Iii", " III")
                    tempstring = Replace(tempstring, " Ii", " II")
                'Goretex/waterproof
                    tempstring = Replace(tempstring, "Gtx", "GTX")
                    tempstring = Replace(tempstring, "Xcr", "XCR")
                    tempstring = Replace(tempstring, " Wp", " WP")
                    tempstring = Replace(tempstring, " Wtpf", " WTPF")
                'short sleeve, Long sleeve
                    tempstring = Replace(tempstring, " Ss", " SS")
                    tempstring = Replace(tempstring, " Ls", " LS")
                'Lightweight, mid weaight, superlight'
                    tempstring = Replace(tempstring, "Lw ", "LW ")
                    tempstring = Replace(tempstring, "Mw ", "MW ")
                    tempstring = Replace(tempstring, "Sl ", "SL ")
                'iPhone, iPod
                    tempstring = Replace(tempstring, "Iphone", "iPhone")
                    tempstring = Replace(tempstring, "Ipod", "iPod")
                'Mary Jane
                    tempstring = Replace(tempstring, " Mj", " MJ")
                'Bike Acronyms
                    tempstring = Replace(tempstring, "Mtb", "MTB")
                    tempstring = Replace(tempstring, "Tlr", "TLR")
                    tempstring = Replace(tempstring, "Xr", "XR")
                    tempstring = Replace(tempstring, "Aw1", "AW1")
                    tempstring = Replace(tempstring, "Xr", "XR")
                 'WP Jacket Layering
                    tempstring = Replace(tempstring, "2l", "2L")
                    tempstring = Replace(tempstring, "2.5l", "2.5L")
                    tempstring = Replace(tempstring, "3l", "3L")
'                End If
                'Put our new modified value back
                .Cells(row, Column).Value = tempstring
                row = row + 1
            Loop
        End With
        
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub
