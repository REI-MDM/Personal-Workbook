Attribute VB_Name = "ZMDCA_Code"
Option Explicit

Sub ZMDCA_Code()
'******************************************************************************
' ZMDCA_Code Macro
' Originally written by AG
'
' This macro is intended to update the Prop 65 values of newly created articles
' in the ZMDCA table in SAP. It starts by connecting to SAP and loading the
' ZMDCA T-Code, then it iterates over the new articles until it encounters
' a new unique generic. It checks for a value in the ZMDCA column (DC), and if
' it finds one it attempts to add the generic and the ZMDCA code to the table
' using the "New Entries" button. If saving is successful, it will color the
' generic cell green and back up to the table. If the generic is already in the
' table, SAP will display an error message, and the script will cancel the new
' addtition. It will then use the "Position" button to locate the generic in the
' table. This should always succeed since it was triggered by the presence of
' the generic to begin with, but if it fails, the generic cell will be colored
' red and an error message will be displayed when the script completes. If the
' generic is located in the table, the script compares the new value with the
' one already in the table. If they are the same, nothing is changed in SAP and
' if the values are different, the value is set to 3*. In both cases, the cell
' is colored green. Once all rows have been processed, the script returns to the
' SAP home screen and displays an error or success message as appropriate.
'
' The results will be in a table under the list of articles on WS_AC
' Successful updates will be listed and colored green
' Successful runs that did not update will be blank and green
' Unsuccessful runs will be colored red
'
' * The reason the value is always 3 is that any combination of different
'   Prop 65 values is equivalent to a value of 3, and any combination of
'   equal values is equivalent to the same value.
'
'        1 2 3
'      - - - -
'    1 | 1 3 3
'    2 | 3 2 3
'    3 | 3 3 3
'
'******************************************************************************
    Dim i As Long
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim dict As Object
    Dim uniquegen As String
    Dim ZMDCAval As String
    Dim Generic As String
    Dim Errors As Boolean
    Dim rowCount As Integer
    Dim tableRow As Integer
    
    Set WB = ActiveWorkbook
    Set WS = WB.Worksheets("AC_Tmpt")
    ' count number of row in job
    rowCount = WS.Cells(10, 3).End(xlDown).row - 7
    Set WS = WB.Worksheets("WS_AC")
    
    'Connect to SAP with SAPCON Sub
    SAPCON
    If session Is Nothing Then
        MsgBox "Looks like we couldn't connect to SAP... Ending Sub"
        Exit Sub
    End If
    
    'type /ZMDCA into command bar and hit enter
    session.FindById("wnd[0]").Maximize
    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nZMDCA"
    session.FindById("wnd[0]").sendVKey 0
    
    uniquegen = " "
    ZMDCAval = " "
    Errors = False
    
    tableRow = rowCount + 7
    ' Set up Output table
    Cells(tableRow, "C").Value = "Prop 65 Generic"
    Cells(tableRow, "D").Value = "Updated Value"
    Cells(tableRow, "C").Font.Bold = True
    Cells(tableRow, "D").Font.Bold = True
    
    'Step through WS_AC sheet and Prop 65 update
    For i = 8 To 8 + rowCount
        'Test for unique generic
        uniquegen = Cells(i, "A").Value
        If uniquegen = "H" Then
            'Test if generic has a ZMDCA value to add
            ZMDCAval = Cells(i, "DC").Value
            If ZMDCAval <> "" Then
                Generic = Cells(i, "C").Value
                ' New Entry Button
                session.FindById("wnd[0]/tbar[1]/btn[5]").press
                ' Article Number
                session.FindById("wnd[0]/usr/tblSAPLZMM_ZMDCA_TBLGTCTRL_ZMDCA/ctxtZMDCA-MATNR[0,0]").Text = Generic
                ' Prop 65 Value
                session.FindById("wnd[0]/usr/tblSAPLZMM_ZMDCA_TBLGTCTRL_ZMDCA/txtZMDCA-CA_INDICATOR[1,0]").Text = ZMDCAval
                ' Save Button
                session.FindById("wnd[0]/tbar[0]/btn[11]").press
                
                'If error message in status bar, generic already exists in table
                If session.FindById("wnd[0]/sbar").MessageType = "E" Then
                    ' Cancel Button
                    session.FindById("wnd[0]/tbar[0]/btn[12]").press
                    ' "Yes" to confirm Cancel
                    session.FindById("wnd[1]/usr/btnSPOP-OPTION1").press
                    ' Back to ZMDCA Table
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                    ' Position Button
                    session.FindById("wnd[0]/usr/btnVIM_POSI_PUSH").press
                    ' Article number in search field
                    session.FindById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").Text = Generic
                    ' Check Mark button
                    session.FindById("wnd[1]/tbar[0]/btn[0]").press
                    
                    'Test if generic is already in table
                    If session.FindById("wnd[0]/usr/tblSAPLZMM_ZMDCA_TBLGTCTRL_ZMDCA/ctxtZMDCA-MATNR[0,0]").Text = Generic Then
                        ' If the update = current entry, do nothing
                        If session.FindById("wnd[0]/usr/tblSAPLZMM_ZMDCA_TBLGTCTRL_ZMDCA/txtZMDCA-CA_INDICATOR[1,0]").Text <> ZMDCAval Then
                            ' If update <> current entry, change entry to 3
                            session.FindById("wnd[0]/usr/tblSAPLZMM_ZMDCA_TBLGTCTRL_ZMDCA/txtZMDCA-CA_INDICATOR[1,0]").Text = 3
                            ' Log Update Success
                            Cells(i + rowCount, "D").Value = 3
                            ' Save Button
                            session.FindById("wnd[0]/tbar[0]/btn[11]").press
                        End If
                        
                        tableRow = tableRow + 1
                        ' Log No Change Success
                        Cells(tableRow, "C").Value = Generic
                        Cells(tableRow, "D").Interior.ColorIndex = 4
                        
                    ' If not found, mark error
                    Else
                        tableRow = tableRow + 1
                        ' Log Error
                        Cells(tableRow, "C").Value = Generic
                        Cells(tableRow, "C").Interior.ColorIndex = 3
                        Errors = True
                    End If
                ' If no error
                Else
                    tableRow = tableRow + 1
                    ' Log New Success
                    Cells(tableRow, "C").Value = Generic
                    Cells(tableRow, "D").Value = ZMDCAval
                    Cells(tableRow, "D").Interior.ColorIndex = 4
                    
                    ' Back to ZMDCA Table
                    session.FindById("wnd[0]/tbar[0]/btn[3]").press
                End If
            End If
        End If
        Next i

    'back to home screen
    session.FindById("wnd[0]").sendVKey 3

    EndSAPCON
    ' Completion Message
    If Errors Then
        MsgBox "ERRORS! Check the log table under the article numbers."
    Else
        Cells(4, 1).Clear 'Clear Prop65 Flag
        MsgBox "All Prop 65 Updates Successful. Check the log table under the article numbers."
    End If
    
End Sub
