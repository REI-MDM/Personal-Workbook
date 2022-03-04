Attribute VB_Name = "CommentFormulas_v000"
    Option Explicit
    Sub CalcShit()
    '******************************************************************************
    ' Assumes you have one header row, and data filled in column "A"
    ' You may have multiple source data columns, and multiple helper formula
    ' columns.
    ' You can have only row 2 formulas populated if you like, as it will
    ' "Drag down" formulas during execution. This will speed up the individual
    ' calculations, as it "Copy/pastes values" on each formula column one at a time
    '    (this also means your formulas must only rely on data "to the left")
    '
    '******************************************************************************
    Dim lastRow As Long
    Dim lastcol As Long
    Dim i As Long
    Dim cmt As Comment
    Dim formulatext As String
        'find the last row - assume column A
        lastRow = Range("A" & Rows.Count).End(xlUp).row
        'find the last column, assume row 1/header row
        lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
        
        'step across our columns
        For i = 1 To lastcol
            'if row 2 has a formula in it, do some stuff
            If Cells(2, i).HasFormula Then
                formulatext = Cells(2, i).Formula
                'delete any pre-existing comments
                On Error Resume Next
                Set cmt = Cells(1, i).Comment
                Cells(1, i).Comment.Delete
                On Error GoTo 0
                're-add a comment with the formula text
                If Cells(1, i).Comment Is Nothing Then
                    Cells(1, i).AddComment formulatext
                End If
                'Autofill down the formula
                Cells(2, i).Select
                Selection.AutoFill Destination:=Range(Cells(2, i).Address & ":" & Cells(lastRow, i).Address)
                'make it calculate
                Range(Cells(2, i).Address & ":" & Cells(lastRow, i).Address).Calculate
                'make rows 3 to end row values instead of formulas
                Range(Cells(3, i).Address & ":" & Cells(lastRow, i).Address).Value = _
                    Range(Cells(3, i).Address & ":" & Cells(lastRow, i).Address).Value
                'Save the book in case it breaks between this calculation and the next
                ActiveWorkbook.Save
            End If
            Set cmt = Nothing
        Next i
    End Sub

'Sub Macro()
'
'Dim FSO As Object
'Set FSO = CreateObject("Scripting.FileSystemObject")
'
'Dim objFile As Object
'Dim objDSO As Object
'
''For Each objFile In FSO.GetFolder("\\ahqnas1.reicorpnet.com\users\mhildru\Profile\Desktop\Temp SAP Work In Progress\").Files
'    Set objDSO = CreateObject("DSOFile.OleDocumentProperties")
'    Set objDSO = Nothing
'    objDSO.Open "\\ahqnas1.reicorpnet.com\users\mhildru\Profile\Desktop\Temp SAP Work In Progress\List Properties.xlsm" 'objFile.Path
'
''    If objDSO.CustomProperties.Item("Template_ID") = ActiveDocument.CustomDocumentProperties("Template_ID").Value Then
''        ActiveDocument.AttachedTemplate = objFile.Path
''        End
''    End If
''Next
'    objDSO.Close
'
'MsgBox ("No matching template found. Please attach the proper template manually."), vbCritical
'
'End Sub

'Sub CommentFormulas()
'Dim lastcol As Long
'Dim i As Long
'Dim formulatext As String '
'    lastcol = Cells(2, Columns.Count).End(xlToLeft).Column
'    For i = 2 To lastcol '
'        If Cells(2, i).HasFormula Then
'            formulatext = Cells(2, i).Formula
'            Cells(1, i).AddComment formulatext
'        End If '
'    Next i
'End Sub

    
