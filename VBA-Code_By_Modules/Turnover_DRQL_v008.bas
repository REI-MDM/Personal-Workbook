Attribute VB_Name = "Turnover_DRQL_v008"
Option Explicit

'Const NetLoc = "http://teamsites.rei.com/merchandising/Article%20and%20Vendor%20Master%20Data/Daily%20Status%20Change%20Request%20List/Daily%20Status%20Change%20Request%20List.xlsm"
Const OldNetLoc = "\\teamsites.rei.com\DavWWWRoot\merchandising\Article and Vendor Master Data\Daily Status Change Request List\"
Const NetLoc = "https://reiweb.sharepoint.com/sites/MDMresources/Daily Status Change Request List/"
Sub TurnoverDailyRequestLists()
    Status_Change
    TurnAssDRQL
End Sub
Private Sub Status_Change()
'******************************************************************************
' Open Daily Status Change, save as archived with date, re-open daily status
' change, clear out, update date, save and close.
'******************************************************************************
Dim fn As String
Dim archBook As Object
Dim resetBook As Object
Dim lastRow As Long
Dim newdate As Date
Dim newpath As String
Dim newname As String

    fn = "Daily Status Change Request List.xlsm"
    'fn = "Daily Status Change Request List - Copy.xlsm"
    'find the date we're going to put on the "rolled Over" form
    If Weekday(Now, vbSunday) = 6 Then
        newdate = Now + 3
    Else
        newdate = Now + 1
    End If
    'MsgBox Format(newdate, "dddd, mmmm dd, yyyy")
    
    If Not Workbooks.CanCheckOut(NetLoc & fn) Then
        MsgBox ("Looks like " & fn & " is checked out Sucka! " & _
            "You gots to wait or go investigate")
    Else
        'Open Book,
        Workbooks.CheckOut (NetLoc & fn)
        Set archBook = Workbooks.Open(Filename:=NetLoc & fn)
        'check if there is no new data
        lastRow = archBook.Worksheets("Sheet1").Range("C" & Rows.Count).End(xlUp).row
        If lastRow = 3 Then
            'if no data, just update date, save and close.
            archBook.Worksheets("Sheet1").Range("A1").Value = Format(newdate, "dddd, mmmm dd, yyyy")
            archBook.SaveAs NetLoc & fn
            'archBook.Close
            archBook.CheckIn Comments:="DRQL"
            'archBook.CheckIn savechanges:=True, Comments:="DRQL", MakePublic:=True
            MsgBox ("No work to do for Status Changes.  You're done.  Good job. Go Team.")
        Else
            'If Data is in there
            'save as archived, leave open
            're-open book, clear data, update date, save and close
            
            'determine our archive path
            newpath = NetLoc & "Archive - Status Change/" & Format(Now, "yyyy") & "/"
            'ensure the archive path exists.
            'If Dir(newpath, vbDirectory) = "" Then MkDir newpath
            'save with newname

            newname = Format(Now, "yyyy-mm-dd ") & fn
            archBook.SaveAs newpath & newname
            'Leave this open to process.

            'Reset the other book.
            'Workbooks.CheckOut (NetLoc & fn)
            Set resetBook = Workbooks.Open(Filename:=NetLoc & fn)
            resetBook.Worksheets("Sheet1").Range("A4:D" & lastRow).ClearContents
            resetBook.Worksheets("Sheet1").Range("F4:F950").ClearContents
            resetBook.Worksheets("Sheet1").Range("A1").Value = Format(newdate, "dddd, mmmm dd, yyyy")
            resetBook.CheckIn Savechanges:=True
            'resetBook.Close True
            MsgBox ("We found some Status Change work for you. Process the visible workbook " & _
                archBook.Name & " and then save and close it." & vbLf & vbLf & "<3 <3 <3")
        End If 'Lastrow > 3 or not
        
    End If  'Checkout Test!
    
End Sub
Private Sub TurnAssDRQL()
'******************************************************************************
' Open Assortment Change, save as archived with date, re-open, clear out,
' update date, save and close.
'******************************************************************************
Dim fn As String
Dim archBook As Object
Dim resetBook As Object
Dim lastRow As Long
Dim newdate As Date
Dim newpath As String
Dim newname As String

    fn = "Assortment Group Change Request List.xlsm"
        
    If Not Workbooks.CanCheckOut(NetLoc & fn) Then
        MsgBox ("Looks like " & fn & " is checked out Sucka! " & _
            "You gots to wait or go investigate")
    Else
        'Open Book,
        Workbooks.CheckOut (NetLoc & fn)
        Set archBook = Workbooks.Open(Filename:=NetLoc & fn)
        'check if there is no new data
        lastRow = archBook.Worksheets("Perm Listings").Range("C" & Rows.Count).End(xlUp).row
        If lastRow = 8 Then
            'if no data, save and close.
            'archBook.undocheckout
            archBook.SaveAs NetLoc & fn
            'archBook.Close
            archBook.CheckIn Comments:="DRQL"
            'no message, since we normally don't have work here.
            'MsgBox ("No Assortment work today. This is a good thing.")
        Else
            'If Data is in there
            'save as archived, leave open
            're-open book, clear data, update date, save and close
            
            'determine our archive path
            newpath = NetLoc & "Archive - Assortment Group Change/" & Format(Now, "yyyy") & "/"
            'ensure the archive path exists.
            'If Dir(newpath, vbDirectory) = "" Then MkDir newpath
            'save with newname
            newname = Format(Now, "yyyy-mm-dd ") & fn
            archBook.SaveAs newpath & newname
            'Leave this open to process.

            'Reset the other book.
            'Workbooks.CheckOut (NetLoc & fn)
            Set resetBook = Workbooks.Open(Filename:=NetLoc & fn)
            resetBook.Worksheets("Perm Listings").Range("A9:P" & lastRow).ClearContents
            resetBook.CheckIn Savechanges:=True
            'resetBook.Close True
            MsgBox ("Holy Carp!  Actually something in the assortment work list!  Have the best time shnookums!")
        End If 'Lastrow > 8
        
    End If  'Checkout Test!
    
End Sub
