Attribute VB_Name = "Promo_Add_Update_v003"
'C:\Users\mhildru\AppData\Roaming\SAP\SAP GUI\Scripts\Script2.vbs
Option Explicit
#If VBA7 Then  '64 bit?
    Private Declare Function _
        CoRegisterMessageFilter Lib "OLE32.DLL" _
        (ByVal lFilterIn As Long, _
        ByRef lPreviousFilter) As Long
#Else
    Private Declare Function _
        CoRegisterMessageFilter Lib "OLE32.DLL" _
        (ByVal lFilterIn As Long, _
        ByRef lPreviousFilter) As Long
#End If

Sub Promo_Add_Item_or_Update_Price()
'******************************************************************************
'proof of concept for working in one promo
'
' Want to evolve this to go through a "mixed bag" Maintain Promo.  Currently it
' is only for one maintain promo, all of the same type.
'
' As currently written, it handles either an "Add Item and Price" or an
' "Update Price"
'
' It enters pricing at the generic, but could/should be extended to enter at
' generic, then GO BACK and enter at variant level where/if needed.
'
' Can also be extended to take care of the WAK16 Activation of the items as well
' But I think that will come at a later date, after we have been able to use
' this a bit more and have had some manual validation steps before jumping to
' Activation

'Need to more robustly call out ranges we deal with, set some sheets
'probably also want to drop in some sort of "log columns"
'******************************************************************************
Dim pws As Worksheet            'Promo worksheet!
Dim i As Long                   'stepping through things
Dim lastRow As Long             'last populated row
Dim proAcDic As Object          'Promo Action Dictionary
Dim activationDic As Object     'store promos we need to activate.  Likely redundant with above
Dim promoAction As String       'Jam promo number and action together for a key
Dim tempdef(0 To 1) As Long     'storing start and stop row for each promoAction
Dim varPriceDic As Object       'store items which should be maintained at variant level
Dim curPrice As Double          'for help in determining if we need to add to varpricedic
Dim gen As String               'six digits for curpricedic
Dim startRow As Long            'start row for a particular promo action
Dim endRow As Long              'end row for a particular promo action
Dim tempstart As Long           'for "flower tool" input - useful if large numbers of adds
Dim tempend As Long             'for "flower tool" input - useful if large numbers of adds
Dim promo As String             'the promo we are dealing with
Dim Action As String            'the action for this promo
Dim lMsgFilter As Long          'some fancy magic Kevin looked up.
Dim tempsplit As Variant        'for breaking apart promo action keys
Dim fe As Variant               'for looping through dictionaries
Dim starttime As Date
Dim totGens As Long
Dim totvars As Long
Dim activate As Boolean
Dim priceBy As String           'Do we price by generic or variant?
Dim copyrange As Range
Dim logrange As Range


'find last row for this promo and action, blah
    starttime = Now
    Set proAcDic = CreateObject("scripting.dictionary")
    Set activationDic = CreateObject("Scripting.dictionary")
    
    Set pws = ActiveSheet
    
    On Error Resume Next
    pws.ShowAllData
    On Error GoTo 0
            
    'decide if we want activation or nor
    'activate = False
    activate = True
    If UCase(pws.Range("AA1").Value) = "FALSE" Then activate = False
    'decide how we're pricing
    priceBy = "Variant"
    If UCase(pws.Range("AB1").Value) = "GENERIC" Then priceBy = "Generic"

            
    lastRow = pws.Range("I" & Rows.Count).End(xlUp).row
    pws.Range("AQ3").Value = Format(starttime, "mm/dd/yyyy hh:mm:ss")
    
    
    'Loop through our full sheet to see how much work there is.
    For i = 6 To lastRow
        'initialize
        promo = pws.Range("C" & i).Value
        Action = pws.Range("A" & i).Value
        startRow = i
        Do While pws.Range("C" & i).Value = promo And _
            pws.Range("A" & i).Value = Action
                i = i + 1
        Loop
        
        'we found a spot where we are dealing with a different action or different promo.
        '"backup a row" - This is dumb
        i = i - 1
        endRow = i
        promoAction = promo & "|" & Action
        If Not proAcDic.Exists(promoAction) Then
            tempdef(0) = startRow
            tempdef(1) = endRow
            proAcDic.Add promoAction, tempdef
        Else
            MsgBox ("It appears your data is not sorted.  Please sort the data and try again")
            Exit Sub
        End If
    Next i
    
    'MsgBox proAcDic.Count
    MsgBox "About to process " & proAcDic.Count & " promo-actions.  Please let me do " & _
        "my thing until I break or I give you the 'Okay' Message box." & _
        vbLf & vbLf & "You will be prompted to enter your SAP system most likely."
    '******************************************************************************
    'start interacting with SAP
    '******************************************************************************
    SAPCON
    'This may help with the OLE message and help to automate the process
    'Source: https://stackoverflow.com/questions/44288799/how-to-deal-with-microsoft-excel-is-waiting-for-another-application-to-complete
    CoRegisterMessageFilter 0&, lMsgFilter
    
    
    'Now loop through all out promo actions!
    For Each fe In proAcDic.keys
        tempsplit = Split(fe, "|")
        promo = tempsplit(0)
        Action = tempsplit(1)
        
        startRow = proAcDic(fe)(0)
        endRow = proAcDic(fe)(1)
        'MsgBox "TEst"
        
                'Abort if we are not adding item and price or updating price!
        If Action <> "Add Item and Price" And _
            Action <> "Update Price" Then
                GoTo NextFE
        Else
            'if we are adding item and price, add this promo to our activation
            'dictionary.  This could probably be elsewhere
            If Not activationDic.Exists(promo) Then
                activationDic.Add promo, True
            End If
        End If
        
        Set varPriceDic = CreateObject("scripting.dictionary")
        'check for items which need to be input at variant level
        
        For i = startRow To endRow
            gen = Left(pws.Range("I" & i).Value, 6)
            'if we are pricing at variant level, everything
            
            If priceBy = "Variant" Then
                If Not varPriceDic.Exists(gen) Then
                    varPriceDic.Add gen, True
                End If 'not already in dic
            Else
            'pricing at generic
                If pws.Range("H" & i).Value <> "" Then
                    'new generic/first variant  Grab price
                    curPrice = pws.Range("P" & i).Value
                Else
                    'not first variant, compare price
                    If pws.Range("P" & i).Value <> curPrice Then
                        If Not varPriceDic.Exists(gen) Then
                            varPriceDic.Add gen, True
                        End If 'not already in dic
                    End If 'curprice
                End If 'generic/variant for price by generic
            End If 'pricing at variant or generic
        Next i
        

        
        'make-a-da-window big
        session.FindById("wnd[0]").resizeWorkingPane 133, 40, False
        'input tcode and press "enter"
        session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwak12"
        session.FindById("wnd[0]").sendVKey 0
        'enter in promo number and press "enter"
        session.FindById("wnd[0]/usr/ctxtWAKHD-AKTNR").Text = promo
        session.FindById("wnd[0]").sendVKey 0
'******************************************************************************
' Load up the promo
'******************************************************************************

        'Debug.Print "Nothing done " & Err.Number
        'push the "load stuff up with filter" button - breaks if not present
        On Error Resume Next
        'Debug.Print "no items " & Err.Number
        'If Err.Number = 0 Then GoTo NoItems
        session.FindById("wnd[1]/tbar[0]/btn[17]").press
        'Debug.Print "no items " & Err.Number
        'push the "multiselect" button?
        session.FindById("wnd[2]/usr/btn%_LT_ARTNR_%_APP_%-VALU_PUSH").press
        If err.Number = 619 Then 'The control could not be found by id.
            'click the button
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Else
            'Copy our "generics" we're adding/maintaining.  Requires that data to be in
            'column "H"
            pws.Range("H" & startRow & ":H" & endRow).Select
            Selection.Copy
           
            'pastey-clipboard button
            session.FindById("wnd[3]/tbar[0]/btn[24]").press
            
            'okay, okay, okay
            session.FindById("wnd[3]/tbar[0]/btn[8]").press
            session.FindById("wnd[2]/tbar[0]/btn[8]").press
            session.FindById("wnd[1]/tbar[0]/btn[0]").press
        End If
        err.Clear
        On Error GoTo 0
        
        'Back in main promo screen
'******************************************************************************
' Add items if Add Item and Price
'******************************************************************************

        'In "Main" Promo Screen.  If "Add Item and Price" we need to add the items!
        If Action = "Add Item and Price" Then
            'We'll use tempstart and tempend to chunk up our promo adds
            'if there are a lot.  SAP breaks somewhere around 30K
            'and "pauses" a lot the larger the number.  5K is a nice
            'round number to process
            tempstart = startRow
        
            Do While tempstart <= endRow
                'adjust our tempend up
                tempend = tempstart + 4999 'arbitrarily grab 5000 lines
                
                'adjust "down" temp end if needed
                If tempend > endRow Then tempend = endRow
                
                'input this batch of 5K
                
                '"flower tool"
                session.FindById("wnd[0]/usr/subBUTTONS:SAPMWAKA:8150/btnSELECT").press
                'Multi-select
                session.FindById("wnd[1]/usr/btn%_LT_MATNR_%_APP_%-VALU_PUSH").press
                
                'clear out any previous entries
                session.FindById("wnd[2]/tbar[0]/btn[16]").press
                
                'copy our variants.
                pws.activate
                pws.Range("I" & tempstart & ":I" & tempend).Select
                Selection.Copy
                
                'pastey button (need items in clipboard!)
                session.FindById("wnd[2]/tbar[0]/btn[24]").press
                
                'okay, okay
                session.FindById("wnd[2]/tbar[0]/btn[8]").press
                session.FindById("wnd[1]/tbar[0]/btn[8]").press
                pws.Range("AQ" & tempstart & ":AQ" & tempend).Value = "Items Added via flower tool at: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
                'bump up our temp start and end
                tempstart = tempend + 1
            Loop
            
        End If
        
'******************************************************************************
' Input Pricing
'******************************************************************************
        'back at main screen
        'find the price we want to maintain
        'do generic level pricing if that is how we are doing it
        If priceBy = "Generic" Then
            For i = startRow To endRow
                If pws.Range("H" & i).Value <> "" Then
                    'set the generic price.  So we can be a little cleaner in the code
                    'but mostly so we can detect needed variant level updates
                    curPrice = pws.Range("P" & i).Value
                    'press searchy button
                    session.FindById("wnd[0]/usr/subBUTTONS:SAPMWAKA:8150/btnSEARCH").press
                    'here, if we have a price warning, we get a "control ID not found" error
                    On Error Resume Next
                    'input the article
                    session.FindById("wnd[1]/usr/ctxtWAKPD-ARTNR").Text = pws.Range("H" & i).Value
                    If err.Number <> 0 Then
                        'we had an error - should try to log it
                        session.FindById("wnd[0]").sendVKey 0
                        session.FindById("wnd[1]/usr/ctxtWAKPD-ARTNR").Text = pws.Range("H" & i).Value
                        pws.Range("AS" & i - 1).Value = "Full price for generic lower than promo price" ' & session.findById("wnd[0]/usr/tblSAPMWAKASCHNERF/txtWAKPD-PLVKP[1,0]").Text
                        err.Clear
                    End If
                    On Error GoTo 0
                        
                    'I don't know what this is doing really
                    'session.findById("wnd[1]/usr/ctxtWAKPD-ARTNR").caretPosition = 6
                    'press the "searchy this item" button
                    session.FindById("wnd[1]/tbar[0]/btn[0]").press
                    'Input the price
                    session.FindById("wnd[0]/usr/tblSAPMWAKASCHNERF/txtWAKPD-PLVKP[5,0]").Text = curPrice
                    pws.Range("AR" & i).Value = "Generic level price input to SAP at:" & Format(Now, "mm/dd/yyyy hh:mm:ss")
                    'Do some crap not sure if this is important.  Probably not.
                    'session.findById("wnd[0]/usr/tblSAPMWAKASCHNERF/txtWAKPD-PLVKP[5,0]").SetFocus
                    'session.findById("wnd[0]/usr/tblSAPMWAKASCHNERF/txtWAKPD-PLVKP[5,0]").caretPosition = 14
                Else
                    'see if we need to add this generic to our varpricedic
                    'check if this lineitem is priced the same as the generic
                    gen = Left(pws.Range("I" & i).Value, 6)
                    If pws.Range("P" & i).Value <> curPrice Then
                        'looks like it does not match
                        If Not varPriceDic.Exists(gen) Then
                            varPriceDic.Add gen, True
                        End If
                    End If
                End If
            Next i
        End If 'Pricing by generic
        
        'kind of dumb, but let's check if we have items in our varpricedic, then loop through EVERYTHING again to variant price
        If varPriceDic.Count > 0 Then
            For i = startRow To endRow
                gen = Left(pws.Range("I" & i).Value, 6)
                curPrice = pws.Range("P" & i).Value
                If varPriceDic.Exists(gen) Then
                    'input variant pricing
                    session.FindById("wnd[0]/usr/subBUTTONS:SAPMWAKA:8150/btnSEARCH").press
                    'here, if we have a price warning, we get a "control ID not found" error
                    On Error Resume Next
                    'input the article
                    session.FindById("wnd[1]/usr/ctxtWAKPD-ARTNR").Text = pws.Range("I" & i).Value
                    If err.Number <> 0 Then
                        'we had an error - should try to log it
                        session.FindById("wnd[0]").sendVKey 0
                        session.FindById("wnd[1]/usr/ctxtWAKPD-ARTNR").Text = pws.Range("I" & i).Value
                        pws.Range("AS" & i - 1).Value = "Full price for Variant lower than promo price"
                        err.Clear
                    End If
                    On Error GoTo 0
                        
                    'I don't know what this is doing really
                    'session.findById("wnd[1]/usr/ctxtWAKPD-ARTNR").caretPosition = 6
                    'press the "searchy this item" button
                    session.FindById("wnd[1]/tbar[0]/btn[0]").press
                    'Input the price
                    session.FindById("wnd[0]/usr/tblSAPMWAKASCHNERF/txtWAKPD-PLVKP[5,0]").Text = curPrice
                    pws.Range("AR" & i).Value = "Variant level price input to SAP at:" & Format(Now, "mm/dd/yyyy hh:mm:ss")
                End If
            Next i
            
        End If
        'erase varpricedic for next round
        Set varPriceDic = Nothing
        
        'press the savey button
        session.FindById("wnd[0]/tbar[0]/btn[11]").press
        If session.FindById("wnd[0]/sbar").MessageType = "W" Then
        session.FindById("wnd[0]").sendVKey 0
        End If
        'Log
        If UCase(session.FindById("wnd[0]/sbar").Text) Like "*SAVED*" Then
        Range("AT" & i - 1).Value = "WAK12 " & session.FindById("wnd[0]/sbar").Text & ": " & Format(Now, "mm/dd/yyyy hh:mm:ss")
        End If

NextFE:
    Next fe
    
'******************************************************************************
' Activate in WAK16
'******************************************************************************
    
    If activate Then
    'Now we can do "bulk" activations
        For Each fe In activationDic.keys
            
            promo = fe

            Set copyrange = pws.Range("I6:I" & lastRow)
            Set logrange = pws.Range("AU6:AU" & lastRow)
            'filter add item and price, and our promo
            pws.Range("$A$5:$AL$" & lastRow).AutoFilter Field:=1, Criteria1:= _
                "=Add Item and Price", Operator:=xlOr, Criteria2:="=Update Price"
            pws.Range("$A$5:$AL$" & lastRow).AutoFilter Field:=3, Criteria1:=promo
            
            'copy our visible cells
            copyrange.SpecialCells(xlCellTypeVisible).Copy
        
            'goto WAK16
            session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nwak16"
            session.FindById("wnd[0]").sendVKey 0
            'toggle preset checkboxes in WAK16
            session.FindById("wnd[0]/usr/chkR_FEHLP").Selected = True
            session.FindById("wnd[0]/usr/chkR_DRUKZ").Selected = False
            'input promo ID
            session.FindById("wnd[0]/usr/ctxtRT_AKTNR").Text = promo
            'probably not needed
            'session.findById("wnd[0]/usr/chkR_DRUKZ").SetFocus
            'press the "input articles" button
            session.FindById("wnd[0]/usr/btn%_RT_MATNR_%_APP_%-VALU_PUSH").press
        
            'press delete to clear any prior entries
            session.FindById("wnd[1]/tbar[0]/btn[16]").press
            'press "paste"
            session.FindById("wnd[1]/tbar[0]/btn[24]").press
            'press Okay"
            session.FindById("wnd[1]/tbar[0]/btn[8]").press
            
            'enable parallel processing if more than 1000 variants (arbitrarily chosen)
            If endRow - startRow > 1000 Then
                session.FindById("wnd[0]/usr/chkR_PARAL").Selected = True
                session.FindById("wnd[0]/usr/ctxtR_GROUP").Text = "parallel_generators"
                session.FindById("wnd[0]/usr/txtR_ANZPO").Text = "20"
                session.FindById("wnd[0]/usr/txtR_MXTSK").Text = "20"
            End If
            
            'execute in background
            session.FindById("wnd[0]/mbar/menu[0]/menu[2]").Select
            'green checkbox
            session.FindById("wnd[1]/tbar[0]/btn[13]").press
            'immediate button (I think)
            session.FindById("wnd[1]/usr/btnSOFORT_PUSH").press
            'save button
            session.FindById("wnd[1]/tbar[0]/btn[11]").press
            
            'Log
            logrange.SpecialCells(xlCellTypeVisible).Value = "WAK16 Activation in background: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
            'pws.Range("AU" & startrow & ":AU" & endrow).Value = "WAK16 Activation in background: " & Format(Now, "mm/dd/yyyy hh:mm:ss")
        
            On Error Resume Next
            pws.ShowAllData
            On Error GoTo 0
        Next fe
    
    End If
    'Go Home!
    session.FindById("wnd[0]/tbar[0]/btn[3]").press
    EndSAPCON

    CoRegisterMessageFilter lMsgFilter, lMsgFilter
    Windows(pws.Parent.Name).activate
    pws.activate
    pws.Range("AQ4").Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
    pws.Range("AQ2").Value = Format(Now - starttime, "hh:mm:ss")
    MsgBox "Okay, all done!"
End Sub


Sub quicktest()
Dim copyrange As Range
Dim lastRow As Long
Dim pws As Worksheet
Dim i As Long
Dim promo As String
Dim activationDic As Object
Dim fe As Variant
Dim Action As String
Dim logrange As Range

    Set activationDic = CreateObject("Scripting.dictionary")

    Set pws = ActiveWorkbook.Worksheets("Maintain_Promo")
    lastRow = 1248
    
    
    On Error Resume Next
    pws.ShowAllData
    On Error GoTo 0
            
    For i = 6 To lastRow
    promo = pws.Range("C" & i).Value
    Action = pws.Range("A" & i).Value
        If Action = "Add Item and Price" Or Action = "Update Price" Then
            If Not activationDic.Exists(promo) Then
                activationDic.Add promo, True
            End If
        End If
    Next i
    
    Set copyrange = pws.Range("I6:I" & lastRow)
    Set logrange = pws.Range("BA6:BA" & lastRow)
    i = 1
    For Each fe In activationDic.keys
            promo = fe

            'filter add item and price, and our promo
            pws.Range("$A$5:$AL$" & lastRow).AutoFilter Field:=1, Criteria1:= _
                "=Add Item and Price", Operator:=xlOr, Criteria2:="=Update Price"
            pws.Range("$A$5:$AL$" & lastRow).AutoFilter Field:=3, Criteria1:=promo
            
            'copy our visible cells
            copyrange.SpecialCells(xlCellTypeVisible).Copy
            
            Worksheets("Test").activate
            Worksheets("Test").Cells(1, i).Value = fe
            i = i + 1
            Worksheets("Test").Cells(1, i).Select
            Worksheets("Test").Paste
            i = i + 1
            pws.activate
            logrange.SpecialCells(xlCellTypeVisible).Value = "Input at the now time: " & Now
            Application.Wait (Now + TimeValue("00:00:05"))
            
            On Error Resume Next
            pws.ShowAllData
            On Error GoTo 0
    Next fe
End Sub
