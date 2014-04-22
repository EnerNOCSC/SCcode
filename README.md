SCcode
======

All of the SC macros code we've got!


OPIE v3.4 Cases and MOLIs 1.6.xlsm

Sub Case_query()
'
' Case_query Macro
'
'

Dim Firstrow, lastRow, lookup As Integer
Dim createShipments, casesEasyPull, shipmentsEasyPull, shipmentsEasyPush, fulfilldate, poackdate, shipmentsFullPush, casesEasyPush As String
Dim poNumber, trackingNumber, quantityShipped, actualShipDate, caseNumber, caseID, NextLine As String

createShipments = "Create Shipments"
casesEasyPull = "Cases - Easy Pull"
casesEasyPush = "Cases - Easy Push"
shipmentsEasyPull = "Shipments - Easy Pull"
shipmentsEasyPush = "Shipments - Easy Push"
shipmentsFullPush = "Shipments - Full Push"

Firstrow = 2

' Find the last populated row
With Sheets(createShipments)
    lastRow = .Cells.Find(what:="*", _
    SearchDirection:=xlPrevious, _
    SearchOrder:=xlByRows).row
End With

Dim i, g, q, t, p, found, c As Integer
i = Firstrow
g = 3
c = 3

For i = Firstrow To lastRow

    poNumber = ""
    trackingNumber = ""
    quantityShipped = ""
    caseNumber = ""
    actualShipDate = ""
    poackdate = ""
    

    poNumber = Sheets(createShipments).Cells(i, 1)
    trackingNumber = Sheets(createShipments).Cells(i, 2)
    quantityShipped = Sheets(createShipments).Cells(i, 3)
    actualShipDate = Sheets(createShipments).Cells(i, 4)
    caseNumber = Sheets(createShipments).Cells(i, 5)
    poackdate = Sheets(createShipments).Cells(i, 6)
     
    
    
        If poackdate = "" And trackingNumber = "" Then
            MsgBox "A tracking number is blank. Bad! Try again."
            Sheets(createShipments).Activate
            Sheets(createShipments).Range(Cells(1, 7), Cells(500, 7)).ClearContents
            Exit Sub
        End If
        
  

    ' Change this to look for caseNumber to get caseID
    If caseNumber <> "" Then
    
        'Query cases
        '----------------------------------------------------------------------------
        Sheets(casesEasyPull).Activate
        Sheets(casesEasyPull).Cells(1, 2) = "Record Type ID"
        Sheets(casesEasyPull).Cells(1, 3) = "equals"
        Sheets(casesEasyPull).Cells(1, 4) = "Order/Return Materials"
        Sheets(casesEasyPull).Cells(1, 5) = "Case Number"
        Sheets(casesEasyPull).Cells(1, 6) = "equals"
        Sheets(casesEasyPull).Cells(1, 7) = caseNumber
        Sheets(casesEasyPull).Cells(1, 1).Select
        
        Run Macro:="sfQuery"
        
        

        
        q = 3
        For q = 3 To 50
        
        If Sheets(casesEasyPull).Cells(q, 1) = "" Then
            If q = 3 Then
            MsgBox "No match found for Case Number " & caseNumber & ". Don't worry about it now, cell is now red."
            Sheets(createShipments).Cells(i, 1).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 2).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 3).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 4).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 5).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 7) = "Nope"
            caseID = "None Found"
            End If
        Else
            caseID = Sheets(casesEasyPull).Cells(q, 1)
            'Sheets("Cases - Easy Push").Range("E" & q) = Sheets(casesEasyPull).Range("C" & q) ''''''''''''''''''need to comment this out in solely MOLI version DLE 4.15.2014
            'Sheets("Cases - Easy Push").Range("F" & q) = Sheets(casesEasyPull).Range("D" & q)
            Sheets(createShipments).Cells(i, 1).Interior.Color = RGB(129, 199, 231)
            Sheets(createShipments).Cells(i, 2).Interior.Color = RGB(129, 199, 231)
            Sheets(createShipments).Cells(i, 3).Interior.Color = RGB(129, 199, 231)
            Sheets(createShipments).Cells(i, 4).Interior.Color = RGB(129, 199, 231)
            Sheets(createShipments).Cells(i, 5).Interior.Color = RGB(129, 199, 231)
            Sheets(createShipments).Cells(i, 7 + 2 * (q - 3)) = "Yep"
            Sheets(createShipments).Cells(i, 8 + 2 * (q - 3)) = caseID
         '   Sheeets(casesEasyPull).Cells(1, 12) = caseID
         '   Sheets(casesEasyPull).Cells(1, 9).Select
        '    Run Macro:="sfQuery"
        '    caseID3 = Sheets(casesEasyPull).Cells(3, 1) '
        '    fulfilldate = Sheets(createShipments).Cells(i, 4)
       '     Sheets(casesEasyPush).Cells(1, 12) = caseID3
        '    Sheets(casesEasyPush).Cells(i, 10) = fulfilldate
            
        End If
        
        Next q
    
    ElseIf poNumber <> "" Then
    
        'Query cases
        '----------------------------------------------------------------------------
        Sheets(casesEasyPull).Activate
        Sheets(casesEasyPull).Cells(1, 2) = "Record Type ID"
        Sheets(casesEasyPull).Cells(1, 3) = "equals"
        Sheets(casesEasyPull).Cells(1, 4) = "Order/Return Materials"
        Sheets(casesEasyPull).Cells(1, 5) = "Purchase Order Number"
        Sheets(casesEasyPull).Cells(1, 6) = "equals"
        Sheets(casesEasyPull).Cells(1, 7) = poNumber
        Sheets(casesEasyPull).Cells(1, 1).Select
        
        Run Macro:="sfQuery"
        q = 3
        For q = 3 To 50
        
        If Sheets(casesEasyPull).Cells(q, 1) = "" Then
            If q = 3 Then
            MsgBox "No match found for PO " & poNumber & ". Don't worry about it now, cell is now red."
            Sheets(createShipments).Cells(i, 1).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 2).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 3).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 4).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 5).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 7) = "Nope"
            caseID = "None Found"
            End If
        Else
            caseID = Sheets(casesEasyPull).Cells(q, 1)
            Sheets(createShipments).Cells(i, 1).Interior.Color = RGB(0, 250, 0)
            Sheets(createShipments).Cells(i, 2).Interior.Color = RGB(0, 250, 0)
            Sheets(createShipments).Cells(i, 3).Interior.Color = RGB(0, 250, 0)
            Sheets(createShipments).Cells(i, 4).Interior.Color = RGB(0, 250, 0)
            Sheets(createShipments).Cells(i, 5).Interior.Color = RGB(0, 250, 0)
            Sheets(createShipments).Cells(i, 7 + 2 * (q - 3)) = "Yep"
            Sheets(createShipments).Cells(i, 8 + 2 * (q - 3)) = caseID
        End If
        
        Next q
        '----------------------------------------------------------------------------
    Else
            MsgBox "No PO and no case ID for one of these guys. Don't worry about it now, cell is now red."
            Sheets(createShipments).Cells(i, 1).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 2).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 3).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 4).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 5).Interior.Color = RGB(250, 0, 0)
            Sheets(createShipments).Cells(i, 7) = "Nope"
            caseID = "None Found"
    End If

' Run through this enough times to get all iterations in

q = 3
For q = 3 To 50
    If Sheets(createShipments).Cells(i, 7 + 2 * (q - 3)) = "Yep" Then
        Sheets(shipmentsEasyPush).Cells(g, 1) = "New"
        Sheets(shipmentsEasyPush).Cells(g, 2) = Sheets(createShipments).Cells(i, 8 + 2 * (q - 3))
        Sheets(shipmentsEasyPush).Cells(g, 3) = quantityShipped
        Sheets(shipmentsEasyPush).Cells(g, 4) = trackingNumber
        Sheets(shipmentsEasyPush).Cells(g, 5) = actualShipDate
  
        
        p = 3
        found = 0
        For p = 3 To (g - 1)
            If Sheets("Cases - Easy Push").Cells(p, 1) = Sheets(createShipments).Cells(i, 8 + 2 * (q - 3)) Then
                found = 1
            End If
        Next p
        
        'seeing if PO is blank, this may not work
        
        ' For p = 3 To (g - 1)
         '   If IsEmpty(Sheets("Cases - Easy Push").Cells(p, 6)) Then
          '      found = 1
           ' End If
        'need to add part about PO not being to boise or UPS
        
        'Next p
        
        If found = 0 Then
            Sheets("Cases - Easy Push").Cells(c, 1) = Sheets(createShipments).Cells(i, 8 + 2 * (q - 3))
            Sheets("Cases - Easy Push").Cells(c, 2) = "Shipped"
            Sheets("Cases - Easy Push").Cells(c, 3) = "Shipment Sent"
            Sheets("Cases - Easy Push").Cells(c, 4) = "None"
            Sheets("Cases - Easy Push").Cells(c, 5) = Sheets(createShipments).Cells(i, 3) 'qty
            c = c + 1
        End If
        
        g = g + 1
    End If
Next q

Next i


Dim caseURL3, caseID3 As String



casesEasyPull = "Cases - Easy Pull"
With Sheets("Create Shipments")
    lastRow = .Cells.Find(what:="*", _
    SearchDirection:=xlPrevious, _
    SearchOrder:=xlByRows).row
End With

lastRow = lastRow + 2

Sheets(casesEasyPull).Activate
Sheets(casesEasyPull).Range("BH3").Select
Run Macro:="sfQuery"

i = 3
For i = 3 To lastRow
    caseURL3 = ""
    If i <> lastRow Then
    
        caseURL3 = Sheets("Create Shipments").Cells(i - 1, 8)
        Sheets(casesEasyPull).Cells(1, 12) = caseURL3
        Sheets(casesEasyPull).Activate
        Range("I1").Select
        
        Run Macro:="sfQuery"
        
        If IsEmpty(ActiveSheet.Range("K3").Value) = True Then
        
         MsgBox "Order Process Type must be populated before fulfillment. Bad! Try again."
            Sheets("Create Shipments").Activate
            Exit Sub
            End If
            
             
            caseID3 = Sheets(casesEasyPull).Cells(3, 9)
            
                Sheets("Create Shipments").Cells(i - 1, 9) = caseID3
                
                Sheets(casesEasyPush).Cells(i, 30) = caseID3
                Sheets(casesEasyPush).Cells(i, 31) = Sheets("Create Shipments").Cells(i - 1, 2) ' Tracking Number
                Sheets(casesEasyPush).Cells(i, 32) = Sheets("Create Shipments").Cells(i - 1, 4) ' fulfillment date
                Sheets(casesEasyPush).Cells(i, 33) = "Shipped" 'status, may need to edit 3/18 DLE
                Sheets(casesEasyPush).Cells(i, 34) = Sheets("Create Shipments").Cells(i - 1, 3) 'qty
                

                
               If Sheets(casesEasyPull).Range("J3") = "a03a000000ETSgTAAX" Or Sheets(casesEasyPull).Range("J3") = "a03a000000ETSgsAAH" Or Sheets(casesEasyPull).Range("J3") = "a03a000000ETSgZAAX" Then
                    Sheets(casesEasyPush).Cells(i, 35) = "Warehouse" 'vendor or warehouse
       
                    Else
                    
                    Sheets(casesEasyPush).Cells(i, 35) = "Vendor" 'vendor or warehouse
                    If Sheets(casesEasyPull).Range("L3") = "01t30000000iGkfAAE" Then
                        GoTo NextLine
                        Else
                    
                        Sheets(casesEasyPull).Activate
                        Sheets("Cases - Easy Pull").Range("V3") = Sheets("Cases - Easy Pull").Range("L3")
                        Sheets(casesEasyPull).Range("V3").Select
                        Run Macro:="sfQueryRow"
                    
                            If Sheets(casesEasyPull).Range("Y3") = "TRUE" Then
                            MsgBox "Product is obsolete. Bad! Try again."
                            Sheets("Create Shipments").Activate
                            Exit Sub
                            End If
                            
                            If Sheets(casesEasyPull).Range("Z3") = 0 Then
                            MsgBox "Standard Cost of product is 0. Bad! Try again."
                            Sheets("Create Shipments").Activate
                            Exit Sub
                            End If
                    Sheets(casesEasyPull).Range("Y3") = Application.WorksheetFunction.Index(Sheets(casesEasyPull).Range("BH:BH"), Application.WorksheetFunction.Match(Sheets(casesEasyPull).Range("X3"), Sheets(casesEasyPull).Range("BJ:BJ"), 0))
                        If Sheets(casesEasyPull).Range("Y3") <> Sheets(casesEasyPull).Range("J3") Then
                        MsgBox "Primary Supplier of product doesn't match Fulfill From. Bad! Try again."
                        Sheets("Create Shipments").Activate
                        Exit Sub
                        End If
    
               End If
    
    End If
NextLine:
Next i

i = 3


If g > 3 Then
    g = g - 1
    c = c - 1
    Sheets(shipmentsEasyPush).Activate
    Sheets(shipmentsEasyPush).Range(Cells(3, 1), Cells(g, 5)).Select
    Run Macro:="sfInsertRow"
    
    Sheets("Cases - Easy Push").Activate
    Sheets("Cases - Easy Push").Range(Cells(3, 1), Cells(c, 5)).Select
    Run Macro:="sfUpdate"
    
    Sheets("Cases - Easy Push").Activate
    Sheets("Cases - Easy Push").Range(Cells(3, 30), Cells(c, 35)).Select
    Run Macro:="sfUpdate"
    
End If

If g > 2 Then
MsgBox "Done creating Shipment records & updating cases. Crikey!"
Else
MsgBox "No Shipment records to create. No updates made to cases."
End If

lookup = Sheets("Cases - Easy Push").Range("A" & Rows.Count).End(xlUp).row

i = 3

For i = 3 To lookup

    Sheets("Cases - Easy Push").Cells(i, 8) = Application.IfError(Application.VLookup(Sheets("Cases - Easy Push").Range("E" & i), Sheets("Supplies").Range("A:D"), 4, 0), "")
    
    If Sheets("Cases - Easy Push").Cells(i, 8) <> "" Then
        Sheets("Cases - Easy Push").Cells(i, 6) = "Do Not Fulfill"
    End If
Next i

'Sheets("Cases - Easy Push").Range("H1").EntireColumn.Delete
'Sheets("Cases - Easy Push").Range("E1").EntireColumn.Hidden = True
MsgBox "If UPS/Boise order, no PO is necessary, fulfill in InvTrak. If the case is a vendor order and it is missing a PO, check to see if the product is tracked in InvTrak."

End Sub
