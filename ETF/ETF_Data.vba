Function stabalize(nxt As Variant)
    Dim i As Integer
    stchng = nxt * 0.05
    i = 0
    stprice = ActiveCell.Offset(2, 1).Value
    While i < 5
     'MsgBox (nxt)
     'MsgBox (stprice)
     'MsgBox (nxt - stchng)
     'MsgBox (nxt + stchng)
     If stprice < nxt - stchng Then
        stabalize = "Not Stable"
        Exit Function
     End If
     'If stprice > nxt + stchng Then
     '   stabalize = "Not Stable"
     '   Exit Function
     'End If
     i = i + 1
     stprice = ActiveCell.Offset(i + 2, 1).Value
    Wend
    stablize = stprice
    


End Function





Sub etf()


Dim fend, trade As Boolean
Dim strdate, trdate, enddate As Date
Dim rmquest, oldrmquest As String
Dim quest As String
Dim chng1, chng As Variant
Dim strprice, tradeprice, endprice As Variant


        Cells(2, 1).Activate
        trade = False
        strprice = ActiveCell.Offset(0, 1).Value
        strdate = ActiveCell.Value
        nxtprice = ActiveCell.Offset(1, 1).Value
        nxtdate = ActiveCell.Offset(1, 0).Value
        
    While IsEmpty(ActiveCell.Value) = False
        chng1 = nxtprice - strprice
        chng = chng1 / strprice * 100
        'MsgBox (trade)
        If trade = True Then
            If chng >= 30 Then
                trade = False
                ActiveCell.EntireRow.Interior.Color = RGB(0, 255, 0)
            End If
            If chng <= -15 Then
                trade = False
                ActiveCell.EntireRow.Interior.Color = RGB(255, 0, 0)
            End If
        End If
        If trade = False Then
            'MsgBox ("trade")
            If chng >= 15 Then
                strdate = nxtdate
                strprice = nxtprice
            End If
            If chng <= -40 Then
                'MsgBox ("trade")
                If Not stabalize(nxtprice) Like "Not Stable" Then
                    ActiveCell.EntireRow.Interior.Color = RGB(0, 0, 255)
                    strprice = nxtprice
                    strdate = nxtdate
                    trade = True
                End If
            End If
        End If
        ActiveCell.Offset(1, 0).Activate
        nxtprice = ActiveCell.Offset(1, 1).Value
        nxtdate = ActiveCell.Offset(1, 0).Value
        'MsgBox (strprice)
        'MsgBox (nxtprice)
        
     Wend
                
    
    



End Sub


