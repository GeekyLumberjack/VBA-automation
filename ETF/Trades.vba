Sub etf()

Dim ie As InternetExplorer
Dim etf As Collection
Dim fend, trade As Boolean
Dim strdate, trdate, enddate As Date
Dim rmquest, oldrmquest As String
Dim quest As String
Dim chng1, chng As Variant
Dim strprice, tradeprice, endprice As Variant

Set etf = New Collection
etf.Add ("TUR")
etf.Add ("EGPT")
etf.Add ("EWZ")
etf.Add ("EWJ")
etf.Add ("EWY")
etf.Add ("EWG")
etf.Add ("EWA")
etf.Add ("EWC")
etf.Add ("EWH")
etf.Add ("EWT")
etf.Add ("EPOL")
etf.Add ("EWP")
etf.Add ("GREK")
etf.Add ("EZA")
etf.Add ("EWK")
etf.Add ("VNM")
etf.Add ("CYNA")

Set ie = New InternetExplorer
'Dim ele As HTMLElementCollection

For e = 1 To etf.Count
fend = False
trade = False
ie.Visible = True
With ie
'    .visable = True
    .navigate "https://finance.yahoo.com/quote/" + etf(e) + "/history?period1=1206680400&period2=1546063200&interval=1d&filter=history&frequency=1d"
    Do While .Busy Or .readyState <> 4
            DoEvents
    Loop
    Set ele = .document.getElementsByTagName("tr")
  ' MsgBox (ele.Length)
   Dim i As Integer
   i = ele.Length
   i = i - 8
    While i > 0
    If i > 1 Then
        'MsgBox (ele(i).innerText)
        If fend = False Then
        If Not ele(i).innerText Like "*2008*" Then
                .document.parentWindow.execScript "window.scrollBy(0,10000)"
                'MsgBox (e)
                Do While .Busy Or .readyState <> 4
                    DoEvents
                Loop
                Set ele = .document.getElementsByTagName("tr")
                i = ele.Length - 7
               ' MsgBox (ele.Length)
        Else
            'MsgBox (ele(i).innerText)
            fend = True
            oldrmquest = ele(i).getElementsByTagName("td")(0).innerText
            rmquest = Replace(oldrmquest, ChrW(8206), "")
            strdate = rmquest
            strprice = ele(i).getElementsByTagName("td")(4).innerText
            'MsgBox (strprice)
        End If
        End If
        If fend = True Then
        If ele(i - 1).getElementsByTagName("td").Length < 4 Then
            'MsgBox (i)
            nxtprice = ele(i - 2).getElementsByTagName("td")(4).innerText
            nxtdate = ele(i - 2).getElementsByTagName("td")(0).innerText
        Else
            nxtprice = ele(i - 1).getElementsByTagName("td")(4).innerText
            'MsgBox (i)
           ' MsgBox (ele(i - 1).getElementsByTagName("td").Length)
            nxtdate = ele(i - 1).getElementsByTagName("td")(0).innerText
            'MsgBox (ele(i - 2).getElementsByTagName("td")(0).innerText)
        End If
        
        chng1 = nxtprice - strprice
        chng = chng1 / strprice * 100
       ' MsgBox (chng)
        If trade = True Then
        If i < ele.Length / 4 Then
        MsgBox ("3/4 of the way done")
        End If
        If i < ele.Length / 4 * 2 Then
        MsgBox ("3/4 of the way done")
        End If
        If i < ele.Length / 4 * 2 + ele.Length / 4 Then
        MsgBox ("3/4 of the way done")
        End If
        If i < ele.Length - ele.Length / 4 Then
        MsgBox ("3/4 of the way done")
        End If
            If chng >= 30 Then
                trade = False
                ActiveCell.Offset(0, 3).Activate
                ActiveCell.Value = nxtdate
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = nxtprice
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = "Yes"
                ActiveCell.Offset(0, -5).Activate
            End If
            If chng <= -15 Then
                trade = False
                ActiveCell.Offset(0, 3).Activate
                ActiveCell.Value = nxtdate
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = nxtprice
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = "No"
                ActiveCell.Offset(0, -5).Activate
            End If
        If trade = False Then
            If chng >= 15 Then
                strdate = nxtdate
                strprice = nxtprice
            End If
            If chng <= -40 Then
                Cells(1, 1).Select
                While IsEmpty(ActiveCell.Value) = False
                    ActiveCell.Offset(1, 0).Activate
                Wend
                ActiveCell.Value = etf(e)
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = nxtdate
                ActiveCell.Offset(0, 1).Activate
                ActiveCell.Value = nxtprice
                ActiveCell.Offset(0, -2).Activate
                strprice = nxtprice
                strdate = nxtdate
                trade = True
            End If
        End If
        End If
        End If
                
                
        i = i - 1
    End If
    Wend
    .Quit
End With
e = e + 1
Next
'ie.Quit
    



End Sub
