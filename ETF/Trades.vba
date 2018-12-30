Sub ETF()

Dim ie As Object
Dim fend As Boolean
Set ie = CreateObject("InternetExplorer.Application")
'Dim ele As HTMLElementCollection
fend = False
With ie
'    .visable = True
    .navigate "https://finance.yahoo.com/quote/TUR/history?period1=1206680400&period2=1546063200&interval=1d&filter=history&frequency=1d"
    Do While .Busy Or .readyState <> 4
            DoEvents
    Loop
    Set ele = .document.getElementsByTagName("tr")
  ' MsgBox (ele.Length)
   Dim i As Integer
   i = ele.Length
   i = i - 8
    While i > 0
        'MsgBox (ele(i).innerText)
        If fend = False Then
        If Not ele(i).innerText Like "*2008*" Then
                .document.parentWindow.execScript "window.scrollBy(0,10000)"
                Do While .Busy Or .readyState <> 4
                    DoEvents
                Loop
                Set ele = .document.getElementsByTagName("tr")
                i = ele.Length - 7
               ' MsgBox (ele.Length)
        Else
            MsgBox (ele(i).innerText)
            fend = True
        End If
        End If
        i = i - 1
    Wend
    .Quit
End With
