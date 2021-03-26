Function httpGet(Text As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", "https://script.google.com/macros/s/AKfycbzN6LaIiU27BLIySHB2pbeG7PL_IbAJDf9tdOMhw5NR_t2lURADymNv_YAwRfYPE8pj/exec?name=" + Text
        .Send
        httpGet = .ResponseText
    End With
End Function
