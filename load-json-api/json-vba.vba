Function getJSON(JsonUrl As String, key As String)
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", JsonUrl
    MyRequest.Send
    'MsgBox MyRequest.ResponseText
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    'MsgBox Json("title")
    getJSON = Json(key)
End Function
