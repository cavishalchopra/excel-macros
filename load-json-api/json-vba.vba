Function getJSON(JsonUrl As String, key As String)
    Application.Volatile
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", JsonUrl
    MyRequest.Send
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
    
    If Json(key) = None Then
        Application.Caller.Font.ColorIndex = 3
        getJSON = "Data not found"
    Else
        Application.Caller.Font.ColorIndex = 1
        getJSON = Json(key)
    End If
            
End Function
