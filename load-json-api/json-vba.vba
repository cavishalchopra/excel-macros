Function getJSON()
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    'Add json url into the below field
    MyRequest.Open "GET", "https://jsonplaceholder.typicode.com/todos/1"
    
    MyRequest.Send
    'MsgBox MyRequest.ResponseText
    
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(MyRequest.ResponseText)
            
    'MsgBox Json("title")
    getJSON = Json("title")
End Function
