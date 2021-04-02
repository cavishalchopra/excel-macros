Sub RegexMatch()
    Dim stringOne As String
    Dim regexOne As Object
    Dim Text As String

    'Get value from the cell A1
    Text = Worksheets("Sheet1").Range("A1").Value

    Set regexOne = New RegExp
    regexOne.Pattern = "[0-9]{8}"
    regexOne.Global = True
    regexOne.IgnoreCase = IgnoreCase

    Set theMatches = regexOne.Execute(Text)

    For Each Match In theMatches
      Debug.Print Match.Value
    Next
End Sub
