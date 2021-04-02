'It extract only the first match
'Developed by Vishal Chopra (admin[at]vishalchopra.in)


Function RegexMatch(Text As String, Pattern As String)
    'Dim stringOne As String
    Dim regexOne As Object

    Set regexOne = New RegExp
    regexOne.Pattern = Pattern
    regexOne.Global = True
    regexOne.IgnoreCase = IgnoreCase

    Set theMatches = regexOne.Execute(Text)
    
    RegexMatch = theMatches(0)
End Function
