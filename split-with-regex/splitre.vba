Public Function SplitRe(Text As String, Pattern As String, Optional IgnoreCase As Boolean) As String()
    Static re As Object

    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.MultiLine = True
    End If

    re.IgnoreCase = IgnoreCase
    re.Pattern = Pattern
    SplitRe = Strings.Split(re.Replace(Text, ChrW(-1)), ChrW(-1))
End Function
