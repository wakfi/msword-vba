Sub bookmark_this()
	' Create a bookmark from the selected text, using the text as the name and location
	' I have this bound to [ Ctrl + Alt + B ], you are free to choose whatever you like
    Dim regEx As New RegExp
    Dim strPattern As String: strPattern = "\s"
    With regEx
        .Global = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    Dim bookmarkName As String: bookmarkName = regEx.Replace(Replace(Trim(Selection.Text), " ", "_"), "")
    If IsNumeric(Selection.Text) Then
        bookmarkName = "no_" + bookmarkName
    End If
    ActiveDocument.Bookmarks.Add (bookmarkName)
End Sub