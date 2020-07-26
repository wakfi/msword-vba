Sub link_to_bookmark()
	' Hyperlink the currently selected text to a bookmark of the same name (case insensitive).
	' Use this over the full line version whenever your selection doesn't automatically expand 
	' to include the line ending. Use cases include linking mentions of a word or phrase in  
	' text to a corresponding entry in a glossary or appendix. 
	' I have this bound to [ Alt + L ], you are free to choose whatever you like
    Dim regEx As New RegExp
    Dim strPattern As String: strPattern = "\s"
    With regEx
        .Global = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    Dim bookmarkName As String: bookmarkName = regEx.Replace(Replace(Trim(Selection.Text), " ", "_"), "")
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        Selection.Hyperlinks.Add Anchor:=Selection.Range, _
         Address:=ActiveDocument.FullName, SubAddress:=bookmarkName
    Else
        MsgBox "No bookmark with name " + bookmarkName
    End If
End Sub