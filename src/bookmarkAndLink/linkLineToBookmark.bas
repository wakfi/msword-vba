Sub link_line_to_bookmark()
	' Hyperlink the current line to a bookmark of the same name. This is a workaround to deal with Word highlight automatically 
	' expanding to include the newline when drag selecting a line, so that the hyperlink doesn't visually extend past the text.
	' Primary use case for this is a list, for example creating a table of contents.
	' I have this bound to [ Ctrl + Shift + Alt + L ], you are free to choose whatever you like.
    Dim regEx As New RegExp
    Dim strPattern As String: strPattern = "\s"
    With regEx
        .Global = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With
    Selection.MoveStartUntil Cset:=vbNewLine, Count:=wdBackward
    Selection.Collapse
    Selection.MoveEndUntil Cset:=vbNewLine
    Dim bookmarkName As String: bookmarkName = regEx.Replace(Replace(Trim(Selection.Text), " ", "_"), "")
	If Strings.InStr(1, bookmarkName, Chr(94)) = 1 Then
        bookmarkName = Strings.Right(bookmarkName, Strings.Len(bookmarkName) - 1)
    End If
    If IsNumeric(bookmarkName) Then
        bookmarkName = "no_" + bookmarkName
    End If
    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
        Selection.Hyperlinks.Add Anchor:=Selection.Range, _
         Address:=ActiveDocument.FullName, SubAddress:=bookmarkName
    Else
        MsgBox "No bookmark with name " + bookmarkName
    End If
End Sub