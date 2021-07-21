Namespace Novacode
	Public Class Bookmark
		Public Property Name() As String
		Public Property Paragraph() As Paragraph

		Public Sub SetText(ByVal newText As String)
			Paragraph.ReplaceAtBookmark(newText, Name)
		End Sub
	End Class
End Namespace
