Namespace Novacode
	Public Class BookmarkCollection
		Inherits List(Of Bookmark)
		Default Public ReadOnly Property Item(ByVal name As String) As Bookmark
			Get
				Return Me.FirstOrDefault(Function(bookmark) String.Equals(bookmark.Name, name, StringComparison.CurrentCultureIgnoreCase))
			End Get
		End Property
	End Class
End Namespace
