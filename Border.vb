Namespace Novacode
	''' <summary>
	''' Represents a border of a table or table cell
	''' Added by lckuiper @ 20101117
	''' </summary>
	Public Class Border
		Public Property Tcbs() As BorderStyle
		Public Property Size() As BorderSize
		Public Property Space() As Integer
		Public Property Color() As Color
		Public Sub New()
			Me.Tcbs = BorderStyle.Tcbs_single
			Me.Size = BorderSize.one
			Me.Space = 0
			Me.Color = Color.Black
		End Sub

		Public Sub New(ByVal tcbs As BorderStyle, ByVal size As BorderSize, ByVal space As Integer, ByVal color As Color)
			Me.Tcbs = tcbs
			Me.Size = size
			Me.Space = space
			Me.Color = color
		End Sub
	End Class
End Namespace