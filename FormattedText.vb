Namespace Novacode
	Public Class FormattedText
		Implements IComparable
		Public Sub New()

		End Sub

		Public index As Integer
		Public text As String
		Public formatting As Formatting

		Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
			Dim other As FormattedText = CType(obj, FormattedText)
			Dim tf As FormattedText = Me

			If other.formatting Is Nothing OrElse tf.formatting Is Nothing Then
				Return -1
			End If

			Return tf.formatting.CompareTo(other.formatting)
		End Function
	End Class
End Namespace
