Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Novacode


Namespace Novacode
	Public Module ExtensionsHeadings
		<System.Runtime.CompilerServices.Extension> _
		Public Function Heading(ByVal paragraph As Paragraph, ByVal headingType As HeadingType) As Paragraph
			Dim StyleName As String = headingType.EnumDescription()
			paragraph.StyleName = StyleName
			Return paragraph
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Function EnumDescription(ByVal enumValue As [Enum]) As String
			If enumValue Is Nothing OrElse enumValue.ToString() = "0" Then
				Return String.Empty
			End If
			Dim enumInfo As FieldInfo = enumValue.GetType().GetField(enumValue.ToString())
			Dim enumAttributes() As DescriptionAttribute = CType(enumInfo.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
			If enumAttributes.Length > 0 Then
				Return enumAttributes(0).Description
			Else
				Return enumValue.ToString()
			End If
		End Function

		''' <summary>
		''' From: http://stackoverflow.com/questions/4108828/generic-extension-method-to-see-if-an-enum-contains-a-flag
		''' Check to see if a flags enumeration has a specific flag set.
		''' </summary>
		''' <param name="variable">Flags enumeration to check</param>
		''' <param name="value">Flag to check for</param>
		''' <returns></returns>
		<System.Runtime.CompilerServices.Extension> _
		Public Function HasFlag(ByVal variable As [Enum], ByVal value As [Enum]) As Boolean
			If variable Is Nothing Then
				Return False
			End If

			If value Is Nothing Then
				Throw New ArgumentNullException("value")
			End If

			' Not as good as the .NET 4 version of this function, but should be good enough
			If Not System.Enum.IsDefined(variable.GetType(), value) Then
				Throw New ArgumentException(String.Format("Enumeration type mismatch.  The flag is of type '{0}', was expecting '{1}'.", value.GetType(), variable.GetType()))
			End If

			Dim num As ULong = Convert.ToUInt64(value)
			Return ((Convert.ToUInt64(variable) And num) = num)

		End Function
	End Module

End Namespace
