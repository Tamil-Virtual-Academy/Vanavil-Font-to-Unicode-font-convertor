Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Text

Namespace Novacode
	Friend Module Extensions
		<System.Runtime.CompilerServices.Extension> _
		Friend Function ToHex(ByVal source As Color) As String
			Dim red As Byte = source.R
			Dim green As Byte = source.G
			Dim blue As Byte = source.B

			Dim redHex As String = red.ToString("X")
			If redHex.Length < 2 Then
				redHex = "0" & redHex
			End If

			Dim blueHex As String = blue.ToString("X")
			If blueHex.Length < 2 Then
				blueHex = "0" & blueHex
			End If

			Dim greenHex As String = green.ToString("X")
			If greenHex.Length < 2 Then
				greenHex = "0" & greenHex
			End If

			Return String.Format("{0}{1}{2}", redHex, greenHex, blueHex)
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Sub Flatten(ByVal e As XElement, ByVal name As XName, ByVal flat As List(Of XElement))
			' Add this element (without its children) to the flat list.
			Dim clone As XElement = CloneElement(e)
			clone.Elements().Remove()

			' Filter elements using XName.
			If clone.Name = name Then
				flat.Add(clone)
			End If

			' Process the children.
			If e.HasElements Then
				For Each elem As XElement In e.Elements(name) ' Filter elements using XName
					elem.Flatten(name, flat)
				Next elem
			End If
		End Sub

		Private Function CloneElement(ByVal element As XElement) As XElement
			Return New XElement(element.Name, element.Attributes(), element.Nodes().Select(Function(n)
				Dim e As XElement = TryCast(n, XElement)
				If e IsNot Nothing Then
					Return CloneElement(e)
				End If
					Return n
			End Function))
		End Function

		<System.Runtime.CompilerServices.Extension> _
		Public Function GetAttribute(ByVal el As XElement, ByVal name As XName, Optional ByVal defaultValue As String = "") As String
			Dim attr = el.Attribute(name)
			If attr IsNot Nothing Then
				Return attr.Value
			End If
			Return defaultValue
		End Function

		''' <summary>
		''' Sets margin for all the pages in a Dox document in Inches. (Written by Shashwat Tripathi)
		''' </summary>
		''' <param name="document"></param>
		''' <param name="top">Margin from the Top. Leave -1 for no change</param>
		''' <param name="bottom">Margin from the Bottom. Leave -1 for no change</param>
		''' <param name="right">Margin from the Right. Leave -1 for no change</param>
		''' <param name="left">Margin from the Left. Leave -1 for no change</param>
		<System.Runtime.CompilerServices.Extension> _
		Public Sub SetMargin(ByVal document As DocX, ByVal top As Single, ByVal bottom As Single, ByVal right As Single, ByVal left As Single)
			Dim ab As XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
			Dim tempElement = document.PageLayout.Xml.Descendants(ab + "pgMar")
			Dim e = tempElement.GetEnumerator()

			For Each item In tempElement
				If left <> -1 Then
					item.SetAttributeValue(ab + "left", (1440 * left) / 1)
				End If
				If right <> -1 Then
					item.SetAttributeValue(ab + "right", (1440 * right) / 1)
				End If
				If top <> -1 Then
					item.SetAttributeValue(ab + "top", (1440 * top) / 1)
				End If
				If bottom <> -1 Then
					item.SetAttributeValue(ab + "bottom", (1440 * bottom) / 1)
				End If
			Next item
		End Sub
	End Module

End Namespace
