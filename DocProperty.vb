Imports System.Text.RegularExpressions

Namespace Novacode
	''' <summary>
	''' Represents a field of type document property. This field displays the value stored in a custom property.
	''' </summary>
	Public Class DocProperty
		Inherits DocXElement
		Friend extractName As New Regex("DOCPROPERTY  (?<name>.*)  ")
'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not allow class members with the same name:
		Private name_Renamed As String

		''' <summary>
		''' The custom property to display.
		''' </summary>
		Public ReadOnly Property Name() As String
			Get
				Return name_Renamed
			End Get
		End Property

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)
			Dim instr As String = Me.Xml.Attribute(XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Value
			Me.name_Renamed = extractName.Match(instr.Trim()).Groups("name").Value
		End Sub
	End Class
End Namespace
