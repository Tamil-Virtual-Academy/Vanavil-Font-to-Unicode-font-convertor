Imports System.IO
Imports System.Xml

Namespace Novacode
	''' <summary>
	''' Represents a table of contents in the document
	''' </summary>
	Public Class TableOfContents
		Inherits DocXElement
		#Region "TocBaseValues"

		Private Const HeaderStyle As String = "TOCHeading"
		Private Const RightTabPos As Integer = 9350
		#End Region

		Private Sub New(ByVal document As DocX, ByVal xml As XElement, ByVal headerStyle As String)
			MyBase.New(document, xml)
			AssureUpdateField(document)
			AssureStyles(document, headerStyle)
		End Sub

		Friend Shared Function CreateTableOfContents(ByVal document As DocX, ByVal title As String, ByVal switches As TableOfContentsSwitches, Optional ByVal headerStyle As String = Nothing, Optional ByVal lastIncludeLevel As Integer = 3, Optional ByVal rightTabPos? As Integer = Nothing) As TableOfContents
			Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocXmlBase, If(headerStyle, TableOfContents.HeaderStyle), title, If(rightTabPos, TableOfContents.RightTabPos), BuildSwitchString(switches, lastIncludeLevel))))
			Dim xml = XElement.Load(reader)
			Return New TableOfContents(document, xml, headerStyle)
		End Function

		Private Sub AssureUpdateField(ByVal document As DocX)
			If document.settings.Descendants().Any(Function(x) x.Name.Equals(DocX.w + "updateFields")) Then
				Return
			End If

			Dim element = New XElement(XName.Get("updateFields", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", True))
			document.settings.Root.Add(element)
		End Sub

		Private Sub AssureStyles(ByVal document As DocX, ByVal headerStyle As String)
			If Not HasStyle(document, headerStyle, "paragraph") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocHeadingStyleBase, If(headerStyle, Me.HeaderStyle))))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
			If Not HasStyle(document, "TOC1", "paragraph") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocElementStyleBase, "TOC1", "toc 1")))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
			If Not HasStyle(document, "TOC2", "paragraph") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocElementStyleBase, "TOC2", "toc 2")))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
			If Not HasStyle(document, "TOC3", "paragraph") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocElementStyleBase, "TOC3", "toc 3")))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
			If Not HasStyle(document, "TOC4", "paragraph") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocElementStyleBase, "TOC4", "toc 4")))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
			If Not HasStyle(document, "Hyperlink", "character") Then
				Dim reader = XmlReader.Create(New StringReader(String.Format(XmlTemplateBases.TocHyperLinkStyleBase)))
				Dim xml = XElement.Load(reader)
				document.styles.Root.Add(xml)
			End If
		End Sub

		Private Function HasStyle(ByVal document As DocX, ByVal value As String, ByVal type As String) As Boolean
			Return document.styles.Descendants().Any(Function(x) x.Name.Equals(DocX.w + "style") AndAlso (x.Attribute(DocX.w + "type") Is Nothing OrElse x.Attribute(DocX.w + "type").Value.Equals(type)) AndAlso x.Attribute(DocX.w + "styleId") IsNot Nothing AndAlso x.Attribute(DocX.w + "styleId").Value.Equals(value))
		End Function

		Private Shared Function BuildSwitchString(ByVal switches As TableOfContentsSwitches, ByVal lastIncludeLevel As Integer) As String
			Dim allSwitches = System.Enum.GetValues(GetType(TableOfContentsSwitches)).Cast(Of TableOfContentsSwitches)()
			Dim switchString = "TOC"
			For Each s In allSwitches.Where(Function(s) s <> TableOfContentsSwitches.None AndAlso switches.HasFlag(s))
				switchString &= " " & s.EnumDescription()
				If s = TableOfContentsSwitches.O Then
					switchString &= String.Format(" '{0}-{1}'", 1, lastIncludeLevel)
				End If
			Next s

			Return switchString
		End Function

	End Class
End Namespace