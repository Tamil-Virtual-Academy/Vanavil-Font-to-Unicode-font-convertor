Namespace Novacode
	''' <summary>
	''' Represents a List in a document.
	''' </summary>
	Public Class List
		Inherits InsertBeforeOrAfter
		''' <summary>
		''' This is a list of paragraphs that will be added to the document
		''' when the list is inserted into the document.
		''' The paragraph needs a numPr defined to be in this items collection.
		''' </summary>
		Private privateItems As List(Of Paragraph)
		Public Property Items() As List(Of Paragraph)
			Get
				Return privateItems
			End Get
			Private Set(ByVal value As List(Of Paragraph))
				privateItems = value
			End Set
		End Property
		''' <summary>
		''' The numId used to reference the list settings in the numbering.xml
		''' </summary>
		Private privateNumId As Integer
		Public Property NumId() As Integer
			Get
				Return privateNumId
			End Get
			Private Set(ByVal value As Integer)
				privateNumId = value
			End Set
		End Property
		''' <summary>
		''' The ListItemType (bullet or numbered) of the list.
		''' </summary>
		Private privateListType? As ListItemType
		Public Property ListType() As ListItemType?
			Get
				Return privateListType
			End Get
			Private Set(ByVal value? As ListItemType)
				privateListType = value
			End Set
		End Property

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)
			Items = New List(Of Paragraph)()
			ListType = Nothing
		End Sub

		''' <summary>
		''' Adds an item to the list.
		''' </summary>
		''' <param name="paragraph"></param>
		''' <exception cref="InvalidOperationException">
		''' Throws an InvalidOperationException if the item cannot be added to the list.
		''' </exception>
		Public Sub AddItem(ByVal paragraph As Paragraph)
			If paragraph.IsListItem Then
				Dim numIdNode = paragraph.Xml.Descendants().First(Function(s) s.Name.LocalName = "numId")
				Dim numId = Int32.Parse(numIdNode.Attribute(DocX.w + "val").Value)

				If CanAddListItem(paragraph) Then
					Me.NumId = numId
					Items.Add(paragraph)
				Else
					Throw New InvalidOperationException("New list items can only be added to this list if they are have the same numId.")
				End If
			End If
		End Sub

		Public Sub AddItemWithStartValue(ByVal paragraph As Paragraph, ByVal start As Integer)
			'TODO: Update the numbering
			UpdateNumberingForLevelStartNumber(Integer.Parse(paragraph.IndentLevel.ToString()), start)
			If ContainsLevel(start) Then
				Throw New InvalidOperationException("Cannot add a paragraph with a start value if another element already exists in this list with that level.")
			End If
			AddItem(paragraph)
		End Sub

		Private Sub UpdateNumberingForLevelStartNumber(ByVal iLevel As Integer, ByVal start As Integer)
			Dim abstractNum = GetAbstractNum(NumId)
			Dim level = abstractNum.Descendants().First(Function(el) el.Name.LocalName = "lvl" AndAlso el.GetAttribute(DocX.w + "ilvl") = iLevel.ToString())
			level.Descendants().First(Function(el) el.Name.LocalName = "start").SetAttributeValue(DocX.w + "val", start)
		End Sub

		''' <summary>
		''' Determine if it is able to add the item to the list
		''' </summary>
		''' <param name="paragraph"></param>
		''' <returns>
		''' Return true if AddItem(...) will succeed with the given paragraph.
		''' </returns>
		Public Function CanAddListItem(ByVal paragraph As Paragraph) As Boolean
			If paragraph.IsListItem Then
				'var lvlNode = paragraph.Xml.Descendants().First(s => s.Name.LocalName == "ilvl");
				Dim numIdNode = paragraph.Xml.Descendants().First(Function(s) s.Name.LocalName = "numId")
				Dim numId = Int32.Parse(numIdNode.Attribute(DocX.w + "val").Value)

				'Level = Int32.Parse(lvlNode.Attribute(DocX.w + "val").Value);
				If Me.NumId = 0 OrElse (numId Is Me.NumId AndAlso numId > 0) Then
					Return True
				End If
			End If
			Return False
		End Function

		Public Function ContainsLevel(ByVal ilvl As Integer) As Boolean
			Return Items.Any(Function(i) i.ParagraphNumberProperties.Descendants().First(Function(el) el.Name.LocalName = "ilvl").Value = ilvl.ToString())
		End Function

		Friend Sub CreateNewNumberingNumId(Optional ByVal level As Integer = 0, Optional ByVal listType As ListItemType = ListItemType.Numbered, Optional ByVal startNumber? As Integer = Nothing, Optional ByVal continueNumbering As Boolean = False)
			ValidateDocXNumberingPartExists()
			If Document.numbering.Root Is Nothing Then
				Throw New InvalidOperationException("Numbering section did not instantiate properly.")
			End If

			Me.ListType = listType

			Dim numId = GetMaxNumId() + 1
			Dim abstractNumId = GetMaxAbstractNumId() + 1

			Dim listTemplate As XDocument
			Select Case listType
				Case ListItemType.Bulleted
					listTemplate = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_bullet_abstract.xml.gz")
				Case ListItemType.Numbered
					listTemplate = HelperFunctions.DecompressXMLResource("Novacode.Resources.numbering.default_decimal_abstract.xml.gz")
				Case Else
					Throw New InvalidOperationException(String.Format("Unable to deal with ListItemType: {0}.", listType.ToString()))
			End Select

			Dim abstractNumTemplate = listTemplate.Descendants().Single(Function(d) d.Name.LocalName = "abstractNum")
			abstractNumTemplate.SetAttributeValue(DocX.w + "abstractNumId", abstractNumId)

			'Fixing an issue where numbering would continue from previous numbered lists. Setting startOverride assures that a numbered list starts on the provided number.
			'The override needs only be on level 0 as this will cascade to the rest of the list.
			Dim abstractNumXml = GetAbstractNumXml(abstractNumId, numId, startNumber, continueNumbering)

			Dim abstractNumNode = Document.numbering.Root.Descendants().LastOrDefault(Function(xElement) xElement.Name.LocalName = "abstractNum")
			Dim numXml = Document.numbering.Root.Descendants().LastOrDefault(Function(xElement) xElement.Name.LocalName = "num")

			If abstractNumNode Is Nothing OrElse numXml Is Nothing Then
				Document.numbering.Root.Add(abstractNumTemplate)
				Document.numbering.Root.Add(abstractNumXml)
			Else
				abstractNumNode.AddAfterSelf(abstractNumTemplate)
				numXml.AddAfterSelf(abstractNumXml)
			End If

			Me.NumId = numId
		End Sub

		Private Function GetAbstractNumXml(ByVal abstractNumId As Integer, ByVal numId As Integer, ByVal startNumber? As Integer, ByVal continueNumbering As Boolean) As XElement
			'Fixing an issue where numbering would continue from previous numbered lists. Setting startOverride assures that a numbered list starts on the provided number.
			'The override needs only be on level 0 as this will cascade to the rest of the list.
			Dim startOverride = New XElement(XName.Get("startOverride", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", If(startNumber, 1)))
			Dim lvlOverride = New XElement(XName.Get("lvlOverride", DocX.w.NamespaceName), New XAttribute(DocX.w + "ilvl", 0), startOverride)
			Dim abstractNumIdElement = New XElement(XName.Get("abstractNumId", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", abstractNumId))
			Return If(continueNumbering, New XElement(XName.Get("num", DocX.w.NamespaceName), New XAttribute(DocX.w + "numId", numId), abstractNumIdElement), New XElement(XName.Get("num", DocX.w.NamespaceName), New XAttribute(DocX.w + "numId", numId), abstractNumIdElement, lvlOverride))
		End Function

		''' <summary>
		''' Method to determine the last numId for a list element. 
		''' Also useful for determining the next numId to use for inserting a new list element into the document.
		''' </summary>
		''' <returns>
		''' 0 if there are no elements in the list already.
		''' Increment the return for the next valid value of a new list element.
		''' </returns>
		Private Function GetMaxNumId() As Integer
			Const defaultValue As Integer = 0
			If Document.numbering Is Nothing Then
				Return defaultValue
			End If

			Dim numlist = Document.numbering.Descendants().Where(Function(d) d.Name.LocalName = "num").ToList()
			If numlist.Any() Then
				Return numlist.Attributes(DocX.w + "numId").Max(Function(e) Integer.Parse(e.Value))
			End If
			Return defaultValue
		End Function

		''' <summary>
		''' Method to determine the last abstractNumId for a list element.
		''' Also useful for determining the next abstractNumId to use for inserting a new list element into the document.
		''' </summary>
		''' <returns>
		''' -1 if there are no elements in the list already.
		''' Increment the return for the next valid value of a new list element.
		''' </returns>
		Private Function GetMaxAbstractNumId() As Integer
			Const defaultValue As Integer = -1

			If Document.numbering Is Nothing Then
				Return defaultValue
			End If

			Dim numlist = Document.numbering.Descendants().Where(Function(d) d.Name.LocalName = "abstractNum").ToList()
			If numlist.Any() Then
				Dim maxAbstractNumId = numlist.Attributes(DocX.w + "abstractNumId").Max(Function(e) Integer.Parse(e.Value))
				Return maxAbstractNumId
			End If
			Return defaultValue
		End Function

		''' <summary>
		''' Get the abstractNum definition for the given numId
		''' </summary>
		''' <param name="numId">The numId on the pPr element</param>
		''' <returns>XElement representing the requested abstractNum</returns>
		Friend Function GetAbstractNum(ByVal numId As Integer) As XElement
			Dim num = Document.numbering.Descendants().First(Function(d) d.Name.LocalName = "num" AndAlso d.GetAttribute(DocX.w + "numId").Equals(numId.ToString()))
			Dim abstractNumId = num.Descendants().First(Function(d) d.Name.LocalName = "abstractNumId")
			Return Document.numbering.Descendants().First(Function(d) d.Name.LocalName = "abstractNum" AndAlso d.GetAttribute("abstractNumId").Equals(abstractNumId.Value))
		End Function

		Private Sub ValidateDocXNumberingPartExists()
			Dim numberingUri = New Uri("/word/numbering.xml", UriKind.Relative)

			' If the internal document contains no /word/numbering.xml create one.
			If Not Document.package.PartExists(numberingUri) Then
				Document.numbering = HelperFunctions.AddDefaultNumberingXml(Document.package)
			End If
		End Sub
	End Class
End Namespace
