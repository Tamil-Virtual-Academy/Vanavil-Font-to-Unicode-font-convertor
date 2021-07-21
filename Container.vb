Imports System.IO
Imports System.IO.Packaging
Imports System.Text.RegularExpressions
Imports System.Collections.ObjectModel

Namespace Novacode
	Public MustInherit Class Container
		Inherits DocXElement
		''' <summary>
		''' Returns a list of all Paragraphs inside this container.
		''' </summary>
		''' <example>
		''' <code>
		'''  Load a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // All Paragraphs in this document.
		'''    <![CDATA[ List<Paragraph> ]]> documentParagraphs = document.Paragraphs;
		'''    
		'''    // Make sure this document contains at least one Table.
		'''    if (document.Tables.Count() > 0)
		'''    {
		'''        // Get the first Table in this document.
		'''        Table t = document.Tables[0];
		'''
		'''        // All Paragraphs in this Table.
		'''        <![CDATA[ List<Paragraph> ]]> tableParagraphs = t.Paragraphs;
		'''    
		'''        // Make sure this Table contains at least one Row.
		'''        if (t.Rows.Count() > 0)
		'''        {
		'''            // Get the first Row in this document.
		'''            Row r = t.Rows[0];
		'''
		'''            // All Paragraphs in this Row.
		'''            <![CDATA[ List<Paragraph> ]]> rowParagraphs = r.Paragraphs;
		'''
		'''            // Make sure this Row contains at least one Cell.
		'''            if (r.Cells.Count() > 0)
		'''            {
		'''                // Get the first Cell in this document.
		'''                Cell c = r.Cells[0];
		'''
		'''                // All Paragraphs in this Cell.
		'''                <![CDATA[ List<Paragraph> ]]> cellParagraphs = c.Paragraphs;
		'''            }
		'''        }
		'''    }
		'''
		'''    // Save all changes to this document.
		'''    document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overridable ReadOnly Property Paragraphs() As ReadOnlyCollection(Of Paragraph)
			Get
'INSTANT VB NOTE: The local variable paragraphs was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim paragraphs_Renamed As List(Of Paragraph) = GetParagraphs()

				For Each p In paragraphs_Renamed
					If (p.Xml.ElementsAfterSelf().FirstOrDefault() IsNot Nothing) AndAlso (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl")) Then
						p.FollowingTable = New Table(Me.Document, p.Xml.ElementsAfterSelf().First())
					End If

					p.ParentContainer = GetParentFromXmlName(p.Xml.Ancestors().First().Name.LocalName)

					If p.IsListItem Then
						GetListItemType(p)
					End If
				Next p

				Return paragraphs_Renamed.AsReadOnly()
			End Get
		End Property
		Public Overridable ReadOnly Property ParagraphsDeepSearch() As ReadOnlyCollection(Of Paragraph)
			Get
				Dim paragraphs As List(Of Paragraph) = GetParagraphs(True)

				For Each p In paragraphs
					If (p.Xml.ElementsAfterSelf().FirstOrDefault() IsNot Nothing) AndAlso (p.Xml.ElementsAfterSelf().First().Name.Equals(DocX.w + "tbl")) Then
						p.FollowingTable = New Table(Me.Document, p.Xml.ElementsAfterSelf().First())
					End If

					p.ParentContainer = GetParentFromXmlName(p.Xml.Ancestors().First().Name.LocalName)

					If p.IsListItem Then
						GetListItemType(p)
					End If
				Next p

				Return paragraphs.AsReadOnly()
			End Get
		End Property
		''' <summary>
		''' Removes paragraph at specified position
		''' </summary>
		''' <param name="index">Index of paragraph to remove</param>
		''' <returns>True if removed</returns>
		Public Function RemoveParagraphAt(ByVal index As Integer) As Boolean
			Dim i As Integer = 0
			For Each paragraph In Xml.Descendants(DocX.w + "p")
				If i = index Then
					paragraph.Remove()
					Return True
				End If
				i += 1

			Next paragraph

			Return False
		End Function

		''' <summary>
		''' Removes paragraph
		''' </summary>
		''' <param name="p">Paragraph to remove</param>
		''' <returns>True if removed</returns>
		Public Function RemoveParagraph(ByVal p As Paragraph) As Boolean
			For Each paragraph In Xml.Descendants(DocX.w + "p")
				If paragraph.Equals(p.Xml) Then
					paragraph.Remove()
					Return True
				End If
			Next paragraph

			Return False
		End Function


		Public Overridable ReadOnly Property Sections() As List(Of Section)
			Get
				Dim allParas = Paragraphs

				Dim parasInASection = New List(Of Paragraph)()
'INSTANT VB NOTE: The local variable sections was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim sections_Renamed = New List(Of Section)()

				For Each para In allParas

					Dim sectionInPara = para.Xml.Descendants().FirstOrDefault(Function(s) s.Name.LocalName = "sectPr")

					If sectionInPara Is Nothing Then
						parasInASection.Add(para)
					Else
						parasInASection.Add(para)
						Dim section = New Section(Document, sectionInPara) With {.SectionParagraphs = parasInASection}
						sections_Renamed.Add(section)
						parasInASection = New List(Of Paragraph)()
					End If

				Next para

				Dim body As XElement = Xml.Element(XName.Get("body", DocX.w.NamespaceName))
				Dim baseSectionXml As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
				Dim baseSection = New Section(Document, baseSectionXml) With {.SectionParagraphs = parasInASection}
				sections_Renamed.Add(baseSection)

				Return sections_Renamed
			End Get
		End Property


		Private Sub GetListItemType(ByVal p As Paragraph)
			Dim ilvlNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(Function(el) el.Name.LocalName = "ilvl")
			Dim ilvlValue = ilvlNode.Attribute(DocX.w + "val").Value

			Dim numIdNode = p.ParagraphNumberProperties.Descendants().FirstOrDefault(Function(el) el.Name.LocalName = "numId")
			Dim numIdValue = numIdNode.Attribute(DocX.w + "val").Value

			'find num node in numbering 
			Dim numNodes = Document.numbering.Descendants().Where(Function(n) n.Name.LocalName = "num")
			Dim numNode As XElement = numNodes.FirstOrDefault(Function(node) node.Attribute(DocX.w + "numId").Value.Equals(numIdValue))

			If numNode IsNot Nothing Then
			   'Get abstractNumId node and its value from numNode
				Dim abstractNumIdNode = numNode.Descendants().First(Function(n) n.Name.LocalName = "abstractNumId")
				Dim abstractNumNodeValue = abstractNumIdNode.Attribute(DocX.w + "val").Value

				Dim abstractNumNodes = Document.numbering.Descendants().Where(Function(n) n.Name.LocalName = "abstractNum")
				Dim abstractNumNode As XElement = abstractNumNodes.FirstOrDefault(Function(node) node.Attribute(DocX.w + "abstractNumId").Value.Equals(abstractNumNodeValue))

				'Find lvl node
				Dim lvlNodes = abstractNumNode.Descendants().Where(Function(n) n.Name.LocalName = "lvl")
				Dim lvlNode As XElement = Nothing
				For Each node As XElement In lvlNodes
					If node.Attribute(DocX.w + "ilvl").Value.Equals(ilvlValue) Then
						lvlNode = node
						Exit For
					End If
				Next node

				   Dim numFmtNode = lvlNode.Descendants().First(Function(n) n.Name.LocalName = "numFmt")
					  p.ListItemType = GetListItemType(numFmtNode.Attribute(DocX.w + "val").Value)
			End If

		End Sub


		Public ParentContainer As ContainerType


		Friend Function GetParagraphs(Optional ByVal deepSearch As Boolean =False) As List(Of Paragraph)
			' Need some memory that can be updated by the recursive search.
			Dim index As Integer = 0
			Dim paragraphs As New List(Of Paragraph)()

			GetParagraphsRecursive(Xml, index, paragraphs, deepSearch)

			Return paragraphs
		End Function

		Friend Sub GetParagraphsRecursive(ByVal Xml As XElement, ByRef index As Integer, ByRef paragraphs As List(Of Paragraph), Optional ByVal deepSearch As Boolean =False)
			' sdtContent are for PageNumbers inside Headers or Footers, don't go any deeper.
			'if (Xml.Name.LocalName == "sdtContent")
			'    return;
			Dim keepSearching = True
			If Xml.Name.LocalName = "p" Then
				paragraphs.Add(New Paragraph(Document, Xml, index))

				index += HelperFunctions.GetText(Xml).Length
				If Not deepSearch Then
					keepSearching = False
				End If
			End If
			If keepSearching AndAlso Xml.HasElements Then
				For Each e As XElement In Xml.Elements()
					GetParagraphsRecursive(e, index, paragraphs, deepSearch)
				Next e
			End If
		End Sub

		Public Overridable ReadOnly Property Tables() As List(Of Table)
			Get
'INSTANT VB NOTE: The local variable tables was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim tables_Renamed As List(Of Table) = (
				    From t In Xml.Descendants(DocX.w + "tbl")
				    Select New Table(Document, t)).ToList()

				Return tables_Renamed
			End Get
		End Property

		Public Overridable ReadOnly Property Lists() As List(Of List)
			Get
'INSTANT VB NOTE: The local variable lists was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim lists_Renamed = New List(Of List)()
				Dim list = New List(Document, Xml)

				For Each paragraph In Paragraphs
					If paragraph.IsListItem Then
						If list.CanAddListItem(paragraph) Then
							list.AddItem(paragraph)
						Else
							lists_Renamed.Add(list)
							list = New List(Document, Xml)
							list.AddItem(paragraph)
						End If
					End If
				Next paragraph

				lists_Renamed.Add(list)

				Return lists_Renamed
			End Get
		End Property

		Public Overridable ReadOnly Property Hyperlinks() As List(Of Hyperlink)
			Get
'INSTANT VB NOTE: The local variable hyperlinks was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim hyperlinks_Renamed As New List(Of Hyperlink)()

				For Each p As Paragraph In Paragraphs
					hyperlinks_Renamed.AddRange(p.Hyperlinks)
				Next p

				Return hyperlinks_Renamed
			End Get
		End Property

		Public Overridable ReadOnly Property Pictures() As List(Of Picture)
			Get
'INSTANT VB NOTE: The local variable pictures was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim pictures_Renamed As New List(Of Picture)()

				For Each p As Paragraph In Paragraphs
					pictures_Renamed.AddRange(p.Pictures)
				Next p

				Return pictures_Renamed
			End Get
		End Property

		''' <summary>
		''' Sets the Direction of content.
		''' </summary>
		''' <param name="direction">Direction either LeftToRight or RightToLeft</param>
		''' <example>
		''' Set the Direction of content in a Paragraph to RightToLeft.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // Get the first Paragraph from this document.
		'''    Paragraph p = document.InsertParagraph();
		'''
		'''    // Set the Direction of this Paragraph.
		'''    p.Direction = Direction.RightToLeft;
		'''
		'''    // Make sure the document contains at lest one Table.
		'''    if (document.Tables.Count() > 0)
		'''    {
		'''        // Get the first Table from this document.
		'''        Table t = document.Tables[0];
		'''
		'''        /* 
		'''         * Set the direction of the entire Table.
		'''         * Note: The same function is available at the Row and Cell level.
		'''         */
		'''        t.SetDirection(Direction.RightToLeft);
		'''    }
		'''
		'''    // Save all changes to this document.
		'''    document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overridable Sub SetDirection(ByVal direction As Direction)
			For Each p As Paragraph In Paragraphs
				p.Direction = direction
			Next p
		End Sub

		Public Overridable Function FindAll(ByVal str As String) As List(Of Integer)
			Return FindAll(str, RegexOptions.None)
		End Function

		Public Overridable Function FindAll(ByVal str As String, ByVal options As RegexOptions) As List(Of Integer)
			Dim list As New List(Of Integer)()

			For Each p As Paragraph In Paragraphs
				Dim indexes As List(Of Integer) = p.FindAll(str, options)

				For i As Integer = 0 To indexes.Count() - 1
					indexes(0) += p.startIndex
				Next i

				list.AddRange(indexes)
			Next p

			Return list
		End Function

		''' <summary>
		''' Find all unique instances of the given Regex Pattern,
		''' returning the list of the unique strings found
		''' </summary>
		''' <param name="pattern"></param>
		''' <param name="options"></param>
		''' <returns></returns>
		Public Overridable Function FindUniqueByPattern(ByVal pattern As String, ByVal options As RegexOptions) As List(Of String)
			Dim rawResults As New List(Of String)()

			For Each p As Paragraph In Paragraphs
				Dim partials As List(Of String) = p.FindAllByPattern(pattern, options)
				rawResults.AddRange(partials)
			Next p

			' this dictionary is used to collect results and test for uniqueness
			Dim uniqueResults As New Dictionary(Of String, Integer)()

			For Each currValue As String In rawResults
				If Not uniqueResults.ContainsKey(currValue) Then
					uniqueResults.Add(currValue, 0)
				End If
			Next currValue

			Return uniqueResults.Keys.ToList() ' return the unique list of results
		End Function

		Public Overridable Sub ReplaceText(ByVal searchValue As String, ByVal newValue As String, Optional ByVal trackChanges As Boolean = False, Optional ByVal options As RegexOptions = RegexOptions.None, Optional ByVal newFormatting As Formatting = Nothing, Optional ByVal matchFormatting As Formatting = Nothing, Optional ByVal formattingOptions As MatchFormattingOptions = MatchFormattingOptions.SubsetMatch, Optional ByVal escapeRegEx As Boolean = True, Optional ByVal useRegExSubstitutions As Boolean = False)
			If String.IsNullOrEmpty(searchValue) Then
				Throw New ArgumentException("oldValue cannot be null or empty", "searchValue")
			End If

			If newValue Is Nothing Then
				Throw New ArgumentException("newValue cannot be null or empty", "newValue")
			End If
			' ReplaceText in Headers of the document.
			Dim headerList = New List(Of Header) From {Document.Headers.first, Document.Headers.even, Document.Headers.odd}
			For Each header In headerList
				If header IsNot Nothing Then
					For Each paragraph In header.Paragraphs
						paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions)
					Next paragraph
				End If
			Next header

			' ReplaceText int main body of document.
			For Each paragraph In Paragraphs
				paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions)
			Next paragraph

			' ReplaceText in Footers of the document.
			Dim footerList = New List(Of Footer) From {Document.Footers.first, Document.Footers.even, Document.Footers.odd}
			For Each footer In footerList
				If footer IsNot Nothing Then
					For Each paragraph In footer.Paragraphs
						paragraph.ReplaceText(searchValue, newValue, trackChanges, options, newFormatting, matchFormatting, formattingOptions, escapeRegEx, useRegExSubstitutions)
					Next paragraph
				End If
			Next footer
		End Sub

		''' <summary>
		''' 
		''' </summary>
		''' <param name="searchValue">Value to find</param>
		''' <param name="regexMatchHandler">A Func that accepts the matching regex search group value and passes it to this to return the replacement string</param>
		''' <param name="trackChanges">Enable trackchanges</param>
		''' <param name="options">Regex options</param>
		''' <param name="newFormatting"></param>
		''' <param name="matchFormatting"></param>
		''' <param name="formattingOptions"></param>
		Public Overridable Sub ReplaceText(ByVal searchValue As String, ByVal regexMatchHandler As Func(Of String,String), Optional ByVal trackChanges As Boolean = False, Optional ByVal options As RegexOptions = RegexOptions.None, Optional ByVal newFormatting As Formatting = Nothing, Optional ByVal matchFormatting As Formatting = Nothing, Optional ByVal formattingOptions As MatchFormattingOptions = MatchFormattingOptions.SubsetMatch)
			If String.IsNullOrEmpty(searchValue) Then
				Throw New ArgumentException("oldValue cannot be null or empty", "searchValue")
			End If

			If regexMatchHandler Is Nothing Then
				Throw New ArgumentException("regexMatchHandler cannot be null", "regexMatchHandler")
			End If

			' ReplaceText in Headers/Footers of the document.
			Dim containerList = New List(Of IParagraphContainer) From {Document.Headers.first, Document.Headers.even, Document.Headers.odd, Document.Footers.first, Document.Footers.even, Document.Footers.odd}
			For Each container In containerList
				If container IsNot Nothing Then
					For Each paragraph In container.Paragraphs
						paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions)
					Next paragraph
				End If
			Next container

			' ReplaceText int main body of document.
			For Each paragraph In Paragraphs
				paragraph.ReplaceText(searchValue, regexMatchHandler, trackChanges, options, newFormatting, matchFormatting, formattingOptions)
			Next paragraph
		End Sub

		''' <summary>
		''' Removes all items with required formatting
		''' </summary>
		''' <returns>Numer of texts removed</returns>
		Public Function RemoveTextInGivenFormat(ByVal matchFormatting As Formatting, Optional ByVal fo As MatchFormattingOptions = MatchFormattingOptions.SubsetMatch) As Integer
			Dim deletedCount = 0
			For Each x In Xml.Elements()
				deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo)
			Next x

			Return deletedCount
		End Function

		Friend Function RemoveTextWithFormatRecursive(ByVal element As XElement, ByVal matchFormatting As Formatting, ByVal fo As MatchFormattingOptions) As Integer
			Dim deletedCount = 0
			For Each x In element.Elements()
				If "rPr".Equals(x.Name.LocalName) Then
					If HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, x, fo) Then
						x.Parent.Remove()
						deletedCount += 1
					End If
				End If

				deletedCount += RemoveTextWithFormatRecursive(x, matchFormatting, fo)
			Next x

			Return deletedCount
		End Function

		Public Overridable Sub InsertAtBookmark(ByVal toInsert As String, ByVal bookmarkName As String)
			If bookmarkName.IsNullOrWhiteSpace() Then
				Throw New ArgumentException("bookmark cannot be null or empty", "bookmarkName")
			End If

			Dim headerCollection = Document.Headers
			Dim headers = New List(Of Header) From {headerCollection.first, headerCollection.even, headerCollection.odd}
			For Each header In headers.Where(Function(x) x IsNot Nothing)
				For Each paragraph In header.Paragraphs
					paragraph.InsertAtBookmark(toInsert, bookmarkName)
				Next paragraph
			Next header

			For Each paragraph In Paragraphs
				paragraph.InsertAtBookmark(toInsert, bookmarkName)
			Next paragraph

			Dim footerCollection = Document.Footers
			Dim footers = New List(Of Footer) From {footerCollection.first, footerCollection.even, footerCollection.odd}
			For Each footer In footers.Where(Function(x) x IsNot Nothing)
				For Each paragraph In footer.Paragraphs
					paragraph.InsertAtBookmark(toInsert, bookmarkName)
				Next paragraph
			Next footer
		End Sub

		Public Function ValidateBookmarks(ParamArray ByVal bookmarkNames() As String) As String()
			Dim headers = {Document.Headers.first, Document.Headers.even, Document.Headers.odd}.Where(Function(h) h IsNot Nothing).ToList()
			Dim footers = {Document.Footers.first, Document.Footers.even, Document.Footers.odd}.Where(Function(f) f IsNot Nothing).ToList()

			Dim nonMatching = New List(Of String)()
			For Each bookmarkName In bookmarkNames
				If headers.SelectMany(Function(h) h.Paragraphs).Any(Function(p) p.ValidateBookmark(bookmarkName)) Then
					Return New String(){}
				End If
				If footers.SelectMany(Function(h) h.Paragraphs).Any(Function(p) p.ValidateBookmark(bookmarkName)) Then
					Return New String(){}
				End If
				If Paragraphs.Any(Function(p) p.ValidateBookmark(bookmarkName)) Then
					Return New String(){}
				End If
				nonMatching.Add(bookmarkName)
			Next bookmarkName

			Return nonMatching.ToArray()
		End Function

		Public Overridable Function InsertParagraph(ByVal index As Integer, ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Return InsertParagraph(index, text, trackChanges, Nothing)
		End Function

		Public Overridable Function InsertParagraph() As Paragraph
			Return InsertParagraph(String.Empty, False)
		End Function

		Public Overridable Function InsertParagraph(ByVal index As Integer, ByVal p As Paragraph) As Paragraph
			Dim newXElement As New XElement(p.Xml)
			p.Xml = newXElement

			Dim paragraph As Paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index)

			If paragraph Is Nothing Then
				Xml.Add(p.Xml)
			Else
				Dim split() As XElement = HelperFunctions.SplitParagraph(paragraph, index - paragraph.startIndex)

				paragraph.Xml.ReplaceWith (split(0), newXElement, split(1))
			End If

			GetParent(p)

			Return p
		End Function

		Public Overridable Function InsertParagraph(ByVal p As Paragraph) As Paragraph
'			#Region "Styles"
			Dim style_document As XDocument

			If p.styles.Count() > 0 Then
				Dim style_package_uri As New Uri("/word/styles.xml", UriKind.Relative)
				If Not Document.package.PartExists(style_package_uri) Then
					Dim style_package As PackagePart = Document.package.CreatePart(style_package_uri, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum)
					Using tw As TextWriter = New StreamWriter(style_package.GetStream())
						style_document = New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), New XElement(XName.Get("styles", DocX.w.NamespaceName)))

						style_document.Save(tw)
					End Using
				End If

				Dim styles_document As PackagePart = Document.package.GetPart(style_package_uri)
				Using tr As TextReader = New StreamReader(styles_document.GetStream())
					style_document = XDocument.Load(tr)
					Dim styles_element As XElement = style_document.Element(XName.Get("styles", DocX.w.NamespaceName))

					Dim ids = From d In styles_element.Descendants(XName.Get("style", DocX.w.NamespaceName))
					          Let a = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
					          Where a IsNot Nothing
					          Select a.Value

					For Each style As XElement In p.styles
						' If styles_element does not contain this element, then add it.

						If Not ids.Contains(style.Attribute(XName.Get("styleId", DocX.w.NamespaceName)).Value) Then
							styles_element.Add(style)
						End If
					Next style
				End Using

				Using tw As TextWriter = New StreamWriter(styles_document.GetStream())
					style_document.Save(tw)
				End Using
			End If
'			#End Region

			Dim newXElement As New XElement(p.Xml)

			Xml.Add(newXElement)

			Dim index As Integer = 0
			If Document.paragraphLookup.Keys.Count() > 0 Then
				index = Document.paragraphLookup.Last().Key

				If Document.paragraphLookup.Last().Value.Text.Length = 0 Then
					index += 1
				Else
					index += Document.paragraphLookup.Last().Value.Text.Length
				End If
			End If

			Dim newParagraph As New Paragraph(Document, newXElement, index)
			Document.paragraphLookup.Add(index, newParagraph)

			GetParent(newParagraph)

			Return newParagraph
		End Function

		Public Overridable Function InsertParagraph(ByVal index As Integer, ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim newParagraph As New Paragraph(Document, New XElement(DocX.w + "p"), index)
			newParagraph.InsertText(0, text, trackChanges, formatting)

			Dim firstPar As Paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index)

			If firstPar IsNot Nothing Then
				Dim splitindex = index - firstPar.startIndex
				If splitindex <= 0 Then
					firstPar.Xml.ReplaceWith(newParagraph.Xml, firstPar.Xml)
				Else
					Dim splitParagraph() As XElement = HelperFunctions.SplitParagraph(firstPar, splitindex)

					firstPar.Xml.ReplaceWith (splitParagraph(0), newParagraph.Xml, splitParagraph(1))
				End If

			Else
				Xml.Add(newParagraph)
			End If

			GetParent(newParagraph)

			Return newParagraph
		End Function


		Private Function GetParentFromXmlName(ByVal xmlName As String) As ContainerType
			Dim parent As ContainerType

			Select Case xmlName
				Case "body"
					parent = ContainerType.Body
				Case "p"
					parent = ContainerType.Paragraph
				Case "tbl"
					parent = ContainerType.Table
				Case "sectPr"
					parent = ContainerType.Section
				Case "tc"
					parent = ContainerType.Cell
				Case Else
					parent = ContainerType.None
			End Select
			Return parent
		End Function

		Private Sub GetParent(ByVal newParagraph As Paragraph)
			Dim containerType = Me.GetType()

			Select Case containerType.Name

				Case "Body"
					newParagraph.ParentContainer = ContainerType.Body
				Case "Table"
					newParagraph.ParentContainer = ContainerType.Table
				Case "TOC"
					newParagraph.ParentContainer = ContainerType.TOC
				Case "Section"
					newParagraph.ParentContainer = ContainerType.Section
				Case "Cell"
					newParagraph.ParentContainer = ContainerType.Cell
				Case "Header"
					newParagraph.ParentContainer = ContainerType.Header
				Case "Footer"
					newParagraph.ParentContainer = ContainerType.Footer
				Case "Paragraph"
					newParagraph.ParentContainer = ContainerType.Paragraph
			End Select
		End Sub


		Private Function GetListItemType(ByVal styleName As String) As ListItemType
			Dim listItemType As ListItemType

			Select Case styleName
				Case "bullet"
					listItemType = ListItemType.Bulleted
				Case Else
					listItemType = ListItemType.Numbered
			End Select

			Return listItemType
		End Function



		Public Overridable Sub InsertSection()
			InsertSection(False)
		End Sub

		Public Overridable Sub InsertSection(ByVal trackChanges As Boolean)
			Dim newParagraphSection = New XElement (XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName), New XElement(XName.Get("sectPr", DocX.w.NamespaceName), New XElement(XName.Get("type", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", "continuous")))))

			If trackChanges Then
				newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, Date.Now, newParagraphSection)
			End If

			Xml.Add(newParagraphSection)
		End Sub

		Public Overridable Sub InsertSectionPageBreak(Optional ByVal trackChanges As Boolean = False)
			Dim newParagraphSection = New XElement (XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName), New XElement(XName.Get("sectPr", DocX.w.NamespaceName))))

			If trackChanges Then
				newParagraphSection = HelperFunctions.CreateEdit(EditType.ins, Date.Now, newParagraphSection)
			End If

			Xml.Add(newParagraphSection)
		End Sub

		Public Overridable Function InsertParagraph(ByVal text As String) As Paragraph
			Return InsertParagraph(text, False, New Formatting())
		End Function

		Public Overridable Function InsertParagraph(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Return InsertParagraph(text, trackChanges, New Formatting())
		End Function

		Public Overridable Function InsertParagraph(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim newParagraph As New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml))

			If trackChanges Then
				newParagraph = HelperFunctions.CreateEdit(EditType.ins, Date.Now, newParagraph)
			End If
			Xml.Add(newParagraph)
			Dim paragraphAdded = New Paragraph(Document, newParagraph, 0)
			If TypeOf Me Is Cell Then
				Dim cell = TryCast(Me, Cell)
				paragraphAdded.PackagePart = cell.mainPart
			ElseIf TypeOf Me Is DocX Then
				paragraphAdded.PackagePart = Document.mainPart
			ElseIf TypeOf Me Is Footer Then
				Dim f = TryCast(Me, Footer)
				paragraphAdded.mainPart = f.mainPart
			ElseIf TypeOf Me Is Header Then
				Dim h = TryCast(Me, Header)
				paragraphAdded.mainPart = h.mainPart
			Else
				Console.WriteLine("No idea what we are {0}", Me)
				paragraphAdded.PackagePart = Document.mainPart
			End If


			GetParent(paragraphAdded)

			Return paragraphAdded
		End Function

		Public Overridable Function InsertEquation(ByVal equation As String) As Paragraph
			Dim p As Paragraph = InsertParagraph()
			p.AppendEquation(equation)
			Return p
		End Function

		Public Overridable Function InsertBookmark(ByVal bookmarkName As String) As Paragraph
			Dim p = InsertParagraph()
			p.AppendBookmark(bookmarkName)
			Return p
		End Function

		Public Overridable Function InsertTable(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table 'Dmitchern, changed to virtual, and overrided in Table.Cell
			Dim newTable As XElement = HelperFunctions.CreateTable(rowCount, columnCount)
			Xml.Add(newTable)

			Return New Table(Document, newTable) With {.mainPart = mainPart}
		End Function

		Public Function InsertTable(ByVal index As Integer, ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			Dim newTable As XElement = HelperFunctions.CreateTable(rowCount, columnCount)

			Dim p As Paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index)

			If p Is Nothing Then
				Xml.Elements().First().AddFirst(newTable)

			Else
				Dim split() As XElement = HelperFunctions.SplitParagraph(p, index - p.startIndex)

				p.Xml.ReplaceWith (split(0), newTable, split(1))
			End If


			Return New Table(Document, newTable) With {.mainPart = mainPart}
		End Function

		Public Function InsertTable(ByVal t As Table) As Table
			Dim newXElement As New XElement(t.Xml)
			Xml.Add(newXElement)

			Dim newTable As New Table(Document, newXElement) With {.mainPart = mainPart, .Design = t.Design}

			Return newTable
		End Function

		Public Function InsertTable(ByVal index As Integer, ByVal t As Table) As Table
			Dim p As Paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index)

			Dim split() As XElement = HelperFunctions.SplitParagraph(p, index - p.startIndex)
			Dim newXElement As New XElement(t.Xml)
			p.Xml.ReplaceWith (split(0), newXElement, split(1))

			Dim newTable As New Table(Document, newXElement) With {.mainPart = mainPart, .Design = t.Design}

			Return newTable
		End Function
		Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)

		End Sub

		Public Function InsertList(ByVal list As List) As List
			For Each item In list.Items
				Xml.Add(item.Xml)
			Next item

			Return list
		End Function
		Public Function InsertList(ByVal list As List, ByVal fontSize As Double) As List
			For Each item In list.Items
				item.FontSize(fontSize)
				Xml.Add(item.Xml)
			Next item
			Return list
		End Function

		Public Function InsertList(ByVal list As List, ByVal fontFamily As FontFamily, ByVal fontSize As Double) As List
			For Each item In list.Items
				item.Font(fontFamily)
				item.FontSize(fontSize)
				Xml.Add(item.Xml)
			Next item
			Return list
		End Function

		Public Function InsertList(ByVal index As Integer, ByVal list As List) As List
			Dim p As Paragraph = HelperFunctions.GetFirstParagraphEffectedByInsert(Document, index)

			Dim split() As XElement = HelperFunctions.SplitParagraph(p, index - p.startIndex)
			Dim elements = New List(Of XElement) From {split(0)}
			elements.AddRange(list.Items.Select(Function(i) New XElement(i.Xml)))
			elements.Add(split(1))
			p.Xml.ReplaceWith(elements.ToArray())

			Return list
		End Function
	End Class
End Namespace
