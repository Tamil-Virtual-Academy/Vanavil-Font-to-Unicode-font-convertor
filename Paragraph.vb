Imports System.IO
Imports System.Collections
Imports System.IO.Packaging
Imports System.Globalization
Imports System.Security.Principal
Imports System.Text.RegularExpressions


Namespace Novacode
	''' <summary>
	''' Represents a document paragraph.
	''' </summary>
	Public Class Paragraph
		Inherits InsertBeforeOrAfter

		' The Append family of functions use this List to apply style.
		Friend runs As List(Of XElement)

		' This paragraphs text alignment
'INSTANT VB NOTE: The variable alignment was renamed since Visual Basic does not allow class members with the same name:
		Private alignment_Renamed As Alignment

		Public ParentContainer As ContainerType

		Private Property ParagraphNumberPropertiesBacker() As XElement
		''' <summary>
		''' Fetch the paragraph number properties for a list element.
		''' </summary>
		Public ReadOnly Property ParagraphNumberProperties() As XElement
			Get
				If ParagraphNumberPropertiesBacker IsNot Nothing Then
					Return ParagraphNumberPropertiesBacker
				Else
					ParagraphNumberPropertiesBacker = GetParagraphNumberProperties()
					Return ParagraphNumberPropertiesBacker
				End If
			End Get
		End Property

		Private Function GetParagraphNumberProperties() As XElement
			Dim numPrNode = Xml.Descendants().FirstOrDefault(Function(el) el.Name.LocalName = "numPr")
			If numPrNode IsNot Nothing Then
				Dim numIdNode = numPrNode.Descendants().First(Function(numId) numId.Name.LocalName = "numId")
				Dim numIdAttribute = numIdNode.Attribute(DocX.w + "val")
				If numIdAttribute IsNot Nothing AndAlso numIdAttribute.Value.Equals("0") Then
					Return Nothing
				End If
			End If

			Return numPrNode
		End Function

		Private Property IsListItemBacker() As Boolean?
		''' <summary>
		''' Determine if this paragraph is a list element.
		''' </summary>
		Public ReadOnly Property IsListItem() As Boolean
			Get
				IsListItemBacker = If(IsListItemBacker, (ParagraphNumberProperties IsNot Nothing))
				Return CBool(IsListItemBacker)
			End Get
		End Property

		Private Property IndentLevelBacker() As Integer?
		''' <summary>
		''' If this element is a list item, get the indentation level of the list item.
		''' </summary>
		Public ReadOnly Property IndentLevel() As Integer?
			Get
				If Not IsListItem Then
					Return Nothing
				End If
				If IndentLevelBacker.HasValue Then
					Return IndentLevelBacker
				Else
					Return IndentLevelBacker = Integer.Parse(ParagraphNumberProperties.Descendants().First(Function(el) el.Name.LocalName = "ilvl").GetAttribute(DocX.w + "val"))
				End If
			End Get
		End Property

		''' <summary>
		''' Determine if the list element is a numbered list of bulleted list element
		''' </summary>
		Public ListItemType As ListItemType

		Friend startIndex, endIndex As Integer

		''' <summary>
		''' Returns a list of all Pictures in a Paragraph.
		''' </summary>
		''' <example>
		''' Returns a list of all Pictures in a Paragraph.
		''' <code>
		''' <![CDATA[
		''' // Create a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // Get the first Paragraph in a document.
		'''    Paragraph p = document.Paragraphs[0];
		''' 
		'''    // Get all of the Pictures in this Paragraph.
		'''    List<Picture> pictures = p.Pictures;
		'''
		'''    // Save this document.
		'''    document.Save();
		''' }
		''' ]]>
		''' </code>
		''' </example>
		Public ReadOnly Property Pictures() As List(Of Picture)
			Get
'INSTANT VB NOTE: The local variable pictures was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim pictures_Renamed As List(Of Picture) = (
				    From p In Xml.Descendants()
				    Where (p.Name.LocalName = "drawing")
				    Let id = (
				        From e In p.Descendants()
				        Where e.Name.LocalName.Equals("blip")
				        Select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()
				    Where id IsNot Nothing
				    Let img = New Image(Document, mainPart.GetRelationship(id))
				    Select New Picture(Document, p, img)).ToList()

				Dim shapes As List(Of Picture) = (
				    From p In Xml.Descendants()
				    Where (p.Name.LocalName = "pict")
				    Let id = (
				        From e In p.Descendants()
				        Where e.Name.LocalName.Equals("imagedata")
				        Select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()
				    Where id IsNot Nothing
				    Let img = New Image(Document, mainPart.GetRelationship(id))
				    Select New Picture(Document, p, img)).ToList()

				For Each p As Picture In shapes
					pictures_Renamed.Add(p)
				Next p


				Return pictures_Renamed
			End Get
		End Property

		''' <summary>
		''' Returns a list of Hyperlinks in this Paragraph.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // Get the first Paragraph in this document.
		'''    Paragraph p = document.Paragraphs[0];
		'''    
		'''    // Get all of the hyperlinks in this Paragraph.
		'''    <![CDATA[ List<Hyperlink> ]]> hyperlinks = paragraph.Hyperlinks;
		'''    
		'''    // Change the first hyperlinks text and Uri
		'''    Hyperlink h0 = hyperlinks[0];
		'''    h0.Text = "DocX";
		'''    h0.Uri = new Uri("http://docx.codeplex.com");
		'''
		'''    // Save this document.
		'''    document.Save();
		''' }
		''' </code>
		''' </example>
		Public ReadOnly Property Hyperlinks() As List(Of Hyperlink)
			Get
'INSTANT VB NOTE: The local variable hyperlinks was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim hyperlinks_Renamed As New List(Of Hyperlink)()

				Dim hyperlink_elements As List(Of XElement) = (
				    From h In Xml.Descendants()
				    Where (h.Name.LocalName = "hyperlink" OrElse h.Name.LocalName = "instrText")
				    Select h).ToList()

				For Each he As XElement In hyperlink_elements
					If he.Name.LocalName = "hyperlink" Then
						Try
							Dim h As New Hyperlink(Document, mainPart, he)
							h.mainPart = mainPart
							hyperlinks_Renamed.Add(h)

						Catch e1 As Exception
						End Try

					Else
						' Find the parent run, no matter how deeply nested we are.
						Dim e As XElement = he
						Do While e.Name.LocalName <> "r"
							e = e.Parent
						Loop

						' Take every element until we reach w:fldCharType="end"
						Dim hyperlink_runs As New List(Of XElement)()
						For Each r As XElement In e.ElementsAfterSelf(XName.Get("r", DocX.w.NamespaceName))
							' Add this run to the list.
							hyperlink_runs.Add(r)

							Dim fldChar As XElement = r.Descendants(XName.Get("fldChar", DocX.w.NamespaceName)).SingleOrDefault(Of XElement)()
							If fldChar IsNot Nothing Then
								Dim fldCharType As XAttribute = fldChar.Attribute(XName.Get("fldCharType", DocX.w.NamespaceName))
								If fldCharType IsNot Nothing AndAlso fldCharType.Value.Equals("end", StringComparison.CurrentCultureIgnoreCase) Then
									Try
										Dim h As New Hyperlink(Document, he, hyperlink_runs)
										h.mainPart = mainPart
										hyperlinks_Renamed.Add(h)

									Catch e2 As Exception
									End Try

									Exit For
								End If
							End If
						Next r
					End If
				Next he

				Return hyperlinks_Renamed
			End Get
		End Property

		'''<summary>
		''' The style name of the paragraph.
		'''</summary>
		Public Property StyleName() As String
			Get
				Dim element = Me.GetOrCreate_pPr()
				Dim styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName))
				If styleElement IsNot Nothing Then
					Dim attr = styleElement.Attribute(XName.Get("val", DocX.w.NamespaceName))
					If attr IsNot Nothing AndAlso (Not String.IsNullOrEmpty(attr.Value)) Then
						Return attr.Value
					End If
				End If
				Return "Normal"
			End Get
			Set(ByVal value As String)
				If String.IsNullOrEmpty(value) Then
					value = "Normal"
				End If
				Dim element = Me.GetOrCreate_pPr()
				Dim styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName))
				If styleElement Is Nothing Then
					element.Add(New XElement(XName.Get("pStyle", DocX.w.NamespaceName)))
					styleElement = element.Element(XName.Get("pStyle", DocX.w.NamespaceName))
				End If
				styleElement.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), value)
			End Set
		End Property

		' A collection of field type DocProperty.
		Private docProperties As List(Of DocProperty)

		Friend styles As New List(Of XElement)()

		''' <summary>
		''' Returns a list of field type DocProperty in this document.
		''' </summary>
		Public ReadOnly Property DocumentProperties() As List(Of DocProperty)
			Get
				Return docProperties
			End Get
		End Property

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement, ByVal startIndex As Integer, Optional ByVal parent As ContainerType = ContainerType.None)
			MyBase.New(document, xml)
			ParentContainer = parent
			Me.startIndex = startIndex
			Me.endIndex = startIndex + GetElementTextLength(xml)

			RebuildDocProperties()

			' As per Unused code affecting performance (Wiki Link: [discussion:454191]) and coffeycathal suggestion no longer requeried
			'#region It's possible that a Paragraph may have pStyle references
			'// Check if this Paragraph references any pStyle elements.
			'var stylesElements = xml.Descendants(XName.Get("pStyle", DocX.w.NamespaceName));

			'// If one or more pStyles are referenced.
			'if (stylesElements.Count() > 0)
			'{
			'    Uri style_package_uri = new Uri("/word/styles.xml", UriKind.Relative);
			'    PackagePart styles_document = document.package.GetPart(style_package_uri);

			'    using (TextReader tr = new StreamReader(styles_document.GetStream()))
			'    {
			'        XDocument style_document = XDocument.Load(tr);
			'        XElement styles_element = style_document.Element(XName.Get("styles", DocX.w.NamespaceName));

			'        var styles_element_ids = stylesElements.Select(e => e.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value);

			'        //foreach(string id in styles_element_ids)
			'        //{
			'        //    var style = 
			'        //    (
			'        //        from d in styles_element.Descendants()
			'        //        let styleId = d.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
			'        //        let type = d.Attribute(XName.Get("type", DocX.w.NamespaceName))
			'        //        where type != null && type.Value == "paragraph" && styleId != null && styleId.Value == id
			'        //        select d
			'        //    ).First();

			'        //    styles.Add(style);
			'        //} 
			'    }
			'}
			'#endregion

			Me.runs = Me.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList()
		End Sub

		''' <summary>
		''' Insert a new Table before this Paragraph, this Table can be from this document or another document.
		''' </summary>
		''' <param name="t">The Table t to be inserted.</param>
		''' <returns>A new Table inserted before this Paragraph.</returns>
		''' <example>
		''' Insert a new Table before this Paragraph.
		''' <code>
		''' // Place holder for a Table.
		''' Table t;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"a.docx"))
		''' {
		'''     // Get the first Table from this document.
		'''     t = documentA.Tables[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"b.docx"))
		''' {
		'''     // Get the first Paragraph in document b.
		'''     Paragraph p2 = documentB.Paragraphs[0];
		'''
		'''     // Insert the Table from document a before this Paragraph.
		'''     Table newTable = p2.InsertTableBeforeSelf(t);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertTableBeforeSelf(ByVal t As Table) As Table
			t = MyBase.InsertTableBeforeSelf(t)
			t.mainPart = mainPart
			Return t
		End Function

'INSTANT VB NOTE: The variable direction was renamed since Visual Basic does not allow class members with the same name:
		Private direction_Renamed As Direction
		''' <summary>
		''' Gets or Sets the Direction of content in this Paragraph.
		''' <example>
		''' Create a Paragraph with content that flows right to left. Default is left to right.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Create a new Paragraph with the text "Hello World".
		'''     Paragraph p = document.InsertParagraph("Hello World.");
		''' 
		'''     // Make this Paragraph flow right to left. Default is left to right.
		'''     p.Direction = Direction.RightToLeft;
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' </summary>
		Public Property Direction() As Direction
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim bidi As XElement = pPr.Element(XName.Get("bidi", DocX.w.NamespaceName))

				If bidi Is Nothing Then
					Return Direction.LeftToRight

				Else
					Return Direction.RightToLeft
				End If
			End Get

			Set(ByVal value As Direction)
				direction_Renamed = value

				Dim pPr As XElement = GetOrCreate_pPr()
				Dim bidi As XElement = pPr.Element(XName.Get("bidi", DocX.w.NamespaceName))

				If direction_Renamed = Direction.RightToLeft Then
					If bidi Is Nothing Then
						pPr.Add(New XElement(XName.Get("bidi", DocX.w.NamespaceName)))
					End If

				Else
					If bidi IsNot Nothing Then
						bidi.Remove()
					End If
				End If
			End Set
		End Property

		Public ReadOnly Property IsKeepWithNext() As Boolean

			Get
				Dim pPr = GetOrCreate_pPr()
				Dim keepWithNextE = pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName))
				If keepWithNextE Is Nothing Then
					Return False
				End If
				Return True
			End Get
		End Property
		''' <summary>
		''' This paragraph will be kept on the same page as the next paragraph
		''' </summary>
		''' <example>
		''' Create a Paragraph that will stay on the same page as the paragraph that comes next
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' 
		''' {
		'''     // Create a new Paragraph with the text "Hello World".
		'''     Paragraph p = document.InsertParagraph("Hello World.");
		'''     p.KeepWithNext();
		'''     document.InsertParagraph("Previous paragraph will appear on the same page as this paragraph");
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <param name="keepWithNext_Renamed"></param>
		''' <returns></returns>

'INSTANT VB NOTE: The parameter keepWithNext was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function KeepWithNext(Optional ByVal keepWithNext_Renamed As Boolean = True) As Paragraph
			Dim pPr = GetOrCreate_pPr()
			Dim keepWithNextE = pPr.Element(XName.Get("keepNext", DocX.w.NamespaceName))
			If keepWithNextE Is Nothing AndAlso keepWithNext_Renamed Then
				pPr.Add(New XElement(XName.Get("keepNext", DocX.w.NamespaceName)))
			End If
			If (Not keepWithNext_Renamed) AndAlso keepWithNextE IsNot Nothing Then
				keepWithNextE.Remove()
			End If
			Return Me

		End Function
		''' <summary>
		''' Keep all lines in this paragraph together on a page
		''' </summary>
		''' <example>
		''' Create a Paragraph whose lines will stay together on a single page
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Create a new Paragraph with the text "Hello World".
		'''     Paragraph p = document.InsertParagraph("All lines of this paragraph will appear on the same page...\nLine 2\nLine 3\nLine 4\nLine 5\nLine 6...");
		'''     p.KeepLinesTogether();
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <param name="keepTogether"></param>
		''' <returns></returns>
		Public Function KeepLinesTogether(Optional ByVal keepTogether As Boolean = True) As Paragraph
			Dim pPr = GetOrCreate_pPr()
			Dim keepLinesE = pPr.Element(XName.Get("keepLines", DocX.w.NamespaceName))
			If keepLinesE Is Nothing AndAlso keepTogether Then
				pPr.Add(New XElement(XName.Get("keepLines", DocX.w.NamespaceName)))
			End If
			If (Not keepTogether) AndAlso keepLinesE IsNot Nothing Then
				keepLinesE.Remove()
			End If
			Return Me
		End Function
		''' <summary>
		''' If the pPr element doesent exist it is created, either way it is returned by this function.
		''' </summary>
		''' <returns>The pPr element for this Paragraph.</returns>
		Friend Function GetOrCreate_pPr() As XElement
			' Get the element.
			Dim pPr As XElement = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName))

			' If it dosen't exist, create it.
			If pPr Is Nothing Then
				Xml.AddFirst(New XElement(XName.Get("pPr", DocX.w.NamespaceName)))
				pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName))
			End If

			' Return the pPr element for this Paragraph.
			Return pPr
		End Function

		''' <summary>
		''' If the ind element doesent exist it is created, either way it is returned by this function.
		''' </summary>
		''' <returns>The ind element for this Paragraphs pPr.</returns>
		Friend Function GetOrCreate_pPr_ind() As XElement
			' Get the element.
			Dim pPr As XElement = GetOrCreate_pPr()
			Dim ind As XElement = pPr.Element(XName.Get("ind", DocX.w.NamespaceName))

			' If it dosen't exist, create it.
			If ind Is Nothing Then
				pPr.Add(New XElement(XName.Get("ind", DocX.w.NamespaceName)))
				ind = pPr.Element(XName.Get("ind", DocX.w.NamespaceName))
			End If

			' Return the pPr element for this Paragraph.
			Return ind
		End Function

'INSTANT VB NOTE: The variable indentationFirstLine was renamed since Visual Basic does not allow class members with the same name:
		Private indentationFirstLine_Renamed As Single
		''' <summary>
		''' Get or set the indentation of the first line of this Paragraph.
		''' </summary>
		''' <example>
		''' Indent only the first line of a Paragraph.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Create a new Paragraph.
		'''     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		''' 
		'''     // Indent only the first line of the Paragraph.
		'''     p.IndentationFirstLine = 2.0f;
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Property IndentationFirstLine() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim ind As XElement = GetOrCreate_pPr_ind()
				Dim firstLine As XAttribute = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName))

				If firstLine IsNot Nothing Then
					Return Single.Parse(firstLine.Value)
				End If

				Return 0.0f
			End Get

			Set(ByVal value As Single)
				If IndentationFirstLine <> value Then
					indentationFirstLine_Renamed = value

					Dim pPr As XElement = GetOrCreate_pPr()
					Dim ind As XElement = GetOrCreate_pPr_ind()

					' Paragraph can either be firstLine or hanging (Remove hanging).
					Dim hanging As XAttribute = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName))
					If hanging IsNot Nothing Then
						hanging.Remove()
					End If

					Dim indentation As String = ((indentationFirstLine_Renamed / 0.1) * 57).ToString()
					Dim firstLine As XAttribute = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName))
					If firstLine IsNot Nothing Then
						firstLine.Value = indentation
					Else
						ind.Add(New XAttribute(XName.Get("firstLine", DocX.w.NamespaceName), indentation))
					End If
				End If
			End Set
		End Property

'INSTANT VB NOTE: The variable indentationHanging was renamed since Visual Basic does not allow class members with the same name:
		Private indentationHanging_Renamed As Single
		''' <summary>
		''' Get or set the indentation of all but the first line of this Paragraph.
		''' </summary>
		''' <example>
		''' Indent all but the first line of a Paragraph.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Create a new Paragraph.
		'''     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		''' 
		'''     // Indent all but the first line of the Paragraph.
		'''     p.IndentationHanging = 1.0f;
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Property IndentationHanging() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim ind As XElement = GetOrCreate_pPr_ind()
				Dim hanging As XAttribute = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName))

				If hanging IsNot Nothing Then
					Return Single.Parse(hanging.Value) / (57 * 10)
				End If

				Return 0.0f
			End Get

			Set(ByVal value As Single)
				If IndentationHanging <> value Then
					indentationHanging_Renamed = value

					Dim pPr As XElement = GetOrCreate_pPr()
					Dim ind As XElement = GetOrCreate_pPr_ind()

					' Paragraph can either be firstLine or hanging (Remove firstLine).
					Dim firstLine As XAttribute = ind.Attribute(XName.Get("firstLine", DocX.w.NamespaceName))
					If firstLine IsNot Nothing Then
						firstLine.Remove()
					End If

					Dim indentation As String = ((indentationHanging_Renamed / 0.1) * 57).ToString()
					Dim hanging As XAttribute = ind.Attribute(XName.Get("hanging", DocX.w.NamespaceName))
					If hanging IsNot Nothing Then
						hanging.Value = indentation
					Else
						ind.Add(New XAttribute(XName.Get("hanging", DocX.w.NamespaceName), indentation))
					End If
				End If
			End Set
		End Property

'INSTANT VB NOTE: The variable indentationBefore was renamed since Visual Basic does not allow class members with the same name:
		Private indentationBefore_Renamed As Single
		''' <summary>
		''' Set the before indentation in cm for this Paragraph.
		''' </summary>
		''' <example>
		''' // Indent an entire Paragraph from the left.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''    // Create a new Paragraph.
		'''    Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		'''
		'''    // Indent this entire Paragraph from the left.
		'''    p.IndentationBefore = 2.0f;
		'''    
		'''    // Save all changes made to this document.
		'''    document.Save();
		'''}
		''' </code>
		''' </example>
		Public Property IndentationBefore() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim ind As XElement = GetOrCreate_pPr_ind()

				Dim left As XAttribute = ind.Attribute(XName.Get("left", DocX.w.NamespaceName))
				If left IsNot Nothing Then
					Return Single.Parse(left.Value) / (57 * 10)
				End If

				Return 0.0f
			End Get

			Set(ByVal value As Single)
				If IndentationBefore <> value Then
					indentationBefore_Renamed = value

					Dim pPr As XElement = GetOrCreate_pPr()
					Dim ind As XElement = GetOrCreate_pPr_ind()

					Dim indentation As String = ((indentationBefore_Renamed / 0.1) * 57).ToString()

					Dim left As XAttribute = ind.Attribute(XName.Get("left", DocX.w.NamespaceName))
					If left IsNot Nothing Then
						left.Value = indentation
					Else
						ind.Add(New XAttribute(XName.Get("left", DocX.w.NamespaceName), indentation))
					End If
				End If
			End Set
		End Property

'INSTANT VB NOTE: The variable indentationAfter was renamed since Visual Basic does not allow class members with the same name:
		Private indentationAfter_Renamed As Single = 0.0f
		''' <summary>
		''' Set the after indentation in cm for this Paragraph.
		''' </summary>
		''' <example>
		''' // Indent an entire Paragraph from the right.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Create a new Paragraph.
		'''     Paragraph p = document.InsertParagraph("Line 1\nLine 2\nLine 3");
		''' 
		'''     // Make the content of this Paragraph flow right to left.
		'''     p.Direction = Direction.RightToLeft;
		''' 
		'''     // Indent this entire Paragraph from the right.
		'''     p.IndentationAfter = 2.0f;
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Property IndentationAfter() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim ind As XElement = GetOrCreate_pPr_ind()

				Dim right As XAttribute = ind.Attribute(XName.Get("right", DocX.w.NamespaceName))
				If right IsNot Nothing Then
					Return Single.Parse(right.Value)
				End If

				Return 0.0f
			End Get

			Set(ByVal value As Single)
				If IndentationAfter <> value Then
					indentationAfter_Renamed = value

					Dim pPr As XElement = GetOrCreate_pPr()
					Dim ind As XElement = GetOrCreate_pPr_ind()

					Dim indentation As String = ((indentationAfter_Renamed / 0.1) * 57).ToString()

					Dim right As XAttribute = ind.Attribute(XName.Get("right", DocX.w.NamespaceName))
					If right IsNot Nothing Then
						right.Value = indentation
					Else
						ind.Add(New XAttribute(XName.Get("right", DocX.w.NamespaceName), indentation))
					End If
				End If
			End Set
		End Property

		''' <summary>
		''' Insert a new Table into this document before this Paragraph.
		''' </summary>
		''' <param name="rowCount">The number of rows this Table should have.</param>
		''' <param name="columnCount">The number of columns this Table should have.</param>
		''' <returns>A new Table inserted before this Paragraph.</returns>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     //Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("Hello World", false);
		'''
		'''     // Insert a new Table before this Paragraph.
		'''     Table newTable = p.InsertTableBeforeSelf(2, 2);
		'''     newTable.Design = TableDesign.LightShadingAccent2;
		'''     newTable.Alignment = Alignment.center;
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertTableBeforeSelf(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			Return MyBase.InsertTableBeforeSelf(rowCount, columnCount)
		End Function

		''' <summary>
		''' Insert a new Table after this Paragraph.
		''' </summary>
		''' <param name="t">The Table t to be inserted.</param>
		''' <returns>A new Table inserted after this Paragraph.</returns>
		''' <example>
		''' Insert a new Table after this Paragraph.
		''' <code>
		''' // Place holder for a Table.
		''' Table t;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"a.docx"))
		''' {
		'''     // Get the first Table from this document.
		'''     t = documentA.Tables[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"b.docx"))
		''' {
		'''     // Get the first Paragraph in document b.
		'''     Paragraph p2 = documentB.Paragraphs[0];
		'''
		'''     // Insert the Table from document a after this Paragraph.
		'''     Table newTable = p2.InsertTableAfterSelf(t);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertTableAfterSelf(ByVal t As Table) As Table
			t = MyBase.InsertTableAfterSelf(t)
			t.mainPart = mainPart
			Return t
		End Function

		''' <summary>
		''' Insert a new Table into this document after this Paragraph.
		''' </summary>
		''' <param name="rowCount">The number of rows this Table should have.</param>
		''' <param name="columnCount">The number of columns this Table should have.</param>
		''' <returns>A new Table inserted after this Paragraph.</returns>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     //Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("Hello World", false);
		'''
		'''     // Insert a new Table after this Paragraph.
		'''     Table newTable = p.InsertTableAfterSelf(2, 2);
		'''     newTable.Design = TableDesign.LightShadingAccent2;
		'''     newTable.Alignment = Alignment.center;
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertTableAfterSelf(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			Return MyBase.InsertTableAfterSelf(rowCount, columnCount)
		End Function

		''' <summary>
		''' Insert a Paragraph before this Paragraph, this Paragraph may have come from the same or another document.
		''' </summary>
		''' <param name="p">The Paragraph to insert.</param>
		''' <returns>The Paragraph now associated with this document.</returns>
		''' <example>
		''' Take a Paragraph from document a, and insert it into document b before this Paragraph.
		''' <code>
		''' // Place holder for a Paragraph.
		''' Paragraph p;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"a.docx"))
		''' {
		'''     // Get the first paragraph from this document.
		'''     p = documentA.Paragraphs[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"b.docx"))
		''' {
		'''     // Get the first Paragraph in document b.
		'''     Paragraph p2 = documentB.Paragraphs[0];
		'''
		'''     // Insert the Paragraph from document a before this Paragraph.
		'''     Paragraph newParagraph = p2.InsertParagraphBeforeSelf(p);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code> 
		''' </example>
		Public Overrides Function InsertParagraphBeforeSelf(ByVal p As Paragraph) As Paragraph
			Dim p2 As Paragraph = MyBase.InsertParagraphBeforeSelf(p)
			p2.PackagePart = mainPart
			Return p2
		End Function

		''' <summary>
		''' Insert a new Paragraph before this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <returns>A new Paragraph inserted before this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph before the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.");
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphBeforeSelf(ByVal text As String) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphBeforeSelf(text)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a new Paragraph before this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <param name="trackChanges">Should this insertion be tracked as a change?</param>
		''' <returns>A new Paragraph inserted before this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph before the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false);
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphBeforeSelf(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphBeforeSelf(text, trackChanges)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a new Paragraph before this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <param name="trackChanges">Should this insertion be tracked as a change?</param>
		''' <param name="formatting">The formatting to apply to this insertion.</param>
		''' <returns>A new Paragraph inserted before this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph before the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     Formatting boldFormatting = new Formatting();
		'''     boldFormatting.Bold = true;
		'''
		'''     p.InsertParagraphBeforeSelf("I was inserted before the next Paragraph.", false, boldFormatting);
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphBeforeSelf(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphBeforeSelf(text, trackChanges, formatting)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a page break before a Paragraph.
		''' </summary>
		''' <example>
		''' Insert 2 Paragraphs into a document with a page break between them.
		''' <code>
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Insert a new Paragraph.
		'''    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
		'''       
		'''    // Insert a new Paragraph.
		'''    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
		'''    
		'''    // Insert a page break before Paragraph two.
		'''    p2.InsertPageBreakBeforeSelf();
		'''    
		'''    // Save this document.
		'''    document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Sub InsertPageBreakBeforeSelf()
			MyBase.InsertPageBreakBeforeSelf()
		End Sub

		''' <summary>
		''' Insert a page break after a Paragraph.
		''' </summary>
		''' <example>
		''' Insert 2 Paragraphs into a document with a page break between them.
		''' <code>
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Insert a new Paragraph.
		'''    Paragraph p1 = document.InsertParagraph("Paragraph 1", false);
		'''       
		'''    // Insert a page break after this Paragraph.
		'''    p1.InsertPageBreakAfterSelf();
		'''       
		'''    // Insert a new Paragraph.
		'''    Paragraph p2 = document.InsertParagraph("Paragraph 2", false);
		'''
		'''    // Save this document.
		'''    document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Overrides Sub InsertPageBreakAfterSelf()
			MyBase.InsertPageBreakAfterSelf()
		End Sub

		<Obsolete("Instead use: InsertHyperlink(Hyperlink h, int index)")>
		Public Function InsertHyperlink(ByVal index As Integer, ByVal h As Hyperlink) As Paragraph
			Return InsertHyperlink(h, index)
		End Function

		''' <summary>
		''' This function inserts a hyperlink into a Paragraph at a specified character index.
		''' </summary>
		''' <param name="index">The index to insert at.</param>
		''' <param name="h">The hyperlink to insert.</param>
		''' <returns>The Paragraph with the Hyperlink inserted at the specified index.</returns>
		''' <!-- 
		''' This function was added by Brian Campbell aka chickendelicious on Jun 16 2010
		''' Thank you Brian.
		''' -->
		Public Function InsertHyperlink(ByVal h As Hyperlink, Optional ByVal index As Integer = 0) As Paragraph
			' Convert the path of this mainPart to its equilivant rels file path.
			Dim path As String = mainPart.Uri.OriginalString.Replace("/word/", "")
			Dim rels_path As New Uri(String.Format("/word/_rels/{0}.rels", path), UriKind.Relative)

			' Check to see if the rels file exists and create it if not.
			If Not Document.package.PartExists(rels_path) Then
				HelperFunctions.CreateRelsPackagePart(Document, rels_path)
			End If

			' Check to see if a rel for this Picture exists, create it if not.
			Dim Id = GetOrGenerateRel(h)

			Dim h_xml As XElement
			If index = 0 Then
				' Add this hyperlink as the last element.
				Xml.AddFirst(h.Xml)

				' Extract the picture back out of the DOM.
				h_xml = CType(Xml.FirstNode, XElement)

			Else
				' Get the first run effected by this Insert
				Dim run As Run = GetFirstRunEffectedByEdit(index)

				If run Is Nothing Then
					' Add this hyperlink as the last element.
					Xml.Add(h.Xml)

					' Extract the picture back out of the DOM.
					h_xml = CType(Xml.LastNode, XElement)

				Else
					' Split this run at the point you want to insert
					Dim splitRun() As XElement = Run.SplitRun(run, index)

					' Replace the origional run.
					run.Xml.ReplaceWith (splitRun(0), h.Xml, splitRun(1))

					' Get the first run effected by this Insert
					run = GetFirstRunEffectedByEdit(index)

					' The picture has to be the next element, extract it back out of the DOM.
					h_xml = CType(run.Xml.NextNode, XElement)
				End If

				h_xml.SetAttributeValue(DocX.r + "id", Id)
			End If

			Return Me
		End Function

		''' <summary>
		''' Remove the Hyperlink at the provided index. The first hyperlink is at index 0.
		''' Using a negative index or an index greater than the index of the last hyperlink will cause an ArgumentOutOfRangeException() to be thrown.
		''' </summary>
		''' <param name="index">The index of the hyperlink to be removed.</param>
		''' <example>
		''' <code>
		''' // Crete a new document.
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''     // Add a Hyperlink into this document.
		'''     Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
		'''
		'''     // Insert a new Paragraph into the document.
		'''     Paragraph p1 = document.InsertParagraph("AC");
		'''     
		'''     // Insert the hyperlink into this Paragraph.
		'''     p1.InsertHyperlink(1, h);
		'''     Assert.IsTrue(p1.Text == "AlinkC"); // Make sure the hyperlink was inserted correctly;
		'''     
		'''     // Remove the hyperlink
		'''     p1.RemoveHyperlink(0);
		'''     Assert.IsTrue(p1.Text == "AC"); // Make sure the hyperlink was removed correctly;
		''' }
		''' </code>
		''' </example>
		Public Sub RemoveHyperlink(ByVal index As Integer)
			' Dosen't make sense to remove a Hyperlink at a negative index.
			If index < 0 Then
				Throw New ArgumentOutOfRangeException()
			End If

			' Need somewhere to store the count.
			Dim count As Integer = 0
			Dim found As Boolean = False
			RemoveHyperlinkRecursive(Xml, index, count, found)

			' If !found then the user tried to remove a hyperlink at an index greater than the last. 
			If Not found Then
				Throw New ArgumentOutOfRangeException()
			End If
		End Sub

		Friend Sub RemoveHyperlinkRecursive(ByVal xml As XElement, ByVal index As Integer, ByRef count As Integer, ByRef found As Boolean)
			If xml.Name.LocalName.Equals("hyperlink", StringComparison.CurrentCultureIgnoreCase) Then
				' This is the hyperlink to be removed.
				If count = index Then
					found = True
					xml.Remove()

				Else
					count += 1
				End If
			End If

			If xml.HasElements Then
				For Each e As XElement In xml.Elements()
					If Not found Then
						RemoveHyperlinkRecursive(e, index, count, found)
					End If
				Next e
			End If
		End Sub

		''' <summary>
		''' Insert a Paragraph after this Paragraph, this Paragraph may have come from the same or another document.
		''' </summary>
		''' <param name="p">The Paragraph to insert.</param>
		''' <returns>The Paragraph now associated with this document.</returns>
		''' <example>
		''' Take a Paragraph from document a, and insert it into document b after this Paragraph.
		''' <code>
		''' // Place holder for a Paragraph.
		''' Paragraph p;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"a.docx"))
		''' {
		'''     // Get the first paragraph from this document.
		'''     p = documentA.Paragraphs[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"b.docx"))
		''' {
		'''     // Get the first Paragraph in document b.
		'''     Paragraph p2 = documentB.Paragraphs[0];
		'''
		'''     // Insert the Paragraph from document a after this Paragraph.
		'''     Paragraph newParagraph = p2.InsertParagraphAfterSelf(p);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code> 
		''' </example>
		Public Overrides Function InsertParagraphAfterSelf(ByVal p As Paragraph) As Paragraph
			Dim p2 As Paragraph = MyBase.InsertParagraphAfterSelf(p)
			p2.PackagePart = mainPart
			Return p2
		End Function

		''' <summary>
		''' Insert a new Paragraph after this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <param name="trackChanges">Should this insertion be tracked as a change?</param>
		''' <param name="formatting">The formatting to apply to this insertion.</param>
		''' <returns>A new Paragraph inserted after this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph after the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     Formatting boldFormatting = new Formatting();
		'''     boldFormatting.Bold = true;
		'''
		'''     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false, boldFormatting);
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphAfterSelf(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphAfterSelf(text, trackChanges, formatting)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a new Paragraph after this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <param name="trackChanges">Should this insertion be tracked as a change?</param>
		''' <returns>A new Paragraph inserted after this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph after the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.", false);
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphAfterSelf(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphAfterSelf(text, trackChanges)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a new Paragraph after this Paragraph.
		''' </summary>
		''' <param name="text">The initial text for this new Paragraph.</param>
		''' <returns>A new Paragraph inserted after this Paragraph.</returns>
		''' <example>
		''' Insert a new paragraph after the first Paragraph in this document.
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("I am a Paragraph", false);
		'''
		'''     p.InsertParagraphAfterSelf("I was inserted after the previous Paragraph.");
		'''
		'''     // Save all changes made to this new document.
		'''     document.Save();
		'''    }// Release this new document form memory.
		''' </code>
		''' </example>
		Public Overrides Function InsertParagraphAfterSelf(ByVal text As String) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraphAfterSelf(text)
			p.PackagePart = mainPart
			Return p
		End Function

		Private Sub RebuildDocProperties()
			docProperties = (
			    From xml In Xml.Descendants(XName.Get("fldSimple", DocX.w.NamespaceName))
			    Select New DocProperty(Document, xml)).ToList()
		End Sub

		''' <summary>
		''' Gets or set this Paragraphs text alignment.
		''' </summary>
		Public Property Alignment() As Alignment
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim jc As XElement = pPr.Element(XName.Get("jc", DocX.w.NamespaceName))

				If jc IsNot Nothing Then
					Dim a As XAttribute = jc.Attribute(XName.Get("val", DocX.w.NamespaceName))

					Select Case a.Value.ToLower()
						Case "left"
							Return Novacode.Alignment.left
						Case "right"
							Return Novacode.Alignment.right
						Case "center"
							Return Novacode.Alignment.center
						Case "both"
							Return Novacode.Alignment.both
					End Select
				End If

				Return Novacode.Alignment.left
			End Get

			Set(ByVal value As Alignment)
				alignment_Renamed = value

				Dim pPr As XElement = GetOrCreate_pPr()
				Dim jc As XElement = pPr.Element(XName.Get("jc", DocX.w.NamespaceName))

				If alignment_Renamed <> Novacode.Alignment.left Then
					If jc Is Nothing Then
						pPr.Add(New XElement(XName.Get("jc", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), alignment_Renamed.ToString())))
					Else
						jc.Attribute(XName.Get("val", DocX.w.NamespaceName)).Value = alignment_Renamed.ToString()
					End If

				Else
					If jc IsNot Nothing Then
						jc.Remove()
					End If
				End If
			End Set
		End Property

		''' <summary>
		''' Remove this Paragraph from the document.
		''' </summary>
		''' <param name="trackChanges">Should this remove be tracked as a change?</param>
		''' <example>
		''' Remove a Paragraph from a document and track it as a change.
		''' <code>
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Create and Insert a new Paragraph into this document.
		'''     Paragraph p = document.InsertParagraph("Hello", false);
		'''
		'''     // Remove the Paragraph and track this as a change.
		'''     p.Remove(true);
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Sub Remove(ByVal trackChanges As Boolean)
			If trackChanges Then
				Dim now As Date = Date.Now.ToUniversalTime()

				Dim elements As List(Of XElement) = Xml.Elements().ToList()
				Dim temp As New List(Of XElement)()
				For i As Integer = 0 To elements.Count() - 1
					Dim e As XElement = elements(i)

					If e.Name.LocalName <> "del" Then
						temp.Add(e)
						e.Remove()

					Else
						If temp.Count() > 0 Then
							e.AddBeforeSelf(CreateEdit(EditType.del, now, temp.Elements()))
							temp.Clear()
						End If
					End If
				Next i

				If temp.Count() > 0 Then
					Xml.Add(CreateEdit(EditType.del, now, temp))
				End If

			Else
				' If this is the only Paragraph in the Cell then we cannot remove it.
				If Xml.Parent.Name.LocalName = "tc" AndAlso Xml.Parent.Elements(XName.Get("p", DocX.w.NamespaceName)).Count() = 1 Then
					Xml.Value = String.Empty

				Else
					' Remove this paragraph from the document
					Xml.Remove()
					Xml = Nothing
				End If
			End If
		End Sub

		''' <summary>
		''' Gets the text value of this Paragraph.
		''' </summary>
		Public ReadOnly Property Text() As String
			' Returns the underlying XElement's Value property.
			Get
				Return HelperFunctions.GetText(Xml)
			End Get
		End Property

		''' <summary>
		''' Gets the formatted text value of this Paragraph.
		''' </summary>
		Public ReadOnly Property MagicText() As List(Of FormattedText)
			' Returns the underlying XElement's Value property.
			Get
				Try
					Return HelperFunctions.GetFormattedText(Xml)

				Catch e1 As Exception
					Return Nothing
				End Try

			End Get
		End Property

		'public Picture InsertPicture(Picture picture)
		'{
		'    Picture newPicture = picture;
		'    newPicture.i = new XElement(picture.i);

		'    xml.Add(newPicture.i);
		'    pictures.Add(newPicture);
		'    return newPicture;  
		'}

		' <summary>
		' Insert a Picture at the end of this paragraph.
		' </summary>
		' <param name="description">A string to describe this Picture.</param>
		' <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
		' <param name="name">The name of this image.</param>
		' <returns>A Picture.</returns>
		' <example>
		' <code>
		' // Create a document using a relative filename.
		' using (DocX document = DocX.Create(@"Test.docx"))
		' {
		'     // Add a new Paragraph to this document.
		'     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
		'
		'     // Add an Image to this document.
		'     Novacode.Image img = document.AddImage(@"Image.jpg");
		'
		'     // Insert pic at the end of Paragraph p.
		'     Picture pic = p.InsertPicture(img.Id, "Photo 31415", "A pie I baked.");
		'
		'     // Rotate the Picture clockwise by 30 degrees. 
		'     pic.Rotation = 30;
		'
		'     // Resize the Picture.
		'     pic.Width = 400;
		'     pic.Height = 300;
		'
		'     // Set the shape of this Picture to be a cube.
		'     pic.SetPictureShape(BasicShapes.cube);
		'
		'     // Flip the Picture Horizontally.
		'     pic.FlipHorizontal = true;
		'
		'     // Save all changes made to this document.
		'     document.Save();
		' }// Release this document from memory.
		' </code>
		' </example>
		' Removed to simplify the API.
		'public Picture InsertPicture(string imageID, string name, string description)
		'{
		'    Picture p = CreatePicture(Document, imageID, name, description);
		'    Xml.Add(p.Xml);
		'    return p;
		'}

		' Removed because it confusses the API.
		'public Picture InsertPicture(string imageID)
		'{
		'    return InsertPicture(imageID, string.Empty, string.Empty);
		'}

		'public Picture InsertPicture(int index, Picture picture)
		'{
		'    Picture p = picture;
		'    p.i = new XElement(picture.i);

		'    Run run = GetFirstRunEffectedByEdit(index);

		'    if (run == null)
		'        xml.Add(p.i);
		'    else
		'    {
		'        // Split this run at the point you want to insert
		'        XElement[] splitRun = Run.SplitRun(run, index);

		'        // Replace the origional run
		'        run.Xml.ReplaceWith
		'        (
		'            splitRun[0],
		'            p.i,
		'            splitRun[1]
		'        );
		'    }

		'    // Rebuild the run lookup for this paragraph
		'    runLookup.Clear();
		'    BuildRunLookup(xml);
		'    DocX.RenumberIDs(document);
		'    return p;
		'}

		' <summary>
		' Insert a Picture into this Paragraph at a specified index.
		' </summary>
		' <param name="description">A string to describe this Picture.</param>
		' <param name="imageID">The unique id that identifies the Image this Picture represents.</param>
		' <param name="name">The name of this image.</param>
		' <param name="index">The index to insert this Picture at.</param>
		' <returns>A Picture.</returns>
		' <example>
		' <code>
		' // Create a document using a relative filename.
		' using (DocX document = DocX.Create(@"Test.docx"))
		' {
		'     // Add a new Paragraph to this document.
		'     Paragraph p = document.InsertParagraph("Here is Picture 1", false);
		'
		'     // Add an Image to this document.
		'     Novacode.Image img = document.AddImage(@"Image.jpg");
		'
		'     // Insert pic at the start of Paragraph p.
		'     Picture pic = p.InsertPicture(0, img.Id, "Photo 31415", "A pie I baked.");
		'
		'     // Rotate the Picture clockwise by 30 degrees. 
		'     pic.Rotation = 30;
		'
		'     // Resize the Picture.
		'     pic.Width = 400;
		'     pic.Height = 300;
		'
		'     // Set the shape of this Picture to be a cube.
		'     pic.SetPictureShape(BasicShapes.cube);
		'
		'     // Flip the Picture Horizontally.
		'     pic.FlipHorizontal = true;
		'
		'     // Save all changes made to this document.
		'     document.Save();
		' }// Release this document from memory.
		' </code>
		' </example>
		' Removed to simplify API.
		'public Picture InsertPicture(int index, string imageID, string name, string description)
		'{
		'    Picture picture = CreatePicture(Document, imageID, name, description);

		'    Run run = GetFirstRunEffectedByEdit(index);

		'    if (run == null)
		'        Xml.Add(picture.Xml);
		'    else
		'    {
		'        // Split this run at the point you want to insert
		'        XElement[] splitRun = Run.SplitRun(run, index);

		'        // Replace the origional run
		'        run.Xml.ReplaceWith
		'        (
		'            splitRun[0],
		'            picture.Xml,
		'            splitRun[1]
		'        );
		'    }

		'    HelperFunctions.RenumberIDs(Document);
		'    return picture;
		'}

		''' <summary>
		''' Create a new Picture.
		''' </summary>
		''' <param name="document"></param>
		''' <param name="id">A unique id that identifies an Image embedded in this document.</param>
		''' <param name="name">The name of this Picture.</param>
		''' <param name="descr">The description of this Picture.</param>
		Friend Shared Function CreatePicture(ByVal document As DocX, ByVal id As String, ByVal name As String, ByVal descr As String) As Picture
			Dim part As PackagePart = document.package.GetPart(document.mainPart.GetRelationship(id).TargetUri)

			Dim newDocPrId As Integer = 1
			Dim existingIds As New List(Of String)()
			For Each bookmarkId In document.Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName))
				Dim idAtt = bookmarkId.Attributes().FirstOrDefault(Function(x) x.Name.LocalName = "id")
				If idAtt IsNot Nothing Then
					existingIds.Add(idAtt.Value)
				End If
			Next bookmarkId

			Do While existingIds.Contains(newDocPrId.ToString())
				newDocPrId += 1
			Loop


			Dim cx, cy As Integer

			Using img As Image = Image.FromStream(part.GetStream())
				cx = img.Width * 9526
				cy = img.Height * 9526
			End Using

			Dim e As New XElement(DocX.w + "drawing")

			Dim xml As XElement = XElement.Parse(String.Format("<w:r xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                    <w:drawing xmlns = ""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                        <wp:inline distT=""0"" distB=""0"" distL=""0"" distR=""0"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"">" & ControlChars.CrLf & "                            <wp:extent cx=""{0}"" cy=""{1}"" />" & ControlChars.CrLf & "                            <wp:effectExtent l=""0"" t=""0"" r=""0"" b=""0"" />" & ControlChars.CrLf & "                            <wp:docPr id=""{5}"" name=""{3}"" descr=""{4}"" />" & ControlChars.CrLf & "                            <wp:cNvGraphicFramePr>" & ControlChars.CrLf & "                                <a:graphicFrameLocks xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" noChangeAspect=""1"" />" & ControlChars.CrLf & "                            </wp:cNvGraphicFramePr>" & ControlChars.CrLf & "                            <a:graphic xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">" & ControlChars.CrLf & "                                <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/picture"">" & ControlChars.CrLf & "                                    <pic:pic xmlns:pic=""http://schemas.openxmlformats.org/drawingml/2006/picture"">" & ControlChars.CrLf & "                                        <pic:nvPicPr>" & ControlChars.CrLf & "                                        <pic:cNvPr id=""0"" name=""{3}"" />" & ControlChars.CrLf & "                                            <pic:cNvPicPr />" & ControlChars.CrLf & "                                        </pic:nvPicPr>" & ControlChars.CrLf & "                                        <pic:blipFill>" & ControlChars.CrLf & "                                            <a:blip r:embed=""{2}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""/>" & ControlChars.CrLf & "                                            <a:stretch>" & ControlChars.CrLf & "                                                <a:fillRect />" & ControlChars.CrLf & "                                            </a:stretch>" & ControlChars.CrLf & "                                        </pic:blipFill>" & ControlChars.CrLf & "                                        <pic:spPr>" & ControlChars.CrLf & "                                            <a:xfrm>" & ControlChars.CrLf & "                                                <a:off x=""0"" y=""0"" />" & ControlChars.CrLf & "                                                <a:ext cx=""{0}"" cy=""{1}"" />" & ControlChars.CrLf & "                                            </a:xfrm>" & ControlChars.CrLf & "                                            <a:prstGeom prst=""rect"">" & ControlChars.CrLf & "                                                <a:avLst />" & ControlChars.CrLf & "                                            </a:prstGeom>" & ControlChars.CrLf & "                                        </pic:spPr>" & ControlChars.CrLf & "                                    </pic:pic>" & ControlChars.CrLf & "                                </a:graphicData>" & ControlChars.CrLf & "                            </a:graphic>" & ControlChars.CrLf & "                        </wp:inline>" & ControlChars.CrLf & "                    </w:drawing></w:r>" & ControlChars.CrLf & "                    ", cx, cy, id, name, descr, newDocPrId.ToString()))

			Return New Picture(document, xml, New Image(document, document.mainPart.GetRelationship(id)))
		End Function

		' Removed because it confusses the API.
		'public Picture InsertPicture(int index, string imageID)
		'{
		'    return InsertPicture(index, imageID, string.Empty, string.Empty);
		'}

		''' <summary>
		''' Creates an Edit either a ins or a del with the specified content and date
		''' </summary>
		''' <param name="t">The type of this edit (ins or del)</param>
		''' <param name="edit_time">The time stamp to use for this edit</param>
		''' <param name="content">The initial content of this edit</param>
		''' <returns></returns>
		Friend Shared Function CreateEdit(ByVal t As EditType, ByVal edit_time As Date, ByVal content As Object) As XElement
			If t = EditType.del Then
				For Each o As Object In CType(content, IEnumerable(Of XElement))
					If TypeOf o Is XElement Then
						Dim e As XElement = (TryCast(o, XElement))
						Dim ts As IEnumerable(Of XElement) = e.DescendantsAndSelf(XName.Get("t", DocX.w.NamespaceName))

						For i As Integer = 0 To ts.Count() - 1
							Dim text As XElement = ts.ElementAt(i)
							text.ReplaceWith(New XElement(DocX.w + "delText", text.Attributes(), text.Value))
						Next i
					End If
				Next o
			End If

			Return (New XElement(DocX.w + t.ToString(), New XAttribute(DocX.w + "id", 0), New XAttribute(DocX.w + "author", WindowsIdentity.GetCurrent().Name), New XAttribute(DocX.w + "date", edit_time), content))
		End Function

		Friend Function GetFirstRunEffectedByEdit(ByVal index As Integer, Optional ByVal type As EditType = EditType.ins) As Run
			Dim len As Integer = HelperFunctions.GetText(Xml).Length

			' Make sure we are looking within an acceptable index range.
			If index < 0 OrElse ((type = EditType.ins AndAlso index > len) OrElse (type = EditType.del AndAlso index >= len)) Then
				Throw New ArgumentOutOfRangeException()
			End If

			' Need some memory that can be updated by the recursive search for the XElement to Split.
			Dim count As Integer = 0
			Dim theOne As Run = Nothing

			GetFirstRunEffectedByEditRecursive(Xml, index, count, theOne, type)

			Return theOne
		End Function

		Friend Sub GetFirstRunEffectedByEditRecursive(ByVal Xml As XElement, ByVal index As Integer, ByRef count As Integer, ByRef theOne As Run, ByVal type As EditType)
			count += HelperFunctions.GetSize(Xml)

			' If the EditType is deletion then we must return the next blah
			If count > 0 AndAlso ((type = EditType.del AndAlso count > index) OrElse (type = EditType.ins AndAlso count >= index)) Then
				' Correct the index
				For Each e As XElement In Xml.ElementsBeforeSelf()
					count -= HelperFunctions.GetSize(e)
				Next e

				count -= HelperFunctions.GetSize(Xml)

				' We have found the element, now find the run it belongs to.
				Do While (Xml.Name.LocalName <> "r") AndAlso (Xml.Name.LocalName <> "pPr")
					Xml = Xml.Parent
				Loop

				theOne = New Run(Document, Xml, count)
				Return
			End If

			If Xml.HasElements Then
				For Each e As XElement In Xml.Elements()
					If theOne Is Nothing Then
						GetFirstRunEffectedByEditRecursive(e, index, count, theOne, type)
					End If
				Next e
			End If
		End Sub

		''' <!-- 
		''' Bug found and fixed by krugs525 on August 12 2009.
		''' Use TFS compare to see exact code change.
		''' -->
		Friend Shared Function GetElementTextLength(ByVal run As XElement) As Integer
			Dim count As Integer = 0

			If run Is Nothing Then
				Return count
			End If

			For Each d In run.Descendants()
				Select Case d.Name.LocalName
					Case "tab"
						If d.Parent.Name.LocalName <> "tabs" Then
							GoTo CaseLabel1
						End If
					Case "br"
					CaseLabel1:
						count += 1
					Case "t"
						GoTo CaseLabel2
					Case "delText"
					CaseLabel2:
						count += d.Value.Length
					Case Else
				End Select
			Next d
			Return count
		End Function

		Friend Function SplitEdit(ByVal edit As XElement, ByVal index As Integer, ByVal type As EditType) As XElement()
			Dim run As Run = GetFirstRunEffectedByEdit(index, type)

			Dim splitRun() As XElement = Run.SplitRun(run, index, type)

			Dim splitLeft As New XElement(edit.Name, edit.Attributes(), run.Xml.ElementsBeforeSelf(), splitRun(0))
			If GetElementTextLength(splitLeft) = 0 Then
				splitLeft = Nothing
			End If

			Dim splitRight As New XElement(edit.Name, edit.Attributes(), splitRun(1), run.Xml.ElementsAfterSelf())
			If GetElementTextLength(splitRight) = 0 Then
				splitRight = Nothing
			End If

			Return (New XElement() { splitLeft, splitRight })
		End Function

		''' <summary>
		''' Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
		''' </summary>
		''' <example>
		''' <code> 
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Create a text formatting.
		'''     Formatting f = new Formatting();
		'''     f.FontColor = Color.Red;
		'''     f.Size = 30;
		'''
		'''     // Iterate through the Paragraphs in this document.
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
		'''         p.InsertText("Start: ", true, f);
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <example>
		''' Inserting tabs using the \t switch.
		''' <code>  
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''      // Create a text formatting.
		'''      Formatting f = new Formatting();
		'''      f.FontColor = Color.Red;
		'''      f.Size = 30;
		'''        
		'''      // Iterate through the paragraphs in this document.
		'''      foreach (Paragraph p in document.Paragraphs)
		'''      {
		'''          // Insert the string "\tEnd" at the end of every paragraph and flag it as a change.
		'''          p.InsertText("\tEnd", true, f);
		'''      }
		'''       
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="Paragraph.RemoveText(int, bool)"/>
		''' <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		''' <param name="value">The System.String to insert.</param>
		''' <param name="trackChanges">Flag this insert as a change.</param>
		''' <param name="formatting">The text formatting.</param>
		Public Sub InsertText(ByVal value As String, Optional ByVal trackChanges As Boolean = False, Optional ByVal formatting As Formatting = Nothing)
			' Default values for optional parameters must be compile time constants.
			' Would have like to write 'public void InsertText(string value, bool trackChanges = false, Formatting formatting = new Formatting())
			If formatting Is Nothing Then
				formatting = New Formatting()
			End If

			Dim newRuns As List(Of XElement) = HelperFunctions.FormatInput(value, formatting.Xml)
			Xml.Add(newRuns)

			HelperFunctions.RenumberIDs(Document)
		End Sub

		''' <summary>
		''' Inserts a specified instance of System.String into a Novacode.DocX.Paragraph at a specified index position.
		''' </summary>
		''' <example>
		''' <code> 
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Create a text formatting.
		'''     Formatting f = new Formatting();
		'''     f.FontColor = Color.Red;
		'''     f.Size = 30;
		'''
		'''     // Iterate through the Paragraphs in this document.
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         // Insert the string "Start: " at the begining of every Paragraph and flag it as a change.
		'''         p.InsertText(0, "Start: ", true, f);
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <example>
		''' Inserting tabs using the \t switch.
		''' <code>  
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Create a text formatting.
		'''     Formatting f = new Formatting();
		'''     f.FontColor = Color.Red;
		'''     f.Size = 30;
		'''
		'''     // Iterate through the paragraphs in this document.
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         // Insert the string "\tStart:\t" at the begining of every paragraph and flag it as a change.
		'''         p.InsertText(0, "\tStart:\t", true, f);
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="Paragraph.RemoveText(int, bool)"/>
		''' <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		''' <param name="index">The index position of the insertion.</param>
		''' <param name="value">The System.String to insert.</param>
		''' <param name="trackChanges">Flag this insert as a change.</param>
		''' <param name="formatting">The text formatting.</param>
		Public Sub InsertText(ByVal index As Integer, ByVal value As String, Optional ByVal trackChanges As Boolean = False, Optional ByVal formatting As Formatting = Nothing)
			' Timestamp to mark the start of insert
			Dim now As Date = Date.Now
			Dim insert_datetime As New Date(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc)

			' Get the first run effected by this Insert
			Dim run As Run = GetFirstRunEffectedByEdit(index)

			If run Is Nothing Then
				Dim insert As Object
				If formatting IsNot Nothing Then 'not sure how to get original formatting here when run == null
					insert = HelperFunctions.FormatInput(value, formatting.Xml)
				Else
					insert = HelperFunctions.FormatInput(value, Nothing)
				End If

				If trackChanges Then
					insert = CreateEdit(EditType.ins, insert_datetime, insert)
				End If
				Xml.Add(insert)

			Else
				Dim newRuns As Object
				Dim rprel = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName))
				If formatting IsNot Nothing Then
					'merge 2 formattings properly
					Dim finfmt As Formatting = Nothing
					Dim oldfmt As Formatting = Nothing

					If rprel IsNot Nothing Then
						oldfmt = Formatting.Parse(rprel)
					End If

					If oldfmt IsNot Nothing Then
						finfmt = oldfmt.Clone()
						If formatting.Bold.HasValue Then
							finfmt.Bold = formatting.Bold
						End If
						If formatting.CapsStyle.HasValue Then
							finfmt.CapsStyle = formatting.CapsStyle
						End If
						If formatting.FontColor.HasValue Then
							finfmt.FontColor = formatting.FontColor
						End If
						finfmt.FontFamily = formatting.FontFamily
						If formatting.Hidden.HasValue Then
							finfmt.Hidden = formatting.Hidden
						End If
						If formatting.Highlight.HasValue Then
							finfmt.Highlight = formatting.Highlight
						End If
						If formatting.Italic.HasValue Then
							finfmt.Italic = formatting.Italic
						End If
						If formatting.Kerning.HasValue Then
							finfmt.Kerning = formatting.Kerning
						End If
						finfmt.Language = formatting.Language
						If formatting.Misc.HasValue Then
							finfmt.Misc = formatting.Misc
						End If
						If formatting.PercentageScale.HasValue Then
							finfmt.PercentageScale = formatting.PercentageScale
						End If
						If formatting.Position.HasValue Then
							finfmt.Position = formatting.Position
						End If
						If formatting.Script.HasValue Then
							finfmt.Script = formatting.Script
						End If
						If formatting.Size.HasValue Then
							finfmt.Size = formatting.Size
						End If
						If formatting.Spacing.HasValue Then
							finfmt.Spacing = formatting.Spacing
						End If
						If formatting.StrikeThrough.HasValue Then
							finfmt.StrikeThrough = formatting.StrikeThrough
						End If
						If formatting.UnderlineColor.HasValue Then
							finfmt.UnderlineColor = formatting.UnderlineColor
						End If
						If formatting.UnderlineStyle.HasValue Then
							finfmt.UnderlineStyle = formatting.UnderlineStyle
						End If
					Else
						finfmt = formatting
					End If

					newRuns = HelperFunctions.FormatInput(value, finfmt.Xml)
				Else
					newRuns = HelperFunctions.FormatInput(value, rprel)
				End If

				' The parent of this Run
				Dim parentElement As XElement = run.Xml.Parent
				Select Case parentElement.Name.LocalName
					Case "ins"
							' The datetime that this ins was created
							Dim parent_ins_date As Date = Date.Parse(parentElement.Attribute(XName.Get("date", DocX.w.NamespaceName)).Value)

'                             
'                             * Special case: You want to track changes,
'                             * and the first Run effected by this insert
'                             * has a datetime stamp equal to now.
'                            
							If trackChanges AndAlso parent_ins_date.CompareTo(insert_datetime) = 0 Then
'                                
'                                 * Inserting into a non edit and this special case, is the same procedure.
'                                
								GoTo CaseLabel1
							End If

'                            
'                             * If not the special case above, 
'                             * then inserting into an ins or a del, is the same procedure.
'                            
							GoTo CaseLabel2

					Case "del"
					CaseLabel2:
							Dim insert As Object = newRuns
							If trackChanges Then
								insert = CreateEdit(EditType.ins, insert_datetime, newRuns)
							End If

							' Split this Edit at the point you want to insert
							Dim splitEdit() As XElement = Me.SplitEdit(parentElement, index, EditType.ins)

							' Replace the origional run
							parentElement.ReplaceWith (splitEdit(0), insert, splitEdit(1))

							Exit Select

					Case Else
					CaseLabel1:
							Dim insert As Object = newRuns
							If trackChanges AndAlso (Not parentElement.Name.LocalName.Equals("ins")) Then
								insert = CreateEdit(EditType.ins, insert_datetime, newRuns)

							' Special case to deal with Page Number elements.
							'if (parentElement.Name.LocalName.Equals("fldSimple"))
							'    parentElement.AddBeforeSelf(insert);

							Else
								' Split this run at the point you want to insert
								Dim splitRun() As XElement = Run.SplitRun(run, index)

								' Replace the origional run
								run.Xml.ReplaceWith (splitRun(0), insert, splitRun(1))
							End If

							Exit Select
				End Select
			End If

			HelperFunctions.RenumberIDs(Document)
		End Sub

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <returns>This Paragraph in curent culture</returns>
		''' <example>
		''' Add a new Paragraph with russian text to this document and then set language of text to local culture.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph with russian text and set curent local culture to it.
		'''     Paragraph p = document.InsertParagraph("Привет мир!").CurentCulture();
		'''       
		'''     // Save this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function CurentCulture() As Paragraph
			ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture.Name))
			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="culture_Renamed">The CultureInfo for text</param>
		''' <returns>This Paragraph in curent culture</returns>
		''' <example>
		''' Add a new Paragraph with russian text to this document and then set language of text to local culture.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph with russian text and set specific culture to it.
		'''     Paragraph p = document.InsertParagraph("Привет мир").Culture(CultureInfo.CreateSpecificCulture("ru-RU"));
		'''       
		'''     // Save this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter culture was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Culture(ByVal culture_Renamed As CultureInfo) As Paragraph
			ApplyTextFormattingProperty(XName.Get("lang", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), culture_Renamed.Name))
			Return Me
		End Function

		''' <summary>
		''' Append text to this Paragraph.
		''' </summary>
		''' <param name="text">The text to append.</param>
		''' <returns>This Paragraph with the new text appened.</returns>
		''' <example>
		''' Add a new Paragraph to this document and then append some text to it.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph and Append some text to it.
		'''     Paragraph p = document.InsertParagraph().Append("Hello World!!!");
		'''       
		'''     // Save this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function Append(ByVal text As String) As Paragraph
			Dim newRuns As List(Of XElement) = HelperFunctions.FormatInput(text, Nothing)
			Xml.Add(newRuns)

			Me.runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(newRuns.Count()).ToList()

			Return Me
		End Function

		''' <summary>
		''' Append a hyperlink to a Paragraph.
		''' </summary>
		''' <param name="h">The hyperlink to append.</param>
		''' <returns>The Paragraph with the hyperlink appended.</returns>
		''' <example>
		''' Creates a Paragraph with some text and a hyperlink.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Add a hyperlink to this document.
		'''    Hyperlink h = document.AddHyperlink("Google", new Uri("http://www.google.com"));
		'''    
		'''    // Add a new Paragraph to this document.
		'''    Paragraph p = document.InsertParagraph();
		'''    p.Append("My favourite search engine is ");
		'''    p.AppendHyperlink(h);
		'''    p.Append(", I think it's great.");
		'''
		'''    // Save all changes made to this document.
		'''    document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function AppendHyperlink(ByVal h As Hyperlink) As Paragraph
			' Convert the path of this mainPart to its equilivant rels file path.
			Dim path As String = mainPart.Uri.OriginalString.Replace("/word/", "")
			Dim rels_path As New Uri("/word/_rels/" & path & ".rels", UriKind.Relative)

			' Check to see if the rels file exists and create it if not.
			If Not Document.package.PartExists(rels_path) Then
				HelperFunctions.CreateRelsPackagePart(Document, rels_path)
			End If

			' Check to see if a rel for this Hyperlink exists, create it if not.
			Dim Id = GetOrGenerateRel(h)

			Xml.Add(h.Xml)
			Xml.Elements().Last().SetAttributeValue(DocX.r + "id", Id)

			Me.runs = Xml.Elements().Last().Elements(XName.Get("r", DocX.w.NamespaceName)).ToList()

			Return Me
		End Function

		''' <summary>
		''' Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		''' </summary>
		''' <param name="p">The Picture to append.</param>
		''' <returns>The Paragraph with the Picture now appended.</returns>
		''' <example>
		''' Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		''' <code>
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''    // Add an image to the document. 
		'''    Image     i = document.AddImage(@"Image.jpg");
		'''    
		'''    // Create a picture i.e. (A custom view of an image)
		'''    Picture   p = i.CreatePicture();
		'''    p.FlipHorizontal = true;
		'''    p.Rotation = 10;
		'''
		'''    // Create a new Paragraph.
		'''    Paragraph par = document.InsertParagraph();
		'''    
		'''    // Append content to the Paragraph.
		'''    par.Append("Here is a cool picture")
		'''       .AppendPicture(p)
		'''       .Append(" don't you think so?");
		'''
		'''    // Save all changes made to this document.
		'''    document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function AppendPicture(ByVal p As Picture) As Paragraph
			' Convert the path of this mainPart to its equilivant rels file path.
			Dim path As String = mainPart.Uri.OriginalString.Replace("/word/", "")
			Dim rels_path As New Uri("/word/_rels/" & path & ".rels", UriKind.Relative)

			' Check to see if the rels file exists and create it if not.
			If Not Document.package.PartExists(rels_path) Then
				HelperFunctions.CreateRelsPackagePart(Document, rels_path)
			End If

			' Check to see if a rel for this Picture exists, create it if not.
			Dim Id = GetOrGenerateRel(p)

			' Add the Picture Xml to the end of the Paragragraph Xml.
			Xml.Add(p.Xml)

			' Extract the attribute id from the Pictures Xml.
			Dim a_id As XAttribute = (
			    From e In Xml.Elements().Last().Descendants()
			    Where e.Name.LocalName.Equals("blip")
			    Select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))).Single()

			' Set its value to the Pictures relationships id.
			a_id.SetValue(Id)

			' For formatting such as .Bold()
			Me.runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Reverse().Take(p.Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).Count()).ToList()

			Return Me
		End Function

		''' <summary>
		''' Add an equation to a document.
		''' </summary>
		''' <param name="equation">The Equation to append.</param>
		''' <returns>The Paragraph with the Equation now appended.</returns>
		''' <example>
		''' Add an equation to a document.
		''' <code>
		''' using (DocX document = DocX.Create("Test.docx"))
		''' {
		'''    // Add an equation to the document. 
		'''    document.AddEquation("x=y+z");
		'''    
		'''    // Save all changes made to this document.
		'''    document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function AppendEquation(ByVal equation As String) As Paragraph
			' Create equation element
			Dim oMathPara As New XElement(XName.Get("oMathPara", DocX.m.NamespaceName), New XElement (XName.Get("oMath", DocX.m.NamespaceName), New XElement (XName.Get("r", DocX.w.NamespaceName), New Formatting() With {.FontFamily = New FontFamily("Cambria Math")}.Xml, New XElement(XName.Get("t", DocX.m.NamespaceName), equation)))) ' create equation string -  create formatting

			' Add equation element into paragraph xml and update runs collection
			Xml.Add(oMathPara)
			runs = Xml.Elements(XName.Get("oMathPara", DocX.m.NamespaceName)).ToList()

			' Return paragraph with equation
			Return Me
		End Function

		Public Function ValidateBookmark(ByVal bookmarkName As String) As Boolean
			Return GetBookmarks().Any(Function(b) b.Name.Equals(bookmarkName))
		End Function

		Public Function AppendBookmark(ByVal bookmarkName As String) As Paragraph
			Dim wBookmarkStart As New XElement(XName.Get("bookmarkStart", DocX.w.NamespaceName), New XAttribute(XName.Get("id", DocX.w.NamespaceName), 0), New XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName))
			Xml.Add(wBookmarkStart)

			Dim wBookmarkEnd As New XElement(XName.Get("bookmarkEnd", DocX.w.NamespaceName), New XAttribute(XName.Get("id", DocX.w.NamespaceName), 0), New XAttribute(XName.Get("name", DocX.w.NamespaceName), bookmarkName))
			Xml.Add(wBookmarkEnd)

			Return Me
		End Function

		Public Function GetBookmarks() As IEnumerable(Of Bookmark)
			Return Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName)).Select(Function(x) x.Attribute(XName.Get("name", DocX.w.NamespaceName))).Select(Function(x) New Bookmark With {.Name = x.Value, .Paragraph = Me})
		End Function

		Public Sub InsertAtBookmark(ByVal toInsert As String, ByVal bookmarkName As String)
			Dim bookmark = Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName)).Where(Function(x) x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value = bookmarkName).SingleOrDefault()
			If bookmark IsNot Nothing Then

				Dim run = HelperFunctions.FormatInput(toInsert, Nothing)
				bookmark.AddBeforeSelf(run)
				runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList()
				HelperFunctions.RenumberIDs(Document)
			End If
		End Sub

		Public Sub ReplaceAtBookmark(ByVal toInsert As String, ByVal bookmarkName As String)
			Dim bookmark As XElement = Xml.Descendants(XName.Get("bookmarkStart", DocX.w.NamespaceName)).Where(Function(x) x.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value = bookmarkName).SingleOrDefault()
			If bookmark Is Nothing Then
				Return
			End If

			Dim nextNode As XNode = bookmark.NextNode
			Dim nextElement As XElement = TryCast(nextNode, XElement)
			Do While nextElement.Equals(Nothing) OrElse nextElement.Name.NamespaceName <> DocX.w.NamespaceName OrElse (nextElement.Name.LocalName <> "r" AndAlso nextElement.Name.LocalName <> "bookmarkEnd")
				nextNode = nextNode.NextNode
				nextElement = TryCast(nextNode, XElement)
			Loop

			' Check if next element is a bookmarkEnd
			If nextElement.Name.LocalName = "bookmarkEnd" Then
				ReplaceAtBookmark_Add(toInsert, bookmark)
				Return
			End If

			Dim contentElement As XElement = nextElement.Elements(XName.Get("t", DocX.w.NamespaceName)).FirstOrDefault()
			If contentElement Is Nothing Then
				ReplaceAtBookmark_Add(toInsert, bookmark)
				Return
			End If

			contentElement.Value = toInsert
		End Sub

		Private Sub ReplaceAtBookmark_Add(ByVal toInsert As String, ByVal bookmark As XElement)
			Dim run = HelperFunctions.FormatInput(toInsert, Nothing)
			bookmark.AddAfterSelf(run)
			runs = Xml.Elements(XName.Get("r", DocX.w.NamespaceName)).ToList()
			HelperFunctions.RenumberIDs(Document)
		End Sub


		Friend Function GetOrGenerateRel(ByVal p As Picture) As String
			Dim image_uri_string As String = p.img.pr.TargetUri.OriginalString

			' Search for a relationship with a TargetUri that points at this Image.
			Dim Id = (
			    From r In mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
			    Where r.TargetUri.OriginalString = image_uri_string
			    Select r.Id).SingleOrDefault()

			' If such a relation dosen't exist, create one.
			If Id Is Nothing Then
				' Check to see if a relationship for this Picture exists and create it if not.
				Dim pr As PackageRelationship = mainPart.CreateRelationship(p.img.pr.TargetUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				Id = pr.Id
			End If
			Return Id
		End Function

		Friend Function GetOrGenerateRel(ByVal h As Hyperlink) As String
			Dim image_uri_string As String = h.Uri.OriginalString

			' Search for a relationship with a TargetUri that points at this Image.
			Dim Id = (
			    From r In mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
			    Where r.TargetUri.OriginalString = image_uri_string
			    Select r.Id).SingleOrDefault()

			' If such a relation dosen't exist, create one.
			If Id Is Nothing Then
				' Check to see if a relationship for this Picture exists and create it if not.
				Dim pr As PackageRelationship = mainPart.CreateRelationship(h.Uri, TargetMode.External, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
				Id = pr.Id
			End If
			Return Id
		End Function

		''' <summary>
		''' Insert a Picture into a Paragraph at the given text index.
		''' If not index is provided defaults to 0.
		''' </summary>
		''' <param name="p">The Picture to insert.</param>
		''' <param name="index">The text index to insert at.</param>
		''' <returns>The modified Paragraph.</returns>
		''' <example>
		''' <code>
		'''Load test document.
		'''using (DocX document = DocX.Create("Test.docx"))
		'''{
		'''    // Add Headers and Footers into this document.
		'''    document.AddHeaders();
		'''    document.AddFooters();
		'''    document.DifferentFirstPage = true;
		'''    document.DifferentOddAndEvenPages = true;
		'''
		'''    // Add an Image to this document.
		'''    Novacode.Image img = document.AddImage(directory_documents + "purple.png");
		'''
		'''    // Create a Picture from this Image.
		'''    Picture pic = img.CreatePicture();
		'''
		'''    // Main document.
		'''    Paragraph p0 = document.InsertParagraph("Hello");
		'''    p0.InsertPicture(pic, 3);
		'''
		'''    // Header first.
		'''    Paragraph p1 = document.Headers.first.InsertParagraph("----");
		'''    p1.InsertPicture(pic, 2);
		'''
		'''    // Header odd.
		'''    Paragraph p2 = document.Headers.odd.InsertParagraph("----");
		'''    p2.InsertPicture(pic, 2);
		'''
		'''    // Header even.
		'''    Paragraph p3 = document.Headers.even.InsertParagraph("----");
		'''    p3.InsertPicture(pic, 2);
		'''
		'''    // Footer first.
		'''    Paragraph p4 = document.Footers.first.InsertParagraph("----");
		'''    p4.InsertPicture(pic, 2);
		'''
		'''    // Footer odd.
		'''    Paragraph p5 = document.Footers.odd.InsertParagraph("----");
		'''    p5.InsertPicture(pic, 2);
		'''
		'''    // Footer even.
		'''    Paragraph p6 = document.Footers.even.InsertParagraph("----");
		'''    p6.InsertPicture(pic, 2);
		'''
		'''    // Save this document.
		'''    document.Save();
		'''}
		''' </code>
		''' </example>
		Public Function InsertPicture(ByVal p As Picture, Optional ByVal index As Integer = 0) As Paragraph
			' Convert the path of this mainPart to its equilivant rels file path.
			Dim path As String = mainPart.Uri.OriginalString.Replace("/word/", "")
			Dim rels_path As New Uri("/word/_rels/" & path & ".rels", UriKind.Relative)

			' Check to see if the rels file exists and create it if not.
			If Not Document.package.PartExists(rels_path) Then
				HelperFunctions.CreateRelsPackagePart(Document, rels_path)
			End If

			' Check to see if a rel for this Picture exists, create it if not.
			Dim Id = GetOrGenerateRel(p)

			Dim p_xml As XElement
			If index = 0 Then
				' Add this hyperlink as the last element.
				Xml.AddFirst(p.Xml)

				' Extract the picture back out of the DOM.
				p_xml = CType(Xml.FirstNode, XElement)

			Else
				' Get the first run effected by this Insert
				Dim run As Run = GetFirstRunEffectedByEdit(index)

				If run Is Nothing Then
					' Add this picture as the last element.
					Xml.Add(p.Xml)

					' Extract the picture back out of the DOM.
					p_xml = CType(Xml.LastNode, XElement)

				Else
					' Split this run at the point you want to insert
					Dim splitRun() As XElement = Run.SplitRun(run, index)

					' Replace the origional run.
					run.Xml.ReplaceWith (splitRun(0), p.Xml, splitRun(1))

					' Get the first run effected by this Insert
					run = GetFirstRunEffectedByEdit(index)

					' The picture has to be the next element, extract it back out of the DOM.
					p_xml = CType(run.Xml.NextNode, XElement)
				End If
			End If
			' Extract the attribute id from the Pictures Xml.
			Dim a_id As XAttribute = (
			    From e In p_xml.Descendants()
			    Where e.Name.LocalName.Equals("blip")
			    Select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))).Single()

			' Set its value to the Pictures relationships id.
			a_id.SetValue(Id)


			Return Me
		End Function

		''' <summary>
		''' Append text on a new line to this Paragraph.
		''' </summary>
		''' <param name="text">The text to append.</param>
		''' <returns>This Paragraph with the new text appened.</returns>
		''' <example>
		''' Add a new Paragraph to this document and then append a new line with some text to it.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph and Append a new line with some text to it.
		'''     Paragraph p = document.InsertParagraph().AppendLine("Hello World!!!");
		'''       
		'''     // Save this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function AppendLine(ByVal text As String) As Paragraph
			Return Append(vbLf & text)
		End Function

		''' <summary>
		''' Append a new line to this Paragraph.
		''' </summary>
		''' <returns>This Paragraph with a new line appeneded.</returns>
		''' <example>
		''' Add a new Paragraph to this document and then append a new line to it.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph and Append a new line with some text to it.
		'''     Paragraph p = document.InsertParagraph().AppendLine();
		'''       
		'''     // Save this document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		Public Function AppendLine() As Paragraph
			Return Append(vbLf)
		End Function

		Friend Sub ApplyTextFormattingProperty(ByVal textFormatPropName As XName, ByVal value As String, ByVal content As Object)
			Dim rPr As XElement = Nothing

			If runs.Count = 0 Then
				Dim pPr As XElement = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName))
				If pPr Is Nothing Then
					Xml.AddFirst(New XElement(XName.Get("pPr", DocX.w.NamespaceName)))
					pPr = Xml.Element(XName.Get("pPr", DocX.w.NamespaceName))
				End If

				rPr = pPr.Element(XName.Get("rPr", DocX.w.NamespaceName))
				If rPr Is Nothing Then
					pPr.AddFirst(New XElement(XName.Get("rPr", DocX.w.NamespaceName)))
					rPr = pPr.Element(XName.Get("rPr", DocX.w.NamespaceName))
				End If

				rPr.SetElementValue(textFormatPropName, value)
				Dim last = rPr.Elements(textFormatPropName).Last()
				If TryCast(content, XAttribute) IsNot Nothing Then 'If content is an attribute
					If last.Attribute((CType(content, XAttribute)).Name) Is Nothing Then
						last.Add(content) 'Add this attribute if element doesn't have it
					Else
						last.Attribute((CType(content, XAttribute)).Name).Value = (CType(content, XAttribute)).Value 'Apply value only if element already has it
					End If
				End If
				Return
			End If

			Dim contentIsListOfFontProperties = False
			Dim fontProps = TryCast(content, IEnumerable)
			If fontProps IsNot Nothing Then
				For Each [property] As Object In fontProps
					contentIsListOfFontProperties = (TryCast([property], XAttribute) IsNot Nothing)
				Next [property]
			End If

			For Each run As XElement In runs
				rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName))
				If rPr Is Nothing Then
					run.AddFirst(New XElement(XName.Get("rPr", DocX.w.NamespaceName)))
					rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName))
				End If

				rPr.SetElementValue(textFormatPropName, value)
				Dim last As XElement = rPr.Elements(textFormatPropName).Last()

				If contentIsListOfFontProperties Then 'if content is a list of attributes, as in the case when specifying a font family
					For Each [property] As Object In fontProps
						If last.Attribute((CType([property], XAttribute)).Name) Is Nothing Then
							last.Add([property]) 'Add this attribute if element doesn't have it
						Else
							last.Attribute((CType([property], XAttribute)).Name).Value = (CType([property], XAttribute)).Value 'Apply value only if element already has it
						End If
					Next [property]
				End If

				If TryCast(content, XAttribute) IsNot Nothing Then 'If content is an attribute
					If last.Attribute((CType(content, XAttribute)).Name) Is Nothing Then
						last.Add(content) 'Add this attribute if element doesn't have it
					Else
						last.Attribute((CType(content, XAttribute)).Name).Value = (CType(content, XAttribute)).Value 'Apply value only if element already has it
					End If
				Else
					'IMPORTANT
					'But what to do if it is not?
				End If
			Next run
		End Sub

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <returns>This Paragraph with the last appended text bold.</returns>
		''' <example>
		''' Append text to this Paragraph and then make it bold.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Bold").Bold()
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function Bold() As Paragraph
			ApplyTextFormattingProperty(XName.Get("b", DocX.w.NamespaceName), String.Empty, Nothing)
			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <returns>This Paragraph with the last appended text italic.</returns>
		''' <example>
		''' Append text to this Paragraph and then make it italic.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Italic").Italic()
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function Italic() As Paragraph
			ApplyTextFormattingProperty(XName.Get("i", DocX.w.NamespaceName), String.Empty, Nothing)
			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="c">A color to use on the appended text.</param>
		''' <returns>This Paragraph with the last appended text colored.</returns>
		''' <example>
		''' Append text to this Paragraph and then color it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Blue").Color(Color.Blue)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function Color(ByVal c As Color) As Paragraph
			ApplyTextFormattingProperty(XName.Get("color", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), c.ToHex()))
			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="underlineStyle_Renamed">The underline style to use for the appended text.</param>
		''' <returns>This Paragraph with the last appended text underlined.</returns>
		''' <example>
		''' Append text to this Paragraph and then underline it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Underlined").UnderlineStyle(UnderlineStyle.doubleLine)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter underlineStyle was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function UnderlineStyle(ByVal underlineStyle_Renamed As UnderlineStyle) As Paragraph
			Dim value As String
			Select Case underlineStyle_Renamed
				Case Novacode.UnderlineStyle.none
					value = String.Empty
				Case Novacode.UnderlineStyle.singleLine
					value = "single"
				Case Novacode.UnderlineStyle.doubleLine
					value = "double"
				Case Else
					value = underlineStyle_Renamed.ToString()
			End Select

			ApplyTextFormattingProperty(XName.Get("u", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), value))
			Return Me
		End Function

'INSTANT VB NOTE: The variable followingTable was renamed since Visual Basic does not allow class members with the same name:
		Private followingTable_Renamed As Table

		'''<summary>
		''' Returns table following the paragraph. Null if the following element isn't table.
		'''</summary>
		Public Property FollowingTable() As Table
			Get
				Return followingTable_Renamed
			End Get
			Friend Set(ByVal value As Table)
				followingTable_Renamed = value
			End Set
		End Property

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="fontSize_Renamed">The font size to use for the appended text.</param>
		''' <returns>This Paragraph with the last appended text resized.</returns>
		''' <example>
		''' Append text to this Paragraph and then resize it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Big").FontSize(20)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter fontSize was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function FontSize(ByVal fontSize_Renamed As Double) As Paragraph
			If fontSize_Renamed - CInt(Fix(fontSize_Renamed)) = 0 Then
				If Not(fontSize_Renamed > 0 AndAlso fontSize_Renamed < 1639) Then
					Throw New ArgumentException("Size", "Value must be in the range 0 - 1638")
				End If

			Else
				Throw New ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5")
			End If

			ApplyTextFormattingProperty(XName.Get("sz", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize_Renamed * 2))
			ApplyTextFormattingProperty(XName.Get("szCs", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), fontSize_Renamed * 2))

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="fontFamily">The font to use for the appended text.</param>
		''' <returns>This Paragraph with the last appended text's font changed.</returns>
		''' <example>
		''' Append text to this Paragraph and then change its font.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Times new roman").Font(new FontFamily("Times new roman"))
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function Font(ByVal fontFamily As FontFamily) As Paragraph
			ApplyTextFormattingProperty(XName.Get("rFonts", DocX.w.NamespaceName), String.Empty, { New XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily.Name), New XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily.Name), New XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily.Name) }) ' Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865 -  Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="capsStyle_Renamed">The caps style to apply to the last appended text.</param>
		''' <returns>This Paragraph with the last appended text's caps style changed.</returns>
		''' <example>
		''' Append text to this Paragraph and then set it to full caps.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("Capitalized").CapsStyle(CapsStyle.caps)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter capsStyle was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function CapsStyle(ByVal capsStyle_Renamed As CapsStyle) As Paragraph
			Select Case capsStyle_Renamed
				Case Novacode.CapsStyle.none

				Case Else
						ApplyTextFormattingProperty(XName.Get(capsStyle_Renamed.ToString(), DocX.w.NamespaceName), String.Empty, Nothing)
						Exit Select
			End Select

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="script_Renamed">The script style to apply to the last appended text.</param>
		''' <returns>This Paragraph with the last appended text's script style changed.</returns>
		''' <example>
		''' Append text to this Paragraph and then set it to superscript.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("superscript").Script(Script.superscript)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter script was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Script(ByVal script_Renamed As Script) As Paragraph
			Select Case script_Renamed
				Case Novacode.Script.none

				Case Else
						ApplyTextFormattingProperty(XName.Get("vertAlign", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), script_Renamed.ToString()))
						Exit Select
			End Select

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		'''<param name="highlight_Renamed">The highlight to apply to the last appended text.</param>
		''' <returns>This Paragraph with the last appended text highlighted.</returns>
		''' <example>
		''' Append text to this Paragraph and then highlight it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("highlighted").Highlight(Highlight.green)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter highlight was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Highlight(ByVal highlight_Renamed As Highlight) As Paragraph
			Select Case highlight_Renamed
				Case Novacode.Highlight.none

				Case Else
						ApplyTextFormattingProperty(XName.Get("highlight", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight_Renamed.ToString()))
						Exit Select
			End Select

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="misc_Renamed">The miscellaneous property to set.</param>
		''' <returns>This Paragraph with the last appended text changed by a miscellaneous property.</returns>
		''' <example>
		''' Append text to this Paragraph and then apply a miscellaneous property.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("outlined").Misc(Misc.outline)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter misc was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Misc(ByVal misc_Renamed As Misc) As Paragraph
			Select Case misc_Renamed
				Case Novacode.Misc.none

				Case Novacode.Misc.outlineShadow
						ApplyTextFormattingProperty(XName.Get("outline", DocX.w.NamespaceName), String.Empty, Nothing)
						ApplyTextFormattingProperty(XName.Get("shadow", DocX.w.NamespaceName), String.Empty, Nothing)

						Exit Select

				Case Novacode.Misc.engrave
						ApplyTextFormattingProperty(XName.Get("imprint", DocX.w.NamespaceName), String.Empty, Nothing)

						Exit Select

				Case Else
						ApplyTextFormattingProperty(XName.Get(misc_Renamed.ToString(), DocX.w.NamespaceName), String.Empty, Nothing)

						Exit Select
			End Select

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="strikeThrough_Renamed">The strike through style to used on the last appended text.</param>
		''' <returns>This Paragraph with the last appended text striked.</returns>
		''' <example>
		''' Append text to this Paragraph and then strike it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("striked").StrikeThrough(StrikeThrough.doubleStrike)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter strikeThrough was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function StrikeThrough(ByVal strikeThrough_Renamed As StrikeThrough) As Paragraph
			Dim value As String
			Select Case strikeThrough_Renamed
				Case Novacode.StrikeThrough.strike
					value = "strike"
				Case Novacode.StrikeThrough.doubleStrike
					value = "dstrike"
				Case Else
					Return Me
			End Select

			ApplyTextFormattingProperty(XName.Get(value, DocX.w.NamespaceName), String.Empty, Nothing)

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <param name="underlineColor_Renamed">The underline color to use, if no underline is set, a single line will be used.</param>
		''' <returns>This Paragraph with the last appended text underlined in a color.</returns>
		''' <example>
		''' Append text to this Paragraph and then underline it using a color.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("color underlined").UnderlineStyle(UnderlineStyle.dotted).UnderlineColor(Color.Orange)
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
'INSTANT VB NOTE: The parameter underlineColor was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function UnderlineColor(ByVal underlineColor_Renamed As Color) As Paragraph
			For Each run As XElement In runs
				Dim rPr As XElement = run.Element(XName.Get("rPr", DocX.w.NamespaceName))
				If rPr Is Nothing Then
					run.AddFirst(New XElement(XName.Get("rPr", DocX.w.NamespaceName)))
					rPr = run.Element(XName.Get("rPr", DocX.w.NamespaceName))
				End If

				Dim u As XElement = rPr.Element(XName.Get("u", DocX.w.NamespaceName))
				If u Is Nothing Then
					rPr.SetElementValue(XName.Get("u", DocX.w.NamespaceName), String.Empty)
					u = rPr.Element(XName.Get("u", DocX.w.NamespaceName))
					u.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), "single")
				End If

				u.SetAttributeValue(XName.Get("color", DocX.w.NamespaceName), underlineColor_Renamed.ToHex())
			Next run

			Return Me
		End Function

		''' <summary>
		''' For use with Append() and AppendLine()
		''' </summary>
		''' <returns>This Paragraph with the last appended text hidden.</returns>
		''' <example>
		''' Append text to this Paragraph and then hide it.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Insert a new Paragraph.
		'''     Paragraph p = document.InsertParagraph();
		'''
		'''     p.Append("I am ")
		'''     .Append("hidden").Hide()
		'''     .Append(" I am not");
		'''        
		'''     // Save this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function Hide() As Paragraph
			ApplyTextFormattingProperty(XName.Get("vanish", DocX.w.NamespaceName), String.Empty, Nothing)

			Return Me
		End Function

		Public Property LineSpacing() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim spacing As XElement = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))

				If spacing IsNot Nothing Then
					Dim line As XAttribute = spacing.Attribute(XName.Get("line", DocX.w.NamespaceName))
					If line IsNot Nothing Then
						Dim f As Single

						If Single.TryParse(line.Value, f) Then
							Return f / 20.0f
						End If
					End If
				End If

				Return 1.1f * 20.0f
			End Get

			Set(ByVal value As Single)
				Spacing(value)
			End Set
		End Property


		''' <summary>
		''' Set the linespacing for this paragraph manually.
		''' </summary>
		''' <param name="spacingType">The type of spacing to be set, can be either Before, After or Line (Standard line spacing).</param>
		''' <param name="spacingFloat">A float value of the amount of spacing. Equals the value that van be set in Word using the "Line and Paragraph spacing" button.</param>
		Public Sub SetLineSpacing(ByVal spacingType As LineSpacingType, ByVal spacingFloat As Single)
			spacingFloat = spacingFloat * 240
			Dim spacingValue As Integer = CInt(Fix(spacingFloat))

			Dim pPr = Me.GetOrCreate_pPr()
			Dim spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))
			If spacing Is Nothing Then
				pPr.Add(New XElement(XName.Get("spacing", DocX.w.NamespaceName)))
				spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))
			End If

			Dim spacingTypeAttribute As String = ""
			Select Case spacingType
				Case LineSpacingType.Line
						spacingTypeAttribute = "line"
						Exit Select
				Case LineSpacingType.Before
						spacingTypeAttribute = "before"
						Exit Select
				Case LineSpacingType.After
						spacingTypeAttribute = "after"
						Exit Select

			End Select

			spacing.SetAttributeValue(XName.Get(spacingTypeAttribute, DocX.w.NamespaceName), spacingValue)
		End Sub

		''' <summary>
		''' Set the linespacing for this paragraph using the Auto value.
		''' </summary>
		''' <param name="spacingType">The type of spacing to be set automatically. Using Auto will set both Before and After. None will remove any linespacing.</param>
		Public Sub SetLineSpacing(ByVal spacingType As LineSpacingTypeAuto)
			Dim spacingValue As Integer = 100

			Dim pPr = Me.GetOrCreate_pPr()
			Dim spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))

			If spacingType.Equals(LineSpacingTypeAuto.None) Then
				If spacing IsNot Nothing Then
					spacing.Remove()
				End If

			Else

				If spacing Is Nothing Then
					pPr.Add(New XElement(XName.Get("spacing", DocX.w.NamespaceName)))
					spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))
				End If

				Dim spacingTypeAttribute As String = ""
				Dim autoSpacingTypeAttribute As String = ""
				Select Case spacingType
					Case LineSpacingTypeAuto.AutoBefore
							spacingTypeAttribute = "before"
							autoSpacingTypeAttribute = "beforeAutospacing"
							Exit Select
					Case LineSpacingTypeAuto.AutoAfter
							spacingTypeAttribute = "after"
							autoSpacingTypeAttribute = "afterAutospacing"
							Exit Select
					Case LineSpacingTypeAuto.Auto
							spacingTypeAttribute = "before"
							autoSpacingTypeAttribute = "beforeAutospacing"
							spacing.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), spacingValue)
							spacing.SetAttributeValue(XName.Get("afterAutospacing", DocX.w.NamespaceName), 1)
							Exit Select

				End Select

				spacing.SetAttributeValue(XName.Get(autoSpacingTypeAttribute, DocX.w.NamespaceName), 1)
				spacing.SetAttributeValue(XName.Get(spacingTypeAttribute, DocX.w.NamespaceName), spacingValue)

			End If

		End Sub


'INSTANT VB NOTE: The parameter spacing was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Spacing(ByVal spacing_Renamed As Double) As Paragraph
			spacing_Renamed *= 20

			If spacing_Renamed - CInt(Fix(spacing_Renamed)) = 0 Then
				If Not(spacing_Renamed > -1585 AndAlso spacing_Renamed < 1585) Then
					Throw New ArgumentException("Spacing", "Value must be in the range: -1584 - 1584")
				End If

			Else
				Throw New ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9")
			End If

			ApplyTextFormattingProperty(XName.Get("spacing", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing_Renamed))

			Return Me
		End Function

'INSTANT VB NOTE: The parameter spacingBefore was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function SpacingBefore(ByVal spacingBefore_Renamed As Double) As Paragraph
			spacingBefore_Renamed *= 20

			Dim pPr = GetOrCreate_pPr()
			Dim spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))
			If spacingBefore_Renamed > 0 Then
				If spacing Is Nothing Then
					spacing = New XElement(XName.Get("spacing", DocX.w.NamespaceName))
					pPr.Add(spacing)
				End If
				Dim attr = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName))
				If attr Is Nothing Then
					spacing.SetAttributeValue(XName.Get("before", DocX.w.NamespaceName), spacingBefore_Renamed)
				Else
					attr.SetValue(spacingBefore_Renamed)
				End If
			End If
			If Math.Abs(spacingBefore_Renamed) < 0.1f AndAlso spacing IsNot Nothing Then
				Dim attr = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName))
				attr.Remove()
				If Not spacing.HasAttributes Then
					spacing.Remove()
				End If
			End If

			Return Me
		End Function

'INSTANT VB NOTE: The parameter spacingAfter was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function SpacingAfter(ByVal spacingAfter_Renamed As Double) As Paragraph
			spacingAfter_Renamed *= 20

			Dim pPr = GetOrCreate_pPr()
			Dim spacing = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))
			If spacingAfter_Renamed > 0 Then
				If spacing Is Nothing Then
					spacing = New XElement(XName.Get("spacing", DocX.w.NamespaceName))
					pPr.Add(spacing)
				End If
				Dim attr = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName))
				If attr Is Nothing Then
					spacing.SetAttributeValue(XName.Get("after", DocX.w.NamespaceName), spacingAfter_Renamed)
				Else
					attr.SetValue(spacingAfter_Renamed)
				End If
			End If
			If Math.Abs(spacingAfter_Renamed) < 0.1f AndAlso spacing IsNot Nothing Then
				Dim attr = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName))
				attr.Remove()
				If Not spacing.HasAttributes Then
					spacing.Remove()
				End If
			End If
			'ApplyTextFormattingProperty(XName.Get("after", DocX.w.NamespaceName), string.Empty, new XAttribute(XName.Get("val", DocX.w.NamespaceName), spacingAfter));

			Return Me
		End Function

'INSTANT VB NOTE: The parameter kerning was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Kerning(ByVal kerning_Renamed As Integer) As Paragraph
			If (Not New Integer?()) { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 }.Contains(kerning) Then
				Throw New ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72")
			End If

			ApplyTextFormattingProperty(XName.Get("kern", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning_Renamed * 2))
			Return Me
		End Function

'INSTANT VB NOTE: The parameter position was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function Position(ByVal position_Renamed As Double) As Paragraph
			If Not(position_Renamed > -1585 AndAlso position_Renamed < 1585) Then
				Throw New ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585")
			End If

			ApplyTextFormattingProperty(XName.Get("position", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), position_Renamed * 2))

			Return Me
		End Function

'INSTANT VB NOTE: The parameter percentageScale was renamed since Visual Basic will not allow parameters with the same name as their enclosing function or property:
		Public Function PercentageScale(ByVal percentageScale_Renamed As Integer) As Paragraph
			If Not(New Integer?() { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(percentageScale) Then
				Throw New ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33")
			End If

			ApplyTextFormattingProperty(XName.Get("w", DocX.w.NamespaceName), String.Empty, New XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale_Renamed))

			Return Me
		End Function

		''' <summary>
		''' Append a field of type document property, this field will display the custom property cp, at the end of this paragraph.
		''' </summary>
		''' <param name="cp">The custom property to display.</param>
		''' <param name="trackChanges"></param>
		''' <param name="f">The formatting to use for this text.</param>
		''' <example>
		''' Create, add and display a custom property in a document.
		''' <code>
		''' // Load a document.
		'''using (DocX document = DocX.Create("CustomProperty_Add.docx"))
		'''{
		'''    // Add a few Custom Properties to this document.
		'''    document.AddCustomProperty(new CustomProperty("fname", "cathal"));
		'''    document.AddCustomProperty(new CustomProperty("age", 24));
		'''    document.AddCustomProperty(new CustomProperty("male", true));
		'''    document.AddCustomProperty(new CustomProperty("newyear2012", new DateTime(2012, 1, 1)));
		'''    document.AddCustomProperty(new CustomProperty("fav_num", 3.141592));
		'''
		'''    // Insert a new Paragraph and append a load of DocProperties.
		'''    Paragraph p = document.InsertParagraph("fname: ")
		'''        .AppendDocProperty(document.CustomProperties["fname"])
		'''        .Append(", age: ")
		'''        .AppendDocProperty(document.CustomProperties["age"])
		'''        .Append(", male: ")
		'''        .AppendDocProperty(document.CustomProperties["male"])
		'''        .Append(", newyear2012: ")
		'''        .AppendDocProperty(document.CustomProperties["newyear2012"])
		'''        .Append(", fav_num: ")
		'''        .AppendDocProperty(document.CustomProperties["fav_num"]);
		'''    
		'''    // Save the changes to the document.
		'''    document.Save();
		'''}
		''' </code>
		''' </example>
		Public Function AppendDocProperty(ByVal cp As CustomProperty, Optional ByVal trackChanges As Boolean = False, Optional ByVal f As Formatting = Nothing) As Paragraph
			Me.InsertDocProperty(cp, trackChanges, f)
			Return Me
		End Function

		''' <summary>
		''' Insert a field of type document property, this field will display the custom property cp, at the end of this paragraph.
		''' </summary>
		''' <param name="cp">The custom property to display.</param>
		''' <param name="trackChanges"></param>
		''' <param name="f">The formatting to use for this text.</param>
		''' <example>
		''' Create, add and display a custom property in a document.
		''' <code>
		''' // Load a document
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Create a custom property.
		'''     CustomProperty name = new CustomProperty("name", "Cathal Coffey");
		'''        
		'''     // Add this custom property to this document.
		'''     document.AddCustomProperty(name);
		'''
		'''     // Create a text formatting.
		'''     Formatting f = new Formatting();
		'''     f.Bold = true;
		'''     f.Size = 14;
		'''     f.StrikeThrough = StrickThrough.strike;
		'''
		'''     // Insert a new paragraph.
		'''     Paragraph p = document.InsertParagraph("Author: ", false, f);
		'''
		'''     // Insert a field of type document property to display the custom property name and track this change.
		'''     p.InsertDocProperty(name, true, f);
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function InsertDocProperty(ByVal cp As CustomProperty, Optional ByVal trackChanges As Boolean = False, Optional ByVal f As Formatting = Nothing) As DocProperty
			Dim f_xml As XElement = Nothing
			If f IsNot Nothing Then
				f_xml = f.Xml
			End If

			Dim e As New XElement(XName.Get("fldSimple", DocX.w.NamespaceName), New XAttribute(XName.Get("instr", DocX.w.NamespaceName), String.Format("DOCPROPERTY {0} \* MERGEFORMAT", cp.Name)), New XElement(XName.Get("r", DocX.w.NamespaceName), New XElement(XName.Get("t", DocX.w.NamespaceName), f_xml, cp.Value)))

			Dim xml As XElement = e
			If trackChanges Then
				Dim now As Date = Date.Now
				Dim insert_datetime As New Date(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc)
				e = CreateEdit(EditType.ins, insert_datetime, e)
			End If

			Me.Xml.Add(e)

			Return New DocProperty(Document, xml)
		End Function

		''' <summary>
		''' Removes characters from a Novacode.DocX.Paragraph.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Iterate through the paragraphs
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         // Remove the first two characters from every paragraph
		'''         p.RemoveText(0, 2, false);
		'''     }
		'''        
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		''' <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		''' <param name="index">The position to begin deleting characters.</param>
		''' <param name="count">The number of characters to delete</param>
		''' <param name="trackChanges">Track changes</param>
		Public Sub RemoveText(ByVal index As Integer, ByVal count As Integer, Optional ByVal trackChanges As Boolean = False)
			' Timestamp to mark the start of insert
			Dim now As Date = Date.Now
			Dim remove_datetime As New Date(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc)

			' The number of characters processed so far
			Dim processed As Integer = 0

			Do
				' Get the first run effected by this Remove
				Dim run As Run = GetFirstRunEffectedByEdit(index, EditType.del)

				' The parent of this Run
				Dim parentElement As XElement = run.Xml.Parent
				Select Case parentElement.Name.LocalName
					Case "ins"
					CaseLabel1:
							Dim splitEditBefore() As XElement = SplitEdit(parentElement, index, EditType.del)
							Dim min As Integer = Math.Min(count - processed, run.Xml.ElementsAfterSelf().Sum(Function(e) GetElementTextLength(e)))
							Dim splitEditAfter() As XElement = SplitEdit(parentElement, index + min, EditType.del)

							Dim temp As XElement = SplitEdit(splitEditBefore(1), index + min, EditType.del)(0)
							Dim middle As Object = CreateEdit(EditType.del, remove_datetime, temp.Elements())
							processed += GetElementTextLength(TryCast(middle, XElement))

							If Not trackChanges Then
								middle = Nothing
							End If

							parentElement.ReplaceWith (splitEditBefore(0), middle, splitEditAfter(1))

							processed += GetElementTextLength(TryCast(middle, XElement))
							Exit Select

					Case "del"
							If trackChanges Then
								' You cannot delete from a deletion, advance processed to the end of this del
								processed += GetElementTextLength(parentElement)

							Else
								GoTo CaseLabel1
							End If

							Exit Select

					Case Else
							Dim splitRunBefore() As XElement = Run.SplitRun(run, index, EditType.del)
							'int min = Math.Min(index + processed + (count - processed), run.EndIndex);
							Dim min As Integer = Math.Min(index + (count - processed), run.EndIndex)
							Dim splitRunAfter() As XElement = Run.SplitRun(run, min, EditType.del)

							Dim middle As Object = CreateEdit(EditType.del, remove_datetime, New List(Of XElement)() From {Run.SplitRun(New Run(Document, splitRunBefore(1), run.StartIndex + GetElementTextLength(splitRunBefore(0))), min, EditType.del)(0)})
							processed += GetElementTextLength(TryCast(middle, XElement))

							If Not trackChanges Then
								middle = Nothing
							End If

							run.Xml.ReplaceWith (splitRunBefore(0), middle, splitRunAfter(1))

							Exit Select
				End Select

				' If after this remove the parent element is empty, remove it.
				If GetElementTextLength(parentElement) = 0 Then
					If parentElement.Parent IsNot Nothing AndAlso parentElement.Parent.Name.LocalName <> "tc" Then
						' Need to make sure there is no drawing element within the parent element.
						' Picture elements contain no text length but they are still content.
						If parentElement.Descendants(XName.Get("drawing", DocX.w.NamespaceName)).Count() = 0 Then
							parentElement.Remove()
						End If
					End If
				End If
			Loop While processed < count

			HelperFunctions.RenumberIDs(Document)
		End Sub


		''' <summary>
		''' Removes characters from a Novacode.DocX.Paragraph.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Iterate through the paragraphs
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         // Remove all but the first 2 characters from this Paragraph.
		'''         p.RemoveText(2, false);
		'''     }
		'''        
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		''' <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		''' <param name="index">The position to begin deleting characters.</param>
		''' <param name="trackChanges">Track changes</param>
		Public Sub RemoveText(ByVal index As Integer, Optional ByVal trackChanges As Boolean = False)
			RemoveText(index, Text.Length - index, trackChanges)
		End Sub

		''' <summary>
		''' Replaces all occurrences of a specified System.String in this instance, with another specified System.String.
		''' </summary>
		''' <example>
		''' <code>
		''' // Load a document using a relative filename.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // The formatting to match.
		'''     Formatting matchFormatting = new Formatting();
		'''     matchFormatting.Size = 10;
		'''     matchFormatting.Italic = true;
		'''     matchFormatting.FontFamily = new FontFamily("Times New Roman");
		'''
		'''     // The formatting to apply to the inserted text.
		'''     Formatting newFormatting = new Formatting();
		'''     newFormatting.Size = 22;
		'''     newFormatting.UnderlineStyle = UnderlineStyle.dotted;
		'''     newFormatting.Bold = true;
		'''
		'''     // Iterate through the paragraphs in this document.
		'''     foreach (Paragraph p in document.Paragraphs)
		'''     {
		'''         /* 
		'''          * Replace all instances of the string "wrong" with the string "right" and ignore case.
		'''          * Each inserted instance of "wrong" should use the Formatting newFormatting.
		'''          * Only replace an instance of "wrong" if it is Size 10, Italic and Times New Roman.
		'''          * SubsetMatch means that the formatting must contain all elements of the match formatting,
		'''          * but it can also contain additional formatting for example Color, UnderlineStyle, etc.
		'''          * ExactMatch means it must not contain additional formatting.
		'''          */
		'''         p.ReplaceText("wrong", "right", false, RegexOptions.IgnoreCase, newFormatting, matchFormatting, MatchFormattingOptions.SubsetMatch);
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="Paragraph.RemoveText(int, int, bool)"/>
		''' <seealso cref="Paragraph.RemoveText(int, bool)"/>
		''' <seealso cref="Paragraph.InsertText(int, string, bool, Formatting)"/>
		''' <seealso cref="Paragraph.InsertText(string, bool, Formatting)"/>
		''' <param name="newValue">A System.String to replace all occurrences of oldValue.</param>
		''' <param name="oldValue">A System.String to be replaced.</param>
		''' <param name="options">A bitwise OR combination of RegexOption enumeration options.</param>
		''' <param name="trackChanges">Track changes</param>
		''' <param name="newFormatting">The formatting to apply to the text being inserted.</param>
		''' <param name="matchFormatting">The formatting that the text must match in order to be replaced.</param>
		''' <param name="fo">How should formatting be matched?</param>
		''' <param name="escapeRegEx">True if the oldValue needs to be escaped, otherwise false. If it represents a valid RegEx pattern this should be false.</param>
		''' <param name="useRegExSubstitutions">True if RegEx-like replace should be performed, i.e. if newValue contains RegEx substitutions. Does not perform named-group substitutions (only numbered groups).</param>
		Public Sub ReplaceText(ByVal oldValue As String, ByVal newValue As String, Optional ByVal trackChanges As Boolean = False, Optional ByVal options As RegexOptions = RegexOptions.None, Optional ByVal newFormatting As Formatting = Nothing, Optional ByVal matchFormatting As Formatting = Nothing, Optional ByVal fo As MatchFormattingOptions = MatchFormattingOptions.SubsetMatch, Optional ByVal escapeRegEx As Boolean = True, Optional ByVal useRegExSubstitutions As Boolean = False)
			Dim tText As String = Text
			Dim mc As MatchCollection = Regex.Matches(tText,If(escapeRegEx, Regex.Escape(oldValue), oldValue), options)

			' Loop through the matches in reverse order
			For Each m As Match In mc.Cast(Of Match)().Reverse()
				' Assume the formatting matches until proven otherwise.
				Dim formattingMatch As Boolean = True

				' Does the user want to match formatting?
				If matchFormatting IsNot Nothing Then
					' The number of characters processed so far
					Dim processed As Integer = 0

					Do
						' Get the next run effected
						Dim run As Run = GetFirstRunEffectedByEdit(m.Index + processed)

						' Get this runs properties
						Dim rPr As XElement = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName))

						If rPr Is Nothing Then
							rPr = New Formatting().Xml
						End If

'                         
'                         * Make sure that every formatting element in f.xml is also in this run,
'                         * if this is not true, then their formatting does not match.
'                         
						If Not HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo) Then
							formattingMatch = False
							Exit Do
						End If

						' We have processed some characters, so update the counter.
						processed += run.Value.Length

					Loop While processed < m.Length
				End If

				' If the formatting matches, do the replace.
				If formattingMatch Then
					Dim repl As String = newValue
					'perform RegEx substitutions. Only named groups are not supported. Everything else is supported. However character escapes are not covered.
					If useRegExSubstitutions AndAlso (Not String.IsNullOrEmpty(repl)) Then
						repl = repl.Replace("$&", m.Value)
						If m.Groups.Count > 0 Then
							Dim lastcap As Integer = 0
							For k As Integer = 0 To m.Groups.Count - 1
								Dim g = m.Groups(k)
								If (g Is Nothing) OrElse (g.Value = "") Then
									Continue For
								End If
								repl = repl.Replace("$" & k.ToString(), g.Value)
								lastcap = k
								'cannot get named groups ATM
							Next k
							repl = repl.Replace("$+", m.Groups(lastcap).Value)
						End If
						If m.Index > 0 Then
							repl = repl.Replace("$`", tText.Substring(0, m.Index))
						End If
						If (m.Index + m.Length) < tText.Length Then
							repl = repl.Replace("$'", tText.Substring(m.Index + m.Length))
						End If
						repl = repl.Replace("$_", tText)
						repl = repl.Replace("$$", "$")
					End If
					If Not String.IsNullOrEmpty(repl) Then
						InsertText(m.Index + m.Length, repl, trackChanges, newFormatting)
					End If
					If m.Length > 0 Then
						RemoveText(m.Index, m.Length, trackChanges)
					End If
				End If
			Next m
		End Sub

		''' <summary>
		''' Find pattern regex must return a group match.
		''' </summary>
		''' <param name="findPattern">Regex pattern that must include one group match. ie (.*)</param>
		''' <param name="regexMatchHandler">A func that accepts the matching find grouping text and returns a replacement value</param>
		''' <param name="trackChanges"></param>
		''' <param name="options"></param>
		''' <param name="newFormatting"></param>
		''' <param name="matchFormatting"></param>
		''' <param name="fo"></param>
		Public Sub ReplaceText(ByVal findPattern As String, ByVal regexMatchHandler As Func(Of String,String), Optional ByVal trackChanges As Boolean = False, Optional ByVal options As RegexOptions = RegexOptions.None, Optional ByVal newFormatting As Formatting = Nothing, Optional ByVal matchFormatting As Formatting = Nothing, Optional ByVal fo As MatchFormattingOptions = MatchFormattingOptions.SubsetMatch)
			Dim matchCollection = Regex.Matches(Text, findPattern, options)

			' Loop through the matches in reverse order
			For Each match In matchCollection.Cast(Of Match)().Reverse()
				' Assume the formatting matches until proven otherwise.
				Dim formattingMatch As Boolean = True

				' Does the user want to match formatting?
				If matchFormatting IsNot Nothing Then
					' The number of characters processed so far
					Dim processed As Integer = 0

					Do
						' Get the next run effected
						Dim run As Run = GetFirstRunEffectedByEdit(match.Index + processed)

						' Get this runs properties
						Dim rPr As XElement = run.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName))

						If rPr Is Nothing Then
							rPr = New Formatting().Xml
						End If

'                         
'                         * Make sure that every formatting element in f.xml is also in this run,
'                         * if this is not true, then their formatting does not match.
'                         
						If Not HelperFunctions.ContainsEveryChildOf(matchFormatting.Xml, rPr, fo) Then
							formattingMatch = False
							Exit Do
						End If

						' We have processed some characters, so update the counter.
						processed += run.Value.Length

					Loop While processed < match.Length
				End If

				' If the formatting matches, do the replace.
				If formattingMatch Then
					Dim newValue = regexMatchHandler.Invoke(match.Groups(1).Value)
					InsertText(match.Index + match.Value.Length, newValue, trackChanges, newFormatting)
					RemoveText(match.Index, match.Value.Length, trackChanges)
				End If
			Next match
		End Sub


		''' <summary>
		''' Find all instances of a string in this paragraph and return their indexes in a List.
		''' </summary>
		''' <param name="str">The string to find</param>
		''' <returns>A list of indexes.</returns>
		''' <example>
		''' Find all instances of Hello in this document and insert 'don't' in frount of them.
		''' <code>
		''' // Load a document
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''     // Loop through the paragraphs in this document.
		'''     foreach(Paragraph p in document.Paragraphs)
		'''     {
		'''         // Find all instances of 'go' in this paragraph.
		'''         <![CDATA[ List<int> ]]> gos = document.FindAll("go");
		'''
		'''         /* 
		'''          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
		'''          * An important trick here is to do the inserting in reverse document order. If you inserted 
		'''          * in document order, every insert would shift the index of the remaining matches.
		'''          */
		'''         gos.Reverse();
		'''         foreach (int index in gos)
		'''         {
		'''             p.InsertText(index, "don't ", false);
		'''         }
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function FindAll(ByVal str As String) As List(Of Integer)
			Return FindAll(str, RegexOptions.None)
		End Function

		''' <summary>
		''' Find all instances of a string in this paragraph and return their indexes in a List.
		''' </summary>
		''' <param name="str">The string to find</param>
		''' <param name="options">The options to use when finding a string match.</param>
		''' <returns>A list of indexes.</returns>
		''' <example>
		''' Find all instances of Hello in this document and insert 'don't' in frount of them.
		''' <code>
		''' // Load a document
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''     // Loop through the paragraphs in this document.
		'''     foreach(Paragraph p in document.Paragraphs)
		'''     {
		'''         // Find all instances of 'go' in this paragraph (Ignore case).
		'''         <![CDATA[ List<int> ]]> gos = document.FindAll("go", RegexOptions.IgnoreCase);
		'''
		'''         /* 
		'''          * Insert 'don't' in frount of every instance of 'go' in this document to produce 'don't go'.
		'''          * An important trick here is to do the inserting in reverse document order. If you inserted 
		'''          * in document order, every insert would shift the index of the remaining matches.
		'''          */
		'''         gos.Reverse();
		'''         foreach (int index in gos)
		'''         {
		'''             p.InsertText(index, "don't ", false);
		'''         }
		'''     }
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Function FindAll(ByVal str As String, ByVal options As RegexOptions) As List(Of Integer)
			Dim mc As MatchCollection = Regex.Matches(Me.Text, Regex.Escape(str), options)

			Dim query = (
			    From m In mc.Cast(Of Match)()
			    Select m.Index).ToList()

			Return query
		End Function

		''' <summary>
		'''  Find all unique instances of the given Regex Pattern
		''' </summary>
		''' <param name="str"></param>
		''' <param name="options"></param>
		''' <returns></returns>
		Public Function FindAllByPattern(ByVal str As String, ByVal options As RegexOptions) As List(Of String)
			Dim mc As MatchCollection = Regex.Matches(Me.Text, str, options)

			Dim query = (
			    From m In mc.Cast(Of Match)()
			    Select m.Value).ToList()

			Return query
		End Function

		''' <summary>
		''' Insert a PageNumber place holder into a Paragraph.
		''' This place holder should only be inserted into a Header or Footer Paragraph.
		''' Word will not automatically update this field if it is inserted into a document level Paragraph.
		''' </summary>
		''' <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		''' <param name="index">The text index to insert this PageNumber place holder at.</param>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add Headers to the document.
		'''     document.AddHeaders();
		'''
		'''     // Get the default Header.
		'''     Header header = document.Headers.odd;
		'''
		'''     // Insert a Paragraph into the Header.
		'''     Paragraph p0 = header.InsertParagraph("Page ( of )");
		'''
		'''     // Insert place holders for PageNumber and PageCount into the Header.
		'''     // Word will replace these with the correct value for each Page.
		'''     p0.InsertPageNumber(PageNumberFormat.normal, 6);
		'''     p0.InsertPageCount(PageNumberFormat.normal, 11);
		'''
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AppendPageCount"/>
		''' <seealso cref="AppendPageNumber"/>
		''' <seealso cref="InsertPageCount"/>
		Public Sub InsertPageNumber(ByVal pnf As PageNumberFormat, Optional ByVal index As Integer = 0)
			Dim fldSimple As New XElement(XName.Get("fldSimple", DocX.w.NamespaceName))

			If pnf = PageNumberFormat.normal Then
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE   \* MERGEFORMAT "))
			Else
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE  \* ROMAN  \* MERGEFORMAT "))
			End If

			Dim content As XElement = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                   <w:rPr>" & ControlChars.CrLf & "                       <w:noProof /> " & ControlChars.CrLf & "                   </w:rPr>" & ControlChars.CrLf & "                   <w:t>1</w:t> " & ControlChars.CrLf & "               </w:r>")

			fldSimple.Add(content)

			If index = 0 Then
				Xml.AddFirst(fldSimple)
			Else
				Dim r As Run = GetFirstRunEffectedByEdit(index, EditType.ins)
				Dim splitEdit() As XElement = Me.SplitEdit(r.Xml, index, EditType.ins)
				r.Xml.ReplaceWith (splitEdit(0), fldSimple, splitEdit(1))
			End If
		End Sub

		''' <summary>
		''' Append a PageNumber place holder onto the end of a Paragraph.
		''' </summary>
		''' <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add Headers to the document.
		'''     document.AddHeaders();
		'''
		'''     // Get the default Header.
		'''     Header header = document.Headers.odd;
		'''
		'''     // Insert a Paragraph into the Header.
		'''     Paragraph p0 = header.InsertParagraph();
		'''
		'''     // Appemd place holders for PageNumber and PageCount into the Header.
		'''     // Word will replace these with the correct value for each Page.
		'''     p0.Append("Page (");
		'''     p0.AppendPageNumber(PageNumberFormat.normal);
		'''     p0.Append(" of ");
		'''     p0.AppendPageCount(PageNumberFormat.normal);
		'''     p0.Append(")");
		''' 
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AppendPageCount"/>
		''' <seealso cref="InsertPageNumber"/>
		''' <seealso cref="InsertPageCount"/>
		Public Sub AppendPageNumber(ByVal pnf As PageNumberFormat)
			Dim fldSimple As New XElement(XName.Get("fldSimple", DocX.w.NamespaceName))

			If pnf = PageNumberFormat.normal Then
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE   \* MERGEFORMAT "))
			Else
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " PAGE  \* ROMAN  \* MERGEFORMAT "))
			End If

			Dim content As XElement = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                   <w:rPr>" & ControlChars.CrLf & "                       <w:noProof /> " & ControlChars.CrLf & "                   </w:rPr>" & ControlChars.CrLf & "                   <w:t>1</w:t> " & ControlChars.CrLf & "               </w:r>")

			fldSimple.Add(content)
			Xml.Add(fldSimple)
		End Sub

		''' <summary>
		''' Insert a PageCount place holder into a Paragraph.
		''' This place holder should only be inserted into a Header or Footer Paragraph.
		''' Word will not automatically update this field if it is inserted into a document level Paragraph.
		''' </summary>
		''' <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		''' <param name="index">The text index to insert this PageCount place holder at.</param>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add Headers to the document.
		'''     document.AddHeaders();
		'''
		'''     // Get the default Header.
		'''     Header header = document.Headers.odd;
		'''
		'''     // Insert a Paragraph into the Header.
		'''     Paragraph p0 = header.InsertParagraph("Page ( of )");
		'''
		'''     // Insert place holders for PageNumber and PageCount into the Header.
		'''     // Word will replace these with the correct value for each Page.
		'''     p0.InsertPageNumber(PageNumberFormat.normal, 6);
		'''     p0.InsertPageCount(PageNumberFormat.normal, 11);
		'''
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AppendPageCount"/>
		''' <seealso cref="AppendPageNumber"/>
		''' <seealso cref="InsertPageNumber"/>
		Public Sub InsertPageCount(ByVal pnf As PageNumberFormat, Optional ByVal index As Integer = 0)
			Dim fldSimple As New XElement(XName.Get("fldSimple", DocX.w.NamespaceName))

			If pnf = PageNumberFormat.normal Then
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES   \* MERGEFORMAT "))
			Else
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES  \* ROMAN  \* MERGEFORMAT "))
			End If

			Dim content As XElement = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                   <w:rPr>" & ControlChars.CrLf & "                       <w:noProof /> " & ControlChars.CrLf & "                   </w:rPr>" & ControlChars.CrLf & "                   <w:t>1</w:t> " & ControlChars.CrLf & "               </w:r>")

			fldSimple.Add(content)

			If index = 0 Then
				Xml.AddFirst(fldSimple)
			Else
				Dim r As Run = GetFirstRunEffectedByEdit(index, EditType.ins)
				Dim splitEdit() As XElement = Me.SplitEdit(r.Xml, index, EditType.ins)
				r.Xml.ReplaceWith (splitEdit(0), fldSimple, splitEdit(1))
			End If
		End Sub

		''' <summary>
		''' Append a PageCount place holder onto the end of a Paragraph.
		''' </summary>
		''' <param name="pnf">The PageNumberFormat can be normal: (1, 2, ...) or Roman: (I, II, ...)</param>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add Headers to the document.
		'''     document.AddHeaders();
		'''
		'''     // Get the default Header.
		'''     Header header = document.Headers.odd;
		'''
		'''     // Insert a Paragraph into the Header.
		'''     Paragraph p0 = header.InsertParagraph();
		'''
		'''     // Appemd place holders for PageNumber and PageCount into the Header.
		'''     // Word will replace these with the correct value for each Page.
		'''     p0.Append("Page (");
		'''     p0.AppendPageNumber(PageNumberFormat.normal);
		'''     p0.Append(" of ");
		'''     p0.AppendPageCount(PageNumberFormat.normal);
		'''     p0.Append(")");
		''' 
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AppendPageNumber"/>
		''' <seealso cref="InsertPageNumber"/>
		''' <seealso cref="InsertPageCount"/>
		Public Sub AppendPageCount(ByVal pnf As PageNumberFormat)
			Dim fldSimple As New XElement(XName.Get("fldSimple", DocX.w.NamespaceName))

			If pnf = PageNumberFormat.normal Then
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES   \* MERGEFORMAT "))
			Else
				fldSimple.Add(New XAttribute(XName.Get("instr", DocX.w.NamespaceName), " NUMPAGES  \* ROMAN  \* MERGEFORMAT "))
			End If

			Dim content As XElement = XElement.Parse("<w:r w:rsidR='001D0226' xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & ControlChars.CrLf & "                   <w:rPr>" & ControlChars.CrLf & "                       <w:noProof /> " & ControlChars.CrLf & "                   </w:rPr>" & ControlChars.CrLf & "                   <w:t>1</w:t> " & ControlChars.CrLf & "               </w:r>")

			fldSimple.Add(content)
			Xml.Add(fldSimple)
		End Sub

		Public Property LineSpacingBefore() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim spacing As XElement = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))

				If spacing IsNot Nothing Then
					Dim line As XAttribute = spacing.Attribute(XName.Get("before", DocX.w.NamespaceName))
					If line IsNot Nothing Then
						Dim f As Single

						If Single.TryParse(line.Value, f) Then
							Return f / 20.0f
						End If
					End If
				End If

				Return 0.0f
			End Get

			Set(ByVal value As Single)
				SpacingBefore(value)
			End Set
		End Property

		Public Property LineSpacingAfter() As Single
			Get
				Dim pPr As XElement = GetOrCreate_pPr()
				Dim spacing As XElement = pPr.Element(XName.Get("spacing", DocX.w.NamespaceName))

				If spacing IsNot Nothing Then
					Dim line As XAttribute = spacing.Attribute(XName.Get("after", DocX.w.NamespaceName))
					If line IsNot Nothing Then
						Dim f As Single

						If Single.TryParse(line.Value, f) Then
							Return f / 20.0f
						End If
					End If
				End If

				Return 10.0f
			End Get


			Set(ByVal value As Single)
				SpacingAfter(value)
			End Set
		End Property
	End Class

	Public Class Run
		Inherits DocXElement
		' A lookup for the text elements in this paragraph
		Private textLookup As New Dictionary(Of Integer, Text)()

'INSTANT VB NOTE: The variable startIndex was renamed since Visual Basic does not allow class members with the same name:
		Private startIndex_Renamed As Integer
'INSTANT VB NOTE: The variable endIndex was renamed since Visual Basic does not allow class members with the same name:
		Private endIndex_Renamed As Integer
		Private text As String

		''' <summary>
		''' Gets the start index of this Text (text length before this text)
		''' </summary>
		Public ReadOnly Property StartIndex() As Integer
			Get
				Return startIndex_Renamed
			End Get
		End Property

		''' <summary>
		''' Gets the end index of this Text (text length before this text + this texts length)
		''' </summary>
		Public ReadOnly Property EndIndex() As Integer
			Get
				Return endIndex_Renamed
			End Get
		End Property

		''' <summary>
		''' The text value of this text element
		''' </summary>
		Friend Property Value() As String
			Set(ByVal value As String)
				text = value
			End Set
			Get
				Return text
			End Get
		End Property

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement, ByVal startIndex As Integer)
			MyBase.New(document, xml)
			Me.startIndex_Renamed = startIndex

			' Get the text elements in this run
			Dim texts As IEnumerable(Of XElement) = xml.Descendants()

			Dim start As Integer = startIndex

			' Loop through each text in this run
			For Each te As XElement In texts
				Select Case te.Name.LocalName
					Case "tab"
							textLookup.Add(start + 1, New Text(Me.Document, te, start))
							text &= vbTab
							start += 1
							Exit Select
					Case "br"
							textLookup.Add(start + 1, New Text(Me.Document, te, start))
							text &= vbLf
							start += 1
							Exit Select
					Case "t"
						GoTo CaseLabel1
					Case "delText"
					CaseLabel1:
							' Only add strings which are not empty
							If te.Value.Length > 0 Then
								textLookup.Add(start + te.Value.Length, New Text(Me.Document, te, start))
								text &= te.Value
								start += te.Value.Length
							End If
							Exit Select
					Case Else
				End Select
			Next te

			endIndex_Renamed = start
		End Sub

		Friend Shared Function SplitRun(ByVal r As Run, ByVal index As Integer, Optional ByVal type As EditType = EditType.ins) As XElement()
			index = index - r.StartIndex

			Dim t As Text = r.GetFirstTextEffectedByEdit(index, type)
			Dim splitText() As XElement = Text.SplitText(t, index)

			Dim splitLeft As New XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), t.Xml.ElementsBeforeSelf().Where(Function(n) n.Name.LocalName <> "rPr"), splitText(0))
			If Paragraph.GetElementTextLength(splitLeft) = 0 Then
				splitLeft = Nothing
			End If

			Dim splitRight As New XElement(r.Xml.Name, r.Xml.Attributes(), r.Xml.Element(XName.Get("rPr", DocX.w.NamespaceName)), splitText(1), t.Xml.ElementsAfterSelf().Where(Function(n) n.Name.LocalName <> "rPr"))
			If Paragraph.GetElementTextLength(splitRight) = 0 Then
				splitRight = Nothing
			End If

			Return (New XElement() { splitLeft, splitRight })
		End Function

		Friend Function GetFirstTextEffectedByEdit(ByVal index As Integer, Optional ByVal type As EditType = EditType.ins) As Text
			' Make sure we are looking within an acceptable index range.
			If index < 0 OrElse index > HelperFunctions.GetText(Xml).Length Then
				Throw New ArgumentOutOfRangeException()
			End If

			' Need some memory that can be updated by the recursive search for the XElement to Split.
			Dim count As Integer = 0
			Dim theOne As Text = Nothing

			GetFirstTextEffectedByEditRecursive(Xml, index, count, theOne, type)

			Return theOne
		End Function

		Friend Sub GetFirstTextEffectedByEditRecursive(ByVal Xml As XElement, ByVal index As Integer, ByRef count As Integer, ByRef theOne As Text, Optional ByVal type As EditType = EditType.ins)
			count += HelperFunctions.GetSize(Xml)
			If count > 0 AndAlso ((type = EditType.del AndAlso count > index) OrElse (type = EditType.ins AndAlso count >= index)) Then
				theOne = New Text(Document, Xml, count - HelperFunctions.GetSize(Xml))
				Return
			End If

			If Xml.HasElements Then
				For Each e As XElement In Xml.Elements()
					If theOne Is Nothing Then
						GetFirstTextEffectedByEditRecursive(e, index, count, theOne)
					End If
				Next e
			End If
		End Sub
	End Class

	Friend Class Text
		Inherits DocXElement
'INSTANT VB NOTE: The variable startIndex was renamed since Visual Basic does not allow class members with the same name:
		Private startIndex_Renamed As Integer
'INSTANT VB NOTE: The variable endIndex was renamed since Visual Basic does not allow class members with the same name:
		Private endIndex_Renamed As Integer
'INSTANT VB NOTE: The variable text was renamed since Visual Basic does not allow class members with the same name:
		Private text_Renamed As String

		''' <summary>
		''' Gets the start index of this Text (text length before this text)
		''' </summary>
		Public ReadOnly Property StartIndex() As Integer
			Get
				Return startIndex_Renamed
			End Get
		End Property

		''' <summary>
		''' Gets the end index of this Text (text length before this text + this texts length)
		''' </summary>
		Public ReadOnly Property EndIndex() As Integer
			Get
				Return endIndex_Renamed
			End Get
		End Property

		''' <summary>
		''' The text value of this text element
		''' </summary>
		Public ReadOnly Property Value() As String
			Get
				Return text_Renamed
			End Get
		End Property

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement, ByVal startIndex As Integer)
			MyBase.New(document, xml)
			Me.startIndex_Renamed = startIndex

			Select Case Me.Xml.Name.LocalName
				Case "t"
						GoTo CaseLabel1

				Case "delText"
				CaseLabel1:
						endIndex_Renamed = startIndex + xml.Value.Length
						text_Renamed = xml.Value
						Exit Select

				Case "br"
						text_Renamed = vbLf
						endIndex_Renamed = startIndex + 1
						Exit Select

				Case "tab"
						text_Renamed = vbTab
						endIndex_Renamed = startIndex + 1
						Exit Select
				Case Else
						Exit Select
			End Select
		End Sub

		Friend Shared Function SplitText(ByVal t As Text, ByVal index As Integer) As XElement()
			If index < t.startIndex_Renamed OrElse index > t.EndIndex Then
				Throw New ArgumentOutOfRangeException("index")
			End If

			Dim splitLeft As XElement = Nothing, splitRight As XElement = Nothing
			If t.Xml.Name.LocalName = "t" OrElse t.Xml.Name.LocalName = "delText" Then
				' The origional text element, now containing only the text before the index point.
				splitLeft = New XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(0, index - t.startIndex_Renamed))
				If splitLeft.Value.Length = 0 Then
					splitLeft = Nothing
				Else
					PreserveSpace(splitLeft)
				End If

				' The origional text element, now containing only the text after the index point.
				splitRight = New XElement(t.Xml.Name, t.Xml.Attributes(), t.Xml.Value.Substring(index - t.startIndex_Renamed, t.Xml.Value.Length - (index - t.startIndex_Renamed)))
				If splitRight.Value.Length = 0 Then
					splitRight = Nothing
				Else
					PreserveSpace(splitRight)
				End If

			Else
				If index = t.EndIndex Then
					splitLeft = t.Xml

				Else
					splitRight = t.Xml
				End If
			End If

			Return (New XElement() { splitLeft, splitRight })
		End Function

		''' <summary>
		''' If a text element or delText element, starts or ends with a space,
		''' it must have the attribute space, otherwise it must not have it.
		''' </summary>
		''' <param name="e">The (t or delText) element check</param>
		Public Shared Sub PreserveSpace(ByVal e As XElement)
			' PreserveSpace should only be used on (t or delText) elements
			If (Not e.Name.Equals(DocX.w + "t")) AndAlso (Not e.Name.Equals(DocX.w + "delText")) Then
				Throw New ArgumentException("SplitText can only split elements of type t or delText", "e")
			End If

			' Check if this w:t contains a space atribute
			Dim space As XAttribute = e.Attributes().Where(Function(a) a.Name.Equals(XNamespace.Xml + "space")).SingleOrDefault()

			' This w:t's text begins or ends with whitespace
			If e.Value.StartsWith(" ") OrElse e.Value.EndsWith(" ") Then
				' If this w:t contains no space attribute, add one.
				If space Is Nothing Then
					e.Add(New XAttribute(XNamespace.Xml + "space", "preserve"))
				End If

			' This w:t's text does not begin or end with a space
			Else
				' If this w:r contains a space attribute, remove it.
				If space IsNot Nothing Then
					space.Remove()
				End If
			End If
		End Sub
	End Class
End Namespace
