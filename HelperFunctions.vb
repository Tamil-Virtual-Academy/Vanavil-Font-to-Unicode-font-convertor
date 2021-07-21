Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.IO.Packaging
Imports System.Reflection
Imports System.Security.Principal
Imports System.Text
Imports System.Xml

Namespace Novacode
	Friend Module HelperFunctions
		Public Const DOCUMENT_DOCUMENTTYPE As String = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
		Public Const TEMPLATE_DOCUMENTTYPE As String = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"

			<System.Runtime.CompilerServices.Extension> _
			Public Function IsNullOrWhiteSpace(ByVal value As String) As Boolean
				If value Is Nothing Then
					Return True
				End If
				Return String.IsNullOrEmpty(value.Trim())
			End Function

		''' <summary>
		''' Checks whether 'toCheck' has all children that 'desired' has and values of 'val' attributes are the same
		''' </summary>
		''' <param name="desired"></param>
		''' <param name="toCheck"></param>
		''' <param name="fo">Matching options whether check if desired attributes are inder a, or a has exactly and only these attributes as b has.</param>
		''' <returns></returns>
		Friend Function ContainsEveryChildOf(ByVal desired As XElement, ByVal toCheck As XElement, ByVal fo As MatchFormattingOptions) As Boolean
			For Each e As XElement In desired.Elements()
				' If a formatting property has the same name and 'val' attribute's value, its considered to be equivalent.
				If Not toCheck.Elements(e.Name).Where(Function(bElement) bElement.GetAttribute(XName.Get("val", DocX.w.NamespaceName)) = e.GetAttribute(XName.Get("val", DocX.w.NamespaceName))).Any() Then
					Return False
				End If
			Next e

			' If the formatting has to be exact, no additionaly formatting must exist.
			If fo = MatchFormattingOptions.ExactMatch Then
				Return desired.Elements().Count() = toCheck.Elements().Count()
			End If

			Return True
		End Function
		Friend Sub CreateRelsPackagePart(ByVal Document As DocX, ByVal uri As Uri)
			Dim pp As PackagePart = Document.package.CreatePart(uri, "application/vnd.openxmlformats-package.relationships+xml", CompressionOption.Maximum)
			Using tw As TextWriter = New StreamWriter(pp.GetStream())
				Dim d As New XDocument(New XDeclaration("1.0", "UTF-8", "yes"), New XElement(XName.Get("Relationships", DocX.rel.NamespaceName)))
				Dim root = d.Root
				d.Save(tw)
			End Using
		End Sub

		Friend Function GetSize(ByVal Xml As XElement) As Integer
			Select Case Xml.Name.LocalName
				Case "tab"
					Return 1
				Case "br"
				CaseLabel2:
					Return 1
				Case "t"
					GoTo CaseLabel1
				Case "delText"
				CaseLabel1:
					Return Xml.Value.Length
				Case "tr"
					GoTo CaseLabel2
				Case "tc"
					GoTo CaseLabel2
				Case Else
					Return 0
			End Select
		End Function

		Friend Function GetText(ByVal e As XElement) As String
			Dim sb As New StringBuilder()
			GetTextRecursive(e, sb)
			Return sb.ToString()
		End Function

		Friend Sub GetTextRecursive(ByVal Xml As XElement, ByRef sb As StringBuilder)
			sb.Append(ToText(Xml))

			If Xml.HasElements Then
				For Each e As XElement In Xml.Elements()
					GetTextRecursive(e, sb)
				Next e
			End If
		End Sub

		Friend Function GetFormattedText(ByVal e As XElement) As List(Of FormattedText)
			Dim alist As New List(Of FormattedText)()
			GetFormattedTextRecursive(e, alist)
			Return alist
		End Function

		Friend Sub GetFormattedTextRecursive(ByVal Xml As XElement, ByRef alist As List(Of FormattedText))
			Dim ft As FormattedText = ToFormattedText(Xml)
			Dim last As FormattedText = Nothing

			If ft IsNot Nothing Then
				If alist.Count() > 0 Then
					last = alist.Last()
				End If

				If last IsNot Nothing AndAlso last.CompareTo(ft) = 0 Then
					' Update text of last entry.
					last.text &= ft.text
				Else
					If last IsNot Nothing Then
						ft.index = last.index + last.text.Length
					End If

					alist.Add(ft)
				End If
			End If

			If Xml.HasElements Then
				For Each e As XElement In Xml.Elements()
					GetFormattedTextRecursive(e, alist)
				Next e
			End If
		End Sub

		Friend Function ToFormattedText(ByVal e As XElement) As FormattedText
			' The text representation of e.
			Dim text As String = ToText(e)
			If text = String.Empty Then
				Return Nothing
			End If

			' e is a w:t element, it must exist inside a w:r element or a w:tabs, lets climb until we find it.
			Do While (Not e.Name.Equals(XName.Get("r", DocX.w.NamespaceName))) AndAlso Not e.Name.Equals(XName.Get("tabs", DocX.w.NamespaceName))
				e = e.Parent
			Loop

			' e is a w:r element, lets find the rPr element.
			Dim rPr As XElement = e.Element(XName.Get("rPr", DocX.w.NamespaceName))

			Dim ft As New FormattedText()
			ft.text = text
			ft.index = 0
			ft.formatting = Nothing

			' Return text with formatting.
			If rPr IsNot Nothing Then
				ft.formatting = Formatting.Parse(rPr)
			End If

			Return ft
		End Function

		Friend Function ToText(ByVal e As XElement) As String
			Select Case e.Name.LocalName
				Case "tab"
				CaseLabel3:
					Return vbTab
				Case "br"
				CaseLabel2:
					Return vbLf
				Case "t"
					GoTo CaseLabel1
				Case "delText"
				CaseLabel1:
						If e.Parent IsNot Nothing AndAlso e.Parent.Name.LocalName = "r" Then
							Dim run As XElement = e.Parent
							Dim rPr = run.Elements().FirstOrDefault(Function(a) a.Name.LocalName = "rPr")
							If rPr IsNot Nothing Then
								Dim caps = rPr.Elements().FirstOrDefault(Function(a) a.Name.LocalName = "caps")

								If caps IsNot Nothing Then
									Return e.Value.ToUpper()
								End If
							End If
						End If

						Return e.Value
				Case "tr"
					GoTo CaseLabel2
				Case "tc"
					GoTo CaseLabel3
				Case Else
					Return ""
			End Select
		End Function

		Friend Function CloneElement(ByVal element As XElement) As XElement
			Return New XElement (element.Name, element.Attributes(), element.Nodes().Select (Function(n)
				Dim e As XElement = TryCast(n, XElement)
				If e IsNot Nothing Then
					Return CloneElement(e)
				End If
					Return n
			End Function))
		End Function

		Friend Function CreateOrGetSettingsPart(ByVal package As Package) As PackagePart
			Dim settingsPart As PackagePart

			Dim settingsUri As New Uri("/word/settings.xml", UriKind.Relative)
			If Not package.PartExists(settingsUri) Then
				settingsPart = package.CreatePart(settingsUri, "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", CompressionOption.Maximum)

				Dim mainDocumentPart As PackagePart = package.GetParts().Single(Function(p) p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) OrElse p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase))

				mainDocumentPart.CreateRelationship(settingsUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings")

				Dim settings As XDocument = XDocument.Parse("<?xml version='1.0' encoding='utf-8' standalone='yes'?>" & ControlChars.CrLf & "                <w:settings xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w10='urn:schemas-microsoft-com:office:word' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>" & ControlChars.CrLf & "                  <w:zoom w:percent='100' />" & ControlChars.CrLf & "                  <w:defaultTabStop w:val='720' />" & ControlChars.CrLf & "                  <w:characterSpacingControl w:val='doNotCompress' />" & ControlChars.CrLf & "                  <w:compat />" & ControlChars.CrLf & "                  <w:rsids>" & ControlChars.CrLf & "                    <w:rsidRoot w:val='00217F62' />" & ControlChars.CrLf & "                    <w:rsid w:val='001915A3' />" & ControlChars.CrLf & "                    <w:rsid w:val='00217F62' />" & ControlChars.CrLf & "                    <w:rsid w:val='00A906D8' />" & ControlChars.CrLf & "                    <w:rsid w:val='00AB5A74' />" & ControlChars.CrLf & "                    <w:rsid w:val='00F071AE' />" & ControlChars.CrLf & "                  </w:rsids>" & ControlChars.CrLf & "                  <m:mathPr>" & ControlChars.CrLf & "                    <m:mathFont m:val='Cambria Math' />" & ControlChars.CrLf & "                    <m:brkBin m:val='before' />" & ControlChars.CrLf & "                    <m:brkBinSub m:val='--' />" & ControlChars.CrLf & "                    <m:smallFrac m:val='off' />" & ControlChars.CrLf & "                    <m:dispDef />" & ControlChars.CrLf & "                    <m:lMargin m:val='0' />" & ControlChars.CrLf & "                    <m:rMargin m:val='0' />" & ControlChars.CrLf & "                    <m:defJc m:val='centerGroup' />" & ControlChars.CrLf & "                    <m:wrapIndent m:val='1440' />" & ControlChars.CrLf & "                    <m:intLim m:val='subSup' />" & ControlChars.CrLf & "                    <m:naryLim m:val='undOvr' />" & ControlChars.CrLf & "                  </m:mathPr>" & ControlChars.CrLf & "                  <w:themeFontLang w:val='en-IE' w:bidi='ar-SA' />" & ControlChars.CrLf & "                  <w:clrSchemeMapping w:bg1='light1' w:t1='dark1' w:bg2='light2' w:t2='dark2' w:accent1='accent1' w:accent2='accent2' w:accent3='accent3' w:accent4='accent4' w:accent5='accent5' w:accent6='accent6' w:hyperlink='hyperlink' w:followedHyperlink='followedHyperlink' />" & ControlChars.CrLf & "                  <w:shapeDefaults>" & ControlChars.CrLf & "                    <o:shapedefaults v:ext='edit' spidmax='2050' />" & ControlChars.CrLf & "                    <o:shapelayout v:ext='edit'>" & ControlChars.CrLf & "                      <o:idmap v:ext='edit' data='1' />" & ControlChars.CrLf & "                    </o:shapelayout>" & ControlChars.CrLf & "                  </w:shapeDefaults>" & ControlChars.CrLf & "                  <w:decimalSymbol w:val='.' />" & ControlChars.CrLf & "                  <w:listSeparator w:val=',' />" & ControlChars.CrLf & "                </w:settings>")

				Dim themeFontLang As XElement = settings.Root.Element(XName.Get("themeFontLang", DocX.w.NamespaceName))
				themeFontLang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture)

				' Save the settings document.
				Using tw As TextWriter = New StreamWriter(settingsPart.GetStream())
					settings.Save(tw)
				End Using
			Else
				settingsPart = package.GetPart(settingsUri)
			End If
			Return settingsPart
		End Function

		Friend Sub CreateCustomPropertiesPart(ByVal document As DocX)
			Dim customPropertiesPart As PackagePart = document.package.CreatePart(New Uri("/docProps/custom.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.custom-properties+xml", CompressionOption.Maximum)

			Dim customPropDoc As New XDocument(New XDeclaration("1.0", "UTF-8", "yes"), New XElement (XName.Get("Properties", DocX.customPropertiesSchema.NamespaceName), New XAttribute(XNamespace.Xmlns + "vt", DocX.customVTypesSchema)))

			Using tw As TextWriter = New StreamWriter(customPropertiesPart.GetStream(FileMode.Create, FileAccess.Write))
				customPropDoc.Save(tw, SaveOptions.None)
			End Using

			document.package.CreateRelationship(customPropertiesPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
		End Sub

		Friend Function DecompressXMLResource(ByVal manifest_resource_name As String) As XDocument
			' XDocument to load the compressed Xml resource into.
			Dim document As XDocument

			' Get a reference to the executing assembly.
			Dim [assembly] As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()

			' Open a Stream to the embedded resource.
			Dim stream As Stream = [assembly].GetManifestResourceStream(manifest_resource_name)

			' Decompress the embedded resource.
			Using zip As New GZipStream(stream, CompressionMode.Decompress)
				' Load this decompressed embedded resource into an XDocument using a TextReader.
				Using sr As TextReader = New StreamReader(zip)
					document = XDocument.Load(sr)
				End Using
			End Using

			' Return the decompressed Xml as an XDocument.
			Return document
		End Function


		''' <summary>
		''' If this document does not contain a /word/numbering.xml add the default one generated by Microsoft Word 
		''' when the default bullet, numbered and multilevel lists are added to a blank document
		''' </summary>
		''' <param name="package"></param>
		''' <returns></returns>
		Friend Function AddDefaultNumberingXml(ByVal package As Package) As XDocument
			Dim numberingDoc As XDocument
			' Create the main document part for this package
			Dim wordNumbering As PackagePart = package.CreatePart(New Uri("/word/numbering.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", CompressionOption.Maximum)

			numberingDoc = DecompressXMLResource("Novacode.Resources.numbering.xml.gz")

			' Save /word/numbering.xml
			Using tw As TextWriter = New StreamWriter(wordNumbering.GetStream(FileMode.Create, FileAccess.Write))
				numberingDoc.Save(tw, SaveOptions.None)
			End Using

			Dim mainDocumentPart As PackagePart = package.GetParts().Single(Function(p) p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) OrElse p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase))

			mainDocumentPart.CreateRelationship(wordNumbering.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
			Return numberingDoc
		End Function



		''' <summary>
		''' If this document does not contain a /word/styles.xml add the default one generated by Microsoft Word.
		''' </summary>
		''' <param name="package"></param>
		''' <returns></returns>
		Friend Function AddDefaultStylesXml(ByVal package As Package) As XDocument
			Dim stylesDoc As XDocument
			' Create the main document part for this package
			Dim word_styles As PackagePart = package.CreatePart(New Uri("/word/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", CompressionOption.Maximum)

			stylesDoc = HelperFunctions.DecompressXMLResource("Novacode.Resources.default_styles.xml.gz")
			Dim lang As XElement = stylesDoc.Root.Element(XName.Get("docDefaults", DocX.w.NamespaceName)).Element(XName.Get("rPrDefault", DocX.w.NamespaceName)).Element(XName.Get("rPr", DocX.w.NamespaceName)).Element(XName.Get("lang", DocX.w.NamespaceName))
			lang.SetAttributeValue(XName.Get("val", DocX.w.NamespaceName), CultureInfo.CurrentCulture)

			' Save /word/styles.xml
			Using tw As TextWriter = New StreamWriter(word_styles.GetStream(FileMode.Create, FileAccess.Write))
				stylesDoc.Save(tw, SaveOptions.None)
			End Using

			Dim mainDocumentPart As PackagePart = package.GetParts().Where (Function(p) p.ContentType.Equals(DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) OrElse p.ContentType.Equals(TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)).Single()

			mainDocumentPart.CreateRelationship(word_styles.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
			Return stylesDoc
		End Function

		Friend Function CreateEdit(ByVal t As EditType, ByVal edit_time As Date, ByVal content As Object) As XElement
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

		Friend Function CreateTable(ByVal rowCount As Integer, ByVal columnCount As Integer) As XElement
			Dim columnWidths(columnCount - 1) As Integer
			For i As Integer = 0 To columnCount - 1
				columnWidths(i) = 2310
			Next i
			Return CreateTable(rowCount, columnWidths)
		End Function

		Friend Function CreateTable(ByVal rowCount As Integer, ByVal columnWidths() As Integer) As XElement
			Dim newTable As New XElement(XName.Get("tbl", DocX.w.NamespaceName), New XElement (XName.Get("tblPr", DocX.w.NamespaceName), New XElement(XName.Get("tblStyle", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), "TableGrid")), New XElement(XName.Get("tblW", DocX.w.NamespaceName), New XAttribute(XName.Get("w", DocX.w.NamespaceName), "5000"), New XAttribute(XName.Get("type", DocX.w.NamespaceName), "auto")), New XElement(XName.Get("tblLook", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), "04A0"))))

'            XElement tableGrid = new XElement(XName.Get("tblGrid", DocX.w.NamespaceName));
'            for (int i = 0; i < columnWidths.Length; i++)
'                tableGrid.Add(new XElement(XName.Get("gridCol", DocX.w.NamespaceName), new XAttribute(XName.Get("w", DocX.w.NamespaceName), XmlConvert.ToString(columnWidths[i]))));
'
'            newTable.Add(tableGrid);

			For i As Integer = 0 To rowCount - 1
				Dim row As New XElement(XName.Get("tr", DocX.w.NamespaceName))

				For j As Integer = 0 To columnWidths.Length - 1
					Dim cell As XElement = CreateTableCell()
					row.Add(cell)
				Next j

				newTable.Add(row)
			Next i
			Return newTable
		End Function

		''' <summary>
		''' Create and return a cell of a table        
		''' </summary>        
		Friend Function CreateTableCell(Optional ByVal w As Double = 2310) As XElement
			Return New XElement (XName.Get("tc", DocX.w.NamespaceName), New XElement(XName.Get("tcPr", DocX.w.NamespaceName), New XElement(XName.Get("tcW", DocX.w.NamespaceName), New XAttribute(XName.Get("w", DocX.w.NamespaceName), w), New XAttribute(XName.Get("type", DocX.w.NamespaceName), "dxa"))), New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName))))
		End Function

		Friend Function CreateItemInList(ByVal list As List, ByVal listText As String, Optional ByVal level As Integer = 0, Optional ByVal listType As ListItemType = ListItemType.Numbered, Optional ByVal startNumber? As Integer = Nothing, Optional ByVal trackChanges As Boolean = False, Optional ByVal continueNumbering As Boolean = False) As List
			If list.NumId = 0 Then
				list.CreateNewNumberingNumId(level, listType, startNumber, continueNumbering)
			End If

			If listText IsNot Nothing Then 'I see no reason why you shouldn't be able to insert an empty element. It simplifies tasks such as populating an item from html.
				Dim newParagraphSection = New XElement (XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName), New XElement(XName.Get("numPr", DocX.w.NamespaceName), New XElement(XName.Get("ilvl", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", level)), New XElement(XName.Get("numId", DocX.w.NamespaceName), New XAttribute(DocX.w + "val", list.NumId)))), New XElement(XName.Get("r", DocX.w.NamespaceName), New XElement(XName.Get("t", DocX.w.NamespaceName), listText)))

				If trackChanges Then
					newParagraphSection = CreateEdit(EditType.ins, Date.Now, newParagraphSection)
				End If

				If startNumber Is Nothing Then
					list.AddItem(New Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph))
				Else
					list.AddItemWithStartValue(New Paragraph(list.Document, newParagraphSection, 0, ContainerType.Paragraph), CInt(Fix(startNumber)))
				End If
			End If

			Return list
		End Function

		Friend Sub RenumberIDs(ByVal document As DocX)
			Dim trackerIDs As IEnumerable(Of XAttribute) = (
			    From d In document.mainDoc.Descendants()
			    Where d.Name.LocalName = "ins" OrElse d.Name.LocalName = "del"
			    Select d.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")))

			For i As Integer = 0 To trackerIDs.Count() - 1
				trackerIDs.ElementAt(i).Value = i.ToString()
			Next i
		End Sub

		Friend Function GetFirstParagraphEffectedByInsert(ByVal document As DocX, ByVal index As Integer) As Paragraph
			' This document contains no Paragraphs and insertion is at index 0
			If document.paragraphLookup.Keys.Count() = 0 AndAlso index = 0 Then
				Return Nothing
			End If

			For Each paragraphEndIndex As Integer In document.paragraphLookup.Keys
				If paragraphEndIndex >= index Then
					Return document.paragraphLookup(paragraphEndIndex)
				End If
			Next paragraphEndIndex

			Throw New ArgumentOutOfRangeException()
		End Function

		Friend Function FormatInput(ByVal text As String, ByVal rPr As XElement) As List(Of XElement)
			Dim newRuns As New List(Of XElement)()
			Dim tabRun As New XElement(DocX.w + "tab")
			Dim breakRun As New XElement(DocX.w + "br")

			Dim sb As New StringBuilder()

			If String.IsNullOrEmpty(text) Then
				Return newRuns 'I dont wanna get an exception if text == null, so just return empy list
			End If

			For Each c As Char In text
				Select Case c
					Case ControlChars.Tab
						If sb.Length > 0 Then
							Dim t As New XElement(DocX.w + "t", sb.ToString())
							Novacode.Text.PreserveSpace(t)
							newRuns.Add(New XElement(DocX.w + "r", rPr, t))
							sb = New StringBuilder()
						End If
						newRuns.Add(New XElement(DocX.w + "r", rPr, tabRun))
					Case ControlChars.Cr, ControlChars.Lf
						If sb.Length > 0 Then
							Dim t As New XElement(DocX.w + "t", sb.ToString())
							Novacode.Text.PreserveSpace(t)
							newRuns.Add(New XElement(DocX.w + "r", rPr, t))
							sb = New StringBuilder()
						End If
						newRuns.Add(New XElement(DocX.w + "r", rPr, breakRun))

					Case Else
						sb.Append(c)
				End Select
			Next c

			If sb.Length > 0 Then
				Dim t As New XElement(DocX.w + "t", sb.ToString())
				Novacode.Text.PreserveSpace(t)
				newRuns.Add(New XElement(DocX.w + "r", rPr, t))
			End If

			Return newRuns
		End Function

		Friend Function SplitParagraph(ByVal p As Paragraph, ByVal index As Integer) As XElement()
			' In this case edit dosent really matter, you have a choice.
			Dim r As Run = p.GetFirstRunEffectedByEdit(index, EditType.ins)

			Dim split() As XElement
			Dim before, after As XElement

			If r.Xml.Parent.Name.LocalName = "ins" Then
				split = p.SplitEdit(r.Xml.Parent, index, EditType.ins)
				before = New XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split(0))
				after = New XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split(1))
			ElseIf r.Xml.Parent.Name.LocalName = "del" Then
				split = p.SplitEdit(r.Xml.Parent, index, EditType.del)

				before = New XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsBeforeSelf(), split(0))
				after = New XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.Parent.ElementsAfterSelf(), split(1))
			Else
				split = Run.SplitRun(r, index)

				before = New XElement(p.Xml.Name, p.Xml.Attributes(), r.Xml.ElementsBeforeSelf(), split(0))
				after = New XElement(p.Xml.Name, p.Xml.Attributes(), split(1), r.Xml.ElementsAfterSelf())
			End If

			If before.Elements().Count() = 0 Then
				before = Nothing
			End If

			If after.Elements().Count() = 0 Then
				after = Nothing
			End If

			Return New XElement() { before, after }
		End Function

		''' <!-- 
		''' Bug found and fixed by trnilse. To see the change, 
		''' please compare this release to the previous release using TFS compare.
		''' -->
		Friend Function IsSameFile(ByVal streamOne As Stream, ByVal streamTwo As Stream) As Boolean
			Dim file1byte, file2byte As Integer

			If streamOne.Length <> streamTwo.Length Then
				' Return false to indicate files are different
				Return False
			End If

			' Read and compare a byte from each file until either a
			' non-matching set of bytes is found or until the end of
			' file1 is reached.
			Do
				' Read one byte from each file.
				file1byte = streamOne.ReadByte()
				file2byte = streamTwo.ReadByte()
			Loop While (file1byte = file2byte) AndAlso (file1byte <> -1)

			' Return the success of the comparison. "file1byte" is 
			' equal to "file2byte" at this point only if the files are 
			' the same.

			streamOne.Position = 0
			streamTwo.Position = 0

			Return ((file1byte - file2byte) = 0)
		End Function

	  Friend Function GetUnderlineStyle(ByVal underlineStyle As String) As UnderlineStyle
		Select Case underlineStyle
		  Case "single"
			Return UnderlineStyle.singleLine
		  Case "double"
			Return UnderlineStyle.doubleLine
		  Case "thick"
			Return UnderlineStyle.thick
		  Case "dotted"
			Return UnderlineStyle.dotted
		  Case "dottedHeavy"
			Return UnderlineStyle.dottedHeavy
		  Case "dash"
			Return UnderlineStyle.dash
		  Case "dashedHeavy"
			Return UnderlineStyle.dashedHeavy
		  Case "dashLong"
			Return UnderlineStyle.dashLong
		  Case "dashLongHeavy"
			Return UnderlineStyle.dashLongHeavy
		  Case "dotDash"
			Return UnderlineStyle.dotDash
		  Case "dashDotHeavy"
			Return UnderlineStyle.dashDotHeavy
		  Case "dotDotDash"
			Return UnderlineStyle.dotDotDash
		  Case "dashDotDotHeavy"
			Return UnderlineStyle.dashDotDotHeavy
		  Case "wave"
			Return UnderlineStyle.wave
		  Case "wavyHeavy"
			Return UnderlineStyle.wavyHeavy
		  Case "wavyDouble"
			Return UnderlineStyle.wavyDouble
		  Case "words"
			Return UnderlineStyle.words
		  Case Else
			Return UnderlineStyle.none
		End Select
	  End Function



	End Module
End Namespace
