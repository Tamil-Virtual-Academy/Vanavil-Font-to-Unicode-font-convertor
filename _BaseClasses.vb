Imports System.IO.Packaging

Namespace Novacode
	''' <summary>
	''' All DocX types are derived from DocXElement. 
	''' This class contains properties which every element of a DocX must contain.
	''' </summary>
	Public MustInherit Class DocXElement
		Friend mainPart As PackagePart
		Public Property PackagePart() As PackagePart
			Get
				Return mainPart
			End Get
			Set(ByVal value As PackagePart)
				mainPart = value
			End Set
		End Property

		''' <summary>
		''' This is the actual Xml that gives this element substance. 
		''' For example, a Paragraphs Xml might look something like the following
		''' <p>
		'''     <r>
		'''         <t>Hello World!</t>
		'''     </r>
		''' </p>
		''' </summary>

		Public Property Xml() As XElement
		''' <summary>
		''' This is a reference to the DocX object that this element belongs to.
		''' Every DocX element is connected to a document.
		''' </summary>
		Friend Property Document() As DocX
		''' <summary>
		''' Store both the document and xml so that they can be accessed by derived types.
		''' </summary>
		''' <param name="document">The document that this element belongs to.</param>
		''' <param name="xml">The Xml that gives this element substance</param>
		Public Sub New(ByVal document As DocX, ByVal xml As XElement)
			Me.Document = document
			Me.Xml = xml
		End Sub
	End Class

	''' <summary>
	''' This class provides functions for inserting new DocXElements before or after the current DocXElement.
	''' Only certain DocXElements can support these functions without creating invalid documents, at the moment these are Paragraphs and Table.
	''' </summary>
	Public MustInherit Class InsertBeforeOrAfter
		Inherits DocXElement
		Public Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)
		End Sub

		Public Overridable Sub InsertPageBreakBeforeSelf()
			Dim p As New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement (XName.Get("r", DocX.w.NamespaceName), New XElement (XName.Get("br", DocX.w.NamespaceName), New XAttribute(XName.Get("type", DocX.w.NamespaceName), "page"))))

			Xml.AddBeforeSelf(p)
		End Sub

		Public Overridable Sub InsertPageBreakAfterSelf()
			Dim p As New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement (XName.Get("r", DocX.w.NamespaceName), New XElement (XName.Get("br", DocX.w.NamespaceName), New XAttribute(XName.Get("type", DocX.w.NamespaceName), "page"))))

			Xml.AddAfterSelf(p)
		End Sub

		Public Overridable Function InsertParagraphBeforeSelf(ByVal p As Paragraph) As Paragraph
			Xml.AddBeforeSelf(p.Xml)
			Dim newlyInserted As XElement = Xml.ElementsBeforeSelf().First()

			p.Xml = newlyInserted

			Return p
		End Function

		Public Overridable Function InsertParagraphAfterSelf(ByVal p As Paragraph) As Paragraph
			Xml.AddAfterSelf(p.Xml)
			Dim newlyInserted As XElement = Xml.ElementsAfterSelf().First()

			'Dmitchern
			If TryCast(Me, Paragraph) IsNot Nothing Then
				Return New Paragraph(Document, newlyInserted, (TryCast(Me, Paragraph)).endIndex)
			Else
				p.Xml = newlyInserted 'IMPORTANT: I think we have return new paragraph in any case, but I dont know what to put as startIndex parameter into Paragraph constructor
				Return p
			End If
		End Function

		Public Overridable Function InsertParagraphBeforeSelf(ByVal text As String) As Paragraph
			Return InsertParagraphBeforeSelf(text, False, New Formatting())
		End Function

		Public Overridable Function InsertParagraphAfterSelf(ByVal text As String) As Paragraph
			Return InsertParagraphAfterSelf(text, False, New Formatting())
		End Function

		Public Overridable Function InsertParagraphBeforeSelf(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Return InsertParagraphBeforeSelf(text, trackChanges, New Formatting())
		End Function

		Public Overridable Function InsertParagraphAfterSelf(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Return InsertParagraphAfterSelf(text, trackChanges, New Formatting())
		End Function

		Public Overridable Function InsertParagraphBeforeSelf(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim newParagraph As New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml))

			If trackChanges Then
				newParagraph = Paragraph.CreateEdit(EditType.ins, Date.Now, newParagraph)
			End If

			Xml.AddBeforeSelf(newParagraph)
			Dim newlyInserted As XElement = Xml.ElementsBeforeSelf().Last()

			Dim p As New Paragraph(Document, newlyInserted, -1)

			Return p
		End Function

		Public Overridable Function InsertParagraphAfterSelf(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim newParagraph As New XElement(XName.Get("p", DocX.w.NamespaceName), New XElement(XName.Get("pPr", DocX.w.NamespaceName)), HelperFunctions.FormatInput(text, formatting.Xml))

			If trackChanges Then
				newParagraph = Paragraph.CreateEdit(EditType.ins, Date.Now, newParagraph)
			End If

			Xml.AddAfterSelf(newParagraph)
			Dim newlyInserted As XElement = Xml.ElementsAfterSelf().First()

			Dim p As New Paragraph(Document, newlyInserted, -1)

			Return p
		End Function

		Public Overridable Function InsertTableAfterSelf(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			Dim newTable As XElement = HelperFunctions.CreateTable(rowCount, columnCount)
			Xml.AddAfterSelf(newTable)
			Dim newlyInserted As XElement = Xml.ElementsAfterSelf().First()

			Return New Table(Document, newlyInserted) With {.mainPart = mainPart}
		End Function

		Public Overridable Function InsertTableAfterSelf(ByVal t As Table) As Table
			Xml.AddAfterSelf(t.Xml)
			Dim newlyInserted As XElement = Xml.ElementsAfterSelf().First()
			'Dmitchern
			Return New Table(Document, newlyInserted) With {.mainPart = mainPart} 'return new table, dont affect parameter table

			't.Xml = newlyInserted;
			'return t;
		End Function

		Public Overridable Function InsertTableBeforeSelf(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			Dim newTable As XElement = HelperFunctions.CreateTable(rowCount, columnCount)
			Xml.AddBeforeSelf(newTable)
			Dim newlyInserted As XElement = Xml.ElementsBeforeSelf().Last()

			Return New Table(Document, newlyInserted) With {.mainPart = mainPart}
		End Function

		Public Overridable Function InsertTableBeforeSelf(ByVal t As Table) As Table
			Xml.AddBeforeSelf(t.Xml)
			Dim newlyInserted As XElement = Xml.ElementsBeforeSelf().Last()

			'Dmitchern
			Return New Table(Document, newlyInserted) With {.mainPart=mainPart} 'return new table, dont affect parameter table

			't.Xml = newlyInserted;
			'return t;
		End Function
	End Class

	Public NotInheritable Class XmlTemplateBases
		#Region "TocXml"
		Public Const TocXmlBase As String = "<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "                  <w:sdtPr>" & ControlChars.CrLf & "                    <w:docPartObj>" & ControlChars.CrLf & "                      <w:docPartGallery w:val='Table of Contents'/>" & ControlChars.CrLf & "                      <w:docPartUnique/>" & ControlChars.CrLf & "                    </w:docPartObj>\" & ControlChars.CrLf & "                  </w:sdtPr>" & ControlChars.CrLf & "                  <w:sdtEndPr>" & ControlChars.CrLf & "                    <w:rPr>" & ControlChars.CrLf & "                      <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>" & ControlChars.CrLf & "                      <w:color w:val='auto'/>" & ControlChars.CrLf & "                      <w:sz w:val='22'/>" & ControlChars.CrLf & "                      <w:szCs w:val='22'/>" & ControlChars.CrLf & "                      <w:lang w:eastAsia='en-US'/>" & ControlChars.CrLf & "                    </w:rPr>" & ControlChars.CrLf & "                  </w:sdtEndPr>" & ControlChars.CrLf & "                  <w:sdtContent>" & ControlChars.CrLf & "                    <w:p>" & ControlChars.CrLf & "                      <w:pPr>" & ControlChars.CrLf & "                        <w:pStyle w:val='{0}'/>" & ControlChars.CrLf & "                      </w:pPr>" & ControlChars.CrLf & "                      <w:r>" & ControlChars.CrLf & "                        <w:t>{1}</w:t>" & ControlChars.CrLf & "                      </w:r>" & ControlChars.CrLf & "                    </w:p>" & ControlChars.CrLf & "                    <w:p>" & ControlChars.CrLf & "                      <w:pPr>" & ControlChars.CrLf & "                        <w:pStyle w:val='TOC1'/>" & ControlChars.CrLf & "                        <w:tabs>" & ControlChars.CrLf & "                          <w:tab w:val='right' w:leader='dot' w:pos='{2}'/>" & ControlChars.CrLf & "                        </w:tabs>" & ControlChars.CrLf & "                        <w:rPr>" & ControlChars.CrLf & "                          <w:noProof/>" & ControlChars.CrLf & "                        </w:rPr>" & ControlChars.CrLf & "                      </w:pPr>" & ControlChars.CrLf & "                      <w:r>" & ControlChars.CrLf & "                        <w:fldChar w:fldCharType='begin' w:dirty='true'/>" & ControlChars.CrLf & "                      </w:r>" & ControlChars.CrLf & "                      <w:r>" & ControlChars.CrLf & "                        <w:instrText xml:space='preserve'> {3} </w:instrText>" & ControlChars.CrLf & "                      </w:r>" & ControlChars.CrLf & "                      <w:r>" & ControlChars.CrLf & "                        <w:fldChar w:fldCharType='separate'/>" & ControlChars.CrLf & "                      </w:r>" & ControlChars.CrLf & "                    </w:p>" & ControlChars.CrLf & "                    <w:p>" & ControlChars.CrLf & "                      <w:r>" & ControlChars.CrLf & "                        <w:rPr>" & ControlChars.CrLf & "                          <w:b/>" & ControlChars.CrLf & "                          <w:bCs/>" & ControlChars.CrLf & "                          <w:noProof/>" & ControlChars.CrLf & "                        </w:rPr>" & ControlChars.CrLf & "                        <w:fldChar w:fldCharType='end'/>" & ControlChars.CrLf & "                      </w:r>" & ControlChars.CrLf & "                    </w:p>" & ControlChars.CrLf & "                  </w:sdtContent>" & ControlChars.CrLf & "                </w:sdt>" & ControlChars.CrLf & "            "
		Public Const TocHeadingStyleBase As String = "<w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "            <w:name w:val='TOC Heading'/>" & ControlChars.CrLf & "            <w:basedOn w:val='Heading1'/>" & ControlChars.CrLf & "            <w:next w:val='Normal'/>" & ControlChars.CrLf & "            <w:uiPriority w:val='39'/>" & ControlChars.CrLf & "            <w:semiHidden/>" & ControlChars.CrLf & "            <w:unhideWhenUsed/>" & ControlChars.CrLf & "            <w:qFormat/>" & ControlChars.CrLf & "            <w:rsid w:val='00E67AA6'/>" & ControlChars.CrLf & "            <w:pPr>" & ControlChars.CrLf & "              <w:outlineLvl w:val='9'/>" & ControlChars.CrLf & "            </w:pPr>" & ControlChars.CrLf & "            <w:rPr>" & ControlChars.CrLf & "              <w:lang w:eastAsia='nb-NO'/>" & ControlChars.CrLf & "            </w:rPr>" & ControlChars.CrLf & "          </w:style>" & ControlChars.CrLf & "        "
		Public Const TocElementStyleBase As String = "  <w:style w:type='paragraph' w:styleId='{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "            <w:name w:val='{1}' />" & ControlChars.CrLf & "            <w:basedOn w:val='Normal' />" & ControlChars.CrLf & "            <w:next w:val='Normal' />" & ControlChars.CrLf & "            <w:autoRedefine />" & ControlChars.CrLf & "            <w:uiPriority w:val='39' />" & ControlChars.CrLf & "            <w:unhideWhenUsed />" & ControlChars.CrLf & "            <w:pPr>" & ControlChars.CrLf & "              <w:spacing w:after='100' />" & ControlChars.CrLf & "              <w:ind w:left='440' />" & ControlChars.CrLf & "            </w:pPr>" & ControlChars.CrLf & "          </w:style>" & ControlChars.CrLf & "        "
		Public Const TocHyperLinkStyleBase As String = "  <w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "            <w:name w:val='Hyperlink' />" & ControlChars.CrLf & "            <w:basedOn w:val='Normal' />" & ControlChars.CrLf & "            <w:uiPriority w:val='99' />" & ControlChars.CrLf & "            <w:unhideWhenUsed />" & ControlChars.CrLf & "            <w:rPr>" & ControlChars.CrLf & "              <w:color w:val='0000FF' w:themeColor='hyperlink' />" & ControlChars.CrLf & "              <w:u w:val='single' />" & ControlChars.CrLf & "            </w:rPr>" & ControlChars.CrLf & "          </w:style>" & ControlChars.CrLf & "        "
		#End Region
	End Class
End Namespace
