Imports System.IO.Packaging
Imports System.Collections.ObjectModel

Namespace Novacode
	Public Class Header
		Inherits Container
		Implements IParagraphContainer
		Public Property PageNumbers() As Boolean
			Get
				Return False
			End Get

			Set(ByVal value As Boolean)
				Dim e As XElement = XElement.Parse("<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "                    <w:sdtPr>" & ControlChars.CrLf & "                      <w:id w:val='157571950' />" & ControlChars.CrLf & "                      <w:docPartObj>" & ControlChars.CrLf & "                        <w:docPartGallery w:val='Page Numbers (Top of Page)' />" & ControlChars.CrLf & "                        <w:docPartUnique />" & ControlChars.CrLf & "                      </w:docPartObj>" & ControlChars.CrLf & "                    </w:sdtPr>" & ControlChars.CrLf & "                    <w:sdtContent>" & ControlChars.CrLf & "                      <w:p w:rsidR='008D2BFB' w:rsidRDefault='008D2BFB'>" & ControlChars.CrLf & "                        <w:pPr>" & ControlChars.CrLf & "                          <w:pStyle w:val='Header' />" & ControlChars.CrLf & "                          <w:jc w:val='center' />" & ControlChars.CrLf & "                        </w:pPr>" & ControlChars.CrLf & "                        <w:fldSimple w:instr=' PAGE \* MERGEFORMAT'>" & ControlChars.CrLf & "                          <w:r>" & ControlChars.CrLf & "                            <w:rPr>" & ControlChars.CrLf & "                              <w:noProof />" & ControlChars.CrLf & "                            </w:rPr>" & ControlChars.CrLf & "                            <w:t>1</w:t>" & ControlChars.CrLf & "                          </w:r>" & ControlChars.CrLf & "                        </w:fldSimple>" & ControlChars.CrLf & "                      </w:p>" & ControlChars.CrLf & "                    </w:sdtContent>" & ControlChars.CrLf & "                  </w:sdt>")

			   Xml.AddFirst(e)

			   PageNumberParagraph = New Paragraph(Document, e.Descendants(XName.Get("p", DocX.w.NamespaceName)).SingleOrDefault(), 0)
			End Set
		End Property

		Public PageNumberParagraph As Paragraph

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement, ByVal mainPart As PackagePart)
			MyBase.New(document, xml)
			Me.mainPart = mainPart
		End Sub

		Public Overrides Function InsertParagraph() As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph()
			p.PackagePart = mainPart
			Return p
		End Function

		Public Overrides Function InsertParagraph(ByVal index As Integer, ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph(index, text, trackChanges)
			p.PackagePart = mainPart
			Return p
		End Function

		Public Overrides Function InsertParagraph(ByVal p As Paragraph) As Paragraph
			p.PackagePart = mainPart
			Return MyBase.InsertParagraph(p)
		End Function

		Public Overrides Function InsertParagraph(ByVal index As Integer, ByVal p As Paragraph) As Paragraph
			p.PackagePart = mainPart
			Return MyBase.InsertParagraph(index, p)
		End Function

		Public Overrides Function InsertParagraph(ByVal index As Integer, ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph(index, text, trackChanges, formatting)
			p.PackagePart = mainPart
			Return p
		End Function

		Public Overrides Function InsertParagraph(ByVal text As String) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph(text)
			p.PackagePart = mainPart
			Return p
		End Function

		Public Overrides Function InsertParagraph(ByVal text As String, ByVal trackChanges As Boolean) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph(text, trackChanges)
			p.PackagePart = mainPart
			Return p
		End Function

		Public Overrides Function InsertParagraph(ByVal text As String, ByVal trackChanges As Boolean, ByVal formatting As Formatting) As Paragraph
			Dim p As Paragraph = MyBase.InsertParagraph(text, trackChanges, formatting)
			p.PackagePart = mainPart

			Return p
		End Function

		Public Overrides Function InsertEquation(ByVal equation As String) As Paragraph
			Dim p As Paragraph = MyBase.InsertEquation(equation)
			p.PackagePart = mainPart
			Return p
		End Function


		Public Overrides ReadOnly Property Paragraphs() As ReadOnlyCollection(Of Paragraph)
			Get
				Dim l As ReadOnlyCollection(Of Paragraph) = MyBase.Paragraphs
				For Each paragraph In l
					paragraph.mainPart = mainPart
				Next paragraph
				Return l
			End Get
		End Property

		Public Overrides ReadOnly Property Tables() As List(Of Table)
			Get
				Dim l As List(Of Table) = MyBase.Tables
				l.ForEach(Function(x) x.mainPart = mainPart)
				Return l
			End Get
		End Property

		Public ReadOnly Property Images() As List(Of Image)
			Get
				Dim imageRelationships As PackageRelationshipCollection = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				If imageRelationships.Count() > 0 Then
					Return (
					    From i In imageRelationships
					    Select New Image(Document, i)).ToList()
				End If

				Return New List(Of Image)()
			End Get
		End Property
		Public Shadows Function InsertTable(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			If rowCount < 1 OrElse columnCount < 1 Then
				Throw New ArgumentOutOfRangeException("Row and Column count must be greater than zero.")
			End If

			Dim t As Table = MyBase.InsertTable(rowCount, columnCount)
			t.mainPart = mainPart
			Return t
		End Function
		Public Shadows Function InsertTable(ByVal index As Integer, ByVal t As Table) As Table
			Dim t2 As Table = MyBase.InsertTable(index, t)
			t2.mainPart = mainPart
			Return t2
		End Function
		Public Shadows Function InsertTable(ByVal t As Table) As Table
			t = MyBase.InsertTable(t)
			t.mainPart = mainPart
			Return t
		End Function
		Public Shadows Function InsertTable(ByVal index As Integer, ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			If rowCount < 1 OrElse columnCount < 1 Then
				Throw New ArgumentOutOfRangeException("Row and Column count must be greater than zero.")
			End If

			Dim t As Table = MyBase.InsertTable(index, rowCount, columnCount)
			t.mainPart = mainPart
			Return t
		End Function

	End Class
End Namespace
