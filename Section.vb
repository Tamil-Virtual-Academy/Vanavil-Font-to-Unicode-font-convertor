Imports System.IO.Packaging

Namespace Novacode
  Public Class Section
	  Inherits Container

	Public SectionBreakType As SectionBreakType

	Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
		MyBase.New(document, xml)
	End Sub

	Public Property SectionParagraphs() As List(Of Paragraph)
  End Class
End Namespace