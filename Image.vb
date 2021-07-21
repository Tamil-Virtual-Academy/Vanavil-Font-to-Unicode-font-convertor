Imports System.IO.Packaging
Imports System.IO

Namespace Novacode
	''' <summary>
	''' Represents an Image embedded in a document.
	''' </summary>
	Public Class Image
		''' <summary>
		''' A unique id which identifies this Image.
		''' </summary>
'INSTANT VB NOTE: The variable id was renamed since Visual Basic does not allow class members with the same name:
		Private id_Renamed As String
		Private document As DocX
		Friend pr As PackageRelationship

		Public Function GetStream(ByVal mode As FileMode, ByVal access As FileAccess) As Stream
			Dim temp As String = pr.SourceUri.OriginalString
			Dim start As String = temp.Remove(temp.LastIndexOf("/"c))
			Dim [end] As String = pr.TargetUri.OriginalString
			Dim full As String = start & "/" & [end]

			Return(document.package.GetPart(New Uri(full, UriKind.Relative)).GetStream(mode, access))
		End Function

		''' <summary>
		''' Returns the id of this Image.
		''' </summary>
		Public ReadOnly Property Id() As String
			Get
				Return id_Renamed
			End Get
		End Property

		Friend Sub New(ByVal document As DocX, ByVal pr As PackageRelationship)
			Me.document = document
			Me.pr = pr
			Me.id_Renamed = pr.Id
		End Sub

		''' <summary>
		''' Add an image to a document, create a custom view of that image (picture) and then insert it into a Paragraph using append.
		''' </summary>
		''' <returns></returns>
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
		Public Function CreatePicture() As Picture
			Return Paragraph.CreatePicture(document, id_Renamed, String.Empty, String.Empty)
		End Function
		Public Function CreatePicture(ByVal height As Integer, ByVal width As Integer) As Picture
			Dim picture As Picture = Paragraph.CreatePicture(document, id_Renamed, String.Empty, String.Empty)
			picture.Height = height
			picture.Width = width
			Return picture
		End Function

	  '''<summary>
	  ''' Returns the name of the image file.
	  '''</summary>
	  Public ReadOnly Property FileName() As String
		Get
		  Return Path.GetFileName(Me.pr.TargetUri.ToString())
		End Get
	  End Property
	End Class
End Namespace
