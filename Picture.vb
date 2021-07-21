Imports System.IO.Packaging

Namespace Novacode
	''' <summary>
	''' Represents a Picture in this document, a Picture is a customized view of an Image.
	''' </summary>
	Public Class Picture
		Inherits DocXElement
		Private Const EmusInPixel As Integer = 9525

		Friend picture_rels As Dictionary(Of PackagePart, PackageRelationship)

		Friend img As Image
'INSTANT VB NOTE: The variable id was renamed since Visual Basic does not allow class members with the same name:
		Private id_Renamed As String
'INSTANT VB NOTE: The variable name was renamed since Visual Basic does not allow class members with the same name:
		Private name_Renamed As String
		Private descr As String
		Private cx, cy As Integer
		'private string fileName;
'INSTANT VB NOTE: The variable rotation was renamed since Visual Basic does not allow class members with the same name:
		Private rotation_Renamed As UInteger
		Private hFlip, vFlip As Boolean
		Private pictureShape As Object
		Private xfrm As XElement
		Private prstGeom As XElement

		''' <summary>
		''' Remove this Picture from this document.
		''' </summary>
		Public Sub Remove()
			Xml.Remove()
		End Sub

		''' <summary>
		''' Wraps an XElement as an Image
		''' </summary>
		''' <param name="document"></param>
		''' <param name="i">The XElement i to wrap</param>
		''' <param name="img"></param>
		Friend Sub New(ByVal document As DocX, ByVal i As XElement, ByVal img As Image)
			MyBase.New(document, i)
			picture_rels = New Dictionary(Of PackagePart, PackageRelationship)()

			Me.img = img

			Me.id_Renamed = (
			    From e In Xml.Descendants()
			    Where e.Name.LocalName.Equals("blip")
			    Select e.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()

			If Me.id_Renamed Is Nothing Then
				Me.id_Renamed = (
				    From e In Xml.Descendants()
				    Where e.Name.LocalName.Equals("imagedata")
				    Select e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")).Value).SingleOrDefault()
			End If

			Me.name_Renamed = (
			    From e In Xml.Descendants()
			    Let a = e.Attribute(XName.Get("name"))
			    Where (a IsNot Nothing)
			    Select a.Value).FirstOrDefault()

			If Me.name_Renamed Is Nothing Then
				Me.name_Renamed = (
				    From e In Xml.Descendants()
				    Let a = e.Attribute(XName.Get("title"))
				    Where (a IsNot Nothing)
				    Select a.Value).FirstOrDefault()
			End If

			Me.descr = (
			    From e In Xml.Descendants()
			    Let a = e.Attribute(XName.Get("descr"))
			    Where (a IsNot Nothing)
			    Select a.Value).FirstOrDefault()

			Me.cx = (
			    From e In Xml.Descendants()
			    Let a = e.Attribute(XName.Get("cx"))
			    Where (a IsNot Nothing)
			    Select Integer.Parse(a.Value)).FirstOrDefault()

			If Me.cx = 0 Then
				Dim style As XAttribute = (
				    From e In Xml.Descendants()
				    Let a = e.Attribute(XName.Get("style"))
				    Where (a IsNot Nothing)
				    Select a).FirstOrDefault()

				Dim fromWidth As String = style.Value.Substring(style.Value.IndexOf("width:") + 6)
				Dim widthInt = ((Double.Parse((fromWidth.Substring(0, fromWidth.IndexOf("pt"))).Replace(".", ","))) / 72.0) * 914400
				cx = Convert.ToInt32(widthInt)
			End If

			Me.cy = (
			    From e In Xml.Descendants()
			    Let a = e.Attribute(XName.Get("cy"))
			    Where (a IsNot Nothing)
			    Select Integer.Parse(a.Value)).FirstOrDefault()

			If Me.cy = 0 Then
				Dim style As XAttribute = (
				    From e In Xml.Descendants()
				    Let a = e.Attribute(XName.Get("style"))
				    Where (a IsNot Nothing)
				    Select a).FirstOrDefault()

				Dim fromHeight As String = style.Value.Substring(style.Value.IndexOf("height:") + 7)
				Dim heightInt = ((Double.Parse((fromHeight.Substring(0, fromHeight.IndexOf("pt"))).Replace(".", ","))) / 72.0) * 914400
				cy = Convert.ToInt32(heightInt)
			End If

			Me.xfrm = (
			    From d In Xml.Descendants()
			    Where d.Name.LocalName.Equals("xfrm")
			    Select d).SingleOrDefault()

			Me.prstGeom = (
			    From d In Xml.Descendants()
			    Where d.Name.LocalName.Equals("prstGeom")
			    Select d).SingleOrDefault()

			If xfrm IsNot Nothing Then
				Me.rotation_Renamed = If(xfrm.Attribute(XName.Get("rot")) Is Nothing, 0, UInteger.Parse(xfrm.Attribute(XName.Get("rot")).Value))
			End If
		End Sub

		Private Sub SetPictureShape(ByVal shape As Object)
			Me.pictureShape = shape

			Dim prst As XAttribute = prstGeom.Attribute(XName.Get("prst"))
			If prst Is Nothing Then
				prstGeom.Add(New XAttribute(XName.Get("prst"), "rectangle"))
			End If

			prstGeom.Attribute(XName.Get("prst")).Value = shape.ToString()
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the BasicShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the BasicShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As BasicShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the RectangleShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the RectangleShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As RectangleShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the BlockArrowShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the BlockArrowShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As BlockArrowShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the EquationShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the EquationShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As EquationShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the FlowchartShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the FlowchartShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As FlowchartShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the StarAndBannerShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the StarAndBannerShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As StarAndBannerShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' Set the shape of this Picture to one in the CalloutShapes enumeration.
		''' </summary>
		''' <param name="shape">A shape from the CalloutShapes enumeration.</param>
		Public Sub SetPictureShape(ByVal shape As CalloutShapes)
			SetPictureShape(CObj(shape))
		End Sub

		''' <summary>
		''' A unique id that identifies an Image embedded in this document.
		''' </summary>
		Public ReadOnly Property Id() As String
			Get
				Return id_Renamed
			End Get
		End Property

		''' <summary>
		''' Flip this Picture Horizontally.
		''' </summary>
		Public Property FlipHorizontal() As Boolean
			Get
				Return hFlip
			End Get

			Set(ByVal value As Boolean)
				hFlip = value

				Dim flipH As XAttribute = xfrm.Attribute(XName.Get("flipH"))
				If flipH Is Nothing Then
					xfrm.Add(New XAttribute(XName.Get("flipH"), "0"))
				End If

				xfrm.Attribute(XName.Get("flipH")).Value = If(hFlip, "1", "0")
			End Set
		End Property

		''' <summary>
		''' Flip this Picture Vertically.
		''' </summary>
		Public Property FlipVertical() As Boolean
			Get
				Return vFlip
			End Get

			Set(ByVal value As Boolean)
				vFlip = value

				Dim flipV As XAttribute = xfrm.Attribute(XName.Get("flipV"))
				If flipV Is Nothing Then
					xfrm.Add(New XAttribute(XName.Get("flipV"), "0"))
				End If

				xfrm.Attribute(XName.Get("flipV")).Value = If(vFlip, "1", "0")
			End Set
		End Property

		''' <summary>
		''' The rotation in degrees of this image, actual value = value % 360
		''' </summary>
		Public Property Rotation() As UInteger
			Get
				Return rotation_Renamed / 60000
			End Get

			Set(ByVal value As UInteger)
				rotation_Renamed = (value Mod 360) * 60000
				Dim xfrm As XElement = (
				    From d In Xml.Descendants()
				    Where d.Name.LocalName.Equals("xfrm")
				    Select d).Single()

				Dim rot As XAttribute = xfrm.Attribute(XName.Get("rot"))
				If rot Is Nothing Then
					xfrm.Add(New XAttribute(XName.Get("rot"), 0))
				End If

				xfrm.Attribute(XName.Get("rot")).Value = rotation_Renamed.ToString()
			End Set
		End Property

		''' <summary>
		''' Gets or sets the name of this Image.
		''' </summary>
		Public Property Name() As String
			Get
				Return name_Renamed
			End Get

			Set(ByVal value As String)
				name_Renamed = value

				For Each a As XAttribute In Xml.Descendants().Attributes(XName.Get("name"))
					a.Value = name_Renamed
				Next a
			End Set
		End Property

		''' <summary>
		''' Gets or sets the description for this Image.
		''' </summary>
		Public Property Description() As String
			Get
				Return descr
			End Get

			Set(ByVal value As String)
				descr = value

				For Each a As XAttribute In Xml.Descendants().Attributes(XName.Get("descr"))
					a.Value = descr
				Next a
			End Set
		End Property

		'''<summary>
		''' Returns the name of the image file for the picture.
		'''</summary>
		Public ReadOnly Property FileName() As String
		  Get
			Return img.FileName
		  End Get
		End Property

		''' <summary>
		''' Get or sets the Width of this Image.
		''' </summary>
		Public Property Width() As Integer
			Get
				Return cx \ EmusInPixel
			End Get

			Set(ByVal value As Integer)
				cx = value

				For Each a As XAttribute In Xml.Descendants().Attributes(XName.Get("cx"))
					a.Value = (cx * EmusInPixel).ToString()
				Next a
			End Set
		End Property

		''' <summary>
		''' Get or sets the height of this Image.
		''' </summary>
		Public Property Height() As Integer
			Get
				Return cy \ EmusInPixel
			End Get

			Set(ByVal value As Integer)
				cy = value

				For Each a As XAttribute In Xml.Descendants().Attributes(XName.Get("cy"))
					a.Value = (cy * EmusInPixel).ToString()
				Next a
			End Set
		End Property

		'public void Delete()
		'{
		'    // Remove xml
		'    i.Remove();

		'    // Rebuild the image collection for this paragraph
		'    // Requires that every Image have a link to its paragraph

		'}
	End Class
End Namespace
