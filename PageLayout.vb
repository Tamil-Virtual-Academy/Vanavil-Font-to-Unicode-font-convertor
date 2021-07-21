Namespace Novacode
	Public Class PageLayout
		Inherits DocXElement
		Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)

		End Sub


		Public Property Orientation() As Orientation
			Get
'                
'                 * Get the pgSz (page size) element for this Section,
'                 * null will be return if no such element exists.
'                 
				Dim pgSz As XElement = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName))

				If pgSz Is Nothing Then
					Return Orientation.Portrait
				End If

				' Get the attribute of the pgSz element.
				Dim val As XAttribute = pgSz.Attribute(XName.Get("orient", DocX.w.NamespaceName))

				' If val is null, this cell contains no information.
				If val Is Nothing Then
					Return Orientation.Portrait
				End If

				If val.Value.Equals("Landscape", StringComparison.CurrentCultureIgnoreCase) Then
					Return Orientation.Landscape
				Else
					Return Orientation.Portrait
				End If
			End Get

			Set(ByVal value As Orientation)
				' Check if already correct value.
				If Orientation = value Then
					Return
				End If

'                
'                 * Get the pgSz (page size) element for this Section,
'                 * null will be return if no such element exists.
'                 
				Dim pgSz As XElement = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName))

				If pgSz Is Nothing Then
					Xml.SetElementValue(XName.Get("pgSz", DocX.w.NamespaceName), String.Empty)
					pgSz = Xml.Element(XName.Get("pgSz", DocX.w.NamespaceName))
				End If

				pgSz.SetAttributeValue(XName.Get("orient", DocX.w.NamespaceName), value.ToString().ToLower())

				If value = Novacode.Orientation.Landscape Then
					pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "16838")
					pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "11906")

				ElseIf value = Novacode.Orientation.Portrait Then
					pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), "11906")
					pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), "16838")
				End If
			End Set
		End Property
	End Class
End Namespace
