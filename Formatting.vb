Imports System.Globalization
Namespace Novacode
	''' <summary>
	''' A text formatting.
	''' </summary>
	Public Class Formatting
		Implements IComparable
		Private rPr As XElement
'INSTANT VB NOTE: The variable hidden was renamed since Visual Basic does not allow class members with the same name:
		Private hidden_Renamed? As Boolean
'INSTANT VB NOTE: The variable bold was renamed since Visual Basic does not allow class members with the same name:
		Private bold_Renamed? As Boolean
'INSTANT VB NOTE: The variable italic was renamed since Visual Basic does not allow class members with the same name:
		Private italic_Renamed? As Boolean
'INSTANT VB NOTE: The variable strikethrough was renamed since Visual Basic does not allow class members with the same name:
		Private strikethrough_Renamed? As StrikeThrough
'INSTANT VB NOTE: The variable script was renamed since Visual Basic does not allow class members with the same name:
		Private script_Renamed? As Script
'INSTANT VB NOTE: The variable highlight was renamed since Visual Basic does not allow class members with the same name:
		Private highlight_Renamed? As Highlight
'INSTANT VB NOTE: The variable size was renamed since Visual Basic does not allow class members with the same name:
		Private size_Renamed? As Double
'INSTANT VB NOTE: The variable fontColor was renamed since Visual Basic does not allow class members with the same name:
		Private fontColor_Renamed? As Color
'INSTANT VB NOTE: The variable underlineColor was renamed since Visual Basic does not allow class members with the same name:
		Private underlineColor_Renamed? As Color
'INSTANT VB NOTE: The variable underlineStyle was renamed since Visual Basic does not allow class members with the same name:
		Private underlineStyle_Renamed? As UnderlineStyle
'INSTANT VB NOTE: The variable misc was renamed since Visual Basic does not allow class members with the same name:
		Private misc_Renamed? As Misc
'INSTANT VB NOTE: The variable capsStyle was renamed since Visual Basic does not allow class members with the same name:
		Private capsStyle_Renamed? As CapsStyle
'INSTANT VB NOTE: The variable fontFamily was renamed since Visual Basic does not allow class members with the same name:
		Private fontFamily_Renamed As FontFamily
'INSTANT VB NOTE: The variable percentageScale was renamed since Visual Basic does not allow class members with the same name:
		Private percentageScale_Renamed? As Integer
'INSTANT VB NOTE: The variable kerning was renamed since Visual Basic does not allow class members with the same name:
		Private kerning_Renamed? As Integer
'INSTANT VB NOTE: The variable position was renamed since Visual Basic does not allow class members with the same name:
		Private position_Renamed? As Integer
'INSTANT VB NOTE: The variable spacing was renamed since Visual Basic does not allow class members with the same name:
		Private spacing_Renamed? As Double

'INSTANT VB NOTE: The variable language was renamed since Visual Basic does not allow class members with the same name:
		Private language_Renamed As CultureInfo

		''' <summary>
		''' A text formatting.
		''' </summary>
		Public Sub New()
			capsStyle_Renamed = Novacode.CapsStyle.none
			strikethrough_Renamed = Novacode.StrikeThrough.none
			script_Renamed = Novacode.Script.none
			highlight_Renamed = Novacode.Highlight.none
			underlineStyle_Renamed = Novacode.UnderlineStyle.none
			misc_Renamed = Novacode.Misc.none

			' Use current culture by default
			language_Renamed = CultureInfo.CurrentCulture

			rPr = New XElement(XName.Get("rPr", DocX.w.NamespaceName))
		End Sub

		''' <summary>
		''' Text language
		''' </summary>
		Public Property Language() As CultureInfo
			Get
				Return language_Renamed
			End Get

			Set(ByVal value As CultureInfo)
				language_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Returns a new identical instance of Formatting.
		''' </summary>
		''' <returns></returns>
		Public Function Clone() As Formatting
			Dim newf As New Formatting()
			newf.Bold = bold_Renamed
			newf.CapsStyle = capsStyle_Renamed
			newf.FontColor = fontColor_Renamed
			newf.FontFamily = fontFamily_Renamed
			newf.Hidden = hidden_Renamed
			newf.Highlight = highlight_Renamed
			newf.Italic = italic_Renamed
			If kerning_Renamed.HasValue Then
				newf.Kerning = kerning_Renamed
			End If
			newf.Language = language_Renamed
			newf.Misc = misc_Renamed
			If percentageScale_Renamed.HasValue Then
				newf.PercentageScale = percentageScale_Renamed
			End If
			If position_Renamed.HasValue Then
				newf.Position = position_Renamed
			End If
			newf.Script = script_Renamed
			If size_Renamed.HasValue Then
				newf.Size = size_Renamed
			End If
			If spacing_Renamed.HasValue Then
				newf.Spacing = spacing_Renamed
			End If
			newf.StrikeThrough = strikethrough_Renamed
			newf.UnderlineColor = underlineColor_Renamed
			newf.UnderlineStyle = underlineStyle_Renamed
			Return newf
		End Function

		Public Shared Function Parse(ByVal rPr As XElement) As Formatting
			Dim formatting As New Formatting()

			' Build up the Formatting object.
			For Each [option] As XElement In rPr.Elements()
				Select Case [option].Name.LocalName
					Case "lang"
						formatting.Language = New CultureInfo(If(If([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName), Nothing), [option].GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName), Nothing)), [option].GetAttribute(XName.Get("bidi", DocX.w.NamespaceName))))
					Case "spacing"
						formatting.Spacing = Double.Parse([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 20.0
					Case "position"
						formatting.Position = Int32.Parse([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2
					Case "kern"
						formatting.Position = Int32.Parse([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2
					Case "w"
						formatting.PercentageScale = Int32.Parse([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName)))
					' <w:sz w:val="20"/><w:szCs w:val="20"/>
					Case "sz"
						formatting.Size = Int32.Parse([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName))) / 2


					Case "rFonts"
						formatting.FontFamily = New FontFamily(If(If(If([option].GetAttribute(XName.Get("cs", DocX.w.NamespaceName), Nothing), [option].GetAttribute(XName.Get("ascii", DocX.w.NamespaceName), Nothing)), [option].GetAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), Nothing)), [option].GetAttribute(XName.Get("eastAsia", DocX.w.NamespaceName))))
					Case "color"
						Try
							Dim color As String = [option].GetAttribute(XName.Get("val", DocX.w.NamespaceName))
							formatting.FontColor = ColorTranslator.FromHtml(String.Format("#{0}", color))
						Catch
						End Try
					Case "vanish"
						formatting.hidden_Renamed = True
					Case "b"
						formatting.Bold = True
					Case "i"
						formatting.Italic = True
					Case "u"
						formatting.UnderlineStyle = HelperFunctions.GetUnderlineStyle([option].GetAttribute(XName.Get("val", DocX.w.NamespaceName)))
					Case Else
				End Select
			Next [option]


			Return formatting
		End Function

		Friend ReadOnly Property Xml() As XElement
			Get
				rPr = New XElement(XName.Get("rPr", DocX.w.NamespaceName))

				If language_Renamed IsNot Nothing Then
					rPr.Add(New XElement(XName.Get("lang", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), language_Renamed.Name)))
				End If

				If spacing_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("spacing", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), spacing_Renamed.Value * 20)))
				End If

				If position_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("position", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), position_Renamed.Value * 2)))
				End If

				If kerning_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("kern", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), kerning_Renamed.Value * 2)))
				End If

				If percentageScale_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("w", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), percentageScale_Renamed)))
				End If

				If fontFamily_Renamed IsNot Nothing Then
					rPr.Add(New XElement (XName.Get("rFonts", DocX.w.NamespaceName), New XAttribute(XName.Get("ascii", DocX.w.NamespaceName), fontFamily_Renamed.Name), New XAttribute(XName.Get("hAnsi", DocX.w.NamespaceName), fontFamily_Renamed.Name), New XAttribute(XName.Get("cs", DocX.w.NamespaceName), fontFamily_Renamed.Name))) ' Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865 -  Added by Maurits Elbers to support non-standard characters. See http://docx.codeplex.com/Thread/View.aspx?ThreadId=70097&ANCHOR#Post453865
				End If

				If hidden_Renamed.HasValue AndAlso hidden_Renamed.Value Then
					rPr.Add(New XElement(XName.Get("vanish", DocX.w.NamespaceName)))
				End If

				If bold_Renamed.HasValue AndAlso bold_Renamed.Value Then
					rPr.Add(New XElement(XName.Get("b", DocX.w.NamespaceName)))
				End If

				If italic_Renamed.HasValue AndAlso italic_Renamed.Value Then
					rPr.Add(New XElement(XName.Get("i", DocX.w.NamespaceName)))
				End If

				If underlineStyle_Renamed.HasValue Then
					Select Case underlineStyle_Renamed
					Case Novacode.UnderlineStyle.none
					Case Novacode.UnderlineStyle.singleLine
						rPr.Add(New XElement(XName.Get("u", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")))
					Case Novacode.UnderlineStyle.doubleLine
						rPr.Add(New XElement(XName.Get("u", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), "double")))
					Case Else
						rPr.Add(New XElement(XName.Get("u", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), underlineStyle_Renamed.ToString())))
					End Select
				End If

				If underlineColor_Renamed.HasValue Then
					' If an underlineColor has been set but no underlineStyle has been set
					If underlineStyle_Renamed = Novacode.UnderlineStyle.none Then
						' Set the underlineStyle to the default
						underlineStyle_Renamed = Novacode.UnderlineStyle.singleLine
						rPr.Add(New XElement(XName.Get("u", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), "single")))
					End If

					rPr.Element(XName.Get("u", DocX.w.NamespaceName)).Add(New XAttribute(XName.Get("color", DocX.w.NamespaceName), underlineColor_Renamed.Value.ToHex()))
				End If

				If strikethrough_Renamed.HasValue Then
					Select Case strikethrough_Renamed
					Case Novacode.StrikeThrough.none
					Case Novacode.StrikeThrough.strike
						rPr.Add(New XElement(XName.Get("strike", DocX.w.NamespaceName)))
					Case Novacode.StrikeThrough.doubleStrike
						rPr.Add(New XElement(XName.Get("dstrike", DocX.w.NamespaceName)))
					Case Else
					End Select
				End If

				If script_Renamed.HasValue Then
					Select Case script_Renamed
					Case Novacode.Script.none
					Case Else
						rPr.Add(New XElement(XName.Get("vertAlign", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), script_Renamed.ToString())))
					End Select
				End If

				If size_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("sz", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), (size_Renamed * 2).ToString())))
					rPr.Add(New XElement(XName.Get("szCs", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), (size_Renamed * 2).ToString())))
				End If

				If fontColor_Renamed.HasValue Then
					rPr.Add(New XElement(XName.Get("color", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), fontColor_Renamed.Value.ToHex())))
				End If

				If highlight_Renamed.HasValue Then
					Select Case highlight_Renamed
					Case Novacode.Highlight.none
					Case Else
						rPr.Add(New XElement(XName.Get("highlight", DocX.w.NamespaceName), New XAttribute(XName.Get("val", DocX.w.NamespaceName), highlight_Renamed.ToString())))
					End Select
				End If

				If capsStyle_Renamed.HasValue Then
					Select Case capsStyle_Renamed
					Case Novacode.CapsStyle.none
					Case Else
						rPr.Add(New XElement(XName.Get(capsStyle_Renamed.ToString(), DocX.w.NamespaceName)))
					End Select
				End If

				If misc_Renamed.HasValue Then
					Select Case misc_Renamed
					Case Novacode.Misc.none
					Case Novacode.Misc.outlineShadow
						rPr.Add(New XElement(XName.Get("outline", DocX.w.NamespaceName)))
						rPr.Add(New XElement(XName.Get("shadow", DocX.w.NamespaceName)))
					Case Novacode.Misc.engrave
						rPr.Add(New XElement(XName.Get("imprint", DocX.w.NamespaceName)))
					Case Else
						rPr.Add(New XElement(XName.Get(misc_Renamed.ToString(), DocX.w.NamespaceName)))
					End Select
				End If

				Return rPr
			End Get
		End Property

		''' <summary>
		''' This formatting will apply Bold.
		''' </summary>
		Public Property Bold() As Boolean?
			Get
				Return bold_Renamed
			End Get
			Set(ByVal value? As Boolean)
				bold_Renamed = value
			End Set
		End Property

		''' <summary>
		''' This formatting will apply Italic.
		''' </summary>
		Public Property Italic() As Boolean?
			Get
				Return italic_Renamed
			End Get
			Set(ByVal value? As Boolean)
				italic_Renamed = value
			End Set
		End Property

		''' <summary>
		''' This formatting will apply StrickThrough.
		''' </summary>
		Public Property StrikeThrough() As StrikeThrough?
			Get
				Return strikethrough_Renamed
			End Get
			Set(ByVal value? As StrikeThrough)
				strikethrough_Renamed = value
			End Set
		End Property

		''' <summary>
		''' The script that this formatting should be, normal, superscript or subscript.
		''' </summary>
		Public Property Script() As Script?
			Get
				Return script_Renamed
			End Get
			Set(ByVal value? As Script)
				script_Renamed = value
			End Set
		End Property

		''' <summary>
		''' The Size of this text, must be between 0 and 1638.
		''' </summary>
		Public Property Size() As Double?
			Get
				Return size_Renamed
			End Get

			Set(ByVal value? As Double)
				Dim temp? As Double = value * 2

				If temp - CInt(Fix(temp)) = 0 Then
					If value > 0 AndAlso value < 1639 Then
						size_Renamed = value
					Else
						Throw New ArgumentException("Size", "Value must be in the range 0 - 1638")
					End If

				Else
					Throw New ArgumentException("Size", "Value must be either a whole or half number, examples: 32, 32.5")
				End If
			End Set
		End Property

		''' <summary>
		''' Percentage scale must be one of the following values 200, 150, 100, 90, 80, 66, 50 or 33.
		''' </summary>
		Public Property PercentageScale() As Integer?
			Get
				Return percentageScale_Renamed
			End Get

			Set(ByVal value? As Integer)
				If (New Integer?() { 200, 150, 100, 90, 80, 66, 50, 33 }).Contains(value) Then
					percentageScale_Renamed = value
				Else
					Throw New ArgumentOutOfRangeException("PercentageScale", "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33")
				End If
			End Set
		End Property

		''' <summary>
		''' The Kerning to apply to this text must be one of the following values 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72.
		''' </summary>
		Public Property Kerning() As Integer?
			Get
				Return kerning_Renamed
			End Get

			Set(ByVal value? As Integer)
				If New Integer?() {8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72}.Contains(value) Then
					kerning_Renamed = value
				Else
					Throw New ArgumentOutOfRangeException("Kerning", "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72")
				End If
			End Set
		End Property

		''' <summary>
		''' Text position must be in the range (-1585 - 1585).
		''' </summary>
		Public Property Position() As Integer?
			Get
				Return position_Renamed
			End Get

			Set(ByVal value? As Integer)
				If value > -1585 AndAlso value < 1585 Then
					position_Renamed = value
				Else
					Throw New ArgumentOutOfRangeException("Position", "Value must be in the range -1585 - 1585")
				End If
			End Set
		End Property

		''' <summary>
		''' Text spacing must be in the range (-1585 - 1585).
		''' </summary>
		Public Property Spacing() As Double?
			Get
				Return spacing_Renamed
			End Get

			Set(ByVal value? As Double)
				Dim temp? As Double = value * 20

				If temp - CInt(Fix(temp)) = 0 Then
					If value > -1585 AndAlso value < 1585 Then
						spacing_Renamed = value
					Else
						Throw New ArgumentException("Spacing", "Value must be in the range: -1584 - 1584")
					End If

				Else
					Throw New ArgumentException("Spacing", "Value must be either a whole or acurate to one decimal, examples: 32, 32.1, 32.2, 32.9")
				End If
			End Set
		End Property

		''' <summary>
		''' The colour of the text.
		''' </summary>
		Public Property FontColor() As Color?
			Get
				Return fontColor_Renamed
			End Get
			Set(ByVal value? As Color)
				fontColor_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Highlight colour.
		''' </summary>
		Public Property Highlight() As Highlight?
			Get
				Return highlight_Renamed
			End Get
			Set(ByVal value? As Highlight)
				highlight_Renamed = value
			End Set
		End Property

		''' <summary>
		''' The Underline style that this formatting applies.
		''' </summary>
		Public Property UnderlineStyle() As UnderlineStyle?
			Get
				Return underlineStyle_Renamed
			End Get
			Set(ByVal value? As UnderlineStyle)
				underlineStyle_Renamed = value
			End Set
		End Property

		''' <summary>
		''' The underline colour.
		''' </summary>
		Public Property UnderlineColor() As Color?
			Get
				Return underlineColor_Renamed
			End Get
			Set(ByVal value? As Color)
				underlineColor_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Misc settings.
		''' </summary>
		Public Property Misc() As Misc?
			Get
				Return misc_Renamed
			End Get
			Set(ByVal value? As Misc)
				misc_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Is this text hidden or visible.
		''' </summary>
		Public Property Hidden() As Boolean?
			Get
				Return hidden_Renamed
			End Get
			Set(ByVal value? As Boolean)
				hidden_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Capitalization style.
		''' </summary>
		Public Property CapsStyle() As CapsStyle?
			Get
				Return capsStyle_Renamed
			End Get
			Set(ByVal value? As CapsStyle)
				capsStyle_Renamed = value
			End Set
		End Property

		''' <summary>
		''' The font family of this formatting.
		''' </summary>
		''' <!-- 
		''' Bug found and fixed by krugs525 on August 12 2009.
		''' Use TFS compare to see exact code change.
		''' -->
		Public Property FontFamily() As FontFamily
			Get
				Return fontFamily_Renamed
			End Get
			Set(ByVal value As FontFamily)
				fontFamily_Renamed = value
			End Set
		End Property

		Public Function CompareTo(ByVal obj As Object) As Integer Implements IComparable.CompareTo
			Dim other As Formatting = CType(obj, Formatting)

			If Not other.hidden_Renamed.Equals(Me.hidden_Renamed) Then
				Return -1
			End If

			If Not other.bold_Renamed.Equals(Me.bold_Renamed) Then
				Return -1
			End If

			If Not other.italic_Renamed.Equals(Me.italic_Renamed) Then
				Return -1
			End If

			If Not other.strikethrough_Renamed.Equals(Me.strikethrough_Renamed) Then
				Return -1
			End If

			If Not other.script_Renamed.Equals(Me.script_Renamed) Then
				Return -1
			End If

			If Not other.highlight_Renamed.Equals(Me.highlight_Renamed) Then
				Return -1
			End If

			If Not other.size_Renamed.Equals(Me.size_Renamed) Then
				Return -1
			End If

			If Not other.fontColor_Renamed.Equals(Me.fontColor_Renamed) Then
				Return -1
			End If

			If Not other.underlineColor_Renamed.Equals(Me.underlineColor_Renamed) Then
				Return -1
			End If

			If Not other.underlineStyle_Renamed.Equals(Me.underlineStyle_Renamed) Then
				Return -1
			End If

			If Not other.misc_Renamed.Equals(Me.misc_Renamed) Then
				Return -1
			End If

			If Not other.capsStyle_Renamed.Equals(Me.capsStyle_Renamed) Then
				Return -1
			End If

			If other.fontFamily_Renamed IsNot Me.fontFamily_Renamed Then
				Return -1
			End If

			If Not other.percentageScale_Renamed.Equals(Me.percentageScale_Renamed) Then
				Return -1
			End If

			If Not other.kerning_Renamed.Equals(Me.kerning_Renamed) Then
				Return -1
			End If

			If Not other.position_Renamed.Equals(Me.position_Renamed) Then
				Return -1
			End If

			If Not other.spacing_Renamed.Equals(Me.spacing_Renamed) Then
				Return -1
			End If

			If Not other.language_Renamed.Equals(Me.language_Renamed) Then
				Return -1
			End If

			Return 0
		End Function
	End Class
End Namespace
