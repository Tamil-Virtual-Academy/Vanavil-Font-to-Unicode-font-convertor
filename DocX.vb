Imports System.IO
Imports System.IO.Packaging
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Collections.ObjectModel

Namespace Novacode
	''' <summary>
	''' Represents a document.
	''' </summary>
	Public Class DocX
		Inherits Container
		Implements IDisposable
		#Region "Namespaces"
		Friend Shared w As XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
		Friend Shared rel As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"

		Friend Shared r As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		Friend Shared m As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math"
		Friend Shared customPropertiesSchema As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
		Friend Shared customVTypesSchema As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

		Friend Shared wp As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
		Friend Shared a As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
		Friend Shared c As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart"

		Friend Shared v As XNamespace = "urn:schemas-microsoft-com:vml"

		Friend Shared n As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
		#End Region

		Friend Function getMarginAttribute(ByVal name As XName) As Single
			Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
			Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
			If sectPr IsNot Nothing Then
				Dim pgMar As XElement = sectPr.Element(XName.Get("pgMar", DocX.w.NamespaceName))
				If pgMar IsNot Nothing Then
					Dim top As XAttribute = pgMar.Attribute(name)
					If top IsNot Nothing Then
						Dim f As Single
						If Single.TryParse(top.Value, f) Then
							Return CInt(Fix(f / 20.0f))
						End If
					End If
				End If
			End If

			Return 0
		End Function

		Friend Sub setMarginAttribute(ByVal xName As XName, ByVal value As Single)
			Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
			Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
			If sectPr IsNot Nothing Then
				Dim pgMar As XElement = sectPr.Element(XName.Get("pgMar", DocX.w.NamespaceName))
				If pgMar IsNot Nothing Then
					Dim top As XAttribute = pgMar.Attribute(xName)
					If top IsNot Nothing Then
						top.SetValue(value * 20)
					End If
				End If
			End If
		End Sub

		Public ReadOnly Property Bookmarks() As BookmarkCollection
			Get
'INSTANT VB NOTE: The local variable bookmarks was renamed since Visual Basic will not allow local variables with the same name as their enclosing function or property:
				Dim bookmarks_Renamed As New BookmarkCollection()
				For Each paragraph As Paragraph In Paragraphs
					bookmarks_Renamed.AddRange(paragraph.GetBookmarks())
				Next paragraph
				Return bookmarks_Renamed
			End Get
		End Property

		''' <summary>
		''' Top margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property MarginTop() As Single
			Get
				Return getMarginAttribute(XName.Get("top", DocX.w.NamespaceName))
			End Get

			Set(ByVal value As Single)
				setMarginAttribute(XName.Get("top", DocX.w.NamespaceName), value)
			End Set
		End Property

		''' <summary>
		''' Bottom margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property MarginBottom() As Single
			Get
				Return getMarginAttribute(XName.Get("bottom", DocX.w.NamespaceName))
			End Get

			Set(ByVal value As Single)
				setMarginAttribute(XName.Get("bottom", DocX.w.NamespaceName), value)
			End Set
		End Property

		''' <summary>
		''' Left margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property MarginLeft() As Single
			Get
				Return getMarginAttribute(XName.Get("left", DocX.w.NamespaceName))
			End Get

			Set(ByVal value As Single)
				setMarginAttribute(XName.Get("left", DocX.w.NamespaceName), value)
			End Set
		End Property

		''' <summary>
		''' Right margin value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property MarginRight() As Single
			Get
				Return getMarginAttribute(XName.Get("right", DocX.w.NamespaceName))
			End Get

			Set(ByVal value As Single)
				setMarginAttribute(XName.Get("right", DocX.w.NamespaceName), value)
			End Set
		End Property

		''' <summary>
		''' Page width value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property PageWidth() As Single
			Get
				Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
				Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
				If sectPr IsNot Nothing Then
					Dim pgSz As XElement = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName))

					If pgSz IsNot Nothing Then
						Dim w As XAttribute = pgSz.Attribute(XName.Get("w", DocX.w.NamespaceName))
						If w IsNot Nothing Then
							Dim f As Single
							If Single.TryParse(w.Value, f) Then
								Return CInt(Fix(f / 20.0f))
							End If
						End If
					End If
				End If

				Return (12240.0f / 20.0f)
			End Get

			Set(ByVal value As Single)
				Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))

				If body IsNot Nothing Then
					Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))

					If sectPr IsNot Nothing Then
						Dim pgSz As XElement = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName))

						If pgSz IsNot Nothing Then
							pgSz.SetAttributeValue(XName.Get("w", DocX.w.NamespaceName), value * 20)
						End If
					End If
				End If
			End Set
		End Property

		''' <summary>
		''' Page height value in points. 1pt = 1/72 of an inch. Word internally writes docx using units = 1/20th of a point.
		''' </summary>
		Public Property PageHeight() As Single
			Get
				Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
				Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
				If sectPr IsNot Nothing Then
					Dim pgSz As XElement = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName))

					If pgSz IsNot Nothing Then
						Dim w As XAttribute = pgSz.Attribute(XName.Get("h", DocX.w.NamespaceName))
						If w IsNot Nothing Then
							Dim f As Single
							If Single.TryParse(w.Value, f) Then
								Return CInt(Fix(f / 20.0f))
							End If
						End If
					End If
				End If

				Return (15840.0f / 20.0f)
			End Get

			Set(ByVal value As Single)
				Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))

				If body IsNot Nothing Then
					Dim sectPr As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))

					If sectPr IsNot Nothing Then
						Dim pgSz As XElement = sectPr.Element(XName.Get("pgSz", DocX.w.NamespaceName))

						If pgSz IsNot Nothing Then
							pgSz.SetAttributeValue(XName.Get("h", DocX.w.NamespaceName), value * 20)
						End If
					End If
				End If
			End Set
		End Property
		''' <summary>
		''' Returns true if any editing restrictions are imposed on this document.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     if(document.isProtected)
		'''         Console.WriteLine("Protected");
		'''     else
		'''         Console.WriteLine("Not protected");
		'''         
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AddProtection(EditRestrictions)"/>
		''' <seealso cref="RemoveProtection"/>
		''' <seealso cref="GetProtectionType"/>
		Public ReadOnly Property isProtected() As Boolean
			Get
				Return settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).Count() > 0
			End Get
		End Property

		''' <summary>
		''' Returns the type of editing protection imposed on this document.
		''' </summary>
		''' <returns>The type of editing protection imposed on this document.</returns>
		''' <example>
		''' <code>
		''' Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Make sure the document is protected before checking the protection type.
		'''     if (document.isProtected)
		'''     {
		'''         EditRestrictions protection = document.GetProtectionType();
		'''         Console.WriteLine("Document is protected using " + protection.ToString());
		'''     }
		'''
		'''     else
		'''         Console.WriteLine("Document is not protected.");
		'''
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AddProtection(EditRestrictions)"/>
		''' <seealso cref="RemoveProtection"/>
		''' <seealso cref="isProtected"/>
		Public Function GetProtectionType() As EditRestrictions
			If isProtected Then
				Dim documentProtection As XElement = settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).FirstOrDefault()
				Dim edit_type As String = documentProtection.Attribute(XName.Get("edit", DocX.w.NamespaceName)).Value
				Return CType(System.Enum.Parse(GetType(EditRestrictions), edit_type), EditRestrictions)
			End If

			Return EditRestrictions.none
		End Function

		''' <summary>
		''' Add editing protection to this document. 
		''' </summary>
		''' <param name="er">The type of protection to add to this document.</param>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Allow no editing, only the adding of comment.
		'''     document.AddProtection(EditRestrictions.comments);
		'''     
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="RemoveProtection"/>
		''' <seealso cref="GetProtectionType"/>
		''' <seealso cref="isProtected"/>
		Public Sub AddProtection(ByVal er As EditRestrictions)
			' Call remove protection before adding a new protection element.
			RemoveProtection()

			If er = EditRestrictions.none Then
				Return
			End If

			Dim documentProtection As New XElement(XName.Get("documentProtection", DocX.w.NamespaceName))
			documentProtection.Add(New XAttribute(XName.Get("edit", DocX.w.NamespaceName), er.ToString()))
			documentProtection.Add(New XAttribute(XName.Get("enforcement", DocX.w.NamespaceName), "1"))

			settings.Root.AddFirst(documentProtection)
		End Sub

		Public Sub AddProtection(ByVal er As EditRestrictions, ByVal strPassword As String)
			' http://blogs.msdn.com/b/vsod/archive/2010/04/05/how-to-set-the-editing-restrictions-in-word-using-open-xml-sdk-2-0.aspx
			' Call remove protection before adding a new protection element.
			RemoveProtection()

			If er = EditRestrictions.none Then
				Return
			End If

			Dim documentProtection As New XElement(XName.Get("documentProtection", DocX.w.NamespaceName))
			documentProtection.Add(New XAttribute(XName.Get("edit", DocX.w.NamespaceName), er.ToString()))
			documentProtection.Add(New XAttribute(XName.Get("enforcement", DocX.w.NamespaceName), "1"))

			Dim InitialCodeArray() As Integer = { &HE1F0, &H1D0F, &HCC9C, &H84C0, &H110C, &HE10, &HF1CE, &H313E, &H1872, &HE139, &HD40F, &H84F9, &H280C, &HA96A, &H4EC3 }
			Dim EncryptionMatrix(,) As Integer = { {&HAEFC, &H4DD9, &H9BB2, &H2745, &H4E8A, &H9D14, &H2A09}, {&H7B61, &HF6C2, &HFDA5, &HEB6B, &HC6F7, &H9DCF, &H2BBF}, {&H4563, &H8AC6, &H5AD, &HB5A, &H16B4, &H2D68, &H5AD0}, {&H375, &H6EA, &HDD4, &H1BA8, &H3750, &H6EA0, &HDD40}, {&HD849, &HA0B3, &H5147, &HA28E, &H553D, &HAA7A, &H44D5}, {&H6F45, &HDE8A, &HAD35, &H4A4B, &H9496, &H390D, &H721A}, {&HEB23, &HC667, &H9CEF, &H29FF, &H53FE, &HA7FC, &H5FD9}, {&H47D3, &H8FA6, &HF6D, &H1EDA, &H3DB4, &H7B68, &HF6D0}, {&HB861, &H60E3, &HC1C6, &H93AD, &H377B, &H6EF6, &HDDEC}, {&H45A0, &H8B40, &H6A1, &HD42, &H1A84, &H3508, &H6A10}, {&HAA51, &H4483, &H8906, &H22D, &H45A, &H8B4, &H1168}, {&H76B4, &HED68, &HCAF1, &H85C3, &H1BA7, &H374E, &H6E9C}, {&H3730, &H6E60, &HDCC0, &HA9A1, &H4363, &H86C6, &H1DAD}, {&H3331, &H6662, &HCCC4, &H89A9, &H373, &H6E6, &HDCC}, {&H1021, &H2042, &H4084, &H8108, &H1231, &H2462, &H48C4} }
		' char 1  
		' char 2  
		' char 3  
		' char 4  
		' char 5  
		' char 6  
		' char 7  
		' char 8  
		' char 9  
		' char 10 
		' char 11 
		' char 12 
		' char 13 
		' char 14 
		' char 15 

			' Generate the Salt
			Dim arrSalt(15) As Byte
			Dim rand As RandomNumberGenerator = New RNGCryptoServiceProvider()
			rand.GetNonZeroBytes(arrSalt)

			'Array to hold Key Values
			Dim generatedKey(3) As Byte

			'Maximum length of the password is 15 chars.
			Dim intMaxPasswordLength As Integer = 15

			If Not String.IsNullOrEmpty(strPassword) Then
				strPassword = strPassword.Substring(0, Math.Min(strPassword.Length, intMaxPasswordLength))

				Dim arrByteChars(strPassword.Length - 1) As Byte

				For intLoop As Integer = 0 To strPassword.Length - 1
					Dim intTemp As Integer = Convert.ToInt32(strPassword.Chars(intLoop))
					arrByteChars(intLoop) = Convert.ToByte(intTemp And &HFF)
					If arrByteChars(intLoop) = 0 Then
						arrByteChars(intLoop) = Convert.ToByte((intTemp And &HFF00) >> 8)
					End If
				Next intLoop

				Dim intHighOrderWord As Integer = InitialCodeArray(arrByteChars.Length - 1)

				For intLoop As Integer = 0 To arrByteChars.Length - 1
					Dim tmp As Integer = intMaxPasswordLength - arrByteChars.Length + intLoop
					For intBit As Integer = 0 To 6
						If (arrByteChars(intLoop) And (&H1 << intBit)) <> 0 Then
							intHighOrderWord = intHighOrderWord Xor EncryptionMatrix(tmp, intBit)
						End If
					Next intBit
				Next intLoop

				Dim intLowOrderWord As Integer = 0

				' For each character in the strPassword, going backwards
				For intLoopChar As Integer = arrByteChars.Length - 1 To 0 Step -1
					intLowOrderWord = (((intLowOrderWord >> 14) And &H1) Or ((intLowOrderWord << 1) And &H7FFF)) Xor arrByteChars(intLoopChar)
				Next intLoopChar

				intLowOrderWord = (((intLowOrderWord >> 14) And &H1) Or ((intLowOrderWord << 1) And &H7FFF)) Xor arrByteChars.Length Xor &HCE4B

				' Combine the Low and High Order Word
				Dim intCombinedkey As Integer = (intHighOrderWord << 16) + intLowOrderWord

				' The byte order of the result shall be reversed [Example: 0x64CEED7E becomes 7EEDCE64. end example],
				' and that value shall be hashed as defined by the attribute values.

				For intTemp As Integer = 0 To 3
					generatedKey(intTemp) = Convert.ToByte((CUInt(intCombinedkey And (&HFF << (intTemp * 8)))) >> (intTemp * 8))
				Next intTemp
			End If

			Dim sb As New StringBuilder()
			For intTemp As Integer = 0 To 3
				sb.Append(Convert.ToString(generatedKey(intTemp), 16))
			Next intTemp
			generatedKey = Encoding.Unicode.GetBytes(sb.ToString().ToUpper())

			Dim tmpArray1() As Byte = generatedKey
			Dim tmpArray2() As Byte = arrSalt
			Dim tempKey(tmpArray1.Length + tmpArray2.Length - 1) As Byte
			Buffer.BlockCopy(tmpArray2, 0, tempKey, 0, tmpArray2.Length)
			Buffer.BlockCopy(tmpArray1, 0, tempKey, tmpArray2.Length, tmpArray1.Length)
			generatedKey = tempKey


			Dim iterations As Integer = 100000

			Dim sha1 As HashAlgorithm = New SHA1Managed()
			generatedKey = sha1.ComputeHash(generatedKey)
			Dim iterator(3) As Byte
			For intTmp As Integer = 0 To iterations - 1

				iterator(0) = Convert.ToByte((intTmp And &HFF) >> 0)
				iterator(1) = Convert.ToByte((intTmp And &HFF00) >> 8)
				iterator(2) = Convert.ToByte((intTmp And &HFF0000) >> 16)
				iterator(3) = Convert.ToByte((intTmp And &HFF000000L) >> 24)

				generatedKey = concatByteArrays(iterator, generatedKey)
				generatedKey = sha1.ComputeHash(generatedKey)
			Next intTmp

			documentProtection.Add(New XAttribute(XName.Get("cryptProviderType", DocX.w.NamespaceName), "rsaFull"))
			documentProtection.Add(New XAttribute(XName.Get("cryptAlgorithmClass", DocX.w.NamespaceName), "hash"))
			documentProtection.Add(New XAttribute(XName.Get("cryptAlgorithmType", DocX.w.NamespaceName), "typeAny"))
			documentProtection.Add(New XAttribute(XName.Get("cryptAlgorithmSid", DocX.w.NamespaceName), "4")) ' SHA1
			documentProtection.Add(New XAttribute(XName.Get("cryptSpinCount", DocX.w.NamespaceName), iterations.ToString()))
			documentProtection.Add(New XAttribute(XName.Get("hash", DocX.w.NamespaceName), Convert.ToBase64String(generatedKey)))
			documentProtection.Add(New XAttribute(XName.Get("salt", DocX.w.NamespaceName), Convert.ToBase64String(arrSalt)))

			settings.Root.AddFirst(documentProtection)
		End Sub

		Private Function concatByteArrays(ByVal array1() As Byte, ByVal array2() As Byte) As Byte()
			Dim result(array1.Length + array2.Length - 1) As Byte
			Buffer.BlockCopy(array2, 0, result, 0, array2.Length)
			Buffer.BlockCopy(array1, 0, result, array2.Length, array1.Length)
			Return result
		End Function

		''' <summary>
		''' Remove editing protection from this document.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a new document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Remove any editing restrictions that are imposed on this document.
		'''     document.RemoveProtection();
		'''
		'''     // Save the document.
		'''     document.Save();
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AddProtection(EditRestrictions)"/>
		''' <seealso cref="GetProtectionType"/>
		''' <seealso cref="isProtected"/>
		Public Sub RemoveProtection()
			' Remove every node of type documentProtection.
			settings.Descendants(XName.Get("documentProtection", DocX.w.NamespaceName)).Remove()
		End Sub

		Public ReadOnly Property PageLayout() As PageLayout
			Get
				Dim sectPr As XElement = Xml.Element(XName.Get("sectPr", DocX.w.NamespaceName))
				If sectPr Is Nothing Then
					Xml.SetElementValue(XName.Get("sectPr", DocX.w.NamespaceName), String.Empty)
					sectPr = Xml.Element(XName.Get("sectPr", DocX.w.NamespaceName))
				End If

				Return New PageLayout(Me, sectPr)
			End Get
		End Property

		''' <summary>
		''' Returns a collection of Headers in this Document.
		''' A document typically contains three Headers.
		''' A default one (odd), one for the first page and one for even pages.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Add header support to this document.
		'''    document.AddHeaders();
		'''
		'''    // Get a collection of all headers in this document.
		'''    Headers headers = document.Headers;
		'''
		'''    // The header used for the first page of this document.
		'''    Header first = headers.first;
		'''
		'''    // The header used for odd pages of this document.
		'''    Header odd = headers.odd;
		'''
		'''    // The header used for even pages of this document.
		'''    Header even = headers.even;
		''' }
		''' </code>
		''' </example>
		Public ReadOnly Property Headers() As Headers
			Get
				Return headers_Renamed
			End Get
		End Property
'INSTANT VB NOTE: The variable headers was renamed since Visual Basic does not allow class members with the same name:
		Private headers_Renamed As Headers

		''' <summary>
		''' Returns a collection of Footers in this Document.
		''' A document typically contains three Footers.
		''' A default one (odd), one for the first page and one for even pages.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Add footer support to this document.
		'''    document.AddFooters();
		'''
		'''    // Get a collection of all footers in this document.
		'''    Footers footers = document.Footers;
		'''
		'''    // The footer used for the first page of this document.
		'''    Footer first = footers.first;
		'''
		'''    // The footer used for odd pages of this document.
		'''    Footer odd = footers.odd;
		'''
		'''    // The footer used for even pages of this document.
		'''    Footer even = footers.even;
		''' }
		''' </code>
		''' </example>
		Public ReadOnly Property Footers() As Footers
			Get
				Return footers_Renamed
			End Get
		End Property

'INSTANT VB NOTE: The variable footers was renamed since Visual Basic does not allow class members with the same name:
		Private footers_Renamed As Footers

		''' <summary>
		''' Should the Document use different Headers and Footers for odd and even pages?
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add header support to this document.
		'''     document.AddHeaders();
		'''
		'''     // Get a collection of all headers in this document.
		'''     Headers headers = document.Headers;
		'''
		'''     // The header used for odd pages of this document.
		'''     Header odd = headers.odd;
		'''
		'''     // The header used for even pages of this document.
		'''     Header even = headers.even;
		'''
		'''     // Force the document to use a different header for odd and even pages.
		'''     document.DifferentOddAndEvenPages = true;
		'''
		'''     // Content can be added to the Headers in the same manor that it would be added to the main document.
		'''     Paragraph p1 = odd.InsertParagraph();
		'''     p1.Append("This is the odd pages header.");
		'''     
		'''     Paragraph p2 = even.InsertParagraph();
		'''     p2.Append("This is the even pages header.");
		'''
		'''     // Save all changes to this document.
		'''     document.Save();    
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Property DifferentOddAndEvenPages() As Boolean
			Get
				Dim settings As XDocument
				Using tr As TextReader = New StreamReader(settingsPart.GetStream())
					settings = XDocument.Load(tr)
				End Using

				Dim evenAndOddHeaders As XElement = settings.Root.Element(w + "evenAndOddHeaders")

				Return evenAndOddHeaders IsNot Nothing
			End Get

			Set(ByVal value As Boolean)
				Dim settings As XDocument
				Using tr As TextReader = New StreamReader(settingsPart.GetStream())
					settings = XDocument.Load(tr)
				End Using

				Dim evenAndOddHeaders As XElement = settings.Root.Element(w + "evenAndOddHeaders")
				If evenAndOddHeaders Is Nothing Then
					If value Then
						settings.Root.AddFirst(New XElement(w + "evenAndOddHeaders"))
					End If
				Else
					If Not value Then
						evenAndOddHeaders.Remove()
					End If
				End If

				Using tw As TextWriter = New StreamWriter(settingsPart.GetStream())
					settings.Save(tw)
				End Using
			End Set
		End Property

		''' <summary>
		''' Should the Document use an independent Header and Footer for the first page?
		''' </summary>
		''' <example>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add header support to this document.
		'''     document.AddHeaders();
		'''
		'''     // The header used for the first page of this document.
		'''     Header first = document.Headers.first;
		'''
		'''     // Force the document to use a different header for first page.
		'''     document.DifferentFirstPage = true;
		'''     
		'''     // Content can be added to the Headers in the same manor that it would be added to the main document.
		'''     Paragraph p = first.InsertParagraph();
		'''     p.Append("This is the first pages header.");
		'''
		'''     // Save all changes to this document.
		'''     document.Save();    
		''' }// Release this document from memory.
		''' </example>
		Public Property DifferentFirstPage() As Boolean
			Get
				Dim body As XElement = mainDoc.Root.Element(w + "body")
				Dim sectPr As XElement = body.Element(w + "sectPr")

				If sectPr IsNot Nothing Then
					Dim titlePg As XElement = sectPr.Element(w + "titlePg")
					If titlePg IsNot Nothing Then
						Return True
					End If
				End If

				Return False
			End Get

			Set(ByVal value As Boolean)
				Dim body As XElement = mainDoc.Root.Element(w + "body")
				Dim sectPr As XElement = Nothing
				Dim titlePg As XElement = Nothing

				If sectPr Is Nothing Then
					body.Add(New XElement(w + "sectPr", String.Empty))
				End If

				sectPr = body.Element(w + "sectPr")

				titlePg = sectPr.Element(w + "titlePg")
				If titlePg Is Nothing Then
					If value Then
						sectPr.Add(New XElement(w + "titlePg", String.Empty))
					End If
				Else
					If Not value Then
						titlePg.Remove()
					End If
				End If
			End Set
		End Property

		Private Function GetHeaderByType(ByVal type As String) As Header
			Return CType(GetHeaderOrFooterByType(type, True), Header)
		End Function

		Private Function GetFooterByType(ByVal type As String) As Footer
			Return CType(GetHeaderOrFooterByType(type, False), Footer)
		End Function

		Private Function GetHeaderOrFooterByType(ByVal type As String, ByVal isHeader As Boolean) As Object
			' Switch which handles either case Header\Footer, this just cuts down on code duplication.
			Dim reference As String = "footerReference"
			If isHeader Then
				reference = "headerReference"
			End If

			' Get the Id of the [default, even or first] [Header or Footer]

			Dim Id As String = (
			    From e In mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
			    Where (e.Name.LocalName = reference) AndAlso (e.Attribute(w + "type").Value = type)
			    Select e.Attribute(r + "id").Value).LastOrDefault()

			If Id IsNot Nothing Then
				' Get the Xml file for this Header or Footer.
				Dim partUri As Uri = mainPart.GetRelationship(Id).TargetUri

				' Weird problem with PackaePart API.
				If Not partUri.OriginalString.StartsWith("/word/") Then
					partUri = New Uri("/word/" & partUri.OriginalString, UriKind.Relative)
				End If

				' Get the Part and open a stream to get the Xml file.
				Dim part As PackagePart = package.GetPart(partUri)

				Dim doc As XDocument
				Using tr As TextReader = New StreamReader(part.GetStream())
					doc = XDocument.Load(tr)

					' Header and Footer extend Container.
					Dim c As Container
					If isHeader Then
						c = New Header(Me, doc.Element(w + "hdr"), part)
					Else
						c = New Footer(Me, doc.Element(w + "ftr"), part)
					End If

					Return c
				End Using
			End If

			' If we got this far something went wrong.
			Return Nothing
		End Function



		Public Function GetSections() As List(Of Section)

			Dim allParas = Paragraphs

			Dim parasInASection = New List(Of Paragraph)()
			Dim sections = New List(Of Section)()

			For Each para In allParas

				Dim sectionInPara = para.Xml.Descendants().FirstOrDefault(Function(s) s.Name.LocalName = "sectPr")

				If sectionInPara Is Nothing Then
					parasInASection.Add(para)
				Else
					parasInASection.Add(para)
					Dim section = New Section(Document, sectionInPara) With {.SectionParagraphs = parasInASection}
					sections.Add(section)
					parasInASection = New List(Of Paragraph)()
				End If

			Next para

			Dim body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
			Dim baseSectionXml As XElement = body.Element(XName.Get("sectPr", DocX.w.NamespaceName))
			Dim baseSection = New Section(Document, baseSectionXml) With {.SectionParagraphs = parasInASection}
			sections.Add(baseSection)

			Return sections
		End Function


		' Get the word\settings.xml part
		Friend settingsPart As PackagePart
		Friend endnotesPart As PackagePart
		Friend footnotesPart As PackagePart
		Friend stylesPart As PackagePart
		Friend stylesWithEffectsPart As PackagePart
		Friend numberingPart As PackagePart
		Friend fontTablePart As PackagePart

		#Region "Internal variables defined foreach DocX object"
		' Object representation of the .docx
		Friend package As Package

		' The mainDocument is loaded into a XDocument object for easy querying and editing
		Friend mainDoc As XDocument
		Friend settings As XDocument
		Friend endnotes As XDocument
		Friend footnotes As XDocument
		Friend styles As XDocument
		Friend stylesWithEffects As XDocument
		Friend numbering As XDocument
		Friend fontTable As XDocument
		Friend header1 As XDocument
		Friend header2 As XDocument
		Friend header3 As XDocument

		' A lookup for the Paragraphs in this document.
		Friend paragraphLookup As New Dictionary(Of Integer, Paragraph)()
		' Every document is stored in a MemoryStream, all edits made to a document are done in memory.
		Friend memoryStream As MemoryStream
		' The filename that this document was loaded from
		Friend filename As String
		' The stream that this document was loaded from
		Friend stream As Stream
		#End Region

		Friend Sub New(ByVal document As DocX, ByVal xml As XElement)
			MyBase.New(document, xml)

		End Sub

		''' <summary>
		''' Returns a list of Images in this document.
		''' </summary>
		''' <example>
		''' Get the unique Id of every Image in this document.
		''' <code>
		''' // Load a document.
		''' DocX document = DocX.Load(@"C:\Example\Test.docx");
		'''
		''' // Loop through each Image in this document.
		''' foreach (Novacode.Image i in document.Images)
		''' {
		'''     // Get the unique Id which identifies this Image.
		'''     string uniqueId = i.Id;
		''' }
		'''
		''' </code>
		''' </example>
		''' <seealso cref="AddImage(string)"/>
		''' <seealso cref="AddImage(Stream)"/>
		''' <seealso cref="Paragraph.Pictures"/>
		''' <seealso cref="Paragraph.InsertPicture"/>
		Public ReadOnly Property Images() As List(Of Image)
			Get
				Dim imageRelationships As PackageRelationshipCollection = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
				If imageRelationships.Count() > 0 Then
					Return (
					    From i In imageRelationships
					    Select New Image(Me, i)).ToList()
				End If

				Return New List(Of Image)()
			End Get
		End Property

		''' <summary>
		''' Returns a list of custom properties in this document.
		''' </summary>
		''' <example>
		''' Method 1: Get the name, type and value of each CustomProperty in this document.
		''' <code>
		''' // Load Example.docx
		''' DocX document = DocX.Load(@"C:\Example\Test.docx");
		'''
		''' /*
		'''  * No two custom properties can have the same name,
		'''  * so a Dictionary is the perfect data structure to store them in.
		'''  * Each custom property can be accessed using its name.
		'''  */
		''' foreach (string name in document.CustomProperties.Keys)
		''' {
		'''     // Grab a custom property using its name.
		'''     CustomProperty cp = document.CustomProperties[name];
		'''
		'''     // Write this custom properties details to Console.
		'''     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
		''' }
		'''
		''' Console.WriteLine("Press any key...");
		'''
		''' // Wait for the user to press a key before closing the Console.
		''' Console.ReadKey();
		''' </code>
		''' </example>
		''' <example>
		''' Method 2: Get the name, type and value of each CustomProperty in this document.
		''' <code>
		''' // Load Example.docx
		''' DocX document = DocX.Load(@"C:\Example\Test.docx");
		''' 
		''' /*
		'''  * No two custom properties can have the same name,
		'''  * so a Dictionary is the perfect data structure to store them in.
		'''  * The values of this Dictionary are CustomProperties.
		'''  */
		''' foreach (CustomProperty cp in document.CustomProperties.Values)
		''' {
		'''     // Write this custom properties details to Console.
		'''     Console.WriteLine(string.Format("Name: '{0}', Value: {1}", cp.Name, cp.Value));
		''' }
		'''
		''' Console.WriteLine("Press any key...");
		'''
		''' // Wait for the user to press a key before closing the Console.
		''' Console.ReadKey();
		''' </code>
		''' </example>
		''' <seealso cref="AddCustomProperty"/>
		Public ReadOnly Property CustomProperties() As Dictionary(Of String, CustomProperty)
			Get
				If package.PartExists(New Uri("/docProps/custom.xml", UriKind.Relative)) Then
					Dim docProps_custom As PackagePart = package.GetPart(New Uri("/docProps/custom.xml", UriKind.Relative))
					Dim customPropDoc As XDocument
					Using tr As TextReader = New StreamReader(docProps_custom.GetStream(FileMode.Open, FileAccess.Read))
						customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace)
					End Using

					' Get all of the custom properties in this document
					Return (
					    From p In customPropDoc.Descendants(XName.Get("property", customPropertiesSchema.NamespaceName))
					    Let Name = p.Attribute(XName.Get("name")).Value
					    Let Type = p.Descendants().Single().Name.LocalName
					    Let Value = p.Descendants().Single().Value
					    Select New CustomProperty(Name, Type, Value)).ToDictionary(Function(p) p.Name, StringComparer.CurrentCultureIgnoreCase)
				End If

				Return New Dictionary(Of String, CustomProperty)()
			End Get
		End Property

		'''<summary>
		''' Returns the list of document core properties with corresponding values.
		'''</summary>
		Public ReadOnly Property CoreProperties() As Dictionary(Of String, String)
			Get
				If package.PartExists(New Uri("/docProps/core.xml", UriKind.Relative)) Then
					Dim docProps_Core As PackagePart = package.GetPart(New Uri("/docProps/core.xml", UriKind.Relative))
					Dim corePropDoc As XDocument
					Using tr As TextReader = New StreamReader(docProps_Core.GetStream(FileMode.Open, FileAccess.Read))
						corePropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace)
					End Using

					' Get all of the core properties in this document
					Return (
					    From docProperty In corePropDoc.Root.Elements()
					    Select New KeyValuePair(Of String, String)(String.Format("{0}:{1}", corePropDoc.Root.GetPrefixOfNamespace(docProperty.Name.Namespace), docProperty.Name.LocalName), docProperty.Value)).ToDictionary(Function(p) p.Key, Function(v) v.Value)
				End If

				Return New Dictionary(Of String, String)()
			End Get
		End Property

		''' <summary>
		''' Get the Text of this document.
		''' </summary>
		''' <example>
		''' Write to Console the Text from this document.
		''' <code>
		''' // Load a document
		''' DocX document = DocX.Load(@"C:\Example\Test.docx");
		'''
		''' // Get the text of this document.
		''' string text = document.Text;
		'''
		''' // Write the text of this document to Console.
		''' Console.Write(text);
		'''
		''' // Wait for the user to press a key before closing the console window.
		''' Console.ReadKey();
		''' </code>
		''' </example>
		Public ReadOnly Property Text() As String
			Get
				Return HelperFunctions.GetText(Xml)
			End Get
		End Property
		 ''' <summary>
		 ''' Get the text of each footnote from this document
		 ''' </summary>
		 Public ReadOnly Property FootnotesText() As IEnumerable(Of String)
			 Get
				For Each footnote As XElement In footnotes.Root.Elements(w + "footnote")
'INSTANT VB TODO TASK: VB does not support iterators and has no equivalent to the C# 'yield' keyword:
					yield Return HelperFunctions.GetText(footnote)
				Next footnote
			 End Get
		 End Property

		''' <summary>
		''' Get the text of each endnote from this document
		''' </summary>
		Public ReadOnly Property EndnotesText() As IEnumerable(Of String)
			Get
				For Each endnote As XElement In endnotes.Root.Elements(w + "endnote")
'INSTANT VB TODO TASK: VB does not support iterators and has no equivalent to the C# 'yield' keyword:
					yield Return HelperFunctions.GetText(endnote)
				Next endnote
			End Get
		End Property



		Friend Function GetCollectiveText(ByVal list As List(Of PackagePart)) As String
			Dim text As String = String.Empty

			For Each hp In list
				Using tr As TextReader = New StreamReader(hp.GetStream())
					Dim d As XDocument = XDocument.Load(tr)

					Dim sb As New StringBuilder()

					' Loop through each text item in this run
					For Each descendant As XElement In d.Descendants()
						Select Case descendant.Name.LocalName
							Case "tab"
								sb.Append(vbTab)
							Case "br"
								sb.Append(vbLf)
							Case "t"
								GoTo CaseLabel1
							Case "delText"
							CaseLabel1:
								sb.Append(descendant.Value)
							Case Else
						End Select
					Next descendant

					text &= vbLf & sb.ToString()
				End Using
			Next hp

			Return text
		End Function

		''' <summary>
		''' Insert the contents of another document at the end of this document. 
		''' </summary>
		''' <param name="remote_document">The document to insert at the end of this document.</param>
		''' <param name="append">If true, document is inserted at the end, otherwise document is inserted at the beginning.</param>
		''' <example>
		''' Create a new document and insert an old document into it.
		''' <code>
		''' // Create a new document.
		''' using (DocX newDocument = DocX.Create(@"NewDocument.docx"))
		''' {
		'''     // Load an old document.
		'''     using (DocX oldDocument = DocX.Load(@"OldDocument.docx"))
		'''     {
		'''         // Insert the old document into the new document.
		'''         newDocument.InsertDocument(oldDocument);
		'''
		'''         // Save the new document.
		'''         newDocument.Save();
		'''     }// Release the old document from memory.
		''' }// Release the new document from memory.
		''' </code>
		''' <remarks>
		''' If the document being inserted contains Images, CustomProperties and or custom styles, these will be correctly inserted into the new document. In the case of Images, new ID's are generated for the Images being inserted to avoid ID conflicts. CustomProperties with the same name will be ignored not replaced.
		''' </remarks>
		''' </example>
		Public Sub InsertDocument(ByVal remote_document As DocX, Optional ByVal append As Boolean = True)
			' We don't want to effect the origional XDocument, so create a new one from the old one.
			Dim remote_mainDoc As New XDocument(remote_document.mainDoc)

			Dim remote_footnotes As XDocument = Nothing
			If remote_document.footnotes IsNot Nothing Then
				remote_footnotes = New XDocument(remote_document.footnotes)
			End If

			Dim remote_endnotes As XDocument = Nothing
			If remote_document.endnotes IsNot Nothing Then
				remote_endnotes = New XDocument(remote_document.endnotes)
			End If

			' Remove all header and footer references.
			remote_mainDoc.Descendants(XName.Get("headerReference", DocX.w.NamespaceName)).Remove()
			remote_mainDoc.Descendants(XName.Get("footerReference", DocX.w.NamespaceName)).Remove()

			' Get the body of the remote document.
			Dim remote_body As XElement = remote_mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))

			' Every file that is missing from the local document will have to be copied, every file that already exists will have to be merged.
			Dim ppc As PackagePartCollection = remote_document.package.GetParts()

			Dim ignoreContentTypes As New List(Of String)() From {"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", "application/vnd.openxmlformats-package.core-properties+xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", "application/vnd.openxmlformats-package.relationships+xml"}

			Dim imageContentTypes As New List(Of String)() From {"image/jpeg", "image/jpg", "image/png", "image/bmp", "image/gif", "image/tiff", "image/icon", "image/pcx", "image/emf", "image/wmf"}
			' Check if each PackagePart pp exists in this document.
			For Each remote_pp As PackagePart In ppc
				If ignoreContentTypes.Contains(remote_pp.ContentType) OrElse imageContentTypes.Contains(remote_pp.ContentType) Then
					Continue For
				End If

				' If this external PackagePart already exits then we must merge them.
				If package.PartExists(remote_pp.Uri) Then
					Dim local_pp As PackagePart = package.GetPart(remote_pp.Uri)
					Select Case remote_pp.ContentType
						Case "application/vnd.openxmlformats-officedocument.custom-properties+xml"
							merge_customs(remote_pp, local_pp, remote_mainDoc)

						' Merge footnotes (and endnotes) before merging styles, then set the remote_footnotes to the just updated footnotes
						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
							merge_footnotes(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes)
							remote_footnotes = footnotes

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
							merge_endnotes(remote_pp, local_pp, remote_mainDoc, remote_document, remote_endnotes)
							remote_endnotes = endnotes

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
							merge_styles(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes)

						' Merge styles after merging the footnotes, so the changes will be applied to the correct document/footnotes
						Case "application/vnd.ms-word.stylesWithEffects+xml"
							merge_styles(remote_pp, local_pp, remote_mainDoc, remote_document, remote_footnotes, remote_endnotes)

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
							merge_fonts(remote_pp, local_pp, remote_mainDoc, remote_document)

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
							merge_numbering(remote_pp, local_pp, remote_mainDoc, remote_document)

						Case Else
					End Select

				' If this external PackagePart does not exits in the internal document then we can simply copy it.
				Else
					Dim packagePart = clonePackagePart(remote_pp)
					Select Case remote_pp.ContentType
						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
							endnotesPart = packagePart
							endnotes = remote_endnotes

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
							footnotesPart = packagePart
							footnotes = remote_footnotes

						Case "application/vnd.openxmlformats-officedocument.custom-properties+xml"

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
							stylesPart = packagePart
							Using tr As TextReader = New StreamReader(stylesPart.GetStream())
								styles = XDocument.Load(tr)
							End Using

						Case "application/vnd.ms-word.stylesWithEffects+xml"
							stylesWithEffectsPart = packagePart
							Using tr As TextReader = New StreamReader(stylesWithEffectsPart.GetStream())
								stylesWithEffects = XDocument.Load(tr)
							End Using

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"
							fontTablePart = packagePart
							Using tr As TextReader = New StreamReader(fontTablePart.GetStream())
								fontTable = XDocument.Load(tr)
							End Using

						Case "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
							numberingPart = packagePart
							Using tr As TextReader = New StreamReader(numberingPart.GetStream())
								numbering = XDocument.Load(tr)
							End Using

					End Select

					clonePackageRelationship(remote_document, remote_pp, remote_mainDoc)
				End If
			Next remote_pp

			For Each hyperlink_rel In remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
				Dim old_rel_Id = hyperlink_rel.Id
				Dim new_rel_Id = mainPart.CreateRelationship(hyperlink_rel.TargetUri, hyperlink_rel.TargetMode, hyperlink_rel.RelationshipType).Id
				Dim hyperlink_refs = remote_mainDoc.Descendants(XName.Get("hyperlink", DocX.w.NamespaceName))
				For Each hyperlink_ref In hyperlink_refs
					Dim a0 As XAttribute = hyperlink_ref.Attribute(XName.Get("id", DocX.r.NamespaceName))
					If a0 IsNot Nothing AndAlso a0.Value = old_rel_Id Then
						a0.SetValue(new_rel_Id)
					End If
				Next hyperlink_ref
			Next hyperlink_rel

			'//ole object links
			For Each oleObject_rel In remote_document.mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject")
				Dim old_rel_Id = oleObject_rel.Id
				Dim new_rel_Id = mainPart.CreateRelationship(oleObject_rel.TargetUri, oleObject_rel.TargetMode, oleObject_rel.RelationshipType).Id
				Dim oleObject_refs = remote_mainDoc.Descendants(XName.Get("OLEObject", "urn:schemas-microsoft-com:office:office"))
				For Each oleObject_ref In oleObject_refs
					Dim a0 As XAttribute = oleObject_ref.Attribute(XName.Get("id", DocX.r.NamespaceName))
					If a0 IsNot Nothing AndAlso a0.Value = old_rel_Id Then
						a0.SetValue(new_rel_Id)
					End If
				Next oleObject_ref
			Next oleObject_rel


			For Each remote_pp As PackagePart In ppc
				If imageContentTypes.Contains(remote_pp.ContentType) Then
					merge_images(remote_pp, remote_document, remote_mainDoc, remote_pp.ContentType)
				End If
			Next remote_pp

			Dim id As Integer = 0
			Dim local_docPrs = mainDoc.Root.Descendants(XName.Get("docPr", DocX.wp.NamespaceName))
			For Each local_docPr In local_docPrs
				Dim a_id As XAttribute = local_docPr.Attribute(XName.Get("id"))
				Dim a_id_value As Integer
				If a_id IsNot Nothing AndAlso Integer.TryParse(a_id.Value, a_id_value) Then
					If a_id_value > id Then
						id = a_id_value
					End If
				End If
			Next local_docPr
			id += 1

			' docPr must be sequential
			Dim docPrs = remote_body.Descendants(XName.Get("docPr", DocX.wp.NamespaceName))
			For Each docPr In docPrs
				docPr.SetAttributeValue(XName.Get("id"), id)
				id += 1
			Next docPr

			' Add the remote documents contents to this document.
			Dim local_body As XElement = mainDoc.Root.Element(XName.Get("body", DocX.w.NamespaceName))
			If append Then
				local_body.Add(remote_body.Elements())
			Else
				local_body.AddFirst(remote_body.Elements())
			End If

			' Copy any missing root attributes to the local document.
			For Each a As XAttribute In remote_mainDoc.Root.Attributes()
				If mainDoc.Root.Attribute(a.Name) Is Nothing Then
					mainDoc.Root.SetAttributeValue(a.Name, a.Value)
				End If
			Next a

		End Sub

		Private Sub merge_images(ByVal remote_pp As PackagePart, ByVal remote_document As DocX, ByVal remote_mainDoc As XDocument, ByVal contentType As String)
			' Before doing any other work, check to see if this image is actually referenced in the document.
			' In my testing I have found cases of Images inside documents that are not referenced
			Dim remote_rel = remote_document.mainPart.GetRelationships().Where(Function(r) r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault()
			If remote_rel Is Nothing Then
				remote_rel = remote_document.mainPart.GetRelationships().Where(Function(r) r.TargetUri.OriginalString.Equals(remote_pp.Uri.OriginalString)).FirstOrDefault()
				If remote_rel Is Nothing Then
					Return
				End If
			End If
			Dim remote_Id As String = remote_rel.Id

			Dim remote_hash As String = ComputeMD5HashString(remote_pp.GetStream())
			Dim image_parts = package.GetParts().Where(Function(pp) pp.ContentType.Equals(contentType))

			Dim found As Boolean = False
			For Each part In image_parts
				Dim local_hash As String = ComputeMD5HashString(part.GetStream())
				If local_hash.Equals(remote_hash) Then
					' This image already exists in this document.
					found = True

					Dim local_rel = mainPart.GetRelationships().Where(Function(r) r.TargetUri.OriginalString.Equals(part.Uri.OriginalString.Replace("/word/", ""))).FirstOrDefault()
					If local_rel Is Nothing Then
						local_rel = mainPart.GetRelationships().Where(Function(r) r.TargetUri.OriginalString.Equals(part.Uri.OriginalString)).FirstOrDefault()
					End If
					If local_rel IsNot Nothing Then
						Dim new_Id As String = local_rel.Id

						' Replace all instances of remote_Id in the local document with local_Id
						Dim elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName))
						For Each elem In elems
							Dim embed As XAttribute = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName))
							If embed IsNot Nothing AndAlso embed.Value = remote_Id Then
								embed.SetValue(new_Id)
							End If
						Next elem

						' Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
						Dim v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName))
						For Each elem In v_elems
							Dim id As XAttribute = elem.Attribute(XName.Get("id", DocX.r.NamespaceName))
							If id IsNot Nothing AndAlso id.Value = remote_Id Then
								id.SetValue(new_Id)
							End If
						Next elem
					End If

					Exit For
				End If
			Next part

			' This image does not exist in this document.
			If Not found Then
				Dim new_uri As String = remote_pp.Uri.OriginalString
				new_uri = new_uri.Remove(new_uri.LastIndexOf("/"))
				'new_uri = new_uri.Replace("word/", "");
				new_uri &= "/" & Guid.NewGuid().ToString() & contentType.Replace("image/", ".")
				If Not new_uri.StartsWith("/") Then
					new_uri = "/" & new_uri
				End If

				Dim new_pp As PackagePart = package.CreatePart(New Uri(new_uri, UriKind.Relative), remote_pp.ContentType, CompressionOption.Normal)

				Using s_read As Stream = remote_pp.GetStream()
					Using s_write As Stream = new_pp.GetStream(FileMode.Create)
						Dim buffer(32767) As Byte
						Dim read As Integer
						read = s_read.Read(buffer, 0, buffer.Length)
						Do While read > 0
							s_write.Write(buffer, 0, read)
							read = s_read.Read(buffer, 0, buffer.Length)
						Loop
					End Using
				End Using

				Dim pr As PackageRelationship = mainPart.CreateRelationship(New Uri(new_uri, UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

				Dim new_Id As String = pr.Id

				'Check if the remote relationship id is a default rId from Word
				Dim defRelId As Match = Regex.Match(remote_Id, "rId\d+", RegexOptions.IgnoreCase)

				' Replace all instances of remote_Id in the local document with local_Id
				Dim elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName))
				For Each elem In elems
					Dim embed As XAttribute = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName))
					If embed IsNot Nothing AndAlso embed.Value = remote_Id Then
						embed.SetValue(new_Id)
					End If
				Next elem

				If Not defRelId.Success Then
					   ' Replace all instances of remote_Id in the local document with local_Id
					Dim elems_local = mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName))
					For Each elem In elems_local
						Dim embed As XAttribute = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName))
						If embed IsNot Nothing AndAlso embed.Value = remote_Id Then
							embed.SetValue(new_Id)
						End If
					Next elem


					' Replace all instances of remote_Id in the local document with local_Id
					Dim v_elems_local = mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName))
					For Each elem In v_elems_local
						Dim id As XAttribute = elem.Attribute(XName.Get("id", DocX.r.NamespaceName))
						If id IsNot Nothing AndAlso id.Value = remote_Id Then
							id.SetValue(new_Id)
						End If
					Next elem
				End If


				' Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
				Dim v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName))
				For Each elem In v_elems
					Dim id As XAttribute = elem.Attribute(XName.Get("id", DocX.r.NamespaceName))
					If id IsNot Nothing AndAlso id.Value = remote_Id Then
						id.SetValue(new_Id)
					End If
				Next elem
			End If
		End Sub

		Private Function ComputeMD5HashString(ByVal stream As Stream) As String
			Dim md5 As MD5 = MD5.Create()
			Dim hash() As Byte = md5.ComputeHash(stream)
			Dim sb As New StringBuilder()
			For i As Integer = 0 To hash.Length - 1
				sb.Append(hash(i).ToString("X2"))
			Next i
			Return sb.ToString()
		End Function

		Private Sub merge_endnotes(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument, ByVal remote As DocX, ByVal remote_endnotes As XDocument)
			Dim ids As IEnumerable(Of Integer) = (
			    From d In endnotes.Root.Descendants()
			    Where d.Name.LocalName = "endnote"
			    Select Integer.Parse(d.Attribute(XName.Get("id", DocX.w.NamespaceName)).Value))

			Dim max_id As Integer = ids.Max() + 1
			Dim endnoteReferences = remote_mainDoc.Descendants(XName.Get("endnoteReference", DocX.w.NamespaceName))

			For Each endnote In remote_endnotes.Root.Elements().OrderBy(Function(fr) fr.Attribute(XName.Get("id", DocX.r.NamespaceName))).Reverse()
				Dim id As XAttribute = endnote.Attribute(XName.Get("id", DocX.w.NamespaceName))
				Dim i As Integer
				If id IsNot Nothing AndAlso Integer.TryParse(id.Value, i) Then
					If i > 0 Then
						For Each endnoteRef In endnoteReferences
							Dim a As XAttribute = endnoteRef.Attribute(XName.Get("id", DocX.w.NamespaceName))
							If a IsNot Nothing AndAlso Integer.Parse(a.Value).Equals(i) Then
								a.SetValue(max_id)
							End If
						Next endnoteRef

						' We care about copying this footnote.
						endnote.SetAttributeValue(XName.Get("id", DocX.w.NamespaceName), max_id)
						endnotes.Root.Add(endnote)
						max_id += 1
					End If
				End If
			Next endnote
		End Sub

		Private Sub merge_footnotes(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument, ByVal remote As DocX, ByVal remote_footnotes As XDocument)
			Dim ids As IEnumerable(Of Integer) = (
			    From d In footnotes.Root.Descendants()
			    Where d.Name.LocalName = "footnote"
			    Select Integer.Parse(d.Attribute(XName.Get("id", DocX.w.NamespaceName)).Value))

			Dim max_id As Integer = ids.Max() + 1
			Dim footnoteReferences = remote_mainDoc.Descendants(XName.Get("footnoteReference", DocX.w.NamespaceName))

			For Each footnote In remote_footnotes.Root.Elements().OrderBy(Function(fr) fr.Attribute(XName.Get("id", DocX.r.NamespaceName))).Reverse()
				Dim id As XAttribute = footnote.Attribute(XName.Get("id", DocX.w.NamespaceName))
				Dim i As Integer
				If id IsNot Nothing AndAlso Integer.TryParse(id.Value, i) Then
					If i > 0 Then
						For Each footnoteRef In footnoteReferences
							Dim a As XAttribute = footnoteRef.Attribute(XName.Get("id", DocX.w.NamespaceName))
							If a IsNot Nothing AndAlso Integer.Parse(a.Value).Equals(i) Then
								a.SetValue(max_id)
							End If
						Next footnoteRef

						' We care about copying this footnote.
						footnote.SetAttributeValue(XName.Get("id", DocX.w.NamespaceName), max_id)
						footnotes.Root.Add(footnote)
						max_id += 1
					End If
				End If
			Next footnote
		End Sub

		Private Sub merge_customs(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument)
			' Get the remote documents custom.xml file.
			Dim remote_custom_document As XDocument
			Using tr As TextReader = New StreamReader(remote_pp.GetStream())
				remote_custom_document = XDocument.Load(tr)
			End Using

			' Get the local documents custom.xml file.
			Dim local_custom_document As XDocument
			Using tr As TextReader = New StreamReader(local_pp.GetStream())
				local_custom_document = XDocument.Load(tr)
			End Using

			Dim pids As IEnumerable(Of Integer) = (
			    From d In remote_custom_document.Root.Descendants()
			    Where d.Name.LocalName = "property"
			    Select Integer.Parse(d.Attribute(XName.Get("pid")).Value))

			Dim pid As Integer = pids.Max() + 1

			For Each remote_property As XElement In remote_custom_document.Root.Elements()
				Dim found As Boolean = False
				For Each local_property As XElement In local_custom_document.Root.Elements()
					Dim remote_property_name As XAttribute = remote_property.Attribute(XName.Get("name"))
					Dim local_property_name As XAttribute = local_property.Attribute(XName.Get("name"))

					If remote_property IsNot Nothing AndAlso local_property_name IsNot Nothing AndAlso remote_property_name.Value.Equals(local_property_name.Value) Then
						found = True
					End If
				Next local_property

				If Not found Then
					remote_property.SetAttributeValue(XName.Get("pid"), pid)
					local_custom_document.Root.Add(remote_property)

					pid += 1
				End If
			Next remote_property

			' Save the modified local custom styles.xml file.
			Using tw As TextWriter = New StreamWriter(local_pp.GetStream(FileMode.Create, FileAccess.Write))
				local_custom_document.Save(tw, SaveOptions.None)
			End Using
		End Sub

		Private Sub merge_numbering(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument, ByVal remote As DocX)
			' Add each remote numbering to this document.
			Dim remote_abstractNums As IEnumerable(Of XElement) = remote.numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName))
			Dim guidd As Integer = 0
			For Each an In remote_abstractNums
				Dim a As XAttribute = an.Attribute(XName.Get("abstractNumId", DocX.w.NamespaceName))
				If a IsNot Nothing Then
					Dim i As Integer
					If Integer.TryParse(a.Value, i) Then
						If i > guidd Then
							guidd = i
						End If
					End If
				End If
			Next an
			guidd += 1

			Dim remote_nums As IEnumerable(Of XElement) = remote.numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName))
			Dim guidd2 As Integer = 0
			For Each an In remote_nums
				Dim a As XAttribute = an.Attribute(XName.Get("numId", DocX.w.NamespaceName))
				If a IsNot Nothing Then
					Dim i As Integer
					If Integer.TryParse(a.Value, i) Then
						If i > guidd2 Then
							guidd2 = i
						End If
					End If
				End If
			Next an
			guidd2 += 1

			For Each remote_abstractNum As XElement In remote_abstractNums
				Dim abstractNumId As XAttribute = remote_abstractNum.Attribute(XName.Get("abstractNumId", DocX.w.NamespaceName))
				If abstractNumId IsNot Nothing Then
					Dim abstractNumIdValue As String = abstractNumId.Value
					abstractNumId.SetValue(guidd)

					For Each remote_num As XElement In remote_nums
						Dim numIds = remote_mainDoc.Descendants(XName.Get("numId", DocX.w.NamespaceName))
						For Each numId In numIds
							Dim attr As XAttribute = numId.Attribute(XName.Get("val", DocX.w.NamespaceName))
							If attr IsNot Nothing AndAlso attr.Value.Equals(remote_num.Attribute(XName.Get("numId", DocX.w.NamespaceName)).Value) Then
								attr.SetValue(guidd2)
							End If

						Next numId
						remote_num.SetAttributeValue(XName.Get("numId", DocX.w.NamespaceName), guidd2)

						Dim e As XElement = remote_num.Element(XName.Get("abstractNumId", DocX.w.NamespaceName))
						If e IsNot Nothing Then
							Dim a2 As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
							If a2 IsNot Nothing AndAlso a2.Value.Equals(abstractNumIdValue) Then
								a2.SetValue(guidd)
							End If
						End If

						guidd2 += 1
					Next remote_num
				End If

				guidd += 1
			Next remote_abstractNum

			' Checking whether there were more than 0 elements, helped me get rid of exceptions thrown while using InsertDocument
			If numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName)).Count() > 0 Then
				numbering.Root.Elements(XName.Get("abstractNum", DocX.w.NamespaceName)).Last().AddAfterSelf(remote_abstractNums)
			End If

			If numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName)).Count() > 0 Then
				numbering.Root.Elements(XName.Get("num", DocX.w.NamespaceName)).Last().AddAfterSelf(remote_nums)
			End If
		End Sub

		Private Sub merge_fonts(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument, ByVal remote As DocX)
			' Add each remote font to this document.
			Dim remote_fonts As IEnumerable(Of XElement) = remote.fontTable.Root.Elements(XName.Get("font", DocX.w.NamespaceName))
			Dim local_fonts As IEnumerable(Of XElement) = fontTable.Root.Elements(XName.Get("font", DocX.w.NamespaceName))

			For Each remote_font As XElement In remote_fonts
				Dim flag_addFont As Boolean = True
				For Each local_font As XElement In local_fonts
					If local_font.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value = remote_font.Attribute(XName.Get("name", DocX.w.NamespaceName)).Value Then
						flag_addFont = False
						Exit For
					End If
				Next local_font

				If flag_addFont Then
					fontTable.Root.Add(remote_font)
				End If
			Next remote_font
		End Sub

		Private Sub merge_styles(ByVal remote_pp As PackagePart, ByVal local_pp As PackagePart, ByVal remote_mainDoc As XDocument, ByVal remote As DocX, ByVal remote_footnotes As XDocument, ByVal remote_endnotes As XDocument)
			Dim local_styles As New Dictionary(Of String, String)()
			For Each local_style As XElement In styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName))
				Dim temp As New XElement(local_style)
				Dim styleId As XAttribute = temp.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
				Dim value As String = styleId.Value
				styleId.Remove()
				Dim key As String = Regex.Replace(temp.ToString(), "\s+", "")
				If Not local_styles.ContainsKey(key) Then
					local_styles.Add(key, value)
				End If
			Next local_style

			' Add each remote style to this document.
			Dim remote_styles As IEnumerable(Of XElement) = remote.styles.Root.Elements(XName.Get("style", DocX.w.NamespaceName))
			For Each remote_style As XElement In remote_styles
				Dim temp As New XElement(remote_style)
				Dim styleId As XAttribute = temp.Attribute(XName.Get("styleId", DocX.w.NamespaceName))
				Dim value As String = styleId.Value
				styleId.Remove()
				Dim key As String = Regex.Replace(temp.ToString(), "\s+", "")
				Dim guuid As String

				' Check to see if the local document already contains the remote style.
				If local_styles.ContainsKey(key) Then
					Dim local_value As String
					local_styles.TryGetValue(key, local_value)

					' If the styleIds are the same then nothing needs to be done.
					If local_value = value Then
						Continue For

					' All we need to do is update the styleId.
					Else
						guuid = local_value
					End If
				Else
					guuid = Guid.NewGuid().ToString()
					' Set the styleId in the remote_style to this new Guid
					' [Fixed the issue that my document referred to a new Guid while my styles still had the old value ("Titel")]
					remote_style.SetAttributeValue(XName.Get("styleId", DocX.w.NamespaceName), guuid)
				End If

				For Each e As XElement In remote_mainDoc.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName))
					Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
					If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
						e_styleId.SetValue(guuid)
					End If
				Next e

				For Each e As XElement In remote_mainDoc.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName))
					Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
					If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
						e_styleId.SetValue(guuid)
					End If
				Next e

				For Each e As XElement In remote_mainDoc.Root.Descendants(XName.Get("tblStyle", DocX.w.NamespaceName))
					Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
					If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
						e_styleId.SetValue(guuid)
					End If
				Next e

				If remote_endnotes IsNot Nothing Then
					For Each e As XElement In remote_endnotes.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName))
						Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
						If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
							e_styleId.SetValue(guuid)
						End If
					Next e

					For Each e As XElement In remote_endnotes.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName))
						Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
						If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
							e_styleId.SetValue(guuid)
						End If
					Next e
				End If

				If remote_footnotes IsNot Nothing Then
					For Each e As XElement In remote_footnotes.Root.Descendants(XName.Get("rStyle", DocX.w.NamespaceName))
						Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
						If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
							e_styleId.SetValue(guuid)
						End If
					Next e

					For Each e As XElement In remote_footnotes.Root.Descendants(XName.Get("pStyle", DocX.w.NamespaceName))
						Dim e_styleId As XAttribute = e.Attribute(XName.Get("val", DocX.w.NamespaceName))
						If e_styleId IsNot Nothing AndAlso e_styleId.Value.Equals(styleId.Value) Then
							e_styleId.SetValue(guuid)
						End If
					Next e
				End If

				' Make sure they don't clash by using a uuid.
				styleId.SetValue(guuid)
				styles.Root.Add(remote_style)
			Next remote_style
		End Sub

		Protected Sub clonePackageRelationship(ByVal remote_document As DocX, ByVal pp As PackagePart, ByVal remote_mainDoc As XDocument)
			Dim url As String = pp.Uri.OriginalString.Replace("/", "")
			Dim remote_rels = remote_document.mainPart.GetRelationships()
			For Each remote_rel In remote_rels
				If url.Equals("word" & remote_rel.TargetUri.OriginalString.Replace("/", "")) Then
					Dim remote_Id As String = remote_rel.Id
					Dim local_Id As String = mainPart.CreateRelationship(remote_rel.TargetUri, remote_rel.TargetMode, remote_rel.RelationshipType).Id

					' Replace all instances of remote_Id in the local document with local_Id
					Dim elems = remote_mainDoc.Descendants(XName.Get("blip", DocX.a.NamespaceName))
					For Each elem In elems
						Dim embed As XAttribute = elem.Attribute(XName.Get("embed", DocX.r.NamespaceName))
						If embed IsNot Nothing AndAlso embed.Value = remote_Id Then
							embed.SetValue(local_Id)
						End If
					Next elem

					 ' Replace all instances of remote_Id in the local document with local_Id (for shapes as well)
					Dim v_elems = remote_mainDoc.Descendants(XName.Get("imagedata", DocX.v.NamespaceName))
					For Each elem In v_elems
						Dim id As XAttribute = elem.Attribute(XName.Get("id", DocX.r.NamespaceName))
						If id IsNot Nothing AndAlso id.Value = remote_Id Then
							id.SetValue(local_Id)
						End If
					Next elem
					Exit For
				End If
			Next remote_rel
		End Sub

		Protected Function clonePackagePart(ByVal pp As PackagePart) As PackagePart
			Dim new_pp As PackagePart = package.CreatePart(pp.Uri, pp.ContentType, CompressionOption.Normal)

			Using s_read As Stream = pp.GetStream()
				Using s_write As Stream = new_pp.GetStream(FileMode.Create)
					Dim buffer(32767) As Byte
					Dim read As Integer
					read = s_read.Read(buffer, 0, buffer.Length)
					Do While read > 0
						s_write.Write(buffer, 0, read)
						read = s_read.Read(buffer, 0, buffer.Length)
					Loop
				End Using
			End Using

			Return new_pp
		End Function

		Protected Function GetMD5HashFromStream(ByVal stream As Stream) As String
			Dim md5 As MD5 = New MD5CryptoServiceProvider()
			Dim retVal() As Byte = md5.ComputeHash(stream)

			Dim sb As New StringBuilder()
			For i As Integer = 0 To retVal.Length - 1
				sb.Append(retVal(i).ToString("x2"))
			Next i
			Return sb.ToString()
		End Function

		''' <summary>
		''' Insert a new Table at the end of this document.
		''' </summary>
		''' <param name="columnCount">The number of columns to create.</param>
		''' <param name="rowCount">The number of rows to create.</param>
		''' <returns>A new Table.</returns>
		''' <example>
		''' Insert a new Table with 2 columns and 3 rows, at the end of a document.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"C:\Example\Test.docx"))
		''' {
		'''     // Create a new Table with 2 columns and 3 rows.
		'''     Table newTable = document.InsertTable(2, 3);
		'''
		'''     // Set the design of this Table.
		'''     newTable.Design = TableDesign.LightShadingAccent2;
		'''
		'''     // Set the column names.
		'''     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
		'''     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
		'''
		'''     // Fill row 1
		'''     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
		'''     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
		'''
		'''     // Fill row 2
		'''     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
		'''     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
		'''
		'''     // Save all changes made to document b.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Shadows Function InsertTable(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			If rowCount < 1 OrElse columnCount < 1 Then
				Throw New ArgumentOutOfRangeException("Row and Column count must be greater than zero.")
			End If

			Dim t As Table = MyBase.InsertTable(rowCount, columnCount)
			t.mainPart = mainPart
			Return t
		End Function

		Public Function AddTable(ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			If rowCount < 1 OrElse columnCount < 1 Then
				Throw New ArgumentOutOfRangeException("Row and Column count must be greater than zero.")
			End If

			Dim t As New Table(Me, HelperFunctions.CreateTable(rowCount, columnCount))
			t.mainPart = mainPart
			Return t
		End Function

		''' <summary>
		''' Create a new list with a list item.
		''' </summary>
		''' <param name="listText">The text of the first element in the created list.</param>
		''' <param name="level">The indentation level of the element in the list.</param>
		''' <param name="listType">The type of list to be created: Bulleted or Numbered.</param>
		''' <param name="startNumber">The number start number for the list. </param>
		''' <param name="trackChanges">Enable change tracking</param>
		''' <param name="continueNumbering">Set to true if you want to continue numbering from the previous numbered list</param>
		''' <returns>
		''' The created List. Call AddListItem(...) to add more elements to the list.
		''' Write the list to the Document with InsertList(...) once the list has all the desired 
		''' elements, otherwise the list will not be included in the working Document.
		''' </returns>
		Public Function AddList(Optional ByVal listText As String = Nothing, Optional ByVal level As Integer = 0, Optional ByVal listType As ListItemType = ListItemType.Numbered, Optional ByVal startNumber? As Integer = Nothing, Optional ByVal trackChanges As Boolean = False, Optional ByVal continueNumbering As Boolean = False) As List
			Return AddListItem(New List(Me, Nothing), listText, level, listType, startNumber, trackChanges, continueNumbering)
		End Function

		''' <summary>
		''' Add a list item to an already existing list.
		''' </summary>
		''' <param name="list">The list to add the new list item to.</param>
		''' <param name="listText">The run text that should be in the new list item.</param>
		''' <param name="level">The indentation level of the new list element.</param>
		''' <param name="startNumber">The number start number for the list. </param>
		''' <param name="trackChanges">Enable change tracking</param>
		''' <param name="listType">Numbered or Bulleted list type. </param>
		''' /// <param name="continueNumbering">Set to true if you want to continue numbering from the previous numbered list</param>
		''' <returns>
		''' The created List. Call AddListItem(...) to add more elements to the list.
		''' Write the list to the Document with InsertList(...) once the list has all the desired 
		''' elements, otherwise the list will not be included in the working Document.
		''' </returns>
		Public Function AddListItem(ByVal list As List, ByVal listText As String, Optional ByVal level As Integer = 0, Optional ByVal listType As ListItemType = ListItemType.Numbered, Optional ByVal startNumber? As Integer = Nothing, Optional ByVal trackChanges As Boolean = False, Optional ByVal continueNumbering As Boolean = False) As List
			If startNumber.HasValue AndAlso continueNumbering Then
				Throw New InvalidOperationException("Cannot specify a start number and at the same time continue numbering from another list")
			End If
			Dim listToReturn = HelperFunctions.CreateItemInList(list, listText, level, listType, startNumber, trackChanges, continueNumbering)
			Dim lastItem = listToReturn.Items.LastOrDefault()
			If lastItem IsNot Nothing Then
				lastItem.PackagePart = mainPart
			End If
			Return listToReturn

		End Function

		''' <summary>
		''' Insert list into the document.
		''' </summary>
		''' <param name="list">The list to insert into the document.</param>
		''' <returns>The list that was inserted into the document.</returns>
		Public Shadows Function InsertList(ByVal list As List) As List
			MyBase.InsertList(list)
			Return list
		End Function
		Public Shadows Function InsertList(ByVal list As List, ByVal fontFamily As FontFamily, ByVal fontSize As Double) As List
			MyBase.InsertList(list, fontFamily, fontSize)
			Return list
		End Function
		Public Shadows Function InsertList(ByVal list As List, ByVal fontSize As Double) As List
			MyBase.InsertList(list, fontSize)
			Return list
		End Function

		''' <summary>
		''' Insert a list at an index location in the document.
		''' </summary>
		''' <param name="index">Index in document to insert the list.</param>
		''' <param name="list">The list that was inserted into the document.</param>
		''' <returns></returns>
		Public Shadows Function InsertList(ByVal index As Integer, ByVal list As List) As List
			MyBase.InsertList(index, list)
			Return list
		End Function

		Friend Function AddStylesForList() As XDocument
			Dim wordStylesUri = New Uri("/word/styles.xml", UriKind.Relative)

			' If the internal document contains no /word/styles.xml create one.
			If Not package.PartExists(wordStylesUri) Then
				HelperFunctions.AddDefaultStylesXml(package)
			End If

			' Load the styles.xml into memory.
			Dim wordStyles As XDocument
			Using tr As TextReader = New StreamReader(package.GetPart(wordStylesUri).GetStream())
				wordStyles = XDocument.Load(tr)
			End Using

			Dim listStyleExists As Boolean = (
			    From s In wordStyles.Element(w + "styles").Elements()
			    Let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
			    Where (styleId IsNot Nothing AndAlso styleId.Value = "ListParagraph")
			    Select s).Any()

			If Not listStyleExists Then
				Dim style = New XElement (w + "style", New XAttribute(w + "type", "paragraph"), New XAttribute(w + "styleId", "ListParagraph"), New XElement(w + "name", New XAttribute(w + "val", "List Paragraph")), New XElement(w + "basedOn", New XAttribute(w + "val", "Normal")), New XElement(w + "uiPriority", New XAttribute(w + "val", "34")), New XElement(w + "qformat"), New XElement(w + "rsid", New XAttribute(w + "val", "00832EE1")), New XElement (w + "rPr", New XElement(w + "ind", New XAttribute(w + "left", "720")), New XElement (w + "contextualSpacing")))
				wordStyles.Element(w + "styles").Add(style)

				' Save the styles document.
				Using tw As TextWriter = New StreamWriter(package.GetPart(wordStylesUri).GetStream())
					wordStyles.Save(tw)
				End Using
			End If

			Return wordStyles
		End Function

		''' <summary>
		''' Insert a Table into this document. The Table's source can be a completely different document.
		''' </summary>
		''' <param name="t">The Table to insert.</param>
		''' <param name="index">The index to insert this Table at.</param>
		''' <returns>The Table now associated with this document.</returns>
		''' <example>
		''' Extract a Table from document a and insert it into document b, at index 10.
		''' <code>
		''' // Place holder for a Table.
		''' Table t;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
		''' {
		'''     // Get the first Table from this document.
		'''     t = documentA.Tables[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
		''' {
		'''     /* 
		'''      * Insert the Table that was extracted from document a, into document b. 
		'''      * This creates a new Table that is now associated with document b.
		'''      */
		'''     Table newTable = documentB.InsertTable(10, t);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Shadows Function InsertTable(ByVal index As Integer, ByVal t As Table) As Table
			Dim t2 As Table = MyBase.InsertTable(index, t)
			t2.mainPart = mainPart
			Return t2
		End Function

		''' <summary>
		''' Insert a Table into this document. The Table's source can be a completely different document.
		''' </summary>
		''' <param name="t">The Table to insert.</param>
		''' <returns>The Table now associated with this document.</returns>
		''' <example>
		''' Extract a Table from document a and insert it at the end of document b.
		''' <code>
		''' // Place holder for a Table.
		''' Table t;
		'''
		''' // Load document a.
		''' using (DocX documentA = DocX.Load(@"C:\Example\a.docx"))
		''' {
		'''     // Get the first Table from this document.
		'''     t = documentA.Tables[0];
		''' }
		'''
		''' // Load document b.
		''' using (DocX documentB = DocX.Load(@"C:\Example\b.docx"))
		''' {
		'''     /* 
		'''      * Insert the Table that was extracted from document a, into document b. 
		'''      * This creates a new Table that is now associated with document b.
		'''      */
		'''     Table newTable = documentB.InsertTable(t);
		'''
		'''     // Save all changes made to document b.
		'''     documentB.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Shadows Function InsertTable(ByVal t As Table) As Table
			t = MyBase.InsertTable(t)
			t.mainPart = mainPart
			Return t
		End Function

		''' <summary>
		''' Insert a new Table at the end of this document.
		''' </summary>
		''' <param name="columnCount">The number of columns to create.</param>
		''' <param name="rowCount">The number of rows to create.</param>
		''' <param name="index">The index to insert this Table at.</param>
		''' <returns>A new Table.</returns>
		''' <example>
		''' Insert a new Table with 2 columns and 3 rows, at index 37 in this document.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Create a new Table with 3 rows and 2 columns. Insert this Table at index 37.
		'''     Table newTable = document.InsertTable(37, 3, 2);
		'''
		'''     // Set the design of this Table.
		'''     newTable.Design = TableDesign.LightShadingAccent3;
		'''
		'''     // Set the column names.
		'''     newTable.Rows[0].Cells[0].Paragraph.InsertText("Ice Cream", false);
		'''     newTable.Rows[0].Cells[1].Paragraph.InsertText("Price", false);
		'''
		'''     // Fill row 1
		'''     newTable.Rows[1].Cells[0].Paragraph.InsertText("Chocolate", false);
		'''     newTable.Rows[1].Cells[1].Paragraph.InsertText("€3:50", false);
		'''
		'''     // Fill row 2
		'''     newTable.Rows[2].Cells[0].Paragraph.InsertText("Vanilla", false);
		'''     newTable.Rows[2].Cells[1].Paragraph.InsertText("€3:00", false);
		'''
		'''     // Save all changes made to document b.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Shadows Function InsertTable(ByVal index As Integer, ByVal rowCount As Integer, ByVal columnCount As Integer) As Table
			If rowCount < 1 OrElse columnCount < 1 Then
				Throw New ArgumentOutOfRangeException("Row and Column count must be greater than zero.")
			End If

			Dim t As Table = MyBase.InsertTable(index, rowCount, columnCount)
			t.mainPart = mainPart
			Return t
		End Function

		''' <summary>
		''' Creates a document using a Stream.
		''' </summary>
		''' <param name="stream">The Stream to create the document from.</param>
		''' <param name="documentType"></param>
		''' <returns>Returns a DocX object which represents the document.</returns>
		''' <example>
		''' Creating a document from a FileStream.
		''' <code>
		''' // Use a FileStream fs to create a new document.
		''' using(FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Create))
		''' {
		'''     // Load the document using fs
		'''     using (DocX document = DocX.Create(fs))
		'''     {
		'''         // Do something with the document here.
		'''
		'''         // Save all changes made to this document.
		'''         document.Save();
		'''     }// Release this document from memory.
		''' }
		''' </code>
		''' </example>
		''' <example>
		''' Creating a document in a SharePoint site.
		''' <code>
		''' using(SPSite mySite = new SPSite("http://server/sites/site"))
		''' {
		'''     // Open a connection to the SharePoint site
		'''     using(SPWeb myWeb = mySite.OpenWeb())
		'''     {
		'''         // Create a MemoryStream ms.
		'''         using (MemoryStream ms = new MemoryStream())
		'''         {
		'''             // Create a document using ms.
		'''             using (DocX document = DocX.Create(ms))
		'''             {
		'''                 // Do something with the document here.
		'''
		'''                 // Save all changes made to this document.
		'''                 document.Save();
		'''             }// Release this document from memory
		'''
		'''             // Add the document to the SharePoint site
		'''             web.Files.Add("filename", ms.ToArray(), true);
		'''         }
		'''     }
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Load(string)"/>
		''' <seealso cref="DocX.Save()"/>
		Public Shared Function Create(ByVal stream As Stream, Optional ByVal documentType As DocumentTypes = DocumentTypes.Document) As DocX
			' Store this document in memory
			Dim ms As New MemoryStream()

			' Create the docx package
			Dim package As Package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite)

			PostCreation(package, documentType)
			Dim document As DocX = DocX.Load(ms)
			document.stream = stream
			Return document
		End Function

		''' <summary>
		''' Creates a document using a fully qualified or relative filename.
		''' </summary>
		''' <param name="filename">The fully qualified or relative filename.</param>
		''' <param name="documentType"></param>
		''' <returns>Returns a DocX object which represents the document.</returns>
		''' <example>
		''' <code>
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Create(@"..\Test.docx"))
		''' {
		'''     // Do something with the document here.
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory
		''' </code>
		''' <code>
		''' // Create a document using a relative filename.
		''' using (DocX document = DocX.Create(@"..\Test.docx"))
		''' {
		'''     // Do something with the document here.
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory
		''' </code>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Load(string)"/>
		''' <seealso cref="DocX.Save()"/>
		''' </example>
		Public Shared Function Create(ByVal filename As String, Optional ByVal documentType As DocumentTypes = DocumentTypes.Document) As DocX
			' Store this document in memory
			Dim ms As New MemoryStream()

			' Create the docx package
			'WordprocessingDocument wdDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
			Dim package As Package = Package.Open(ms, FileMode.Create, FileAccess.ReadWrite)

			PostCreation(package, documentType)
			Dim document As DocX = DocX.Load(ms)
			document.filename = filename
			Return document
		End Function

		Friend Shared Sub PostCreation(ByVal package As Package, Optional ByVal documentType As DocumentTypes = DocumentTypes.Document)
			Dim mainDoc, stylesDoc, numberingDoc As XDocument

'			#Region "MainDocumentPart"
			' Create the main document part for this package
			Dim mainDocumentPart As PackagePart
			If documentType = DocumentTypes.Document Then
				mainDocumentPart = package.CreatePart(New Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", CompressionOption.Normal)
			Else
				mainDocumentPart = package.CreatePart(New Uri("/word/document.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml", CompressionOption.Normal)
			End If
			package.CreateRelationship(mainDocumentPart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")

			' Load the document part into a XDocument object
			Using tr As TextReader = New StreamReader(mainDocumentPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
				mainDoc = XDocument.Parse("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & ControlChars.CrLf & "                   <w:document xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">" & ControlChars.CrLf & "                   <w:body>" & ControlChars.CrLf & "                    <w:sectPr w:rsidR=""003E25F4"" w:rsidSect=""00FC3028"">" & ControlChars.CrLf & "                        <w:pgSz w:w=""11906"" w:h=""16838""/>" & ControlChars.CrLf & "                        <w:pgMar w:top=""1440"" w:right=""1440"" w:bottom=""1440"" w:left=""1440"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>" & ControlChars.CrLf & "                        <w:cols w:space=""708""/>" & ControlChars.CrLf & "                        <w:docGrid w:linePitch=""360""/>" & ControlChars.CrLf & "                    </w:sectPr>" & ControlChars.CrLf & "                   </w:body>" & ControlChars.CrLf & "                   </w:document>")
			End Using

			' Save the main document
			Using tw As TextWriter = New StreamWriter(mainDocumentPart.GetStream(FileMode.Create, FileAccess.Write))
				mainDoc.Save(tw, SaveOptions.None)
			End Using
'			#End Region

'			#Region "StylePart"
			stylesDoc = HelperFunctions.AddDefaultStylesXml(package)
'			#End Region

'			#Region "NumberingPart"
			numberingDoc = HelperFunctions.AddDefaultNumberingXml(package)
'			#End Region

			package.Close()
		End Sub

		Friend Shared Function PostLoad(ByRef package As Package) As DocX
			Dim document As New DocX(Nothing, Nothing)
			document.package = package
			document.Document = document

'			#Region "MainDocumentPart"
			document.mainPart = package.GetParts().Where (Function(p) p.ContentType.Equals(HelperFunctions.DOCUMENT_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase) OrElse p.ContentType.Equals(HelperFunctions.TEMPLATE_DOCUMENTTYPE, StringComparison.CurrentCultureIgnoreCase)).Single()

			Using tr As TextReader = New StreamReader(document.mainPart.GetStream(FileMode.Open, FileAccess.Read))
				document.mainDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace)
			End Using
'			#End Region

			PopulateDocument(document, package)

			Using tr As TextReader = New StreamReader(document.settingsPart.GetStream())
				document.settings = XDocument.Load(tr)
			End Using

			document.paragraphLookup.Clear()
			For Each paragraph In document.Paragraphs
				If Not document.paragraphLookup.ContainsKey(paragraph.endIndex) Then
					document.paragraphLookup.Add(paragraph.endIndex, paragraph)
				End If
			Next paragraph

			Return document
		End Function

		Private Shared Sub PopulateDocument(ByVal document As DocX, ByVal package As Package)
			Dim headers As New Headers()
			headers.odd = document.GetHeaderByType("default")
			headers.even = document.GetHeaderByType("even")
			headers.first = document.GetHeaderByType("first")

			Dim footers As New Footers()
			footers.odd = document.GetFooterByType("default")
			footers.even = document.GetFooterByType("even")
			footers.first = document.GetFooterByType("first")

			'// Get the sectPr for this document.
			'XElement sect = document.mainDoc.Descendants(XName.Get("sectPr", DocX.w.NamespaceName)).Single();

			'if (sectPr != null)
			'{
			'    // Extract the even header reference
			'    var header_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
			'    string id = header_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res = document.mainPart.GetRelationship(id);
			'    string ans = res.SourceUri.OriginalString;
			'    headers.even.xml_filename = ans;

			'    // Extract the odd header reference
			'    var header_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "headerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
			'    string id2 = header_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res2 = document.mainPart.GetRelationship(id2);
			'    string ans2 = res2.SourceUri.OriginalString;
			'    headers.odd.xml_filename = ans2;

			'    // Extract the first header reference
			'    var header_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "h
			'eaderReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
			'    string id3 = header_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res3 = document.mainPart.GetRelationship(id3);
			'    string ans3 = res3.SourceUri.OriginalString;
			'    headers.first.xml_filename = ans3;

			'    // Extract the even footer reference
			'    var footer_even_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "even");
			'    string id4 = footer_even_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res4 = document.mainPart.GetRelationship(id4);
			'    string ans4 = res4.SourceUri.OriginalString;
			'    footers.even.xml_filename = ans4;

			'    // Extract the odd footer reference
			'    var footer_odd_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "default");
			'    string id5 = footer_odd_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res5 = document.mainPart.GetRelationship(id5);
			'    string ans5 = res5.SourceUri.OriginalString;
			'    footers.odd.xml_filename = ans5;

			'    // Extract the first footer reference
			'    var footer_first_ref = sectPr.Elements().SingleOrDefault(x => x.Name.LocalName == "footerReference" && x.Attribute(XName.Get("type", DocX.w.NamespaceName)) != null && x.Attribute(XName.Get("type", DocX.w.NamespaceName)).Value == "first");
			'    string id6 = footer_first_ref.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value;
			'    var res6 = document.mainPart.GetRelationship(id6);
			'    string ans6 = res6.SourceUri.OriginalString;
			'    footers.first.xml_filename = ans6;

			'}

			document.Xml = document.mainDoc.Root.Element(w + "body")
			document.headers_Renamed = headers
			document.footers_Renamed = footers
			document.settingsPart = HelperFunctions.CreateOrGetSettingsPart(package)

			Dim ps = package.GetParts()

			'document.endnotesPart = HelperFunctions.GetPart();

			For Each rel In document.mainPart.GetRelationships()
				Select Case rel.RelationshipType
					Case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
						document.endnotesPart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.endnotesPart.GetStream())
							document.endnotes = XDocument.Load(tr)
						End Using

					Case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
						document.footnotesPart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.footnotesPart.GetStream())
							document.footnotes = XDocument.Load(tr)
						End Using

					Case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
						document.stylesPart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.stylesPart.GetStream())
							document.styles = XDocument.Load(tr)
						End Using

					Case "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"
						document.stylesWithEffectsPart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.stylesWithEffectsPart.GetStream())
							document.stylesWithEffects = XDocument.Load(tr)
						End Using

					Case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
						document.fontTablePart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.fontTablePart.GetStream())
							document.fontTable = XDocument.Load(tr)
						End Using

					Case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
						document.numberingPart = package.GetPart(New Uri("/word/" & rel.TargetUri.OriginalString.Replace("/word/", ""), UriKind.Relative))
						Using tr As TextReader = New StreamReader(document.numberingPart.GetStream())
							document.numbering = XDocument.Load(tr)
						End Using

					Case Else
				End Select
			Next rel
		End Sub

		''' <summary>
		''' Saves and copies the document into a new DocX object
		''' </summary>
		''' <returns>
		''' Returns a new DocX object with an identical document
		''' </returns>
		''' <example>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Save()"/>
		''' </example>
		Public Function Copy() As DocX
			Dim ms As New MemoryStream()
			SaveAs(ms)
			ms.Seek(0, SeekOrigin.Begin)

			Return DocX.Load(ms)
		End Function

		''' <summary>
		''' Loads a document into a DocX object using a Stream.
		''' </summary>
		''' <param name="stream">The Stream to load the document from.</param>
		''' <returns>
		''' Returns a DocX object which represents the document.
		''' </returns>
		''' <example>
		''' Loading a document from a FileStream.
		''' <code>
		''' // Open a FileStream fs to a document.
		''' using (FileStream fs = new FileStream(@"C:\Example\Test.docx", FileMode.Open))
		''' {
		'''     // Load the document using fs.
		'''     using (DocX document = DocX.Load(fs))
		'''     {
		'''         // Do something with the document here.
		'''            
		'''         // Save all changes made to the document.
		'''         document.Save();
		'''     }// Release this document from memory.
		''' }
		''' </code>
		''' </example>
		''' <example>
		''' Loading a document from a SharePoint site.
		''' <code>
		''' // Get the SharePoint site that you want to access.
		''' using (SPSite mySite = new SPSite("http://server/sites/site"))
		''' {
		'''     // Open a connection to the SharePoint site
		'''     using (SPWeb myWeb = mySite.OpenWeb())
		'''     {
		'''         // Grab a document stored on this site.
		'''         SPFile file = web.GetFile("Source_Folder_Name/Source_File");
		'''
		'''         // DocX.Load requires a Stream, so open a Stream to this document.
		'''         Stream str = new MemoryStream(file.OpenBinary());
		'''
		'''         // Load the file using the Stream str.
		'''         using (DocX document = DocX.Load(str))
		'''         {
		'''             // Do something with the document here.
		'''
		'''             // Save all changes made to the document.
		'''             document.Save();
		'''         }// Release this document from memory.
		'''     }
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="DocX.Load(string)"/>
		''' <seealso cref="DocX.Save()"/>
		Public Shared Function Load(ByVal stream As Stream) As DocX
			Dim ms As New MemoryStream()

			stream.Position = 0
			Dim data(stream.Length - 1) As Byte
			stream.Read(data, 0, CInt(stream.Length))
			ms.Write(data, 0, CInt(stream.Length))

			' Open the docx package
			Dim package As Package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite)

			Dim document As DocX = PostLoad(package)
			document.package = package
			document.memoryStream = ms
			document.stream = stream
			Return document
		End Function

		''' <summary>
		''' Loads a document into a DocX object using a fully qualified or relative filename.
		''' </summary>
		''' <param name="filename">The fully qualified or relative filename.</param>
		''' <returns>
		''' Returns a DocX object which represents the document.
		''' </returns>
		''' <example>
		''' <code>
		''' // Load a document using its fully qualified filename
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Do something with the document here
		'''
		'''     // Save all changes made to document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' <code>
		''' // Load a document using its relative filename.
		''' using(DocX document = DocX.Load(@"..\..\Test.docx"))
		''' { 
		'''     // Do something with the document here.
		'''                
		'''     // Save all changes made to document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Save()"/>
		''' </example>
		Public Shared Function Load(ByVal filename As String) As DocX
			If Not File.Exists(filename) Then
				Throw New FileNotFoundException(String.Format("File could not be found {0}", filename))
			End If

			Dim ms As New MemoryStream()

			Using fs As New FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read)
				Dim data(fs.Length - 1) As Byte
				fs.Read(data, 0, CInt(fs.Length))
				ms.Write(data, 0, CInt(fs.Length))
			End Using

			' Open the docx package
			Dim package As Package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite)

			Dim document As DocX = PostLoad(package)
			document.package = package
			document.filename = filename
			document.memoryStream = ms

			Return document
		End Function

		'''<summary>
		''' Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		'''</summary>
		'''<param name="templateFilePath">The path to the document template file.</param>
		'''<exception cref="FileNotFoundException">The document template file not found.</exception>
		Public Sub ApplyTemplate(ByVal templateFilePath As String)
			ApplyTemplate(templateFilePath, True)
		End Sub

		'''<summary>
		''' Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		'''</summary>
		'''<param name="templateFilePath">The path to the document template file.</param>
		'''<param name="includeContent">Whether to copy the document template text content to document.</param>
		'''<exception cref="FileNotFoundException">The document template file not found.</exception>
		Public Sub ApplyTemplate(ByVal templateFilePath As String, ByVal includeContent As Boolean)
			If Not File.Exists(templateFilePath) Then
				Throw New FileNotFoundException(String.Format("File could not be found {0}", templateFilePath))
			End If
			Using packageStream As New FileStream(templateFilePath, FileMode.Open, FileAccess.Read)
				ApplyTemplate(packageStream, includeContent)
			End Using
		End Sub

		'''<summary>
		''' Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		'''</summary>
		'''<param name="templateStream">The stream of the document template file.</param>
		Public Sub ApplyTemplate(ByVal templateStream As Stream)
			ApplyTemplate(templateStream, True)
		End Sub

		'''<summary>
		''' Applies document template to the document. Document template may include styles, headers, footers, properties, etc. as well as text content.
		'''</summary>
		'''<param name="templateStream">The stream of the document template file.</param>
		'''<param name="includeContent">Whether to copy the document template text content to document.</param>
		Public Sub ApplyTemplate(ByVal templateStream As Stream, ByVal includeContent As Boolean)
			Dim templatePackage As Package = Package.Open(templateStream)
			Try
				Dim documentPart As PackagePart = Nothing
				Dim documentDoc As XDocument = Nothing
				For Each packagePart As PackagePart In templatePackage.GetParts()
					Select Case packagePart.Uri.ToString()
						Case "/word/document.xml"
							documentPart = packagePart
							Using xr As XmlReader = XmlReader.Create(packagePart.GetStream(FileMode.Open, FileAccess.Read))
								documentDoc = XDocument.Load(xr)
							End Using
						Case "/_rels/.rels"
							If Not Me.package.PartExists(packagePart.Uri) Then
								Me.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption)
							End If
							Dim globalRelsPart As PackagePart = Me.package.GetPart(packagePart.Uri)
							Using tr As New StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), Encoding.UTF8)
								Using tw As New StreamWriter(globalRelsPart.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8)
									tw.Write(tr.ReadToEnd())
								End Using
							End Using
						Case "/word/_rels/document.xml.rels"
						Case Else
							If Not Me.package.PartExists(packagePart.Uri) Then
								Me.package.CreatePart(packagePart.Uri, packagePart.ContentType, packagePart.CompressionOption)
							End If
							Dim packagePartEncoding As Encoding = Encoding.Default
							If packagePart.Uri.ToString().EndsWith(".xml") OrElse packagePart.Uri.ToString().EndsWith(".rels") Then
								packagePartEncoding = Encoding.UTF8
							End If
							Dim nativePart As PackagePart = Me.package.GetPart(packagePart.Uri)
							Using tr As New StreamReader(packagePart.GetStream(FileMode.Open, FileAccess.Read), packagePartEncoding)
								Using tw As New StreamWriter(nativePart.GetStream(FileMode.Create, FileAccess.Write), tr.CurrentEncoding)
									tw.Write(tr.ReadToEnd())
								End Using
							End Using
					End Select
				Next packagePart
				If documentPart IsNot Nothing Then
					Dim mainContentType As String = documentPart.ContentType.Replace("template.main", "document.main")
					If Me.package.PartExists(documentPart.Uri) Then
						Me.package.DeletePart(documentPart.Uri)
					End If
					Dim documentNewPart As PackagePart = Me.package.CreatePart(documentPart.Uri, mainContentType, documentPart.CompressionOption)
					Using xw As XmlWriter = XmlWriter.Create(documentNewPart.GetStream(FileMode.Create, FileAccess.Write))
						documentDoc.WriteTo(xw)
					End Using
					For Each documentPartRel As PackageRelationship In documentPart.GetRelationships()
						documentNewPart.CreateRelationship(documentPartRel.TargetUri, documentPartRel.TargetMode, documentPartRel.RelationshipType, documentPartRel.Id)
					Next documentPartRel
					Me.mainPart = documentNewPart
					Me.mainDoc = documentDoc
					PopulateDocument(Me, templatePackage)

					' DragonFire: I added next line and recovered ApplyTemplate method. 
					' I do it, becouse  PopulateDocument(...) writes into field "settingsPart" the part of Template's package 
					'  and after line "templatePackage.Close();" in finally, field "settingsPart" becomes not available and method "Save" throw an exception...
					' That's why I recreated settingsParts and unlinked it from Template's package =)
					settingsPart = HelperFunctions.CreateOrGetSettingsPart(package)
				End If
				If Not includeContent Then
					For Each paragraph As Paragraph In Me.Paragraphs
						paragraph.Remove(False)
					Next paragraph
				End If
			Finally
				Me.package.Flush()
				Dim documentRelsPart = Me.package.GetPart(New Uri("/word/_rels/document.xml.rels", UriKind.Relative))
				Using tr As TextReader = New StreamReader(documentRelsPart.GetStream(FileMode.Open, FileAccess.Read))
					tr.Read()
				End Using
				templatePackage.Close()
				PopulateDocument(Document, package)
			End Try
		End Sub

		''' <summary>
		''' Add an Image into this document from a fully qualified or relative filename.
		''' </summary>
		''' <param name="filename">The fully qualified or relative filename.</param>
		''' <returns>An Image file.</returns>
		''' <example>
		''' Add an Image into this document from a fully qualified filename.
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Add an Image from a file.
		'''     document.AddImage(@"C:\Example\Image.png");
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="AddImage(System.IO.Stream)"/>
		''' <seealso cref="Paragraph.InsertPicture"/>
		Public Function AddImage(ByVal filename As String) As Image
			Dim contentType As String = ""

			' The extension this file has will be taken to be its format.
			Select Case Path.GetExtension(filename)
				Case ".tiff"
					contentType = "image/tif"
				Case ".tif"
					contentType = "image/tif"
				Case ".png"
					contentType = "image/png"
				Case ".bmp"
					contentType = "image/png"
				Case ".gif"
					contentType = "image/gif"
				Case ".jpg"
					contentType = "image/jpg"
				Case ".jpeg"
					contentType = "image/jpeg"
				Case Else
					contentType = "image/jpg"
			End Select

			Return AddImage(TryCast(filename, Object), contentType)
		End Function

		''' <summary>
		''' Add an Image into this document from a Stream.
		''' </summary>
		''' <param name="stream">A Stream stream.</param>
		''' <returns>An Image file.</returns>
		''' <example>
		''' Add an Image into a document using a Stream. 
		''' <code>
		''' // Open a FileStream fs to an Image.
		''' using (FileStream fs = new FileStream(@"C:\Example\Image.jpg", FileMode.Open))
		''' {
		'''     // Load a document.
		'''     using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		'''     {
		'''         // Add an Image from a filestream fs.
		'''         document.AddImage(fs);
		'''
		'''         // Save all changes made to this document.
		'''         document.Save();
		'''     }// Release this document from memory.
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="AddImage(string)"/>
		''' <seealso cref="Paragraph.InsertPicture"/>
		Public Function AddImage(ByVal stream As Stream) As Image
			Return AddImage(TryCast(stream, Object))
		End Function

		''' <summary>
		''' Adds a hyperlink to a document and creates a Paragraph which uses it.
		''' </summary>
		''' <param name="text">The text as displayed by the hyperlink.</param>
		''' <param name="uri">The hyperlink itself.</param>
		''' <returns>Returns a hyperlink that can be inserted into a Paragraph.</returns>
		''' <example>
		''' Adds a hyperlink to a document and creates a Paragraph which uses it.
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
		Public Function AddHyperlink(ByVal text As String, ByVal uri As Uri) As Hyperlink
			Dim i As New XElement(XName.Get("hyperlink", DocX.w.NamespaceName), New XAttribute(r + "id", String.Empty), New XAttribute(w + "history", "1"), New XElement(XName.Get("r", DocX.w.NamespaceName), New XElement(XName.Get("rPr", DocX.w.NamespaceName), New XElement(XName.Get("rStyle", DocX.w.NamespaceName), New XAttribute(w + "val", "Hyperlink"))), New XElement(XName.Get("t", DocX.w.NamespaceName), text)))

			Dim h As New Hyperlink(Me, mainPart, i)

			h.text_Renamed = text
			h.uri_Renamed = uri

			AddHyperlinkStyleIfNotPresent()

			Return h
		End Function

		Friend Sub AddHyperlinkStyleIfNotPresent()
			Dim word_styles_Uri As New Uri("/word/styles.xml", UriKind.Relative)

			' If the internal document contains no /word/styles.xml create one.
			If Not package.PartExists(word_styles_Uri) Then
				HelperFunctions.AddDefaultStylesXml(package)
			End If

			' Load the styles.xml into memory.
			Dim word_styles As XDocument
			Using tr As TextReader = New StreamReader(package.GetPart(word_styles_Uri).GetStream())
				word_styles = XDocument.Load(tr)
			End Using

			Dim hyperlinkStyleExists As Boolean = (
			    From s In word_styles.Element(w + "styles").Elements()
			    Let styleId = s.Attribute(XName.Get("styleId", w.NamespaceName))
			    Where (styleId IsNot Nothing AndAlso styleId.Value = "Hyperlink")
			    Select s).Count() > 0

			If Not hyperlinkStyleExists Then
				Dim style As New XElement(w + "style", New XAttribute(w + "type", "character"), New XAttribute(w + "styleId", "Hyperlink"), New XElement(w + "name", New XAttribute(w + "val", "Hyperlink")), New XElement(w + "basedOn", New XAttribute(w + "val", "DefaultParagraphFont")), New XElement(w + "uiPriority", New XAttribute(w + "val", "99")), New XElement(w + "unhideWhenUsed"), New XElement(w + "rsid", New XAttribute(w + "val", "0005416C")), New XElement (w + "rPr", New XElement(w + "color", New XAttribute(w + "val", "0000FF"), New XAttribute(w + "themeColor", "hyperlink")), New XElement (w + "u", New XAttribute(w + "val", "single"))))
				word_styles.Element(w + "styles").Add(style)

				' Save the styles document.
				Using tw As TextWriter = New StreamWriter(package.GetPart(word_styles_Uri).GetStream())
					word_styles.Save(tw)
				End Using
			End If
		End Sub

		Private Function GetNextFreeRelationshipID() As String
			Dim id As Integer = (
			    From r In mainPart.GetRelationships()
			    Where r.Id.Substring(0, 3).Equals("rId")
			    Select Integer.Parse(r.Id.Substring(3))).DefaultIfEmpty().Max()

			' The conventiom for ids is rid01, rid02, etc
			Dim newId As String = id.ToString()
			Dim result As Integer
			If Integer.TryParse(newId, result) Then
				Return ("rId" & (result + 1))
			Else
				Dim guid As String = String.Empty
				Do
					guid = Guid.NewGuid().ToString()
				Loop While Char.IsDigit(guid.Chars(0))
				Return guid
			End If
		End Function

		''' <summary>
		''' Adds three new Headers to this document. One for the first page, one for odd pages and one for even pages.
		''' </summary>
		''' <example>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add header support to this document.
		'''     document.AddHeaders();
		'''
		'''     // Get a collection of all headers in this document.
		'''     Headers headers = document.Headers;
		'''
		'''     // The header used for the first page of this document.
		'''     Header first = headers.first;
		'''
		'''     // The header used for odd pages of this document.
		'''     Header odd = headers.odd;
		'''
		'''     // The header used for even pages of this document.
		'''     Header even = headers.even;
		'''
		'''     // Force the document to use a different header for first, odd and even pages.
		'''     document.DifferentFirstPage = true;
		'''     document.DifferentOddAndEvenPages = true;
		'''
		'''     // Content can be added to the Headers in the same manor that it would be added to the main document.
		'''     Paragraph p = first.InsertParagraph();
		'''     p.Append("This is the first pages header.");
		'''
		'''     // Save all changes to this document.
		'''     document.Save();    
		''' }// Release this document from memory.
		''' </example>
		Public Sub AddHeaders()
			AddHeadersOrFooters(True)

			headers_Renamed.odd = Document.GetHeaderByType("default")
			headers_Renamed.even = Document.GetHeaderByType("even")
			headers_Renamed.first = Document.GetHeaderByType("first")
		End Sub

		''' <summary>
		''' Adds three new Footers to this document. One for the first page, one for odd pages and one for even pages.
		''' </summary>
		''' <example>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''     // Add footer support to this document.
		'''     document.AddFooters();
		'''
		'''     // Get a collection of all footers in this document.
		'''     Footers footers = document.Footers;
		'''
		'''     // The footer used for the first page of this document.
		'''     Footer first = footers.first;
		'''
		'''     // The footer used for odd pages of this document.
		'''     Footer odd = footers.odd;
		'''
		'''     // The footer used for even pages of this document.
		'''     Footer even = footers.even;
		'''
		'''     // Force the document to use a different footer for first, odd and even pages.
		'''     document.DifferentFirstPage = true;
		'''     document.DifferentOddAndEvenPages = true;
		'''
		'''     // Content can be added to the Footers in the same manor that it would be added to the main document.
		'''     Paragraph p = first.InsertParagraph();
		'''     p.Append("This is the first pages footer.");
		'''
		'''     // Save all changes to this document.
		'''     document.Save();    
		''' }// Release this document from memory.
		''' </example>
		Public Sub AddFooters()
			AddHeadersOrFooters(False)

			footers_Renamed.odd = Document.GetFooterByType("default")
			footers_Renamed.even = Document.GetFooterByType("even")
			footers_Renamed.first = Document.GetFooterByType("first")
		End Sub

		''' <summary>
		''' Adds a Header to a document.
		''' If the document already contains a Header it will be replaced.
		''' </summary>
		''' <returns>The Header that was added to the document.</returns>
		Friend Sub AddHeadersOrFooters(ByVal b As Boolean)
			Dim element As String = "ftr"
			Dim reference As String = "footer"
			If b Then
				element = "hdr"
				reference = "header"
			End If

			DeleteHeadersOrFooters(b)

			Dim sectPr As XElement = mainDoc.Root.Element(w + "body").Element(w + "sectPr")

			For i As Integer = 1 To 3
				Dim header_uri As String = String.Format("/word/{0}{1}.xml", reference, i)

				Dim headerPart As PackagePart = package.CreatePart(New Uri(header_uri, UriKind.Relative), String.Format("application/vnd.openxmlformats-officedocument.wordprocessingml.{0}+xml", reference), CompressionOption.Normal)
				Dim headerRelationship As PackageRelationship = mainPart.CreateRelationship(headerPart.Uri, TargetMode.Internal, String.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference))

				Dim header As XDocument

				' Load the document part into a XDocument object
				Using tr As TextReader = New StreamReader(headerPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
					header = XDocument.Parse(String.Format("<?xml version=""1.0"" encoding=""utf-16"" standalone=""yes""?>" & ControlChars.CrLf & "                       <w:{0} xmlns:ve=""http://schemas.openxmlformats.org/markup-compatibility/2006"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"" xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:wp=""http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"" xmlns:w10=""urn:schemas-microsoft-com:office:word"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:wne=""http://schemas.microsoft.com/office/word/2006/wordml"">" & ControlChars.CrLf & "                         <w:p w:rsidR=""009D472B"" w:rsidRDefault=""009D472B"">" & ControlChars.CrLf & "                           <w:pPr>" & ControlChars.CrLf & "                             <w:pStyle w:val=""{1}"" />" & ControlChars.CrLf & "                           </w:pPr>" & ControlChars.CrLf & "                         </w:p>" & ControlChars.CrLf & "                       </w:{0}>", element, reference))
				End Using

				' Save the main document
				Using tw As TextWriter = New StreamWriter(headerPart.GetStream(FileMode.Create, FileAccess.Write))
					header.Save(tw, SaveOptions.None)
				End Using

				Dim type As String
				Select Case i
					Case 1
						type = "default"
					Case 2
						type = "even"
					Case 3
						type = "first"
					Case Else
						Throw New ArgumentOutOfRangeException()
				End Select

				sectPr.Add(New XElement (w + String.Format("{0}Reference", reference), New XAttribute(w + "type", type), New XAttribute(r + "id", headerRelationship.Id)))
			Next i
		End Sub

		Friend Sub DeleteHeadersOrFooters(ByVal b As Boolean)
			Dim reference As String = "footer"
			If b Then
				reference = "header"
			End If

			' Get all header Relationships in this document.
			Dim header_relationships = mainPart.GetRelationshipsByType(String.Format("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{0}", reference))

			For Each header_relationship As PackageRelationship In header_relationships
				' Get the TargetUri for this Part.
				Dim header_uri As Uri = header_relationship.TargetUri

				' Check to see if the document actually contains the Part.
				If Not header_uri.OriginalString.StartsWith("/word/") Then
					header_uri = New Uri("/word/" & header_uri.OriginalString, UriKind.Relative)
				End If

				If package.PartExists(header_uri) Then
					' Delete the Part
					package.DeletePart(header_uri)

					' Get all references to this Relationship in the document.
					Dim query = (
					    From e In mainDoc.Descendants(XName.Get("body", DocX.w.NamespaceName)).Descendants()
					    Where (e.Name.LocalName = String.Format("{0}Reference", reference)) AndAlso (e.Attribute(r + "id").Value = header_relationship.Id)
					    Select e)

					' Remove all references to this Relationship in the document.
					For i As Integer = 0 To query.Count() - 1
						query.ElementAt(i).Remove()
					Next i

					' Delete the Relationship.
					package.DeleteRelationship(header_relationship.Id)
				End If
			Next header_relationship
		End Sub

		Friend Function AddImage(ByVal o As Object, Optional ByVal contentType As String = "image/jpeg") As Image
			' Open a Stream to the new image being added.
			Dim newImageStream As Stream
			If TypeOf o Is String Then
				newImageStream = New FileStream(TryCast(o, String), FileMode.Open, FileAccess.Read)
			Else
				newImageStream = TryCast(o, Stream)
			End If

			' Get all image parts in word\document.xml

			Dim relationshipsByImages As PackageRelationshipCollection = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
			Dim imageParts As List(Of PackagePart) = relationshipsByImages.Select(Function(ir) package.GetParts().FirstOrDefault(Function(p) p.Uri.ToString().EndsWith(ir.TargetUri.ToString()))).Where(Function(e) e IsNot Nothing).ToList()

			For Each relsPart As PackagePart In package.GetParts().Where(Function(part) part.Uri.ToString().Contains("/word/")).Where(Function(part) part.ContentType.Equals("application/vnd.openxmlformats-package.relationships+xml"))
				Dim relsPartContent As XDocument
				Using tr As TextReader = New StreamReader(relsPart.GetStream(FileMode.Open, FileAccess.Read))
					relsPartContent = XDocument.Load(tr)
				End Using

				Dim imageRelationships As IEnumerable(Of XElement) = relsPartContent.Root.Elements().Where (Function(imageRel) imageRel.Attribute(XName.Get("Type")).Value.Equals("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"))

				For Each imageRelationship As XElement In imageRelationships
					If imageRelationship.Attribute(XName.Get("Target")) IsNot Nothing Then
						Dim targetMode As String = String.Empty

						Dim targetModeAttibute As XAttribute = imageRelationship.Attribute(XName.Get("TargetMode"))
						If targetModeAttibute IsNot Nothing Then
							targetMode = targetModeAttibute.Value
						End If

						If Not targetMode.Equals("External") Then
							Dim imagePartUri As String = Path.Combine(Path.GetDirectoryName(relsPart.Uri.ToString()), imageRelationship.Attribute(XName.Get("Target")).Value)
							imagePartUri = Path.GetFullPath(imagePartUri.Replace("\_rels", String.Empty))
							imagePartUri = imagePartUri.Replace(Path.GetFullPath("\"), String.Empty).Replace("\", "/")

							If Not imagePartUri.StartsWith("/") Then
								imagePartUri = "/" & imagePartUri
							End If

							Dim imagePart As PackagePart = package.GetPart(New Uri(imagePartUri, UriKind.Relative))
							imageParts.Add(imagePart)
						End If
					End If
				Next imageRelationship
			Next relsPart

			' Loop through each image part in this document.
			For Each pp As PackagePart In imageParts
				' Open a tempory Stream to this image part.
				Using tempStream As Stream = pp.GetStream(FileMode.Open, FileAccess.Read)
					' Compare this image to the new image being added.
					If HelperFunctions.IsSameFile(tempStream, newImageStream) Then
						' Get the image object for this image part
						Dim id As String = mainPart.GetRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image").Where(Function(r) r.TargetUri Is pp.Uri).Select(Function(r) r.Id).First()

						' Return the Image object
						Return Images.Where(Function(i) i.Id = id).First()
					End If
				End Using
			Next pp

			Dim imgPartUriPath As String = String.Empty
			Dim extension As String = contentType.Substring(contentType.LastIndexOf("/") + 1)
			Do
				' Create a new image part.
				imgPartUriPath = String.Format ("/word/media/{0}.{1}", Guid.NewGuid().ToString(), extension) ' The unique part.

			Loop While package.PartExists(New Uri(imgPartUriPath, UriKind.Relative))

			' We are now guareenteed that imgPartUriPath is unique.
			Dim img As PackagePart = package.CreatePart(New Uri(imgPartUriPath, UriKind.Relative), contentType, CompressionOption.Normal)

			' Create a new image relationship
			Dim rel As PackageRelationship = mainPart.CreateRelationship(img.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")

			' Open a Stream to the newly created Image part.
			Using stream As Stream = img.GetStream(FileMode.Create, FileAccess.Write)
				' Using the Stream to the real image, copy this streams data into the newly create Image part.
				Using newImageStream
					Dim bytes(newImageStream.Length - 1) As Byte
					newImageStream.Read(bytes, 0, CInt(newImageStream.Length))
					stream.Write(bytes, 0, CInt(newImageStream.Length))
				End Using ' Close the Stream to the new image.
			End Using ' Close the Stream to the new image part.

			Return New Image(Me, rel)
		End Function

		''' <summary>
		''' Save this document back to the location it was loaded from.
		''' </summary>
		''' <example>
		''' <code>
		''' // Load a document.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // Add an Image from a file.
		'''     document.AddImage(@"C:\Example\Image.jpg");
		'''
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		''' <seealso cref="DocX.SaveAs(string)"/>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Load(string)"/> 
		''' <!-- 
		''' Bug found and fixed by krugs525 on August 12 2009.
		''' Use TFS compare to see exact code change.
		''' -->
		Public Sub Save()
			Dim headers As Headers = Headers

			' Save the main document
			Using tw As TextWriter = New StreamWriter(mainPart.GetStream(FileMode.Create, FileAccess.Write))
				mainDoc.Save(tw, SaveOptions.None)
			End Using

			If settings Is Nothing Then
				Using tr As TextReader = New StreamReader(settingsPart.GetStream())
					settings = XDocument.Load(tr)
				End Using
			End If

			Dim body As XElement = mainDoc.Root.Element(w + "body")
			Dim sectPr As XElement = body.Descendants(w + "sectPr").FirstOrDefault()

			If sectPr IsNot Nothing Then
				Dim evenHeaderRef = (
				    From e In mainDoc.Descendants(w + "headerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If evenHeaderRef IsNot Nothing Then
					Dim even As XElement = headers.even.Xml

					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(evenHeaderRef).TargetUri)

					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), even), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				Dim oddHeaderRef = (
				    From e In mainDoc.Descendants(w + "headerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If oddHeaderRef IsNot Nothing Then
					Dim odd As XElement = headers.odd.Xml

					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(oddHeaderRef).TargetUri)

					' Save header1
					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), odd), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				Dim firstHeaderRef = (
				    From e In mainDoc.Descendants(w + "headerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If firstHeaderRef IsNot Nothing Then
					Dim first As XElement = headers.first.Xml
					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(firstHeaderRef).TargetUri)

					' Save header3
					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), first), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				Dim oddFooterRef = (
				    From e In mainDoc.Descendants(w + "footerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("default", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If oddFooterRef IsNot Nothing Then
					Dim odd As XElement = footers_Renamed.odd.Xml
					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(oddFooterRef).TargetUri)

					' Save header1
					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), odd), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				Dim evenFooterRef = (
				    From e In mainDoc.Descendants(w + "footerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("even", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If evenFooterRef IsNot Nothing Then
					Dim even As XElement = footers_Renamed.even.Xml
					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(evenFooterRef).TargetUri)

					' Save header2
					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), even), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				Dim firstFooterRef = (
				    From e In mainDoc.Descendants(w + "footerReference")
				    Let type = e.Attribute(w + "type")
				    Where type IsNot Nothing AndAlso type.Value.Equals("first", StringComparison.CurrentCultureIgnoreCase)
				    Select e.Attribute(r + "id").Value).LastOrDefault()

				If firstFooterRef IsNot Nothing Then
					Dim first As XElement = footers_Renamed.first.Xml
					Dim target As Uri = PackUriHelper.ResolvePartUri (mainPart.Uri, mainPart.GetRelationship(firstFooterRef).TargetUri)

					' Save header3
					Using tw As TextWriter = New StreamWriter(package.GetPart(target).GetStream(FileMode.Create, FileAccess.Write))
						CType(New XDocument (New XDeclaration("1.0", "UTF-8", "yes"), first), XDocument).Save(tw, SaveOptions.None)
					End Using
				End If

				' Save the settings document.
				Using tw As TextWriter = New StreamWriter(settingsPart.GetStream(FileMode.Create, FileAccess.Write))
					settings.Save(tw, SaveOptions.None)
				End Using

				If endnotesPart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(endnotesPart.GetStream(FileMode.Create, FileAccess.Write))
						endnotes.Save(tw, SaveOptions.None)
					End Using
				End If

				If footnotesPart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(footnotesPart.GetStream(FileMode.Create, FileAccess.Write))
						footnotes.Save(tw, SaveOptions.None)
					End Using
				End If

				If stylesPart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(stylesPart.GetStream(FileMode.Create, FileAccess.Write))
						styles.Save(tw, SaveOptions.None)
					End Using
				End If

				If stylesWithEffectsPart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(stylesWithEffectsPart.GetStream(FileMode.Create, FileAccess.Write))
						stylesWithEffects.Save(tw, SaveOptions.None)
					End Using
				End If

				If numberingPart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(numberingPart.GetStream(FileMode.Create, FileAccess.Write))
						numbering.Save(tw, SaveOptions.None)
					End Using
				End If

				If fontTablePart IsNot Nothing Then
					Using tw As TextWriter = New StreamWriter(fontTablePart.GetStream(FileMode.Create, FileAccess.Write))
						fontTable.Save(tw, SaveOptions.None)
					End Using
				End If
			End If

			' Close the document so that it can be saved.
			package.Flush()

'			#Region "Save this document back to a file or stream, that was specified by the user at save time."
			If filename IsNot Nothing Then
				Using fs As New FileStream(filename, FileMode.Create)
					fs.Write(memoryStream.ToArray(), 0, CInt(memoryStream.Length))
				End Using
			Else
				If stream.CanSeek Then ' 2013-05-25: Check if stream can be seeked to support System.Web.HttpResponseStream
					' Set the length of this stream to 0
					stream.SetLength(0)

					' Write to the beginning of the stream
					stream.Position = 0
				End If

				memoryStream.WriteTo(stream)
				memoryStream.Flush()
			End If
'			#End Region
		End Sub

		''' <summary>
		''' Save this document to a file.
		''' </summary>
		''' <param name="filename">The filename to save this document as.</param>
		''' <example>
		''' Load a document from one file and save it to another.
		''' <code>
		''' // Load a document using its fully qualified filename.
		''' DocX document = DocX.Load(@"C:\Example\Test1.docx");
		'''
		''' // Insert a new Paragraph
		''' document.InsertParagraph("Hello world!", false);
		'''
		''' // Save the document to a new location.
		''' document.SaveAs(@"C:\Example\Test2.docx");
		''' </code>
		''' </example>
		''' <example>
		''' Load a document from a Stream and save it to a file.
		''' <code>
		''' DocX document;
		''' using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		''' {
		'''     // Load a document using a stream.
		'''     document = DocX.Load(fs1);
		'''
		'''     // Insert a new Paragraph
		'''     document.InsertParagraph("Hello world again!", false);
		''' }
		'''    
		''' // Save the document to a new location.
		''' document.SaveAs(@"C:\Example\Test2.docx");
		''' </code>
		''' </example>
		''' <seealso cref="DocX.Save()"/>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Load(string)"/>
		Public Sub SaveAs(ByVal filename As String)
			Me.filename = filename
			Me.stream = Nothing
			Save()
		End Sub

		''' <summary>
		''' Save this document to a Stream.
		''' </summary>
		''' <param name="stream">The Stream to save this document to.</param>
		''' <example>
		''' Load a document from a file and save it to a Stream.
		''' <code>
		''' // Place holder for a document.
		''' DocX document;
		'''
		''' using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		''' {
		'''     // Load a document using a stream.
		'''     document = DocX.Load(fs1);
		'''
		'''     // Insert a new Paragraph
		'''     document.InsertParagraph("Hello world again!", false);
		''' }
		'''
		''' using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
		''' {
		'''     // Save the document to a different stream.
		'''     document.SaveAs(fs2);
		''' }
		'''
		''' // Release this document from memory.
		''' document.Dispose();
		''' </code>
		''' </example>
		''' <example>
		''' Load a document from one Stream and save it to another.
		''' <code>
		''' DocX document;
		''' using (FileStream fs1 = new FileStream(@"C:\Example\Test1.docx", FileMode.Open))
		''' {
		'''     // Load a document using a stream.
		'''     document = DocX.Load(fs1);
		'''
		'''     // Insert a new Paragraph
		'''     document.InsertParagraph("Hello world again!", false);
		''' }
		''' 
		''' using (FileStream fs2 = new FileStream(@"C:\Example\Test2.docx", FileMode.Create))
		''' {
		'''     // Save the document to a different stream.
		'''     document.SaveAs(fs2);
		''' }
		''' </code>
		''' </example>
		''' <seealso cref="DocX.Save()"/>
		''' <seealso cref="DocX.Load(System.IO.Stream)"/>
		''' <seealso cref="DocX.Load(string)"/>
		Public Sub SaveAs(ByVal stream As Stream)
			Me.filename = Nothing
			Me.stream = stream
			Save()
		End Sub

		''' <summary>
		''' Add a core property to this document. If a core property already exists with the same name it will be replaced. Core property names are case insensitive.
		''' </summary>
		'''<param name="propertyName">The property name.</param>
		'''<param name="propertyValue">The property value.</param>
		'''<example>
		''' Add a core properties of each type to a document.
		''' <code>
		''' // Load Example.docx
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // If this document does not contain a core property called 'forename', create one.
		'''     if (!document.CoreProperties.ContainsKey("forename"))
		'''     {
		'''         // Create a new core property called 'forename' and set its value.
		'''         document.AddCoreProperty("forename", "Cathal");
		'''     }
		'''
		'''     // Get this documents core property called 'forename'.
		'''     string forenameValue = document.CoreProperties["forename"];
		'''
		'''     // Print all of the information about this core property to Console.
		'''     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", "forename", forenameValue));
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' } // Release this document from memory.
		'''
		''' // Wait for the user to press a key before exiting.
		''' Console.ReadKey();
		''' </code>
		''' </example>
		''' <seealso cref="CoreProperties"/>
		''' <seealso cref="CustomProperty"/>
		''' <seealso cref="CustomProperties"/>
		Public Sub AddCoreProperty(ByVal propertyName As String, ByVal propertyValue As String)
			Dim propertyNamespacePrefix As String = If(propertyName.Contains(":"), propertyName.Split( { ":"c })(0), "cp")
			Dim propertyLocalName As String = If(propertyName.Contains(":"), propertyName.Split( { ":"c })(1), propertyName)

			' If this document does not contain a coreFilePropertyPart create one.)
			If Not package.PartExists(New Uri("/docProps/core.xml", UriKind.Relative)) Then
				Throw New Exception("Core properties part doesn't exist.")
			End If

			Dim corePropDoc As XDocument
			Dim corePropPart As PackagePart = package.GetPart(New Uri("/docProps/core.xml", UriKind.Relative))
			Using tr As TextReader = New StreamReader(corePropPart.GetStream(FileMode.Open, FileAccess.Read))
				corePropDoc = XDocument.Load(tr)
			End Using

			Dim corePropElement As XElement = (
			    From propElement In corePropDoc.Root.Elements()
			    Where (propElement.Name.LocalName.Equals(propertyLocalName))
			    Select propElement).SingleOrDefault()
			If corePropElement IsNot Nothing Then
				corePropElement.SetValue(propertyValue)
			Else
				Dim propertyNamespace = corePropDoc.Root.GetNamespaceOfPrefix(propertyNamespacePrefix)
				corePropDoc.Root.Add(New XElement(XName.Get(propertyLocalName, propertyNamespace.NamespaceName), propertyValue))
			End If

			Using tw As TextWriter = New StreamWriter(corePropPart.GetStream(FileMode.Create, FileAccess.Write))
				corePropDoc.Save(tw)
			End Using
			UpdateCorePropertyValue(Me, propertyLocalName, propertyValue)
		End Sub

		Friend Shared Sub UpdateCorePropertyValue(ByVal document As DocX, ByVal corePropertyName As String, ByVal corePropertyValue As String)
			Dim matchPattern As String = String.Format("(DOCPROPERTY)?{0}\\\*MERGEFORMAT", corePropertyName).ToLower()
			For Each e As XElement In document.mainDoc.Descendants(XName.Get("fldSimple", w.NamespaceName))
				Dim attr_value As String = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", String.Empty).Trim().ToLower()

				If Regex.IsMatch(attr_value, matchPattern) Then
					Dim firstRun As XElement = e.Element(w + "r")
					Dim firstText As XElement = firstRun.Element(w + "t")
					Dim rPr As XElement = firstText.Element(w + "rPr")

					' Delete everything and insert updated text value
					e.RemoveNodes()

					Dim t As New XElement(w + "t", rPr, corePropertyValue)
					Novacode.Text.PreserveSpace(t)
					e.Add(New XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t))
				End If
			Next e

'			#Region "Headers"

			Dim headerParts As IEnumerable(Of PackagePart) = From headerPart In document.package.GetParts()
			                                                 Where (Regex.IsMatch(headerPart.Uri.ToString(), "/word/header\d?.xml"))
			                                                 Select headerPart
			For Each pp As PackagePart In headerParts
				Dim header As XDocument = XDocument.Load(New StreamReader(pp.GetStream()))

				For Each e As XElement In header.Descendants(XName.Get("fldSimple", w.NamespaceName))
					Dim attr_value As String = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", String.Empty).Trim().ToLower()
					If Regex.IsMatch(attr_value, matchPattern) Then
						Dim firstRun As XElement = e.Element(w + "r")

						' Delete everything and insert updated text value
						e.RemoveNodes()

						Dim t As New XElement(w + "t", corePropertyValue)
						Novacode.Text.PreserveSpace(t)
						e.Add(New XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t))
					End If
				Next e

				Using tw As TextWriter = New StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write))
					header.Save(tw)
				End Using
			Next pp
'			#End Region

'			#Region "Footers"
			Dim footerParts As IEnumerable(Of PackagePart) = From footerPart In document.package.GetParts()
			                                                 Where (Regex.IsMatch(footerPart.Uri.ToString(), "/word/footer\d?.xml"))
			                                                 Select footerPart
			For Each pp As PackagePart In footerParts
				Dim footer As XDocument = XDocument.Load(New StreamReader(pp.GetStream()))

				For Each e As XElement In footer.Descendants(XName.Get("fldSimple", w.NamespaceName))
					Dim attr_value As String = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", String.Empty).Trim().ToLower()
					If Regex.IsMatch(attr_value, matchPattern) Then
						Dim firstRun As XElement = e.Element(w + "r")

						' Delete everything and insert updated text value
						e.RemoveNodes()

						Dim t As New XElement(w + "t", corePropertyValue)
						Novacode.Text.PreserveSpace(t)
						e.Add(New XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t))
					End If
				Next e

				Using tw As TextWriter = New StreamWriter(pp.GetStream(FileMode.Create, FileAccess.Write))
					footer.Save(tw)
				End Using
			Next pp
'			#End Region
			PopulateDocument(document, document.package)
		End Sub

		''' <summary>
		''' Add a custom property to this document. If a custom property already exists with the same name it will be replace. CustomProperty names are case insensitive.
		''' </summary>
		''' <param name="cp">The CustomProperty to add to this document.</param>
		''' <example>
		''' Add a custom properties of each type to a document.
		''' <code>
		''' // Load Example.docx
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''     // A CustomProperty called forename which stores a string.
		'''     CustomProperty forename;
		'''
		'''     // If this document does not contain a custom property called 'forename', create one.
		'''     if (!document.CustomProperties.ContainsKey("forename"))
		'''     {
		'''         // Create a new custom property called 'forename' and set its value.
		'''         document.AddCustomProperty(new CustomProperty("forename", "Cathal"));
		'''     }
		'''
		'''     // Get this documents custom property called 'forename'.
		'''     forename = document.CustomProperties["forename"];
		'''
		'''     // Print all of the information about this CustomProperty to Console.
		'''     Console.WriteLine(string.Format("Name: '{0}', Value: '{1}'\nPress any key...", forename.Name, forename.Value));
		'''     
		'''     // Save all changes made to this document.
		'''     document.Save();
		''' } // Release this document from memory.
		'''
		''' // Wait for the user to press a key before exiting.
		''' Console.ReadKey();
		''' </code>
		''' </example>
		''' <seealso cref="CustomProperty"/>
		''' <seealso cref="CustomProperties"/>
		Public Sub AddCustomProperty(ByVal cp As CustomProperty)
			' If this document does not contain a customFilePropertyPart create one.
			If Not package.PartExists(New Uri("/docProps/custom.xml", UriKind.Relative)) Then
				HelperFunctions.CreateCustomPropertiesPart(Me)
			End If

			Dim customPropDoc As XDocument
			Dim customPropPart As PackagePart = package.GetPart(New Uri("/docProps/custom.xml", UriKind.Relative))
			Using tr As TextReader = New StreamReader(customPropPart.GetStream(FileMode.Open, FileAccess.Read))
				customPropDoc = XDocument.Load(tr, LoadOptions.PreserveWhitespace)
			End Using

			' Each custom property has a PID, get the highest PID in this document.
			Dim pids As IEnumerable(Of Integer) = (
			    From d In customPropDoc.Descendants()
			    Where d.Name.LocalName = "property"
			    Select Integer.Parse(d.Attribute(XName.Get("pid")).Value))

			Dim pid As Integer = 1
			If pids.Count() > 0 Then
				pid = pids.Max()
			End If

			' Check if a custom property already exists with this name
			' 2013-05-25: IgnoreCase while searching for custom property as it would produce a currupted docx.
			Dim customProperty = (
			    From d In customPropDoc.Descendants()
			    Where (d.Name.LocalName = "property") AndAlso (d.Attribute(XName.Get("name")).Value.Equals(cp.Name, StringComparison.InvariantCultureIgnoreCase))
			    Select d).SingleOrDefault()

			' If a custom property with this name already exists remove it.
			If customProperty IsNot Nothing Then
				customProperty.Remove()
			End If

			Dim propertiesElement As XElement = customPropDoc.Element(XName.Get("Properties", customPropertiesSchema.NamespaceName))
			propertiesElement.Add(New XElement (XName.Get("property", customPropertiesSchema.NamespaceName), New XAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"), New XAttribute("pid", pid + 1), New XAttribute("name", cp.Name), New XElement(customVTypesSchema & cp.Type, If(cp.Value, ""))))

			' Save the custom properties
			Using tw As TextWriter = New StreamWriter(customPropPart.GetStream(FileMode.Create, FileAccess.Write))
				customPropDoc.Save(tw, SaveOptions.None)
			End Using

			' Refresh all fields in this document which display this custom property.
			UpdateCustomPropertyValue(Me, cp.Name, (If(cp.Value, "")).ToString())
		End Sub

		''' <summary>
		''' Update the custom properties inside the document
		''' </summary>
		''' <param name="document">The DocX document</param>
		''' <param name="customPropertyName">The property used inside the document</param>
		''' <param name="customPropertyValue">The new value for the property</param>
		''' <remarks>Different version of Word create different Document XML.</remarks>
		Friend Shared Sub UpdateCustomPropertyValue(ByVal document As DocX, ByVal customPropertyName As String, ByVal customPropertyValue As String)
			' A list of documents, which will contain, The Main Document and if they exist: header1, header2, header3, footer1, footer2, footer3.
			Dim documents As New List(Of XElement)() From {document.mainDoc.Root}

			' Check if each header exists and add if if so.
'			#Region "Headers"
			Dim headers As Headers = document.Headers
			If headers.first IsNot Nothing Then
				documents.Add(headers.first.Xml)
			End If
			If headers.odd IsNot Nothing Then
				documents.Add(headers.odd.Xml)
			End If
			If headers.even IsNot Nothing Then
				documents.Add(headers.even.Xml)
			End If
'			#End Region

			' Check if each footer exists and add if if so.
'			#Region "Footers"
			Dim footers As Footers = document.Footers
			If footers.first IsNot Nothing Then
				documents.Add(footers.first.Xml)
			End If
			If footers.odd IsNot Nothing Then
				documents.Add(footers.odd.Xml)
			End If
			If footers.even IsNot Nothing Then
				documents.Add(footers.even.Xml)
			End If
'			#End Region

		   Dim matchCustomPropertyName = customPropertyName
		   If customPropertyName.Contains(" ") Then
			   matchCustomPropertyName = """" & customPropertyName & """"
		   End If
		   Dim match_value As String = String.Format("DOCPROPERTY  {0}  \* MERGEFORMAT", matchCustomPropertyName).Replace(" ", String.Empty)

			' Process each document in the list.
			For Each doc As XElement In documents
'				#Region "Word 2010+"
				For Each e As XElement In doc.Descendants(XName.Get("instrText", w.NamespaceName))

					Dim attr_value As String = e.Value.Replace(" ", String.Empty).Trim()

					If attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase) Then
						Dim node As XNode = e.Parent.NextNode
						Dim found As Boolean = False
						Do
							If node.NodeType = XmlNodeType.Element Then
								Dim ele = TryCast(node, XElement)
								Dim match = ele.Descendants(XName.Get("t", w.NamespaceName))
								If match.Count() > 0 Then
									If Not found Then
										match.First().Value = customPropertyValue
										found = True
									Else
										ele.RemoveNodes()
									End If
								Else
									match = ele.Descendants(XName.Get("fldChar", w.NamespaceName))
									If match.Count() > 0 Then
										Dim endMatch = match.First().Attribute(XName.Get("fldCharType", w.NamespaceName))
										If endMatch IsNot Nothing AndAlso endMatch.Value = "end" Then
											Exit Do
										End If
									End If
								End If
							End If
							node = node.NextNode
						Loop
					End If
				Next e
'				#End Region

'				#Region "< Word 2010"
				For Each e As XElement In doc.Descendants(XName.Get("fldSimple", w.NamespaceName))
					Dim attr_value As String = e.Attribute(XName.Get("instr", w.NamespaceName)).Value.Replace(" ", String.Empty).Trim()

					If attr_value.Equals(match_value, StringComparison.CurrentCultureIgnoreCase) Then
						Dim firstRun As XElement = e.Element(w + "r")
						Dim firstText As XElement = firstRun.Element(w + "t")
						Dim rPr As XElement = firstText.Element(w + "rPr")

						' Delete everything and insert updated text value
						e.RemoveNodes()

						Dim t As New XElement(w + "t", rPr, customPropertyValue)
						Novacode.Text.PreserveSpace(t)
						e.Add(New XElement(firstRun.Name, firstRun.Attributes(), firstRun.Element(XName.Get("rPr", w.NamespaceName)), t))
					End If
				Next e
'				#End Region
			Next doc
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

		Public Function InsertParagraphs(ByVal text As String) As Paragraph()
			Dim textArray() As String = text.Split(ControlChars.Lf)
			Dim paragraphs As New List(Of Paragraph)()
			For Each textForParagraph In textArray
				Dim p As Paragraph = MyBase.InsertParagraph(text)
				p.PackagePart = mainPart
				paragraphs.Add(p)
			Next textForParagraph
			Return paragraphs.ToArray()
		End Function

		Public Overrides ReadOnly Property Paragraphs() As ReadOnlyCollection(Of Paragraph)
			Get
				Dim l As ReadOnlyCollection(Of Paragraph) = MyBase.Paragraphs
				For Each paragraph In l
					paragraph.PackagePart = mainPart
				Next paragraph
				Return l
			End Get
		End Property

		Public Overrides ReadOnly Property Lists() As List(Of List)
			Get
				Dim l As List(Of List) = MyBase.Lists
				l.ForEach(Function(x) x.Items.ForEach(Function(i) i.PackagePart = mainPart))
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


		''' <summary>
		''' Create an equation and insert it in the new paragraph
		''' </summary>        
		Public Overrides Function InsertEquation(ByVal equation As String) As Paragraph
			Dim p As Paragraph = MyBase.InsertEquation(equation)
			p.PackagePart = mainPart
			Return p
		End Function

		''' <summary>
		''' Insert a chart in document
		''' </summary>
		Public Sub InsertChart(ByVal chart As Chart)
			' Create a new chart part uri.
			Dim chartPartUriPath As String = String.Empty
			Dim chartIndex As Int32 = 1
			Do
				chartPartUriPath = String.Format ("/word/charts/chart{0}.xml", chartIndex)
				chartIndex += 1
			Loop While package.PartExists(New Uri(chartPartUriPath, UriKind.Relative))

			' Create chart part.
			Dim chartPackagePart As PackagePart = package.CreatePart(New Uri(chartPartUriPath, UriKind.Relative), "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", CompressionOption.Normal)

			' Create a new chart relationship
			Dim relID As String = GetNextFreeRelationshipID()
			Dim rel As PackageRelationship = mainPart.CreateRelationship(chartPackagePart.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", relID)

			' Save a chart info the chartPackagePart
			Using tw As TextWriter = New StreamWriter(chartPackagePart.GetStream(FileMode.Create, FileAccess.Write))
				chart.Xml.Save(tw)
			End Using

			' Insert a new chart into a paragraph.
			Dim p As Paragraph = InsertParagraph()
			Dim chartElement As New XElement(XName.Get("r", DocX.w.NamespaceName), New XElement(XName.Get("drawing", DocX.w.NamespaceName), New XElement(XName.Get("inline", DocX.wp.NamespaceName), New XElement(XName.Get("extent", DocX.wp.NamespaceName), New XAttribute("cx", "5486400"), New XAttribute("cy", "3200400")), New XElement(XName.Get("effectExtent", DocX.wp.NamespaceName), New XAttribute("l", "0"), New XAttribute("t", "0"), New XAttribute("r", "19050"), New XAttribute("b", "19050")), New XElement(XName.Get("docPr", DocX.wp.NamespaceName), New XAttribute("id", "1"), New XAttribute("name", "chart")), New XElement(XName.Get("graphic", DocX.a.NamespaceName), New XElement(XName.Get("graphicData", DocX.a.NamespaceName), New XAttribute("uri", DocX.c.NamespaceName), New XElement(XName.Get("chart", DocX.c.NamespaceName), New XAttribute(XName.Get("id", DocX.r.NamespaceName), relID)))))))
			p.Xml.Add(chartElement)
		End Sub

		''' <summary>
		''' Inserts a default TOC into the current document.
		''' Title: Table of contents
		''' Swithces will be: TOC \h \o '1-3' \u \z
		''' </summary>
		''' <returns>The inserted TableOfContents</returns>
		Public Function InsertDefaultTableOfContents() As TableOfContents
			Return InsertTableOfContents("Table of contents", TableOfContentsSwitches.O Or TableOfContentsSwitches.H Or TableOfContentsSwitches.Z Or TableOfContentsSwitches.U)
		End Function

		''' <summary>
		''' Inserts a TOC into the current document.
		''' </summary>
		''' <param name="title">The title of the TOC</param>
		''' <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
		''' <param name="headerStyle">Lets you set the style name of the TOC header</param>
		''' <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
		''' <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
		''' <returns>The inserted TableOfContents</returns>
		Public Function InsertTableOfContents(ByVal title As String, ByVal switches As TableOfContentsSwitches, Optional ByVal headerStyle As String = Nothing, Optional ByVal maxIncludeLevel As Integer = 3, Optional ByVal rightTabPos? As Integer = Nothing) As TableOfContents
			Dim toc = TableOfContents.CreateTableOfContents(Me, title, switches, headerStyle, maxIncludeLevel, rightTabPos)
			Xml.Add(toc.Xml)
			Return toc
		End Function

		''' <summary>
		''' Inserts at TOC into the current document before the provided <paramref name="reference"/>
		''' </summary>
		''' <param name="reference">The paragraph to use as reference</param>
		''' <param name="title">The title of the TOC</param>
		''' <param name="switches">Switches to be applied, see: http://officeopenxml.com/WPtableOfContents.php </param>
		''' <param name="headerStyle">Lets you set the style name of the TOC header</param>
		''' <param name="maxIncludeLevel">Lets you specify how many header levels should be included - default is 1-3</param>
		''' <param name="rightTabPos">Lets you override the right tab position - this is not common</param>
		''' <returns>The inserted TableOfContents</returns>
		Public Function InsertTableOfContents(ByVal reference As Paragraph, ByVal title As String, ByVal switches As TableOfContentsSwitches, Optional ByVal headerStyle As String = Nothing, Optional ByVal maxIncludeLevel As Integer = 3, Optional ByVal rightTabPos? As Integer = Nothing) As TableOfContents
			Dim toc = TableOfContents.CreateTableOfContents(Me, title, switches, headerStyle, maxIncludeLevel, rightTabPos)
			reference.Xml.AddBeforeSelf(toc.Xml)
			Return toc
		End Function

		#Region "IDisposable Members"

		''' <summary>
		''' Releases all resources used by this document.
		''' </summary>
		''' <example>
		''' If you take advantage of the using keyword, Dispose() is automatically called for you.
		''' <code>
		''' // Load document.
		''' using (DocX document = DocX.Load(@"C:\Example\Test.docx"))
		''' {
		'''      // The document is only in memory while in this scope.
		'''
		''' }// Dispose() is automatically called at this point.
		''' </code>
		''' </example>
		''' <example>
		''' This example is equilivant to the one above example.
		''' <code>
		''' // Load document.
		''' DocX document = DocX.Load(@"C:\Example\Test.docx");
		''' 
		''' // Do something with the document here.
		'''
		''' // Dispose of the document.
		''' document.Dispose();
		''' </code>
		''' </example>
		Public Sub Dispose() Implements IDisposable.Dispose
			package.Close()
		End Sub

		#End Region
	End Class
End Namespace
