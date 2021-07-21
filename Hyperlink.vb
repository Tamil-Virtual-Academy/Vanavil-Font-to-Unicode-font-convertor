Imports System.Text
Imports System.IO.Packaging

Namespace Novacode
	''' <summary>
	''' Represents a Hyperlink in a document.
	''' </summary>
	Public Class Hyperlink
		Inherits DocXElement
'INSTANT VB NOTE: The variable uri was renamed since Visual Basic does not allow class members with the same name:
		Friend uri_Renamed As Uri
'INSTANT VB NOTE: The variable text was renamed since Visual Basic does not allow class members with the same name:
		Friend text_Renamed As String

		Friend hyperlink_rels As Dictionary(Of PackagePart, PackageRelationship)
		Friend type As Integer
		Friend id As String
		Friend instrText As XElement
		Friend runs As List(Of XElement)

		''' <summary>
		''' Remove a Hyperlink from this Paragraph only.
		''' </summary>
		''' <example>
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Create(@"Test.docx"))
		''' {
		'''    // Add a hyperlink to this document.
		'''    Hyperlink h = document.AddHyperlink("link", new Uri("http://www.google.com"));
		'''
		'''    // Add a Paragraph to this document and insert the hyperlink
		'''    Paragraph p1 = document.InsertParagraph();
		'''    p1.Append("This is a cool ").AppendHyperlink(h).Append(" .");
		'''
		'''    /* 
		'''     * Remove the hyperlink from this Paragraph only. 
		'''     * Note a reference to the hyperlink will still exist in the document and it can thus be reused.
		'''     */
		'''    p1.Hyperlinks[0].Remove();
		'''
		'''    // Add a new Paragraph to this document and reuse the hyperlink h.
		'''    Paragraph p2 = document.InsertParagraph();
		'''    p2.Append("This is the same cool ").AppendHyperlink(h).Append(" .");
		'''
		'''    document.Save();
		''' }// Release this document from memory.
		''' </code>
		''' </example>
		Public Sub Remove()
			Xml.Remove()
		End Sub

		''' <summary>
		''' Change the Text of a Hyperlink.
		''' </summary>
		''' <example>
		''' Change the Text of a Hyperlink.
		''' <code>
		''' // Create a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // Get all of the hyperlinks in this document
		'''    List&lt;Hyperlink&gt; hyperlinks = document.Hyperlinks;
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
		Public Property Text() As String
			Get
				Return Me.text_Renamed
			End Get

			Set(ByVal value As String)
				Dim rPr As New XElement(DocX.w + "rPr", New XElement (DocX.w + "rStyle", New XAttribute(DocX.w + "val", "Hyperlink")))

				' Format and add the new text.
				Dim newRuns As List(Of XElement) = HelperFunctions.FormatInput(value, rPr)

				If type = 0 Then
					' Get all the runs in this Text.
					Dim runs = From r In Xml.Elements()
					           Where r.Name.LocalName = "r"
					           Select r

					' Remove each run.
					For i As Integer = 0 To runs.Count() - 1
						runs.Remove()
					Next i

					Xml.Add(newRuns)

				Else
					Dim separate As XElement = XElement.Parse("" & ControlChars.CrLf & "                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "                        <w:fldChar w:fldCharType='separate'/> " & ControlChars.CrLf & "                    </w:r>")

					Dim [end] As XElement = XElement.Parse("" & ControlChars.CrLf & "                    <w:r xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" & ControlChars.CrLf & "                        <w:fldChar w:fldCharType='end' /> " & ControlChars.CrLf & "                    </w:r>")

					runs.Last().AddAfterSelf(separate, newRuns, [end])
					runs.ForEach(Function(r) r.Remove())
				End If

				Me.text_Renamed = value
			End Set
		End Property

		''' <summary>
		''' Change the Uri of a Hyperlink.
		''' </summary>
		''' <example>
		''' Change the Uri of a Hyperlink.
		''' <code>
		''' <![CDATA[
		''' // Create a document.
		''' using (DocX document = DocX.Load(@"Test.docx"))
		''' {
		'''    // Get all of the hyperlinks in this document
		'''    List<Hyperlink> hyperlinks = document.Hyperlinks;
		'''    
		'''    // Change the first hyperlinks text and Uri
		'''    Hyperlink h0 = hyperlinks[0];
		'''    h0.Text = "DocX";
		'''    h0.Uri = new Uri("http://docx.codeplex.com");
		'''
		'''    // Save this document.
		'''    document.Save();
		''' }
		''' ]]>
		''' </code>
		''' </example>
		Public Property Uri() As Uri
			Get
				If type = 0 AndAlso id <> String.Empty Then
					Dim r As PackageRelationship = mainPart.GetRelationship(id)
					Return r.TargetUri
				End If

				Return Me.uri_Renamed
			End Get

			Set(ByVal value As Uri)
				If type = 0 Then
					Dim r As PackageRelationship = mainPart.GetRelationship(id)

					' Get all of the information about this relationship.
					Dim r_tm As TargetMode = r.TargetMode
					Dim r_rt As String = r.RelationshipType
					Dim r_id As String = r.Id

					' Delete the relationship
					mainPart.DeleteRelationship(r_id)
					mainPart.CreateRelationship(value, r_tm, r_rt, r_id)

				Else
					instrText.Value = "HYPERLINK " & """" & value.ToString() & """"
				End If

				Me.uri_Renamed = value
			End Set
		End Property

		Friend Sub New(ByVal document As DocX, ByVal mainPart As PackagePart, ByVal i As XElement)
			MyBase.New(document, i)
			Me.type = 0
			Me.id = i.Attribute(XName.Get("id", DocX.r.NamespaceName)).Value

			Dim sb As New StringBuilder()
			HelperFunctions.GetTextRecursive(i, sb)
			Me.text_Renamed = sb.ToString()
		End Sub

		Friend Sub New(ByVal document As DocX, ByVal instrText As XElement, ByVal runs As List(Of XElement))
			MyBase.New(document, Nothing)
			Me.type = 1
			Me.instrText = instrText
			Me.runs = runs

			Try
				Dim start As Integer = instrText.Value.IndexOf("HYPERLINK """) & "HYPERLINK """.Length
				Dim [end] As Integer = instrText.Value.IndexOf("""", start)
				If start <> -1 AndAlso [end] <> -1 Then
					Me.uri_Renamed = New Uri(instrText.Value.Substring(start, [end] - start), UriKind.Absolute)

					Dim sb As New StringBuilder()
					HelperFunctions.GetTextRecursive(New XElement(XName.Get("temp", DocX.w.NamespaceName), runs), sb)
					Me.text_Renamed = sb.ToString()
				End If

			Catch e As Exception
				Throw e
			End Try
		End Sub
	End Class
End Namespace
