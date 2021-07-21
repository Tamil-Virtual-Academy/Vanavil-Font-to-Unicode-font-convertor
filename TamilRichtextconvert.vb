Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Drawing.Text
Imports System.Drawing.Printing
Imports System.IO
Imports System.Windows.Forms
Imports System.Threading
Imports System.Text
Imports System.Xml
Imports System.Globalization
Imports System.Diagnostics
Imports GemBox.Document
Imports GemBox.Document.Tables
'Imports GemBox.Spreadsheet
'Imports GemBox.Presentation
'Imports GemBox.Presentation.Tables
'Imports HtmlAgilityPack
Imports System.Collections.Generic
Imports System.Linq
'Imports System.Linq

Public Class TamilRichtextconvert

    'Public isunicodechecked As Boolean = True
    Public drag_drop As Boolean = False
    Public file_path As String
    Public inputdocument As DocumentModel
    'Dim totxmls As Integer = 0
    'Dim savefrom As String
    'Dim saveto As String
    'Dim special As Boolean = False
    'Dim uninumber As Integer
    ''Dim mDocument As Document
    'Dim docfonts() As String
    'Dim src1() As String
    'Dim tgt1() As String
    'Dim src2() As String
    'Dim tgt2() As String
    'Dim uni1() As String
    'Dim uni2() As String
    'Dim srcc1() As String
    'Dim srcc2() As String
    'Dim src_count As Integer
    'Dim tgt_count As Integer
    'Dim uni_count As Integer
    'Dim srcc_count As Int16
    'Dim fromenc() As String
    'Dim toenc() As String
    'Dim fromfont() As String
    'Dim totamfont() As String
    'Dim toengfont() As String
    'Dim totamsizer() As Integer
    'Dim toengsizer() As Integer
    'Dim fromencstatus() As Boolean
    'Dim srcfont() As Boolean
    'Dim srcenc() As String
    'Dim srccum() As String
    'Dim srccount As Integer
    'Dim src(,,) As String
    'Dim srccnt() As Integer
    'Dim loaded As Boolean = False

    Public Sub loadd()

    End Sub
    'Public Sub OpenDocumentfilename(ByVal filename As String)
    '    'Dim fileName As String = SelectOpenFileName()
    '    'If fileName Is Nothing Then
    '    '    Return
    '    'End If

    '    ' This operation can take some time so we set the Cursor to WaitCursor.
    '    '  Application.DoEvents()
    '    '  Dim cursor As Cursor = cursor.Current
    '    '  cursor.Current = Cursors.WaitCursor

    '    ' Load document is put in a try-catch block to handle situations when it fails for some reason.
    '    Try
    '        ' Loads the document into Aspose.Words object model.
    '        mDocument = New Document(filename)
    '        ' viewopendoc_button.Enabled = True
    '        'GroupBox2.Visible = True
    '        Dim i As Integer
    '        ReDim docfonts(mDocument.FontInfos.Count)
    '        For i = 0 To mDocument.FontInfos.Count - 1
    '            docfonts(i) = mDocument.FontInfos(i).Name
    '            ' MsgBox(mDocument.FontInfos(i).Name)
    '            'mDocument.FontInfos(i).
    '        Next
    '        '''''''''''''''
    '        'Dim stylee As IWParagraphStyle = Nothing
    '        'sDocument = New WordDocument(fileName)

    '        'For i = 0 To sDocument.Styles.Count - 1
    '        '    If sDocument.Styles(i).StyleType = DLS.StyleType.CharacterStyle Then
    '        '        stylee = CType(sDocument.Styles(i), IWParagraphStyle)
    '        '        ' MsgBox(stylee.CharacterFormat.FontName)
    '        '    End If

    '        '    Next
    '    Catch ex As Exception
    '        ' CType(New ExceptionDialog(ex), ExceptionDialog).ShowDialog()
    '    End Try

    '    ' Restore cursor.
    '    ' Cursor.Current = Cursor
    'End Sub
    Private Function totext(ByVal hextext As String) As String
        Dim y As Integer
        Dim num As String = ""
        Dim value As String = ""
        For y = 1 To Len(hextext)
            value = value & ChrW(CInt("&h" & Mid(hextext, y, 4)))
            y = y + 3
        Next y
        totext = value
    End Function
    Private Function RemoveDuplicateChars(ByVal key As String) As String
        ' --- Removes duplicate chars using string concats. ---
        ' Store encountered letters in this string.
        Dim table As String = ""

        ' Store the result in this string.
        Dim result As String = ""

        ' Loop over each character.
        For Each value As Char In key
            ' See if character is in the table.
            If table.IndexOf(value) = -1 Then
                ' Append to the table and the result.
                table &= value
                result &= value
            End If
        Next value
        Return result
    End Function


    Public Sub processgemtovanavil(ByVal fname As String)
        Try
            Dim i As Integer

            ReDim fromenc(1)
            ReDim toenc(1)
            ReDim fromfont(1)
            ReDim totamfont(1)
            ReDim toengfont(1)
            ReDim totamsizer(1)
            ReDim toengsizer(1)
            ReDim fromencstatus(1)


            For i = 0 To 0

                fromenc(i) = "Unicode" 'rules_list.Items(i).SubItems(0).Text
                'toenc(i) = "Vanavil" ' rules_list.Items(i).SubItems(1).Text
                toenc(i) = "LT" ' rules_list.Items(i).SubItems(1).Text
                fromfont(i) = "Arial" 'rules_list.Items(i).SubItems(2).Text
                ' totamfont(i) = "VANAVIL-Avvaiyar" 'rules_list.Items(i).SubItems(3).Text
                totamfont(i) = "LAKSHMAN" 'rules_list.Items(i).SubItems(3).Text
                toengfont(i) = "Arial" ' rules_list.Items(i).SubItems(5).Text
                totamsizer(i) = 2 'rules_list.Items(i).SubItems(4).Text
                toengsizer(i) = 0 'rules_list.Items(i).SubItems(6).Text
                fromencstatus(i) = True 'srcfont(Array.IndexOf(srcenc, fromenc(i)))


                'MsgBox(fromfont(i) & "  " & fromenc(i) & "   " & fromencstatus(i))
            Next
            ' mDocument.JoinRunsWithSameFormatting()

            Dim stri As String
            'Dim strj As String
            'Dim tempi As Integer
            'Dim tempj As Integer
            'Dim tempk As Integer

            'Dim booli As Boolean

            GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

            Dim gemboxdocument As DocumentModel = DocumentModel.Load(fname)
            Dim vanavildocument As DocumentModel = New DocumentModel()


            Dim sb As New StringBuilder()

            'Dim para() As Paragraph = gemboxdocument.GetChildElements(True, ElementType.Paragraph)
            ' Dim vsection As Section = New Section(vanavildocument)



            Dim vparagraphs() As Paragraph
            ReDim Preserve vparagraphs(0)
            Dim vparagraphcount As Int64 = 0





            'For Each tablecell As GemBox.Document.Tables.TableCell In gemboxdocument.GetChildElements(True, ElementType.TableCell)
            '    MsgBox(tablecell.Content.ToString)
            'Next



            For Each paragraph As GemBox.Document.Paragraph In gemboxdocument.GetChildElements(True, ElementType.Paragraph)





                ' MsgBox(vparagraphcount)
                Dim vruns() As Run 'Run '= New Run(vanavildocument)
                ReDim Preserve vruns(0)
                Dim vruncount As Int64
                vruncount = 0

                '  For Each paragraph As GemBox.Document.Paragraph In para
                For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
                    stri = run.Text
                    Dim strilength As Int16 = stri.Length
                    Dim si As Int16 = 1
                    Dim from As Int16 = si
                    Dim too As Int16 = si
                    Dim flag As Boolean = isonlyunicode(Mid(stri, si, 1))
                    Dim styles As FontStyle = FontStyle.Regular

                    If (run.CharacterFormat.Bold) Then
                        styles = styles Or FontStyle.Bold
                    End If

                    If (run.CharacterFormat.Italic) Then
                        styles = styles Or FontStyle.Italic
                    End If

                    If (run.CharacterFormat.Strikethrough) Then
                        styles = styles Or FontStyle.Strikeout
                    End If
                    ' Dim newfontstyle As CharacterStyle =
                    While si < strilength
                        si += 1
                        If isonlyunicode(Mid(stri, si, 1)) <> flag Then
                            too = si - 1
                            If flag = True Then
                                vruncount += 1
                                ReDim Preserve vruns(vruncount)


                                vruns(vruncount - 1) = New Run(vanavildocument)
                                '  vruns(vruncount - 1) = run.Clone()
                                vruns(vruncount - 1).CharacterFormat.FontName = totamfont(0)
                                vruns(vruncount - 1).Text = convert(fromenc(0), toenc(0), Mid(stri, from, too - from + 1))
                                vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + totamsizer(0)

                                If (run.CharacterFormat.Bold) Then
                                    vruns(vruncount - 1).CharacterFormat.Bold = True
                                End If

                                If (run.CharacterFormat.Italic) Then
                                    vruns(vruncount - 1).CharacterFormat.Italic = True
                                End If

                                If (run.CharacterFormat.Strikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                                End If
                                If (run.CharacterFormat.DoubleStrikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                                End If
                                If (run.CharacterFormat.Subscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Subscript = True
                                End If
                                If (run.CharacterFormat.Superscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Superscript = True
                                End If
                                If (run.CharacterFormat.Hidden) Then
                                    vruns(vruncount - 1).CharacterFormat.Hidden = True
                                End If

                                vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle




                                ' vruns(vruncount - 1) = vrun

                                ' MsgBox(" out of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " is pure tamil Unicode")
                            Else
                                vruncount += 1
                                ReDim Preserve vruns(vruncount)


                                vruns(vruncount - 1) = New Run(vanavildocument)
                                vruns(vruncount - 1).CharacterFormat.FontName = toengfont(0)
                                vruns(vruncount - 1).Text = Mid(stri, from, too - from + 1)
                                vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + toengsizer(0)

                                If (run.CharacterFormat.Bold) Then
                                    vruns(vruncount - 1).CharacterFormat.Bold = True
                                End If

                                If (run.CharacterFormat.Italic) Then
                                    vruns(vruncount - 1).CharacterFormat.Italic = True
                                End If

                                If (run.CharacterFormat.Strikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                                End If
                                If (run.CharacterFormat.DoubleStrikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                                End If
                                If (run.CharacterFormat.Subscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Subscript = True
                                End If
                                If (run.CharacterFormat.Superscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Superscript = True
                                End If
                                If (run.CharacterFormat.Hidden) Then
                                    vruns(vruncount - 1).CharacterFormat.Hidden = True
                                End If

                                vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle

                                ' MsgBox(" out of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " is pure english")
                            End If
                            flag = isonlyunicode(Mid(stri, si, 1))
                            from = si

                        End If
                    End While
                    too = si '- 1

                    If flag = True Then
                        vruncount += 1
                        ReDim Preserve vruns(vruncount)


                        vruns(vruncount - 1) = New Run(vanavildocument)
                        vruns(vruncount - 1).CharacterFormat.FontName = totamfont(0)
                        vruns(vruncount - 1).Text = convert(fromenc(0), toenc(0), Mid(stri, from, too - from + 1))
                        vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + totamsizer(0)

                        If (run.CharacterFormat.Bold) Then
                            vruns(vruncount - 1).CharacterFormat.Bold = True
                        End If

                        If (run.CharacterFormat.Italic) Then
                            vruns(vruncount - 1).CharacterFormat.Italic = True
                        End If

                        If (run.CharacterFormat.Strikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                        End If
                        If (run.CharacterFormat.DoubleStrikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                        End If
                        If (run.CharacterFormat.Subscript) Then
                            vruns(vruncount - 1).CharacterFormat.Subscript = True
                        End If
                        If (run.CharacterFormat.Superscript) Then
                            vruns(vruncount - 1).CharacterFormat.Superscript = True
                        End If
                        If (run.CharacterFormat.Hidden) Then
                            vruns(vruncount - 1).CharacterFormat.Hidden = True
                        End If

                        vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle
                        ' MsgBox("(ll) out of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " is pure tamil Unicode")
                    Else
                        vruncount += 1
                        ReDim Preserve vruns(vruncount)


                        vruns(vruncount - 1) = New Run(vanavildocument)
                        vruns(vruncount - 1).CharacterFormat.FontName = toengfont(0)
                        vruns(vruncount - 1).Text = Mid(stri, from, too - from + 1)
                        vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + toengsizer(0)

                        If (run.CharacterFormat.Bold) Then
                            vruns(vruncount - 1).CharacterFormat.Bold = True
                        End If

                        If (run.CharacterFormat.Italic) Then
                            vruns(vruncount - 1).CharacterFormat.Italic = True
                        End If

                        If (run.CharacterFormat.Strikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                        End If
                        If (run.CharacterFormat.DoubleStrikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                        End If
                        If (run.CharacterFormat.Subscript) Then
                            vruns(vruncount - 1).CharacterFormat.Subscript = True
                        End If
                        If (run.CharacterFormat.Superscript) Then
                            vruns(vruncount - 1).CharacterFormat.Superscript = True
                        End If
                        If (run.CharacterFormat.Hidden) Then
                            vruns(vruncount - 1).CharacterFormat.Hidden = True
                        End If

                        vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle
                        ' MsgBox("(ll) out of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " is pure english")
                    End If

                Next
                ' Dim vparagraph As Paragraph
                vruns = vruns.Where(Function(c) c IsNot Nothing).ToArray()
                ' MsgBox("vruncount = " & vruncount)
                If vruncount > 0 Then
                    vparagraphcount += 1
                    ReDim Preserve vparagraphs(vparagraphcount)



                    'Dim vrunlist As IEnumerable(Of Inline) = vruns.ToList
                    '   "Unable to cast object of type 'System.Collections.Generic.List`1[GemBox.Document.Element]' to type 'System.Collections.Generic.IEnumerable`1[GemBox.Document.Inline]'."
                    vparagraphs(vparagraphcount - 1) = New Paragraph(vanavildocument, vruns)


                    'If vruncount > 0 Then
                    '    For vvi As Int64 = 1 To vruncount
                    '        vparagraphs(vparagraphcount - 1)
                    '    Next
                    '    'vruns
                    'End If
                End If
            Next
            ' MsgBox("vparagraphcountcount = " & vparagraphcount)
            vparagraphs = vparagraphs.Where(Function(c) c IsNot Nothing).ToArray()
            If vparagraphcount > 0 Then
                Dim vsection As New Section(vanavildocument, vparagraphs)

                vanavildocument.Sections.Add(vsection)

            End If
            If drag_drop = True Then

                vanavildocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
            Else
                If Not Directory.Exists(targetdir) Then
                    Directory.CreateDirectory(targetdir)
                End If
                vanavildocument.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
            End If
            ' gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())


            ' MsgBox("Done")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    'Public Sub processgemtovanavil(ByVal fname As String)
    '    Try
    '        Dim i As Integer

    '        ReDim fromenc(1)
    '        ReDim toenc(1)
    '        ReDim fromfont(1)
    '        ReDim totamfont(1)
    '        ReDim toengfont(1)
    '        ReDim totamsizer(1)
    '        ReDim toengsizer(1)
    '        ReDim fromencstatus(1)


    '        For i = 0 To 0

    '            fromenc(i) = "Unicode" 'rules_list.Items(i).SubItems(0).Text
    '            toenc(i) = "Vanavil" ' rules_list.Items(i).SubItems(1).Text
    '            fromfont(i) = "Arial" 'rules_list.Items(i).SubItems(2).Text
    '            totamfont(i) = "VANAVIL-Avvaiyar" 'rules_list.Items(i).SubItems(3).Text
    '            toengfont(i) = "Arial" ' rules_list.Items(i).SubItems(5).Text
    '            totamsizer(i) = 2 'rules_list.Items(i).SubItems(4).Text
    '            toengsizer(i) = 0 'rules_list.Items(i).SubItems(6).Text
    '            fromencstatus(i) = True 'srcfont(Array.IndexOf(srcenc, fromenc(i)))


    '            'MsgBox(fromfont(i) & "  " & fromenc(i) & "   " & fromencstatus(i))
    '        Next
    '        ' mDocument.JoinRunsWithSameFormatting()

    '        Dim stri As String
    '        Dim strj As String
    '        Dim tempi As Integer
    '        Dim tempj As Integer
    '        Dim tempk As Integer

    '        Dim booli As Boolean

    '        GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

    '        Dim gemboxdocument As DocumentModel = DocumentModel.Load(fname)

    '        Dim sb As New StringBuilder()

    '        For Each paragraph As GemBox.Document.Paragraph In gemboxdocument.GetChildElements(True, ElementType.Paragraph)
    '            For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
    '                'Dim isBold As Boolean = run.CharacterFormat.Bold
    '                'Dim text As String = run.Text
    '                'If run.Text = "தங்களை" Then run.Text = "thangalai"
    '                'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
    '                tempi = -1

    '                tempi = 0

    '                'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

    '                If tempi > -1 Then
    '                    If fromencstatus(tempi) Then
    '                        ' MsgBox(run.Text & "  " & run.Font.Name)

    '                        stri = run.Text
    '                        tempk = 1
    '                        MsgBox(stri & "  " & run.CharacterFormat.FontName & " " & AscW(Mid(stri, tempk, 1)))

    '                        tempj = Len(stri)
    '                        tempk = 1
    '                        booli = morethan127(Mid(stri, tempk, 1))
    '                        strj = Mid(stri, tempk, 1)

    '                        ''''''''''


    '                        If fromenc(tempi) = "Unicode" Then
    '                            If morethan127(run.Text) Then ' If isthereunicode(run.Text) Then
    '                                run.CharacterFormat.FontName = totamfont(tempi)
    '                                ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")
    '                                run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
    '                                run.CharacterFormat.Size += totamsizer(tempi)
    '                            Else
    '                                If IsNumeric(Trim(run.Text)) Then
    '                                    run.CharacterFormat.FontName = totamfont(tempi)
    '                                Else
    '                                    'MsgBox(run.Text)
    '                                    run.CharacterFormat.FontName = toengfont(tempi)
    '                                End If

    '                            End If
    '                        Else
    '                            run.CharacterFormat.FontName = totamfont(tempi)
    '                            '  run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

    '                        End If

    '                        '   run.Font.Name = totamfont(tempi)

    '                        ' run.CharacterFormat.Size += totamsizer(tempi)

    '                        ' run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

    '                    Else
    '                        MsgBox("what? Unicode is not source?")
    '                        If fromenc(tempi) = "Unicode" Then
    '                            If isthereunicode(run.Text) Then
    '                                run.CharacterFormat.FontName = totamfont(tempi)
    '                                ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

    '                            Else
    '                                run.CharacterFormat.FontName = toengfont(tempi)
    '                            End If
    '                        Else
    '                            run.CharacterFormat.FontName = totamfont(tempi)
    '                            ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

    '                        End If
    '                        ' run.Font.Name = totamfont(tempi)
    '                        run.CharacterFormat.Size += totamsizer(tempi)
    '                        run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
    '                    End If
    '                End If
    '            Next

    '        Next
    '        If drag_drop = True Then

    '            gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
    '        Else
    '            If Not Directory.Exists(targetdir) Then
    '                Directory.CreateDirectory(targetdir)
    '            End If
    '            gemboxdocument.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
    '        End If
    '        ' gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())


    '        ' MsgBox("Done")

    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try

    'End Sub
    Public Function processgemdoctovanavil(gemboxdocument As DocumentModel) As MemoryStream
        Try
            Dim i As Integer

            ReDim fromenc(1)
            ReDim toenc(1)
            ReDim fromfont(1)
            ReDim totamfont(1)
            ReDim toengfont(1)
            ReDim totamsizer(1)
            ReDim toengsizer(1)
            ReDim fromencstatus(1)
            Dim privateFonts As New System.Drawing.Text.PrivateFontCollection()
            ' privateFonts.AddFontFile(Application.StartupPath & "\" & "vanavilavvai.ttf")
            privateFonts.AddFontFile(Application.StartupPath & "\" & "LAKSHMAN.ttf")
            ' privateFonts.AddFontFile(Application.StartupPath & "\" & "Vanavil.ttf")
            'privateFonts.AddFontFile(Application.StartupPath & "\" & "vanaviltemp.ttf")
            Dim font As New System.Drawing.Font(privateFonts.Families(0), 12)
            ' MsgBox(font.Name)

            For i = 0 To 0

                fromenc(i) = "Unicode" 'rules_list.Items(i).SubItems(0).Text
                'toenc(i) = "Vanavil" ' rules_list.Items(i).SubItems(1).Text
                toenc(i) = "LT" ' rules_list.Items(i).SubItems(1).Text
                fromfont(i) = "Arial" 'rules_list.Items(i).SubItems(2).Text
                totamfont(i) = font.Name 'tempfont '"VANAVIL-Avvaiyar" 'rules_list.Items(i).SubItems(3).Text
                toengfont(i) = "Arial" ' rules_list.Items(i).SubItems(5).Text
                totamsizer(i) = 2 'rules_list.Items(i).SubItems(4).Text
                toengsizer(i) = 0 'rules_list.Items(i).SubItems(6).Text
                fromencstatus(i) = True 'srcfont(Array.IndexOf(srcenc, fromenc(i)))


                'MsgBox(fromfont(i) & "  " & fromenc(i) & "   " & fromencstatus(i))
            Next
            ' mDocument.JoinRunsWithSameFormatting()

            Dim stri As String
            'Dim strj As String
            'Dim tempi As Integer
            'Dim tempj As Integer
            'Dim tempk As Integer

            'Dim booli As Boolean

            GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

            'Dim gemboxdocument As DocumentModel = DocumentModel.Load(fname)
            Dim vanavildocument As DocumentModel = New DocumentModel()


            Dim sb As New StringBuilder()

            'Dim para() As Paragraph = gemboxdocument.GetChildElements(True, ElementType.Paragraph)
            ' Dim vsection As Section = New Section(vanavildocument)



            Dim vparagraphs() As Paragraph
            ReDim Preserve vparagraphs(0)
            Dim vparagraphcount As Int64 = 0





            'For Each tablecell As GemBox.Document.Tables.TableCell In gemboxdocument.GetChildElements(True, ElementType.TableCell)
            '    MsgBox(tablecell.Content.ToString)
            'Next



            For Each paragraph As GemBox.Document.Paragraph In gemboxdocument.GetChildElements(True, ElementType.Paragraph)





                ' MsgBox(vparagraphcount)
                Dim vruns() As Run 'Run '= New Run(vanavildocument)
                ReDim Preserve vruns(0)
                Dim vruncount As Int64
                vruncount = 0

                '  For Each paragraph As GemBox.Document.Paragraph In para
                For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
                    stri = run.Text
                    Dim strilength As Int16 = stri.Length
                    Dim si As Int16 = 1
                    Dim from As Int16 = si
                    Dim too As Int16 = si
                    Dim flag As Boolean = isonlyunicode(Mid(stri, si, 1))
                    Dim styles As FontStyle = FontStyle.Regular

                    If (run.CharacterFormat.Bold) Then
                        styles = styles Or FontStyle.Bold
                    End If

                    If (run.CharacterFormat.Italic) Then
                        styles = styles Or FontStyle.Italic
                    End If

                    If (run.CharacterFormat.Strikethrough) Then
                        styles = styles Or FontStyle.Strikeout
                    End If
                    ' Dim newfontstyle As CharacterStyle =
                    While si < strilength
                        si += 1
                        If isonlyunicode(Mid(stri, si, 1)) <> flag Then
                            too = si - 1
                            If flag = True Then
                                vruncount += 1
                                ReDim Preserve vruns(vruncount)


                                vruns(vruncount - 1) = New Run(vanavildocument)

                                '  vruns(vruncount - 1) = run.Clone()
                                vruns(vruncount - 1).CharacterFormat.FontName = totamfont(0)
                                vruns(vruncount - 1).Text = convert(fromenc(0), toenc(0), Mid(stri, from, too - from + 1))
                                vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + totamsizer(0)

                                If (run.CharacterFormat.Bold) Then
                                    vruns(vruncount - 1).CharacterFormat.Bold = True
                                End If

                                If (run.CharacterFormat.Italic) Then
                                    vruns(vruncount - 1).CharacterFormat.Italic = True
                                End If

                                If (run.CharacterFormat.Strikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                                End If
                                If (run.CharacterFormat.DoubleStrikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                                End If
                                If (run.CharacterFormat.Subscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Subscript = True
                                End If
                                If (run.CharacterFormat.Superscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Superscript = True
                                End If
                                If (run.CharacterFormat.Hidden) Then
                                    vruns(vruncount - 1).CharacterFormat.Hidden = True
                                End If

                                vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle




                                ' vruns(vruncount - 1) = vrun

                                ' MsgBox(" out Of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " Is pure tamil Unicode")
                            Else
                                vruncount += 1
                                ReDim Preserve vruns(vruncount)


                                vruns(vruncount - 1) = New Run(vanavildocument)
                                vruns(vruncount - 1).CharacterFormat.FontName = toengfont(0)
                                vruns(vruncount - 1).Text = Mid(stri, from, too - from + 1)
                                vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + toengsizer(0)

                                If (run.CharacterFormat.Bold) Then
                                    vruns(vruncount - 1).CharacterFormat.Bold = True
                                End If

                                If (run.CharacterFormat.Italic) Then
                                    vruns(vruncount - 1).CharacterFormat.Italic = True
                                End If

                                If (run.CharacterFormat.Strikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                                End If
                                If (run.CharacterFormat.DoubleStrikethrough) Then
                                    vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                                End If
                                If (run.CharacterFormat.Subscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Subscript = True
                                End If
                                If (run.CharacterFormat.Superscript) Then
                                    vruns(vruncount - 1).CharacterFormat.Superscript = True
                                End If
                                If (run.CharacterFormat.Hidden) Then
                                    vruns(vruncount - 1).CharacterFormat.Hidden = True
                                End If

                                vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                                vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle

                                ' MsgBox(" out Of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " Is pure english")
                            End If
                            flag = isonlyunicode(Mid(stri, si, 1))
                            from = si

                        End If
                    End While
                    too = si '- 1

                    If flag = True Then
                        vruncount += 1
                        ReDim Preserve vruns(vruncount)


                        vruns(vruncount - 1) = New Run(vanavildocument)
                        vruns(vruncount - 1).CharacterFormat.FontName = totamfont(0)
                        vruns(vruncount - 1).Text = convert(fromenc(0), toenc(0), Mid(stri, from, too - from + 1))
                        vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + totamsizer(0)

                        If (run.CharacterFormat.Bold) Then
                            vruns(vruncount - 1).CharacterFormat.Bold = True
                        End If

                        If (run.CharacterFormat.Italic) Then
                            vruns(vruncount - 1).CharacterFormat.Italic = True
                        End If

                        If (run.CharacterFormat.Strikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                        End If
                        If (run.CharacterFormat.DoubleStrikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                        End If
                        If (run.CharacterFormat.Subscript) Then
                            vruns(vruncount - 1).CharacterFormat.Subscript = True
                        End If
                        If (run.CharacterFormat.Superscript) Then
                            vruns(vruncount - 1).CharacterFormat.Superscript = True
                        End If
                        If (run.CharacterFormat.Hidden) Then
                            vruns(vruncount - 1).CharacterFormat.Hidden = True
                        End If

                        vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle
                        ' MsgBox("(ll) out Of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " Is pure tamil Unicode")
                    Else
                        vruncount += 1
                        ReDim Preserve vruns(vruncount)


                        vruns(vruncount - 1) = New Run(vanavildocument)
                        vruns(vruncount - 1).CharacterFormat.FontName = toengfont(0)
                        vruns(vruncount - 1).Text = Mid(stri, from, too - from + 1)
                        vruns(vruncount - 1).CharacterFormat.Size = run.CharacterFormat.Size + toengsizer(0)

                        If (run.CharacterFormat.Bold) Then
                            vruns(vruncount - 1).CharacterFormat.Bold = True
                        End If

                        If (run.CharacterFormat.Italic) Then
                            vruns(vruncount - 1).CharacterFormat.Italic = True
                        End If

                        If (run.CharacterFormat.Strikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.Strikethrough = True
                        End If
                        If (run.CharacterFormat.DoubleStrikethrough) Then
                            vruns(vruncount - 1).CharacterFormat.DoubleStrikethrough = True
                        End If
                        If (run.CharacterFormat.Subscript) Then
                            vruns(vruncount - 1).CharacterFormat.Subscript = True
                        End If
                        If (run.CharacterFormat.Superscript) Then
                            vruns(vruncount - 1).CharacterFormat.Superscript = True
                        End If
                        If (run.CharacterFormat.Hidden) Then
                            vruns(vruncount - 1).CharacterFormat.Hidden = True
                        End If

                        vruns(vruncount - 1).CharacterFormat.FontColor = run.CharacterFormat.FontColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineColor = run.CharacterFormat.UnderlineColor
                        vruns(vruncount - 1).CharacterFormat.UnderlineStyle = run.CharacterFormat.UnderlineStyle
                        ' MsgBox("(ll) out Of " & vbNewLine & stri & vbNewLine & "the word" & vbNewLine & Mid(stri, from, too - from + 1) & vbNewLine & " Is pure english")
                    End If

                Next
                ' Dim vparagraph As Paragraph
                vruns = vruns.Where(Function(c) c IsNot Nothing).ToArray()
                ' MsgBox("vruncount = " & vruncount)
                If vruncount > 0 Then
                    vparagraphcount += 1
                    ReDim Preserve vparagraphs(vparagraphcount)



                    'Dim vrunlist As IEnumerable(Of Inline) = vruns.ToList
                    '   "Unable To cast Object Of type 'System.Collections.Generic.List`1[GemBox.Document.Element]' to type 'System.Collections.Generic.IEnumerable`1[GemBox.Document.Inline]'."
                    vparagraphs(vparagraphcount - 1) = New Paragraph(vanavildocument, vruns)


                    'If vruncount > 0 Then
                    '    For vvi As Int64 = 1 To vruncount
                    '        vparagraphs(vparagraphcount - 1)
                    '    Next
                    '    'vruns
                    'End If
                End If
            Next
            ' MsgBox("vparagraphcountcount = " & vparagraphcount)
            vparagraphs = vparagraphs.Where(Function(c) c IsNot Nothing).ToArray()
            If vparagraphcount > 0 Then
                Dim vsection As New Section(vanavildocument, vparagraphs)

                vanavildocument.Sections.Add(vsection)

            End If
            Dim stream2 As MemoryStream = New MemoryStream()

            ' Save document to RTF stream.
            vanavildocument.Save("test_vanavil" & ".rtf") 'SelectSaveFileName())
            vanavildocument.Save(stream2, SaveOptions.RtfDefault)
            stream2.Seek(0, SeekOrigin.Begin)
            Return stream2
            'Dim reader As StreamReader = New StreamReader(stream2)
            'Dim outputtext As String = reader.ReadToEnd()
            'Return outputtext
            'If drag_drop = True Then

            '    vanavildocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
            'Else
            '    If Not Directory.Exists(targetdir) Then
            '        Directory.CreateDirectory(targetdir)
            '    End If
            '    vanavildocument.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())
            'End If
            ' gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_vanavil" & ".docx") 'SelectSaveFileName())


            ' MsgBox("Done")

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function processgemdoctounicode(ByVal gemboxdocument As DocumentModel) As String
        Try
            Dim i As Integer

            ReDim fromenc(1)
            ReDim toenc(1)
            ReDim fromfont(1)
            ReDim totamfont(1)
            ReDim toengfont(1)
            ReDim totamsizer(1)
            ReDim toengsizer(1)
            ReDim fromencstatus(1)


            For i = 0 To 0

                ' fromenc(i) = "Vanavil" 'rules_list.Items(i).SubItems(0).Text
                fromenc(i) = "LT" 'rules_list.Items(i).SubItems(0).Text
                toenc(i) = "Unicode" ' rules_list.Items(i).SubItems(1).Text
                'fromfont(i) = "VANAVIL-Avvaiyar" 'rules_list.Items(i).SubItems(2).Text
                fromfont(i) = "LAKSHMAN" 'rules_list.Items(i).SubItems(2).Text
                totamfont(i) = "Arial Unicode MS" 'rules_list.Items(i).SubItems(3).Text
                toengfont(i) = "Arial Unicode MS" ' rules_list.Items(i).SubItems(5).Text
                totamsizer(i) = -2 'rules_list.Items(i).SubItems(4).Text
                toengsizer(i) = 0 'rules_list.Items(i).SubItems(6).Text
                fromencstatus(i) = True 'srcfont(Array.IndexOf(srcenc, fromenc(i)))


                'MsgBox(fromfont(i) & "  " & fromenc(i) & "   " & fromencstatus(i))
            Next
            ' mDocument.JoinRunsWithSameFormatting()

            Dim stri As String
            Dim strj As String
            Dim tempi As Integer
            Dim tempj As Integer
            Dim tempk As Integer

            Dim booli As Boolean


            'Select Case Path.GetExtension(fname)
            '    Case ".docx", ".doc", ".rtf"

            GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

            '  Dim gemboxdocument As DocumentModel = DocumentModel.Load(fname)

            Dim sb As New StringBuilder()

            For Each paragraph As GemBox.Document.Paragraph In gemboxdocument.GetChildElements(True, ElementType.Paragraph)
                For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
                    'Dim isBold As Boolean = run.CharacterFormat.Bold
                    'Dim text As String = run.Text
                    'If run.Text = "தங்களை" Then run.Text = "thangalai"
                    'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                    tempi = -1
                    'If (Mid(run.CharacterFormat.FontName, 1, 7).ToLower = "vanavil") Then tempi = 0
                    'If (Mid(run.CharacterFormat.FontName, 1, 7).ToLower = "LT-TM-Roja") Then tempi = 0

                    If run.CharacterFormat.FontName = "LT-TM-Roja" Then tempi = 0
                    'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

                    If tempi > -1 Then
                        If fromencstatus(tempi) Then
                            ' MsgBox(run.Text & "  " & run.Font.Name)

                            stri = run.Text
                            tempk = 1
                            ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

                            tempj = Len(stri)
                            tempk = 1
                            booli = morethan127(Mid(stri, tempk, 1))
                            strj = Mid(stri, tempk, 1)

                            ''''''''''


                            If fromenc(tempi) = "Unicode" Then
                                If isthereunicode(run.Text) Then
                                    run.CharacterFormat.FontName = totamfont(tempi)
                                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                Else
                                    run.CharacterFormat.FontName = toengfont(tempi)
                                End If
                            Else
                                run.CharacterFormat.FontName = totamfont(tempi)
                                run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                            End If

                            '   run.Font.Name = totamfont(tempi)

                            run.CharacterFormat.Size += totamsizer(tempi)
                            ' Dim tttt As String = run.Text

                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

                            'MsgBox(run.Text)
                        Else
                            If fromenc(tempi) = "Unicode" Then
                                If isthereunicode(run.Text) Then
                                    run.CharacterFormat.FontName = totamfont(tempi)
                                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                Else
                                    run.CharacterFormat.FontName = toengfont(tempi)
                                End If
                            Else
                                run.CharacterFormat.FontName = totamfont(tempi)
                                run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                            End If
                            ' run.Font.Name = totamfont(tempi)
                            run.CharacterFormat.Size += totamsizer(tempi)
                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
                        End If
                    End If
                Next

            Next
        Dim stream2 As MemoryStream = New MemoryStream()

        ' Save document to RTF stream.
        gemboxdocument.Save(stream2, SaveOptions.RtfDefault)
        stream2.Seek(0, SeekOrigin.Begin)

        Dim reader As StreamReader = New StreamReader(stream2)
        Dim outputtext As String = reader.ReadToEnd()
        Return outputtext
            'gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
            'If drag_drop = False Then
            '    If Not Directory.Exists(targetdir) Then
            '        Directory.CreateDirectory(targetdir)
            '    End If
            '    gemboxdocument.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
            'Else
            '    gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
            'End If
            ' MsgBox("Done")
            'Case ".xls", ".xlsx", ".ods", ".csv"

            '    SpreadsheetInfo.SetLicense("EQI3-BK2T-UZM5-17XP")


            '    Dim gemboxspreadsheet As ExcelFile = ExcelFile.Load(fname)
            '    For Each sheet As ExcelWorksheet In gemboxspreadsheet.Worksheets

            '        For Each row As ExcelRow In sheet.Rows
            '            For Each cell As ExcelCell In row.AllocatedCells

            '                If cell.ValueType <> CellValueType.Null Then
            '                    Dim str As String = GetHtmlFormattedValue(cell)
            '                    'MsgBox(str)
            '                    Dim resultstr As String = ""
            '                    Dim doc As HtmlAgilityPack.HtmlDocument
            '                    doc = New HtmlAgilityPack.HtmlDocument
            '                    doc.LoadHtml(str)
            '                    Dim q = New Queue(Of HtmlNode)()
            '                    q.Enqueue(doc.DocumentNode)
            '                    Do While q.Count > 0
            '                        Dim item As HtmlNode = q.Dequeue()
            '                        Dim node As HtmlNode = Nothing
            '                        ' MsgBox(item.OuterHtml)
            '                        If item.Name = "span" Then
            '                            Dim node1 As HtmlNode = HtmlTextNode.CreateNode(item.OuterHtml)
            '                            Dim node2 As String = HtmlTextNode.CreateNode(node1.GetAttributeValue("style", "")).InnerHtml

            '                            Dim spanstr As String = item.OuterHtml
            '                            Dim respanstr As String = ""
            '                            Dim innertextwithtag As String = HtmlAgilityPack.HtmlEntity.DeEntitize("<p>" & item.InnerText & "</p>")
            '                            Dim innertext As String = Mid(innertextwithtag, 4, innertextwithtag.Length - 7)
            '                            Dim convertedtext As String = innertext
            '                            Dim fontfamilyindex As Int16 = node2.IndexOf("font-family:")
            '                            Dim semicolonindex As Int16 = node2.IndexOf(";", fontfamilyindex + 5)
            '                            Dim fontfamily As String = Mid(node2, fontfamilyindex + 13, semicolonindex - fontfamilyindex - 12)
            '                            'MsgBox(item.OuterHtml)
            '                            If (Mid(fontfamily, 1, 7).ToLower = "vanavil") Then
            '                                fontfamily = "Times New Roman"
            '                                convertedtext = convert(fromenc(0), toenc(0), innertext)
            '                                respanstr = "<span style=" & Chr(34) & Mid(node2, 1, fontfamilyindex + 12) & fontfamily & Mid(node2, semicolonindex + 1) & Chr(34) & ">" & convertedtext & "</span>"

            '                            Else
            '                                respanstr = item.OuterHtml

            '                            End If

            '                            resultstr = resultstr & respanstr

            '                        Else

            '                            Dim xi As Int16 = 0
            '                            For Each child As HtmlNode In item.ChildNodes
            '                                xi += 1
            '                                q.Enqueue(child)
            '                            Next child
            '                            If xi = 0 Then
            '                                resultstr = item.OuterHtml
            '                            End If
            '                        End If
            '                    Loop
            '                    cell.SetValue(resultstr, GemBox.Spreadsheet.LoadOptions.HtmlDefault)
            '                Else

            '                End If
            '            Next
            '        Next
            '    Next
            '    If drag_drop = False Then
            '        If Not Directory.Exists(targetdir) Then
            '            Directory.CreateDirectory(targetdir)
            '        End If
            '        gemboxspreadsheet.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & Path.GetExtension(fname)) 'SelectSaveFileName())
            '    Else
            '        gemboxspreadsheet.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & Path.GetExtension(fname)) 'SelectSaveFileName())
            '    End If

            'Case ".pptx", ".ppt"

            '    GemBox.Presentation.ComponentInfo.SetLicense("EQI3-BK2T-UZM5-Q5XQ")

            '    Dim gemboxpresentation As PresentationDocument = PresentationDocument.Load(fname)

            '    Dim sb As New StringBuilder()
            '    ' MsgBox(gemboxpresentation.Slides.Count)
            '    For Each Slidee As GemBox.Presentation.Slide In gemboxpresentation.Slides
            '        ' Slidee.Content.DrawingType(DrawingType.UnknownDrawing)

            '        For Each shape As Shape In Slidee.Content.Drawings.OfType(Of Shape)

            '            For Each paragraph As TextParagraph In shape.Text.Paragraphs
            '                For Each run As TextRun In paragraph.Elements.OfType(Of TextRun)
            '                    'Dim isBold As Boolean = run.CharacterFormat.Bold
            '                    'Dim text As String = run.Text
            '                    'If run.Text = "தங்களை" Then run.Text = "thangalai"
            '                    'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
            '                    tempi = -1
            '                    'run.Format.Font
            '                    If (Mid(run.Format.Font, 1, 7).ToLower = "vanavil") Then tempi = 0

            '                    'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

            '                    If tempi > -1 Then
            '                        If fromencstatus(tempi) Then
            '                            ' MsgBox(run.Text & "  " & run.Font.Name)

            '                            stri = run.Text
            '                            tempk = 1
            '                            ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

            '                            tempj = Len(stri)
            '                            tempk = 1
            '                            booli = morethan127(Mid(stri, tempk, 1))
            '                            strj = Mid(stri, tempk, 1)

            '                            ''''''''''


            '                            If fromenc(tempi) = "Unicode" Then
            '                                If isthereunicode(run.Text) Then
            '                                    run.Format.Font = totamfont(tempi)
            '                                    ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '                                Else
            '                                    run.Format.Font = toengfont(tempi)
            '                                End If
            '                            Else
            '                                run.Format.Font = totamfont(tempi)
            '                                'run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '                            End If

            '                            '   run.Font.Name = totamfont(tempi)

            '                            ' run.CharacterFormat.Size += totamsizer(tempi)
            '                            ' Dim tttt As String = run.Text

            '                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

            '                            ' MsgBox(tttt & " " & run.Text)
            '                        Else
            '                            If fromenc(tempi) = "Unicode" Then
            '                                If isthereunicode(run.Text) Then
            '                                    run.Format.Font = totamfont(tempi)
            '                                    'run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '                                Else
            '                                    run.Format.Font = toengfont(tempi)
            '                                End If
            '                            Else
            '                                run.Format.Font = totamfont(tempi)
            '                                ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '                            End If
            '                            ' run.Font.Name = totamfont(tempi)
            '                            ' run.CharacterFormat.Size += totamsizer(tempi)
            '                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
            '                        End If
            '                    End If

            '                Next
            '            Next
            '        Next
            '    Next


            '    'For Each paragraph As GemBox.Document.Paragraph In gemboxpresentation.GetChildElements(True, ElementType.Paragraph)
            '    '    For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
            '    '        'Dim isBold As Boolean = run.CharacterFormat.Bold
            '    '        'Dim text As String = run.Text
            '    '        'If run.Text = "தங்களை" Then run.Text = "thangalai"
            '    '        'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
            '    '        tempi = -1
            '    '        If (Mid(run.CharacterFormat.FontName, 1, 7).ToLower = "vanavil") Then tempi = 0

            '    '        'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

            '    '        If tempi > -1 Then
            '    '            If fromencstatus(tempi) Then
            '    '                ' MsgBox(run.Text & "  " & run.Font.Name)

            '    '                stri = run.Text
            '    '                tempk = 1
            '    '                ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

            '    '                tempj = Len(stri)
            '    '                tempk = 1
            '    '                booli = morethan127(Mid(stri, tempk, 1))
            '    '                strj = Mid(stri, tempk, 1)

            '    '                ''''''''''


            '    '                If fromenc(tempi) = "Unicode" Then
            '    '                    If isthereunicode(run.Text) Then
            '    '                        run.CharacterFormat.FontName = totamfont(tempi)
            '    '                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '    '                    Else
            '    '                        run.CharacterFormat.FontName = toengfont(tempi)
            '    '                    End If
            '    '                Else
            '    '                    run.CharacterFormat.FontName = totamfont(tempi)
            '    '                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '    '                End If

            '    '                '   run.Font.Name = totamfont(tempi)

            '    '                run.CharacterFormat.Size += totamsizer(tempi)
            '    '                ' Dim tttt As String = run.Text

            '    '                run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

            '    '                ' MsgBox(tttt & " " & run.Text)
            '    '            Else
            '    '                If fromenc(tempi) = "Unicode" Then
            '    '                    If isthereunicode(run.Text) Then
            '    '                        run.CharacterFormat.FontName = totamfont(tempi)
            '    '                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '    '                    Else
            '    '                        run.CharacterFormat.FontName = toengfont(tempi)
            '    '                    End If
            '    '                Else
            '    '                    run.CharacterFormat.FontName = totamfont(tempi)
            '    '                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

            '    '                End If
            '    '                ' run.Font.Name = totamfont(tempi)
            '    '                run.CharacterFormat.Size += totamsizer(tempi)
            '    '                run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
            '    '            End If
            '    '        End If
            '    '    Next

            '    'Next
            '    'gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
            '    If drag_drop = False Then
            '        If Not Directory.Exists(targetdir) Then
            '            Directory.CreateDirectory(targetdir)
            '        End If
            '        gemboxpresentation.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".pptx") 'SelectSaveFileName())
            '    Else
            '        gemboxpresentation.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".pptx") 'SelectSaveFileName())
            '    End If
            ' MsgBox("Done")
            ' End Select



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Sub processgemtounicode(ByVal fname As String)
        Try
            Dim i As Integer

            ReDim fromenc(1)
            ReDim toenc(1)
            ReDim fromfont(1)
            ReDim totamfont(1)
            ReDim toengfont(1)
            ReDim totamsizer(1)
            ReDim toengsizer(1)
            ReDim fromencstatus(1)


            For i = 0 To 0

                fromenc(i) = "LT-TM-Lakshman" 'rules_list.Items(i).SubItems(0).Text
                toenc(i) = "Unicode" ' rules_list.Items(i).SubItems(1).Text
                fromfont(i) = "LT-TM-Lakshman" 'rules_list.Items(i).SubItems(2).Text
                totamfont(i) = "Arial" 'rules_list.Items(i).SubItems(3).Text
                toengfont(i) = "Arial" ' rules_list.Items(i).SubItems(5).Text
                totamsizer(i) = -2 'rules_list.Items(i).SubItems(4).Text
                toengsizer(i) = 0 'rules_list.Items(i).SubItems(6).Text
                fromencstatus(i) = True 'srcfont(Array.IndexOf(srcenc, fromenc(i)))


                'MsgBox(fromfont(i) & "  " & fromenc(i) & "   " & fromencstatus(i))
            Next
            ' mDocument.JoinRunsWithSameFormatting()

            Dim stri As String
            Dim strj As String
            Dim tempi As Integer
            Dim tempj As Integer
            Dim tempk As Integer

            Dim booli As Boolean


            Select Case Path.GetExtension(fname)
                Case ".docx", ".doc", ".rtf"

                    GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

                    Dim gemboxdocument As DocumentModel = DocumentModel.Load(fname, LoadOptions.DocxDefault)

                    Dim sb As New StringBuilder()

                    For Each paragraph As GemBox.Document.Paragraph In gemboxdocument.GetChildElements(True, ElementType.Paragraph)
                        For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
                            'Dim isBold As Boolean = run.CharacterFormat.Bold
                            'Dim text As String = run.Text
                            'If run.Text = "தங்களை" Then run.Text = "thangalai"
                            'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                            tempi = -1
                            If (Mid(run.CharacterFormat.FontName, 1, 7).ToLower = "LT-TM-Lakshman") Then tempi = 0

                            'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

                            If tempi > -1 Then
                                If fromencstatus(tempi) Then
                                    ' MsgBox(run.Text & "  " & run.Font.Name)

                                    stri = run.Text
                                    tempk = 1
                                    ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

                                    tempj = Len(stri)
                                    tempk = 1
                                    booli = morethan127(Mid(stri, tempk, 1))
                                    strj = Mid(stri, tempk, 1)

                                    ''''''''''


                                    If fromenc(tempi) = "Unicode" Then
                                        If isthereunicode(run.Text) Then
                                            run.CharacterFormat.FontName = totamfont(tempi)
                                            run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                        Else
                                            run.CharacterFormat.FontName = toengfont(tempi)
                                        End If
                                    Else
                                        run.CharacterFormat.FontName = totamfont(tempi)
                                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                    End If

                                    '   run.Font.Name = totamfont(tempi)

                                    run.CharacterFormat.Size += totamsizer(tempi)
                                    ' Dim tttt As String = run.Text

                                    run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

                                    ' MsgBox(tttt & " " & run.Text)
                                Else
                                    If fromenc(tempi) = "Unicode" Then
                                        If isthereunicode(run.Text) Then
                                            run.CharacterFormat.FontName = totamfont(tempi)
                                            run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                        Else
                                            run.CharacterFormat.FontName = toengfont(tempi)
                                        End If
                                    Else
                                        run.CharacterFormat.FontName = totamfont(tempi)
                                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                                    End If
                                    ' run.Font.Name = totamfont(tempi)
                                    run.CharacterFormat.Size += totamsizer(tempi)
                                    run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
                                End If
                            End If
                        Next

                    Next
                    'gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
                    If drag_drop = False Then
                        If Not Directory.Exists(targetdir) Then
                            Directory.CreateDirectory(targetdir)
                        End If
                        gemboxdocument.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
                    Else
                        gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
                    End If
                    ' MsgBox("Done")
                    'Case ".xls", ".xlsx", ".ods", ".csv"

                    '    SpreadsheetInfo.SetLicense("EQI3-BK2T-UZM5-17XP")


                    '    Dim gemboxspreadsheet As ExcelFile = ExcelFile.Load(fname)
                    '    For Each sheet As ExcelWorksheet In gemboxspreadsheet.Worksheets

                    '        For Each row As ExcelRow In sheet.Rows
                    '            For Each cell As ExcelCell In row.AllocatedCells

                    '                If cell.ValueType <> CellValueType.Null Then
                    '                    Dim str As String = GetHtmlFormattedValue(cell)
                    '                    'MsgBox(str)
                    '                    Dim resultstr As String = ""
                    '                    Dim doc As HtmlAgilityPack.HtmlDocument
                    '                    doc = New HtmlAgilityPack.HtmlDocument
                    '                    doc.LoadHtml(str)
                    '                    Dim q = New Queue(Of HtmlNode)()
                    '                    q.Enqueue(doc.DocumentNode)
                    '                    Do While q.Count > 0
                    '                        Dim item As HtmlNode = q.Dequeue()
                    '                        Dim node As HtmlNode = Nothing
                    '                        ' MsgBox(item.OuterHtml)
                    '                        If item.Name = "span" Then
                    '                            Dim node1 As HtmlNode = HtmlTextNode.CreateNode(item.OuterHtml)
                    '                            Dim node2 As String = HtmlTextNode.CreateNode(node1.GetAttributeValue("style", "")).InnerHtml

                    '                            Dim spanstr As String = item.OuterHtml
                    '                            Dim respanstr As String = ""
                    '                            Dim innertextwithtag As String = HtmlAgilityPack.HtmlEntity.DeEntitize("<p>" & item.InnerText & "</p>")
                    '                            Dim innertext As String = Mid(innertextwithtag, 4, innertextwithtag.Length - 7)
                    '                            Dim convertedtext As String = innertext
                    '                            Dim fontfamilyindex As Int16 = node2.IndexOf("font-family:")
                    '                            Dim semicolonindex As Int16 = node2.IndexOf(";", fontfamilyindex + 5)
                    '                            Dim fontfamily As String = Mid(node2, fontfamilyindex + 13, semicolonindex - fontfamilyindex - 12)
                    '                            'MsgBox(item.OuterHtml)
                    '                            If (Mid(fontfamily, 1, 7).ToLower = "vanavil") Then
                    '                                fontfamily = "Times New Roman"
                    '                                convertedtext = convert(fromenc(0), toenc(0), innertext)
                    '                                respanstr = "<span style=" & Chr(34) & Mid(node2, 1, fontfamilyindex + 12) & fontfamily & Mid(node2, semicolonindex + 1) & Chr(34) & ">" & convertedtext & "</span>"

                    '                            Else
                    '                                respanstr = item.OuterHtml

                    '                            End If

                    '                            resultstr = resultstr & respanstr

                    '                        Else

                    '                            Dim xi As Int16 = 0
                    '                            For Each child As HtmlNode In item.ChildNodes
                    '                                xi += 1
                    '                                q.Enqueue(child)
                    '                            Next child
                    '                            If xi = 0 Then
                    '                                resultstr = item.OuterHtml
                    '                            End If
                    '                        End If
                    '                    Loop
                    '                    cell.SetValue(resultstr, GemBox.Spreadsheet.LoadOptions.HtmlDefault)
                    '                Else

                    '                End If
                    '            Next
                    '        Next
                    '    Next
                    '    If drag_drop = False Then
                    '        If Not Directory.Exists(targetdir) Then
                    '            Directory.CreateDirectory(targetdir)
                    '        End If
                    '        gemboxspreadsheet.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & Path.GetExtension(fname)) 'SelectSaveFileName())
                    '    Else
                    '        gemboxspreadsheet.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & Path.GetExtension(fname)) 'SelectSaveFileName())
                    '    End If

                    'Case ".pptx", ".ppt"

                    '    GemBox.Presentation.ComponentInfo.SetLicense("EQI3-BK2T-UZM5-Q5XQ")

                    '    Dim gemboxpresentation As PresentationDocument = PresentationDocument.Load(fname)

                    '    Dim sb As New StringBuilder()
                    '    ' MsgBox(gemboxpresentation.Slides.Count)
                    '    For Each Slidee As GemBox.Presentation.Slide In gemboxpresentation.Slides
                    '        ' Slidee.Content.DrawingType(DrawingType.UnknownDrawing)

                    '        For Each shape As Shape In Slidee.Content.Drawings.OfType(Of Shape)

                    '            For Each paragraph As TextParagraph In shape.Text.Paragraphs
                    '                For Each run As TextRun In paragraph.Elements.OfType(Of TextRun)
                    '                    'Dim isBold As Boolean = run.CharacterFormat.Bold
                    '                    'Dim text As String = run.Text
                    '                    'If run.Text = "தங்களை" Then run.Text = "thangalai"
                    '                    'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                    '                    tempi = -1
                    '                    'run.Format.Font
                    '                    If (Mid(run.Format.Font, 1, 7).ToLower = "vanavil") Then tempi = 0

                    '                    'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

                    '                    If tempi > -1 Then
                    '                        If fromencstatus(tempi) Then
                    '                            ' MsgBox(run.Text & "  " & run.Font.Name)

                    '                            stri = run.Text
                    '                            tempk = 1
                    '                            ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

                    '                            tempj = Len(stri)
                    '                            tempk = 1
                    '                            booli = morethan127(Mid(stri, tempk, 1))
                    '                            strj = Mid(stri, tempk, 1)

                    '                            ''''''''''


                    '                            If fromenc(tempi) = "Unicode" Then
                    '                                If isthereunicode(run.Text) Then
                    '                                    run.Format.Font = totamfont(tempi)
                    '                                    ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '                                Else
                    '                                    run.Format.Font = toengfont(tempi)
                    '                                End If
                    '                            Else
                    '                                run.Format.Font = totamfont(tempi)
                    '                                'run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '                            End If

                    '                            '   run.Font.Name = totamfont(tempi)

                    '                            ' run.CharacterFormat.Size += totamsizer(tempi)
                    '                            ' Dim tttt As String = run.Text

                    '                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

                    '                            ' MsgBox(tttt & " " & run.Text)
                    '                        Else
                    '                            If fromenc(tempi) = "Unicode" Then
                    '                                If isthereunicode(run.Text) Then
                    '                                    run.Format.Font = totamfont(tempi)
                    '                                    'run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '                                Else
                    '                                    run.Format.Font = toengfont(tempi)
                    '                                End If
                    '                            Else
                    '                                run.Format.Font = totamfont(tempi)
                    '                                ' run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '                            End If
                    '                            ' run.Font.Name = totamfont(tempi)
                    '                            ' run.CharacterFormat.Size += totamsizer(tempi)
                    '                            run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
                    '                        End If
                    '                    End If

                    '                Next
                    '            Next
                    '        Next
                    '    Next


                    '    'For Each paragraph As GemBox.Document.Paragraph In gemboxpresentation.GetChildElements(True, ElementType.Paragraph)
                    '    '    For Each run As GemBox.Document.Run In paragraph.GetChildElements(True, ElementType.Run)
                    '    '        'Dim isBold As Boolean = run.CharacterFormat.Bold
                    '    '        'Dim text As String = run.Text
                    '    '        'If run.Text = "தங்களை" Then run.Text = "thangalai"
                    '    '        'sb.AppendFormat("{0}{1}{2}", If(isBold, "<b>", ""), text, If(isBold, "</b>", ""))
                    '    '        tempi = -1
                    '    '        If (Mid(run.CharacterFormat.FontName, 1, 7).ToLower = "vanavil") Then tempi = 0

                    '    '        'tempi = Array.IndexOf(fromfont, run.CharacterFormat.FontName)

                    '    '        If tempi > -1 Then
                    '    '            If fromencstatus(tempi) Then
                    '    '                ' MsgBox(run.Text & "  " & run.Font.Name)

                    '    '                stri = run.Text
                    '    '                tempk = 1
                    '    '                ' MsgBox(stri & "  " & run.Font.Name & " " & AscW(Mid(stri, tempk, 1)))

                    '    '                tempj = Len(stri)
                    '    '                tempk = 1
                    '    '                booli = morethan127(Mid(stri, tempk, 1))
                    '    '                strj = Mid(stri, tempk, 1)

                    '    '                ''''''''''


                    '    '                If fromenc(tempi) = "Unicode" Then
                    '    '                    If isthereunicode(run.Text) Then
                    '    '                        run.CharacterFormat.FontName = totamfont(tempi)
                    '    '                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '    '                    Else
                    '    '                        run.CharacterFormat.FontName = toengfont(tempi)
                    '    '                    End If
                    '    '                Else
                    '    '                    run.CharacterFormat.FontName = totamfont(tempi)
                    '    '                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '    '                End If

                    '    '                '   run.Font.Name = totamfont(tempi)

                    '    '                run.CharacterFormat.Size += totamsizer(tempi)
                    '    '                ' Dim tttt As String = run.Text

                    '    '                run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)

                    '    '                ' MsgBox(tttt & " " & run.Text)
                    '    '            Else
                    '    '                If fromenc(tempi) = "Unicode" Then
                    '    '                    If isthereunicode(run.Text) Then
                    '    '                        run.CharacterFormat.FontName = totamfont(tempi)
                    '    '                        run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '    '                    Else
                    '    '                        run.CharacterFormat.FontName = toengfont(tempi)
                    '    '                    End If
                    '    '                Else
                    '    '                    run.CharacterFormat.FontName = totamfont(tempi)
                    '    '                    run.CharacterFormat.Language = CultureInfo.GetCultureInfo("ta-IN")

                    '    '                End If
                    '    '                ' run.Font.Name = totamfont(tempi)
                    '    '                run.CharacterFormat.Size += totamsizer(tempi)
                    '    '                run.Text = convert(fromenc(tempi), toenc(tempi), run.Text)
                    '    '            End If
                    '    '        End If
                    '    '    Next

                    '    'Next
                    '    'gemboxdocument.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".docx") 'SelectSaveFileName())
                    '    If drag_drop = False Then
                    '        If Not Directory.Exists(targetdir) Then
                    '            Directory.CreateDirectory(targetdir)
                    '        End If
                    '        gemboxpresentation.Save(targetdir & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".pptx") 'SelectSaveFileName())
                    '    Else
                    '        gemboxpresentation.Save(Path.GetDirectoryName(fname) & "\" & Path.GetFileNameWithoutExtension(fname) & "_unicode" & ".pptx") 'SelectSaveFileName())
                    '    End If
                    ' MsgBox("Done")
            End Select



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    'Private Shared Function GetHtmlFormattedValue(ByVal cell As ExcelCell) As String
    '    Dim tempEf = New ExcelFile()
    '    Dim tempWs = tempEf.Worksheets.Add("Sheet")
    '    Dim tempCell = tempWs.Cells("A1")

    '    tempCell.SetValue(cell.StringValue)

    '    For Each fcr As FormattedCharacterRange In cell.CharacterRanges
    '        Dim fcrFont As ExcelFont = fcr.Font

    '        Dim tempFcr = tempCell.GetCharacters(fcr.StartIndex, fcr.Length)
    '        Dim tempFcrFont = tempFcr.Font

    '        tempFcrFont.Color = fcrFont.Color
    '        tempFcrFont.Italic = fcrFont.Italic
    '        tempFcrFont.Name = fcrFont.Name
    '        tempFcrFont.ScriptPosition = fcrFont.ScriptPosition
    '        tempFcrFont.Size = fcrFont.Size
    '        tempFcrFont.Strikeout = fcrFont.Strikeout
    '        tempFcrFont.UnderlineStyle = fcrFont.UnderlineStyle
    '        tempFcrFont.Weight = fcrFont.Weight
    '    Next fcr

    '    Dim builder As New StringBuilder()
    '    Using writer As XmlWriter = XmlWriter.Create(builder)
    '        tempEf.Save(writer, New GemBox.Spreadsheet.HtmlSaveOptions() With {.HtmlType = GemBox.Spreadsheet.HtmlType.HtmlTable})
    '    End Using

    '    Dim tempXml As New XmlDocument()
    '    tempXml.LoadXml(builder.ToString())

    '    Return tempXml.SelectSingleNode("//td").InnerXml
    'End Function
    Private Function morethan127(ByVal str As String) As Boolean
        Try
            Dim flag As Boolean = False
            Dim i As Integer
            For i = 1 To Len(str)
                If AscW(Mid(str, i, 1)) > 127 Then
                    flag = True
                    Exit For
                End If
            Next
            Return flag
        Catch ex As Exception

        End Try
    End Function
    Private Function isonlyunicode(ByVal str As String) As Boolean
        Try
            Dim flag As Boolean = True
            Dim i As Integer
            For i = 1 To Len(str)
                If AscW(Mid(str, i, 1)) < 2943 Then
                    flag = False
                    Exit For
                End If
            Next
            Return flag
        Catch ex As Exception

        End Try
    End Function
    Private Function isthereunicode(ByVal str As String) As Boolean
        Try
            Dim flag As Boolean = False
            Dim i As Integer
            For i = 1 To Len(str)
                If AscW(Mid(str, i, 1)) > 255 Then
                    flag = True
                    Exit For
                End If
            Next
            Return flag
        Catch ex As Exception

        End Try
    End Function

    Public Function convert(ByVal fromm As String, ByVal too As String, ByVal str As String) As String
        Try
            'this




            '  MsgBox(fromm & "  " & too)

            Dim j As Int16
            Dim srcindex As Integer
            Dim tgtindex As Integer
            srcindex = Array.IndexOf(srcenc, fromm) 'combobox1.Items.IndexOf(fromm)
            tgtindex = Array.IndexOf(srcenc, too) 'combobox2.Items.IndexOf(too)

            'MsgBox(srcindex & "  " & tgtindex)
            'pb1.Visible = True
            Dim cc As Int16
            'pb1.Refresh()
            Dim tt As Integer
            '  str = TextBox1.Text
            'textbox1.Enabled = False
            'textbox1.BackColor = Me.BackColor

            special = True
            savefrom = fromm
            saveto = too
            'fs2.textbox1.Text = prevtext
            ReDim uni1(srccnt(uninumber))
            ReDim uni2(srccnt(uninumber))
            uni_count = srccnt(uninumber)
            For j = 0 To uni_count - 1
                uni1(j) = src(0, uninumber, j)
                uni2(j) = src(1, uninumber, j)
            Next


            ReDim src1(srccnt(srcindex))
            ReDim src2(srccnt(srcindex))

            src_count = srccnt(srcindex)
            For j = 0 To src_count - 1
                src1(j) = src(0, srcindex, j)
                src2(j) = src(1, srcindex, j)
                'MsgBox(src1(j) & " " & src2(j))
            Next


            ReDim tgt1(srccnt(tgtindex))
            ReDim tgt2(srccnt(tgtindex))
            tgt_count = srccnt(tgtindex)

            For j = 0 To tgt_count - 1
                tgt1(j) = src(0, tgtindex, j)
                tgt2(j) = src(1, tgtindex, j)
            Next
            'pb1.Minimum = 0
            'pb1.Maximum = (src_count / 100) + (uni_count / 100)
            'pb1.Value = 0
            'pb1.Caption = "Converting.. Please wait"
            'pb1.Show()
            cc = 0
            'pb1.BringToFront()
            If Trim(str).Length > 0 Then
                For j = 0 To src_count - 1
                    str = str.Replace(src2(j), src1(j))
                Next
            End If
            'pb1.Value = src_count / 100
            'pb1.Refresh()
            'For j = 0 To tgt_count - 1
            '    strr += tgt2(j) & vbTab & tgt1(j) & vbNewLine
            'Next
            cc = 0
            If Trim(str).Length > 0 Then
                For j = 0 To uni_count - 1

                    'If cc = 100 Then
                    '    pb1.Value += 1
                    '    pb1.Refresh()
                    '    cc = 0
                    'Else
                    '    cc += 1
                    'End If

                    tt = Array.IndexOf(tgt1, uni1(j))

                    If tt <> -1 Then
                        ' If uni1(j) = "அ" Then MsgBox(tgt1(tt) & "  " & uni1(j) & "  " & tgt2(tt))
                        'fs2.textbox1.Text = fs2.textbox1.Text.Replace(tgt1(tt), tgt2(tt))
                        str = str.Replace(tgt1(tt), tgt2(tt))
                    End If

                Next
            End If
            'pb1.Value = (src_count / 100) + (uni_count / 100)
            'pb1.Refresh()

            'pb1.BarValue = 0
            ' pb1.Visible = False
            'pb1.Refresh()
            special = False
            Return str

            'textischanged()


            'Button2.Enabled = True

            ' textbox1.Enabled = True
        Catch ex As Exception
            MsgBox(ex.Message)

            Return str
            'textbox1.BackColor = Color.White
            'textbox1.Enabled = True
            'Button2.Enabled = True

            'If t1.ThreadState = ThreadState.Running Or t1.ThreadState = ThreadState.Suspended Then
            '    t1.Abort()
            'End If
            'If t2.ThreadState = ThreadState.Running Or t2.ThreadState = ThreadState.Suspended Then
            '    t2.Abort()
            'End If

        End Try
    End Function

    Public Sub processfile()
        loadd()
        If isunicodechecked Then
            processgemtounicode(file_path)
        Else
            processgemtovanavil(file_path)
        End If



    End Sub
    Public Sub processdocument()
        'loadd()
        If isunicodechecked Then
            processgemdoctounicode(inputdocument)
        Else
            'processgemdoctovanavil(inputdocument)
        End If



    End Sub
    'Public Sub SaveDocument()
    '    If mDocument Is Nothing Then
    '        Return
    '    End If

    '    Dim fileName As String = Path.GetDirectoryName(mDocument.OriginalFileName) & "\" & Path.GetFileNameWithoutExtension(mDocument.OriginalFileName) & "_unicode" & ".docx" 'SelectSaveFileName()
    '    '  MsgBox(fileName)
    '    If fileName Is Nothing Then
    '        Application.Exit()
    '        Return
    '    End If
    '    'mDocument.SaveOptions.ExportPrettyFormat = True nags

    '    ' This operation can take some time so we set the Cursor to WaitCursor.
    '    Application.DoEvents()
    '    Dim cursor As Cursor = Cursor.Current
    '    Cursor.Current = Cursors.WaitCursor

    '    ' This operation is put in try-catch block to handle situations when operation fails for some reason.
    '    Try
    '        If (Not Path.GetExtension(fileName).Equals(".AsposePdf")) Then
    '            ' For most of the save formats it is enough to just invoke Aspose.Words save.
    '            mDocument.Save(fileName)
    '        Else
    '            SavePdfLegacyWayViaAsposePdf(mDocument, fileName)
    '        End If
    '        '   Application.Exit()
    '    Catch ex As Exception
    '        ' CType(New ExceptionDialog(ex), ExceptionDialog).ShowDialog()
    '    End Try

    '    ' Restore cursor.
    '    Cursor.Current = cursor
    'End Sub
    Private mInitialDirectory As String = Application.StartupPath

    'Private Function SelectSaveFileName() As String
    '    Dim dlg As SaveFileDialog = New SaveFileDialog()
    '    Try
    '        dlg.CheckFileExists = False
    '        dlg.CheckPathExists = True
    '        dlg.Title = "Save Document As"
    '        dlg.InitialDirectory = Path.GetDirectoryName(file_path) 'mInitialDirectory
    '        'dlg.Filter = "Word 97-2003 Document (*.doc)|*.doc|" & "Word 2007 OOXML Document (*.docx)|*.docx|" & "Word 2007 OOXML Macro-Enabled Document (*.docm)|*.docm|" & "PDF (*.pdf)|*.pdf|" & "PDF (legacy, via Aspose.Pdf XML) (*.AsposePdf)|*.AsposePdf|" & "OpenDocument Text (*.odt)|*.odt|" & "Web Page (*.html)|*.html|" & "Single File Web Page (*.mht)|*.mht|" & "Rich Text Format (*.rtf)|*.rtf|" & "Word 2003 WordprocessingML (*.xml)|*.xml|" & "Plain Text (*.txt)|*.txt|" & "IDPF EPUB Document (*.epub)|*.epub"
    '        'dlg.Filter = "Word 97-2003 Document (*.doc)|*.doc|" & "Word 2007 OOXML Document (*.docx)|*.docx|" & "Word 2007 OOXML Macro-Enabled Document (*.docm)|*.docm|" & "OpenDocument Text (*.odt)|*.odt|" & "Web Page (*.html)|*.html|" & "Single File Web Page (*.mht)|*.mht|" & "Rich Text Format (*.rtf)|*.rtf|" & "Plain Text (*.txt)|*.txt|"
    '        'dlg.Filter = "Word 2007 OOXML Document (*.docx)|*.docx|"
    '        dlg.FileName = Path.GetFileNameWithoutExtension(mDocument.OriginalFileName) & "_unicode.docx"

    '        Dim dlgResult As DialogResult = dlg.ShowDialog()
    '        ' Optimized to allow automatic conversion to VB.NET
    '        If dlgResult.Equals(System.Windows.Forms.DialogResult.OK) Then
    '            mInitialDirectory = Path.GetDirectoryName(dlg.FileName)
    '            Return dlg.FileName
    '        Else
    '            Return Nothing
    '        End If
    '    Finally
    '        dlg.Dispose()
    '    End Try
    'End Function
End Class
