Imports System
Imports System.Windows.Forms
Imports System.Xml
Imports Microsoft.VisualBasic

Module startup



    Public isunicodechecked As Boolean = True
    '  Public tempfont As String = "NHM_REF_VANAVIL"
    'Public tempfont As String = "vanavilavvai"
    'Public permfont As String = "VANAVIL-Avvaiyar"
    Public tempfont As String = "LAKSHMAN"
    Public permfont As String = "LAKSHMAN"
    Public totxmls As Integer = 0
    Public savefrom As String
    Public saveto As String
    Public special As Boolean = False
    Public uninumber As Integer
    'Public mDocument As Document
    Public docfonts() As String
    Public src1() As String
    Public tgt1() As String
    Public src2() As String
    Public tgt2() As String
    Public uni1() As String
    Public uni2() As String
    Public srcc1() As String
    Public srcc2() As String
    Public src_count As Integer
    Public tgt_count As Integer
    Public uni_count As Integer
    Public srcc_count As Int16
    Public fromenc() As String
    Public toenc() As String
    Public fromfont() As String
    Public totamfont() As String
    Public toengfont() As String
    Public totamsizer() As Integer
    Public toengsizer() As Integer
    Public fromencstatus() As Boolean
    Public srcfont() As Boolean
    Public srcenc() As String
    Public srccum() As String
    Public srccount As Integer
    Public src(,,) As String
    Public srccnt() As Integer
    Public loaded As Boolean = False
    Public targetdir As String = ""
    Public Sub main()


        Dim contentitem As String
        Dim xml_all As XmlTextReader
        Dim xml_conf As XmlTextReader
        Dim i As Int16 = 0
        Dim j As Int16
        Dim k As Int16
        Dim temp As String
        Dim cump As String
        'Dim f1 As String
        'Dim f2 As Single
        'Dim f3 As System.Drawing.FontStyle
        Dim great As Int16 = 0
        ChDir(Application.StartupPath)
        xml_conf = New System.Xml.XmlTextReader("Config.xml")
        xml_conf.WhitespaceHandling = WhitespaceHandling.None
        xml_conf.Read()
        xml_conf.Read()
        xml_conf.Read()
        xml_conf.Read()
        'detect = CBool(xml_conf.ReadElementString("Detect"))
        'from = xml_conf.ReadElementString("From")
        'too = xml_conf.ReadElementString("To")
        'wrap = CBool(xml_conf.ReadElementString("Wordwrap"))
        'lang = xml_conf.ReadElementString("Language")
        xml_conf.Close()
        'combobox1.Items.Clear()
        'combobox2.Items.Clear()
        ChDir(Application.StartupPath & "\Data\Tamil") '& Trim(lang))
        contentitem = Dir("*.xml")
        'combobox1.BeginUpdate()
        'combobox2.BeginUpdate()
        'combobox1.Hide()
        'combobox2.Hide()
        ' ChDir(Application.StartupPath & "\Data")
        contentitem = Dir("*.xml")
        Do Until contentitem = ""
            totxmls += 1
            contentitem = Dir()
        Loop
        contentitem = Dir("*.xml")
        'pb1.Caption = "Loading.. Encoding Definitions"
        'pb1.Maximum = totxmls
        'pb1.Minimum = 0
        'pb1.Value = 0
        ' srccount = 0
        Do Until contentitem = ""
            If (Mid(contentitem, 1, Len(contentitem) - 4).ToLower = "unicode") Then uninumber = i
            xml_all = New System.Xml.XmlTextReader(Mid(contentitem, 1, Len(contentitem) - 4) & ".xml")
            xml_all.WhitespaceHandling = WhitespaceHandling.None
            xml_all.Read()
            xml_all.Read()
            j = 0
            cump = ""
            While Not xml_all.EOF

                xml_all.Read()
                If Not xml_all.IsStartElement() Then
                    Exit While
                End If
                xml_all.Read()
                If j > great Then great = j
                ReDim Preserve src(2, 100, great + 1)
                'MsgBox(xml_all.ReadElementString("Unicode"))
                src(0, i, j) = totext(xml_all.ReadElementString("Unicode")) 'Trim(totext(xml_all.ReadElementString("Unicode")))
                src(1, i, j) = totext(xml_all.ReadElementString("This")) 'Trim(totext(xml_all.ReadElementString("This")))
                cump = cump & src(1, i, j)
                'MsgBox(cump)
                j += 1
            End While
            ReDim Preserve srccnt(i + 1)
            srccnt(i) = j
            xml_all.Close()
            ' MsgBox(Mid(contentitem, 1, Len(contentitem) - 4))
            'combobox1.Items.Add(Mid(contentitem, 1, Len(contentitem) - 4))
            'combobox2.Items.Add(Mid(contentitem, 1, Len(contentitem) - 4))
            'srccount += 1
            ReDim Preserve srcenc(i + 1)
            ReDim Preserve srcfont(i + 1)
            ReDim Preserve srccum(i + 1)
            srcenc(i) = Mid(contentitem, 1, Len(contentitem) - 4)

            srccum(i) = RemoveDuplicateChars(cump)
            ' srcfont(i) = New System.Drawing.Font("Times New Roman", 12, FontStyle.Regular)
            contentitem = Dir()
            i += 1
            ' pb1.Value += 1
            'pb1.Refresh()
            ' If i = 100 Then Exit Do
        Loop


        'srccum(i) = RemoveDuplicateChars("")
        ' srcfont(i) = New System.Drawing.Font("Times New Roman", 12, FontStyle.Regular)
        'contentitem = Dir()
        'i += 1
        'pb1.Value += 1
        'pb1.Refresh()

        '  Me.Visible = True
        srccount = i
        'Clipboard.SetDataObject(TextBox1.Font.ToString)
        ' ReDim srcfont(ComboBox1.Items.Count)


        ChDir(Application.StartupPath)
        xml_conf = New System.Xml.XmlTextReader("Config.xml")
        xml_conf.WhitespaceHandling = WhitespaceHandling.None
        xml_conf.Read()
        xml_conf.Read()

        xml_conf.MoveToNextAttribute()
        xml_conf.Read()
        xml_conf.Read()
        While Not xml_conf.EOF
            'MsgBox(xml_conf.Name)
            If Not xml_conf.IsStartElement() Then
                Exit While
            End If
            xml_conf.Read()

            temp = xml_conf.ReadElementString("Encoding")
            'MsgBox(temp)
            k = Array.IndexOf(srcenc, temp)

            If k > -1 Then
                srcfont(k) = CBool(xml_conf.ReadElementString("Bilingual"))
                '   MsgBox(srcfont(k))
            End If
            xml_conf.Read()
        End While
        xml_conf.Close()

        loaded = True
        ' MsgBox("hi")

        'Application.Run(New Converter)
        Application.Run(New XtraConverter)

        'Catch ex As Exception
        '    MsgBox("nags " & ex.Message)
        'End Try

    End Sub
    Public Function totext(ByVal hextext As String) As String
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
End Module
