Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Threading
Imports System.Windows.Forms
Imports DevExpress.Skins
Imports GemBox.Document
Imports Microsoft.VisualBasic

Public Class XtraConverter


    Private Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        If FolderBrowserDialog.ShowDialog(Me) = DialogResult.OK Then
            parentlocation.Text = FolderBrowserDialog.SelectedPath
            Try
                ProgressBar1.Position = 0
                button_pullfiles.Enabled = False
                LoadFiles(parentlocation.Text, False)
                button_pullfiles.Enabled = True
                button_convert.Enabled = True
                check_fileselection.Checked = False
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                button_pullfiles.Enabled = True
            End Try
        End If
    End Sub
    Private Sub LoadFiles(ByVal initial As String, ByVal bool_includesubdirectory As Boolean)
        ' MessageBox.Show(bool_includesubdirectory)
        lstFiles.Items.Clear()
        Dim searchoption As System.IO.SearchOption
        If bool_includesubdirectory Then
            searchoption = SearchOption.AllDirectories
        Else
            searchoption = SearchOption.TopDirectoryOnly
        End If
        ' This list stores the results.
        Dim result As New List(Of String)
        Dim finalresult As New List(Of String)
        Dim wholestack As New Stack(Of String)
        ' This stack stores the directories to process.
        Dim stack As New Stack(Of String)

        ' Add the initial directory
        stack.Push(initial)
        ' Continue processing for each stacked directory
        Do While (stack.Count > 0)
            ' Get top directory string
            Dim dir As String = stack.Pop
            Try
                ' Add all immediate file paths
                If Not wholestack.Contains(dir) Then
                    result.AddRange(customgetFiles(dir, "*.docx|*.doc|*.rtf|*.xls|*.xlsx|*.ods|*.csv|*.pptx|*.ppt", searchoption))
                    wholestack.Push(dir)

                End If
                ' Loop through all subdirectories and add them to the stack.
                Dim directoryName As String
                For Each directoryName In Directory.GetDirectories(dir, searchoption)
                    stack.Push(directoryName)

                Next
            Catch ex As Exception
            End Try
        Loop

        ' Return the list
        finalresult = result.Distinct.ToList
        For Each s As String In finalresult
            Dim item As New ListViewItem()
            item.UseItemStyleForSubItems = False

            item.SubItems.Add("Not started")
            item.SubItems.Add("")
            item.SubItems(2).BackColor = System.Drawing.Color.White
            item.SubItems.Add("")
            item.Text = s.Replace("%20", " ")
            lstFiles.Items.Add(item)

        Next s
        lstFiles.Refresh()
    End Sub
    Private Sub button_pullfiles_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles button_pullfiles.Click
        Try
            ProgressBar1.Position = 0
            button_pullfiles.Enabled = False
            LoadFiles(parentlocation.Text, False)
            button_pullfiles.Enabled = True
            button_convert.Enabled = True
            check_fileselection.Checked = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            button_pullfiles.Enabled = True
        End Try
    End Sub
    Public Function customgetFiles(ByVal SourceFolder As String, ByVal Filter As String,
 ByVal searchOption As System.IO.SearchOption) As String()
        ' ArrayList will hold all file names
        Dim alFiles As ArrayList = New ArrayList()

        ' Create an array of filter string
        Dim MultipleFilters() As String = Filter.Split("|")

        ' for each filter find mathing file names
        For Each FileFilter As String In MultipleFilters
            ' add found file names to array list
            'MessageBox.Show(FileFilter.Substring(1, FileFilter.Length - 1))
            alFiles.AddRange(Directory.GetFiles(SourceFolder, FileFilter, searchOption))
        Next

        ' returns string array of relevant file names
        Return alFiles.ToArray(Type.GetType("System.String"))
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles check_fileselection.CheckedChanged
        Try

            For Each item As ListViewItem In lstFiles.Items
                item.Checked = check_fileselection.Checked
            Next
        Catch ex As Exception

        End Try
    End Sub
    Private _threads As New ArrayList()
    Private _instances As New ArrayList()
    Private _activeConvertCount As Integer = 0

    Dim checkedfilepaths() As String
    Dim checkedindex() As Int16
    Dim checkedtimetaken() As Double
    Dim checkedcount As Integer
    Private Sub Button_Convertt(sender As Object, e As EventArgs) Handles button_convert.Click
        Try
            Application.UseWaitCursor = True
            targetdir = childlocation.Text
            If Not targetdir = "" Then
                isunicodechecked = unicode_checked.Checked
                ProgressBar1.Properties.Minimum = 0
                ProgressBar1.Properties.Maximum = lstFiles.CheckedItems.Count
                ProgressBar1.Position = 0
                'Dim currentSkin As Skin = CommonSkins.GetSkin(DefaultLookAndFeel1.LookAndFeel)
                'Dim bc As Color = currentSkin.Colors.GetColor(DevExpress.Skins.CommonColors.DisabledControl)
                'ProgressBar1.Properties.StartColor = Color.FromArgb(200, 197, 62)
                'ProgressBar1.Properties.EndColor = Color.FromArgb(147, 145, 43)

                totalprogressvalue = lstFiles.CheckedItems.Count
                If Not totalprogressvalue = 0 Then
                    button_cancel.Visible = True
                    button_convert.Enabled = False
                End If
                checkedcount = 0
                Dim j As Int16 = 0
                For Each item As ListViewItem In lstFiles.Items
                    If item.Checked Then
                        item.SubItems(2).BackColor = System.Drawing.Color.FromArgb(234, 228, 214)
                        item.SubItems(1).Text = "Started"
                        ReDim Preserve checkedfilepaths(checkedcount + 1)
                        ReDim Preserve checkedindex(checkedcount + 1)
                        checkedfilepaths(checkedcount) = item.SubItems(0).Text
                        checkedindex(checkedcount) = j
                        checkedcount += 1
                    End If
                    j += 1
                Next
                ReDim checkedtimetaken(checkedcount)
                If (Not BackgroundWorker1.IsBusy) Then
                    BackgroundWorker1.RunWorkerAsync()
                Else
                    MessageBox.Show("Converter process is still running. Please try after some time.")
                End If
            Else
                Application.UseWaitCursor = False
                MessageBox.Show("Target directory is empty")
            End If


        Catch ex As Exception
            Application.UseWaitCursor = False
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub button_convert_VisibleChanged(sender As Object, e As EventArgs) Handles button_convert.VisibleChanged
        button_cancel.Visible = Not (TryCast(sender, DevExpress.XtraEditors.SimpleButton)).Visible
    End Sub
    Dim totalprogressvalue As Int16
    Dim currentprogressvalue As Int16
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try

            '_instances = New ArrayList()
            '_threads = New ArrayList()
            '_activeConvertCount = 0

            If totalprogressvalue = 0 Then
                MessageBox.Show("Please select one or more files to convert.")
            Else

                '  Dim download As FileDownloader = Nothing

                Dim i As Int16 = 0
                For i = 0 To checkedcount - 1
                    If BackgroundWorker1.CancellationPending Then
                        e.Cancel = True
                        MessageBox.Show("Conversion cancelled")
                        Exit Sub

                    Else
                        Dim stopwatch As Stopwatch = Stopwatch.StartNew()

                        ' item.SubItems(1).Text = "Not started"

                        '  item.Tag = item.SubItems(1).Text
                        Try
                            Dim trc As TamilRichtextconvert
                            trc = New TamilRichtextconvert
                            trc.file_path = checkedfilepaths(i)
                            trc.drag_drop = False
                            'trc.isunicodechecked = isunicodechecked
                            ' item.SubItems(1).Text = "Started"
                            trc.processfile()

                            ' item.SubItems(1).Text = "Converted"

                        Catch
                            ' If the download fails for some reason, flag and error. 
                            'item.SubItems(1).Text = "Error"
                        End Try
                        'i += 1
                        ' progressbar1.Value = i
                        stopwatch.Stop()
                        checkedtimetaken(i) = stopwatch.Elapsed.Milliseconds
                        BackgroundWorker1.ReportProgress(i * 100 / totalprogressvalue, i)
                    End If
                    ' lstFiles.Refresh()
                    ' progressbar1.Refresh()
                Next

                Dim result As Integer = MessageBox.Show("Do you want to open the directory of Converted documents?", "Conversion completed", MessageBoxButtons.YesNo)
                If result = DialogResult.Yes Then
                    System.Diagnostics.Process.Start(targetdir)

                End If

                ' StartDownload()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub button_cancel_Click(sender As Object, e As EventArgs) Handles button_cancel.Click
        BackgroundWorker1.CancelAsync()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Try
            Dim i As Int16 = CInt(e.UserState.ToString)
            ' MessageBox.Show(e.UserState.ToString)
            lstFiles.Items(checkedindex(i)).UseItemStyleForSubItems = False
            lstFiles.Items(checkedindex(i)).SubItems(1).Text = "Converted"

            lstFiles.Items(checkedindex(i)).SubItems(2).BackColor = System.Drawing.Color.FromArgb(153, 187, 232) 'Color.FromArgb(147, 145, 43) 'System.Drawing.Color.FromArgb(147, 119, 104)
            '  lstFiles.Items(checkedindex(i)).SubItems(2).Text = CDbl(checkedtimetaken(i) / 1000)
            If isunicodechecked Then
                lstFiles.Items(checkedindex(i)).SubItems(3).Text = Path.GetDirectoryName(lstFiles.Items(checkedindex(i)).SubItems(0).Text) & "\" & Path.GetFileNameWithoutExtension(lstFiles.Items(checkedindex(i)).SubItems(0).Text) & "_unicode" & ".docx"
            Else
                lstFiles.Items(checkedindex(i)).SubItems(3).Text = Path.GetDirectoryName(lstFiles.Items(checkedindex(i)).SubItems(0).Text) & "\" & Path.GetFileNameWithoutExtension(lstFiles.Items(checkedindex(i)).SubItems(0).Text) & "_vanavil" & ".docx"
            End If
            ProgressBar1.Position = i + 1

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            button_convert.Visible = True
            button_cancel.Visible = False
            Application.UseWaitCursor = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub XtraTabPage2_DragEnter(sender As Object, e As DragEventArgs) Handles XtraTabPage2.DragEnter

    End Sub
    Private Sub TabPage2_DragEnter(sender As Object, e As DragEventArgs) Handles XtraTabPage2.DragEnter ', Panel1.DragEnter 'TabPage2.DragEnter
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                e.Effect = DragDropEffects.Copy
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Dim dragdropfiles() As String
    Private Sub TabPage2_DragDrop(sender As Object, e As DragEventArgs) Handles XtraTabPage2.DragDrop ', Panel1.DragDrop
        Try

            Label10.Text = "Conversion Started" ''dragdropduration
            Label10.Refresh()
            dragdropfiles = e.Data.GetData(DataFormats.FileDrop)
            Application.UseWaitCursor = True
            'Call BackgroundWorker2.RunWorkerAsync()
            If (Not BackgroundWorker2.IsBusy) Then
                BackgroundWorker2.RunWorkerAsync()
            Else
                MessageBox.Show("Converter process is still running. Please try after some time.")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Dim dragdropduration As Double
    Private Sub BackgroundWorker2_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker2.DoWork

        Dim i As Int16 = 0
        Dim stopwatch As Stopwatch = Stopwatch.StartNew()
        For Each Patth As String In dragdropfiles
            i += 1
            'ReDim Preserve Workers(i)
            'Workers(i - 1) = New BackgroundWorker
            'Workers(i - 1).WorkerReportsProgress = True
            'Workers(i - 1).WorkerSupportsCancellation = True
            'AddHandler Workers(i - 1).DoWork, AddressOf bw_DoWork
            'AddHandler Workers(i - 1).ProgressChanged, AddressOf bw_ProgressChanged
            'AddHandler Workers(i - 1).RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
            'Workers(i - 1).RunWorkerAsync()
            'Workers(i - 1).

            '  If i = 1 Then

            Dim trc As TamilRichtextconvert
            trc = New TamilRichtextconvert

            isunicodechecked = unicode_checked1.Checked

            trc.file_path = Patth
            trc.drag_drop = True
            trc.processfile()

        Next
        stopwatch.Stop()
        dragdropduration = stopwatch.Elapsed.Milliseconds


    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try

            If unicode_checked1.Checked Then
                Label10.Text = "Conversion completed. Converted Files will be available in the same directory from where you dropped the files and they will have " & Strings.ChrW(34) & "_" & "Unicode" & Strings.ChrW(34) & " appended to the filename." ''dragdropduration
            Else
                Label10.Text = "Conversion completed. Converted Files will be available in the same directory from where you dropped the files and they will have " & Strings.ChrW(34) & "_" & "Vanavil" & Strings.ChrW(34) & " appended to the filename." ''dragdropduration
            End If
            Label10.Refresh()
            Application.UseWaitCursor = False

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            TextBox1.Clear()
            TextBox2.Clear()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            ' TextBox1.Text = Clipboard.GetText
            TextBox1.Paste()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            'Dim trc As TamilRichtextconvert
            'trc = New TamilRichtextconvert
            'trc.loadd()
            'If unicode_checked2.Checked Then
            '    TextBox2.Text = trc.convert("Vanavil", "Unicode", TextBox1.Text)
            'Else
            '    TextBox2.Text = trc.convert("Unicode", "Vanavil", TextBox1.Text)
            'End If 
            GemBox.Document.ComponentInfo.SetLicense("DW1R-R1HW-JKVY-N4DH")

            Dim stream1 As MemoryStream = New MemoryStream()
            ' Save RichTextBox content to RTF stream.

            Me.TextBox1.SaveFile(stream1, RichTextBoxStreamType.RichText)
            stream1.Seek(0, SeekOrigin.Begin)
            ' Load document from RTF stream and prepend or append clipboard content to it.
            Dim document As DocumentModel = DocumentModel.Load(stream1, LoadOptions.RtfDefault)
            Dim trc As TamilRichtextconvert
            trc = New TamilRichtextconvert
            trc.loadd()
            trc.inputdocument = document
            Dim outputrtf As String
            Dim outputstream As MemoryStream = New MemoryStream()

            If unicode_checked2.Checked Then

                outputrtf = trc.processgemdoctounicode(document)
                Me.TextBox2.Rtf = outputrtf

                'TextBox2.Text = trc.convert("Vanavil", "Unicode", TextBox1.Text)
            Else
                outputstream = trc.processgemdoctovanavil(document)
                outputstream.Seek(0, SeekOrigin.Begin)
                Dim reader As StreamReader = New StreamReader(outputstream)
                outputrtf = reader.ReadToEnd()

                'Dim filewriter As System.IO.StreamWriter = New StreamWriter("temp.rtf")
                'filewriter.Write(outputrtf)
                'filewriter.Close()
                'TextBox2.LoadFile("temp.rtf")
                'File.Delete("temp.rtf")



                TextBox2.Rtf = outputrtf
                'Clipboard.SetText(outputrtf, TextDataFormat.Rtf)


                'DocumentModel.Load(outputstream, LoadOptions.RtfDefault).Content.SaveToClipboard()
                'TextBox2.Rtf = Clipboard.GetText(TextDataFormat.Rtf)

                'Me.TextBox2.LoadFile(outputstream, RichTextBoxStreamType.RichText)
                'Me.TextBox2.Refresh()

                ' Me.TextBox2.LoadFile("test_vanavil" & ".rtf")
                'MsgBox("hi")
                'TextBox2.Text = trc.convert("Unicode", "Vanavil", TextBox1.Text)
                ' Me.TextBox2.Clear()
                ' Me.TextBox2.ResetText()
                'Me.TextBox2.LoadFile(outputstream, RichTextBoxStreamType.RichText)
            End If


        Catch ex As System.InvalidOperationException
            MsgBox("Highly rich objects like Tables not supported in this conversion. Please use File Conversion in previous Tab for highly rich Documents.")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            ' TextBox2.Copy()
            'Clipboard.SetText(TextBox2.Text)
            'Using stream As MemoryStream = New MemoryStream()
            '    ' Save RichTextBox selection to RTF stream.
            '    Dim writer As StreamWriter = New StreamWriter(stream)
            '    writer.Write(Me.TextBox2.SelectedRtf)
            '    writer.Flush()
            '    stream.Seek(0, SeekOrigin.Begin)
            '    ' stream.Position = 0

            '    ' Save RTF stream to clipboard.
            '    DocumentModel.Load(stream, LoadOptions.RtfDefault).Content.SaveToClipboard()
            'End Using
            Clipboard.SetText(TextBox2.Rtf.Replace(tempfont, permfont), TextDataFormat.Rtf)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub unicode_checked_CheckedChanged(sender As Object, e As EventArgs) Handles unicode_checked.CheckedChanged, unicode_checked1.CheckedChanged, unicode_checked2.CheckedChanged
        Try
            unicode_checked.Checked = TryCast(sender, RadioButton).Checked
            unicode_checked1.Checked = unicode_checked.Checked
            unicode_checked2.Checked = unicode_checked.Checked
            isunicodechecked = unicode_checked.Checked
            'MessageBox.Show(isunicodechecked)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Vanavil_checked_CheckedChanged(sender As Object, e As EventArgs) Handles Vanavil_checked.CheckedChanged, Vanavil_checked1.CheckedChanged, Vanavil_checked2.CheckedChanged
        Try
            Vanavil_checked.Checked = TryCast(sender, RadioButton).Checked
            Vanavil_checked1.Checked = Vanavil_checked.Checked
            Vanavil_checked2.Checked = Vanavil_checked.Checked
            isunicodechecked = unicode_checked.Checked
            ' MessageBox.Show(isunicodechecked)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Converter_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            'If (Date.Now.Year > 2017) Or (Date.Now.Month > 1) Then
            '    MessageBox.Show("This must be older build. Check with Tamil Virtual Academy or Indic Labs for newer build.")
            '    Application.Exit()
            'End If
            SizeLastColumn(lstFiles)

            ' MessageBox.Show("hello..")
            DevExpress.Skins.SkinManager.EnableFormSkins()
            DevExpress.UserSkins.BonusSkins.Register()
            DevExpress.LookAndFeel.UserLookAndFeel.Default.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Skin
            'DevExpress.LookAndFeel.UserLookAndFeel.Default.SetSkinStyle("Xmas 2008 Blue")

            ' DevExpress.LookAndFeel.UserLookAndFeel.Default.SkinName = "Blue"

            ' DevExpress.LookAndFeel.UserLookAndFeel.Default.SetStyle(DevExpress.LookAndFeel.LookAndFeelStyle.Office2003, False, False)

            Application.EnableVisualStyles()

            'Application.SetCompatibleTextRenderingDefault(False)
            ' Dim lookandfeel1 As DevExpress.LookAndFeel.DefaultLookAndFeel

        Catch ex As Exception
            MessageBox.Show("hello.." & ex.Message)
        End Try
    End Sub

    Private Sub button7_Click(sender As Object, e As EventArgs) Handles button7.Click
        If FolderBrowserDialog1.ShowDialog(Me) = DialogResult.OK Then
            childlocation.Text = FolderBrowserDialog1.SelectedPath

        End If
    End Sub

    Private Sub button_clearfiles_Click(sender As Object, e As EventArgs) Handles button_clearfiles.Click
        Try
            Application.UseWaitCursor = False
            ProgressBar1.Position = 0
            ' button_pullfiles.Enabled = False
            parentlocation.Text = ""
            childlocation.Text = ""
            button_pullfiles.Enabled = True
            button_convert.Enabled = True
            check_fileselection.Checked = False
            lstFiles.Items.Clear()
            unicode_checked.Checked = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            button_pullfiles.Enabled = True
        End Try
    End Sub
    Private Sub lstfiles_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstFiles.Resize
        SizeLastColumn(CType(sender, ListView))
    End Sub

    Private Sub SizeLastColumn(ByVal lv As ListView)
        lv.Columns(lv.Columns.Count - 1).Width = -2
    End Sub
    Private Sub lstfiles_DrawColumnHeader(ByVal sender As Object, ByVal e As DrawListViewColumnHeaderEventArgs) Handles lstFiles.DrawColumnHeader
        e.Graphics.FillRectangle(Brushes.Red, e.Bounds)
        e.DrawText()
    End Sub
    Private Sub lstfiles_DrawItem(ByVal sender As Object, ByVal e As DrawListViewItemEventArgs) Handles lstFiles.DrawItem
        e.DrawDefault = True
    End Sub


End Class