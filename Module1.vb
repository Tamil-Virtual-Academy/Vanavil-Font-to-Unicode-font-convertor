'Imports System.IO
'Imports System.Text
'Imports System.Windows.Forms
'Imports System

'Module Module1
'    Public totxmls As Integer = 0
'    Public savefrom As String
'    Public saveto As String
'    Public special As Boolean = False
'    Public aboutform As Boolean = False
'    Public uninumber As Integer
'    Public src(,,) As String
'    Public srccnt() As Integer
'    Public backk As Boolean = False
'    Public loaded As Boolean = False
'    Public prevtext As String
'    Public nexttext As String
'    Public casetext As String
'    'Public detect As Boolean
'    'Public from As String
'    'Public too As String
'    Public fromindex As Integer = 0
'    Public tooindex As Integer = 0
'    'Public wrap As Boolean
'    'Public lang As String = ""
'    Public srcfont() As Boolean
'    Public srcenc() As String
'    Public srccum() As String
'    Public srccount As Integer
'    Public Sub main()
'        '' Dim resid As Integer
'        ''Antidebugging() 'Antidebugging or AntiTracing Protection

'        'Dim ApplicationName As String
'        'Dim ApplicationKey As String

'        '' Set variable to allow key from you
'        'ApplicationName = "Indiclabs RTF Converter"
'        'ApplicationKey = "ampula1020i"

'        'Call SS_R("Smart Solutions", "L86QOpXVtLpKmFJO994JdqUxkIkoYjH4gB5Y79te6DNE1Y/y7aUcZ/kAX28wBZNC66vEfdlZzjj4o4p7Fim3QQ==")
'        'Call SetApplicationInfo(ApplicationName, ApplicationKey)
'        ''Call SS_DefaultKey("Demo", "25P8RZE4-8XK4TCEZ-TAM9J4HY-RJ45ERQJ-CGYVYL9P-CT8EVDLL", "42009601")
'        'Call SS_DefaultKey("Demo", "TL-CB3C93DDEJDeq2YnIyT60Z7Yjmabpp5Lpqm3eumA", "")

'        'Call SS_Initialize()
'        ''MessageBox.Show(SSUser(SS_GetUserName, SS_GetUserKey, ""))
'        'If SS_TrialMode = 99 Then
'        '    'MsgBox("Your system clock has been moved back ! please restore to correct date to use software.")
'        '    MessageBox.Show("Clock moved back detected!")
'        '    Application.Exit()
'        'End If
'        ''  Threading.Thread.Sleep(2000)
'        '' MessageBox.Show(SSUser(SS_GetUserName, SS_GetUserKey, "42009601"))

'        ''MessageBox.Show("SS_IsUnlocked() " & SS_IsUnlocked.ToString & Environment.NewLine & "SS_TrialExpired() " & SS_TrialExpired.ToString & Environment.NewLine & "SS_TrialMode() " & SS_TrialMode.ToString & Environment.NewLine & "SS_GetUserName() " & SS_GetUserName.ToString & Environment.NewLine & "SS_GetUserKey() " & SS_GetUserKey.ToString & Environment.NewLine) ' & "ResultID " & ResultID)
'        ''MessageBox.Show(SS_RemoveKey)
'        'If SS_IsUnlocked = False Then

'        '    ' Check if the trial period has been expired
'        '    If SS_TrialExpired() = False Then

'        '        ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
'        '        ' You can add your personal text with your language

'        '        'Trial Mode = Days Expiration = 1
'        '        If SS_TrialMode = -1 Then
'        '            'name_box.Text = SS_GetUserName
'        '            'key_box.Text = SS_GetUserKey
'        '            ' MessageBox.Show(SS_LicenseInfo() & " day(s) left")
'        '            Application.Run(New Form1)
'        '        End If
'        '        If SS_TrialMode = 1 Then
'        '            'name_box.Text = SS_GetUserName
'        '            'key_box.Text = SS_GetUserKey
'        '            MessageBox.Show(SS_LicenseInfo() & " day(s) left")
'        '            Application.Run(New DocumentExplorer.MainForm)
'        '        End If

'        '        'Trial Mode = DateExpiration = 2, return day remain before date expiration
'        '        If SS_TrialMode = 2 Then
'        '            'name_box.Text = SS_GetUserName
'        '            'key_box.Text = SS_GetUserKey
'        '            MessageBox.Show(SS_LicenseInfo() & " day(s) left before expiration")
'        '            Application.Run(New DocumentExplorer.MainForm)
'        '        End If

'        '        'Trial Mode = Run Count = 3
'        '        If SS_TrialMode = 3 Then
'        '            ' name_box.Text = SS_GetUserName
'        '            'key_box.Text = SS_GetUserKey
'        '            MessageBox.Show(SS_LicenseInfo() & " run(s) left")
'        '            Application.Run(New Form1)
'        '        End If

'        '        'Trial Mode = FreeMode  = 4
'        '        If SS_TrialMode = 4 Then
'        '            'name_box.Text = SS_GetUserName
'        '            'key_box.Text = SS_GetUserKey
'        '            MessageBox.Show("Free Mode")
'        '            Application.Run(New Form1)
'        '        End If


'        '    End If

'        '    'Check if the Trial Period has been expired
'        '    If SS_TrialExpired() = True Then
'        '        'MsgBox("Your trial period has been expired, please order now mysoftware name (e.g.)")
'        '        MessageBox.Show("Trial period has been expired !")
'        '        Application.Run(New Form1)
'        '    End If

'        'Else

'        '    MessageBox.Show("Registered Version")
'        '    Application.Run(New Form1)
'        '    'Retrieve User and Key informations
'        '    'name_box.Text = SS_GetUserName
'        '    '            key_box.Text = SS_GetUserKey

'        'End If
'        '' Set SerialShield for your application parameters and initialization

'        ''resid = SSUser(SS_GetUserName, SS_GetUserKey, "")
'        ''MessageBox.Show(resid & "  " & SS_GetUserName & "  " & SS_GetUserKey)
'        ''If resid = 1 Then
'        ''    ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
'        ''    ' You can add your personal text with your language

'        ''    'Trial Mode = Days Expiration = 1
'        ''    If SS_TrialMode = 1 Then

'        ''        MessageBox.Show(SS_LicenseInfo() & " day(s) left")
'        ''        Application.Run(New DocumentExplorer.MainForm)
'        ''    End If
'        ''ElseIf resid = 3 Then
'        ''    MessageBox.Show("Trial over")
'        ''    Application.Run(New Form1)
'        ''End If
'        ''If SS_IsUnlocked = False Then
'        ''    Try
'        ''        ' MessageBox.Show(SS_TrialMode)
'        ''        Application.Run(New Form1)
'        ''    Catch ex As Exception

'        ''    End Try
'        ''Else
'        ''    Application.Run(New DocumentExplorer.MainForm)
'        ''End If
'        If Date.Now > My.MySettings.Default.myData Then
'            ' MessageBox.Show("Application can't run")
'            'Me.Close()
'            Application.Exit()
'        End If
'        Application.Run(New DocumentExplorer.MainForm)
'    End Sub
'    Public Sub startup()

'        ' SerialShield DLL Licensing Protection

'        ' Antidebugging() 'Antidebugging or AntiTracing Protection


'    End Sub
'    Public Function RemoveDuplicateChars(ByVal key As String) As String
'        ' --- Removes duplicate chars using string concats. ---
'        ' Store encountered letters in this string.
'        Dim table As String = ""

'        ' Store the result in this string.
'        Dim result As String = ""

'        ' Loop over each character.
'        For Each value As Char In key
'            ' See if character is in the table.
'            If table.IndexOf(value) = -1 Then
'                ' Append to the table and the result.
'                table &= value
'                result &= value
'            End If
'        Next value
'        Return result
'    End Function


'End Module
