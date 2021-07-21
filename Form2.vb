Imports System
Imports System.Windows.Forms
Public Class Form2
    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'Form2
        '
        Me.ClientSize = New System.Drawing.Size(1005, 695)
        Me.Name = "Form2"
        Me.ResumeLayout(False)

    End Sub

    'Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    '    Try
    '        Antidebugging() 'Antidebugging or AntiTracing Protection

    '        Dim ApplicationName As String
    '        Dim ApplicationKey As String

    '        ' Set variable to allow key from you
    '        ApplicationName = "Indiclabs RTF Converter"
    '        ApplicationKey = "ampula1020i"

    '        Call SS_R("Smart Solutions", "L86QOpXVtLpKmFJO994JdqUxkIkoYjH4gB5Y79te6DNE1Y/y7aUcZ/kAX28wBZNC66vEfdlZzjj4o4p7Fim3QQ==")
    '        Call SetApplicationInfo(ApplicationName, ApplicationKey)
    '        Call SS_Initialize()

    '        mac_id_box.Text = GetMachineID()
    '        ' Antidebugging() 'Antidebugging or AntiTracing Protection

    '        Dim ResultID As Integer
    '        message_box.Text = ""
    '        ' ResultID = SSUser(name_box.Text, key_box.Text, "")
    '        'message_box.Text = ("SS_IsUnlocked() " & SS_IsUnlocked.ToString & Environment.NewLine & "SS_TrialExpired() " & SS_TrialExpired.ToString & Environment.NewLine & "SS_TrialMode() " & SS_TrialMode.ToString & Environment.NewLine &  "SS_GetUserName() " & SS_GetUserName.ToString & Environment.NewLine &  "SS_GetUserKey() " & SS_GetUserKey.ToString & Environment.NewLine & "ResultID " & ResultID)

    '        'Retrieve User and Key informations
    '        'name_box.Text = SS_GetUserName
    '        'key_box.Text = SS_GetUserKey

    '        If SS_IsUnlocked = False Then

    '            ' Check if the trial period has been expired
    '            If SS_TrialExpired() = False Then

    '                ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
    '                ' You can add your personal text with your language

    '                'Trial Mode = Days Expiration = 1
    '                If SS_TrialMode = 1 Then
    '                    name_box.Text = SS_GetUserName
    '                    key_box.Text = SS_GetUserKey
    '                    message_box.Text = SS_LicenseInfo() & " day(s) left"
    '                End If

    '                'Trial Mode = DateExpiration = 2, return day remain before date expiration
    '                If SS_TrialMode = 2 Then
    '                    name_box.Text = SS_GetUserName
    '                    key_box.Text = SS_GetUserKey
    '                    message_box.Text = SS_LicenseInfo() & " day(s) left before expiration"
    '                End If

    '                'Trial Mode = Run Count = 3
    '                If SS_TrialMode = 3 Then
    '                    name_box.Text = SS_GetUserName
    '                    key_box.Text = SS_GetUserKey
    '                    message_box.Text = SS_LicenseInfo() & " run(s) left"
    '                End If

    '                'Trial Mode = FreeMode  = 4
    '                If SS_TrialMode = 4 Then
    '                    name_box.Text = SS_GetUserName
    '                    key_box.Text = SS_GetUserKey
    '                    message_box.Text = "Free Mode"
    '                End If


    '            End If

    '            'Check if the Trial Period has been expired
    '            If SS_TrialExpired() = True Then
    '                '  MsgBox("Your trial period has been expired, please order now mysoftware name (e.g.)")
    '                message_box.Text = "Trial period has been expired !"
    '            End If

    '        Else
    '            message_box.Text = "Registered Version"
    '            'Retrieve User and Key informations
    '            name_box.Text = SS_GetUserName
    '            key_box.Text = SS_GetUserKey

    '        End If
    '        ' MessageBox.Show(SSUser(SS_GetUserName, SS_GetUserKey, "42009601"))
    '        If SSUser(SS_GetUserName, SS_GetUserKey, "42009601") = 1 Then
    '            ' SS_RemoveKey()
    '            Me.Close()
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub
    'Public Sub startup()

    '    ' SerialShield DLL Licensing Protection

    '    Antidebugging() 'Antidebugging or AntiTracing Protection

    '    Dim ApplicationName As String
    '    Dim ApplicationKey As String

    '    ' Set variable to allow key from you
    '    ApplicationName = "Indiclabs RTF Converter"
    '    ApplicationKey = "ampula1020i"

    '    Call SS_R("Smart Solutions", "L86QOpXVtLpKmFJO994JdqUxkIkoYjH4gB5Y79te6DNE1Y/y7aUcZ/kAX28wBZNC66vEfdlZzjj4o4p7Fim3QQ==")
    '    Call SetApplicationInfo(ApplicationName, ApplicationKey)
    '    Call SS_Initialize()

    '    'Check if system clock has been moved back
    '    'If SS_TrialMode = 99 Then
    '    '    MsgBox("Your system clock has been moved back ! please restore to correct date to use software.")
    '    '    message_box.Text = "Clock moved back detected!"
    '    '    Exit Sub
    '    'End If

    '    'result_box.Text = ("SS_IsUnlocked() " & SS_IsUnlocked.ToString & vbNewLine & _
    '    '"SS_TrialExpired() " & SS_TrialExpired.ToString & vbNewLine & _
    '    '"SS_TrialMode() " & SS_TrialMode.ToString & vbNewLine & _
    '    '"SS_GetUserName() " & SS_GetUserName.ToString & vbNewLine & _
    '    '"SS_GetUserKey() " & SS_GetUserKey.ToString)


    '    '' Check if your software is unlocked mode
    '    'If SS_IsUnlocked = False Then

    '    '    ' Check if the trial period has been expired
    '    '    If SS_TrialExpired() = False Then

    '    '        ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
    '    '        ' You can add your personal text with your language

    '    '        'Trial Mode = Days Expiration = 1
    '    '        If SS_TrialMode = 1 Then
    '    '            name_box.Text = SS_GetUserName
    '    '            key_box.Text = SS_GetUserKey
    '    '            message_box.Text = SS_LicenseInfo() & " day(s) left"
    '    '        End If

    '    '        'Trial Mode = DateExpiration = 2, return day remain before date expiration
    '    '        If SS_TrialMode = 2 Then
    '    '            name_box.Text = SS_GetUserName
    '    '            key_box.Text = SS_GetUserKey
    '    '            message_box.Text = SS_LicenseInfo() & " day(s) left before expiration"
    '    '        End If

    '    '        'Trial Mode = Run Count = 3
    '    '        If SS_TrialMode = 3 Then
    '    '            name_box.Text = SS_GetUserName
    '    '            key_box.Text = SS_GetUserKey
    '    '            message_box.Text = SS_LicenseInfo() & " run(s) left"
    '    '        End If

    '    '        'Trial Mode = FreeMode  = 4
    '    '        If SS_TrialMode = 4 Then
    '    '            name_box.Text = SS_GetUserName
    '    '            key_box.Text = SS_GetUserKey
    '    '            message_box.Text = "Free Mode"
    '    '        End If


    '    '    End If

    '    '    'Check if the Trial Period has been expired
    '    '    If SS_TrialExpired() = True Then
    '    '        MsgBox("Your trial period has been expired, please order now mysoftware name (e.g.)")
    '    '        message_box.Text = "Trial period has been expired !"
    '    '    End If

    '    'Else
    '    '    message_box.Text = "Registered Version"
    '    '    'Retrieve User and Key informations
    '    '    name_box.Text = SS_GetUserName
    '    '    key_box.Text = SS_GetUserKey

    '    'End If
    'End Sub
    'Private Sub remove_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles remove_btn.Click
    '    Try
    '        If SS_RemoveKey = True Then
    '            ' MsgBox("The current key has been removed.")
    '            message_box.Text = "The key has been removed."

    '            'Retrieve User and Key informations
    '            name_box.Text = SS_GetUserName
    '            key_box.Text = SS_GetUserKey
    '            '            result_box.Text = ("SS_IsUnlocked() " & SS_IsUnlocked.ToString & vbNewLine & _
    '            '"SS_TrialExpired() " & SS_TrialExpired.ToString & vbNewLine & _
    '            '"SS_TrialMode() " & SS_TrialMode.ToString & vbNewLine & _
    '            '"SS_GetUserName() " & SS_GetUserName.ToString & vbNewLine & _
    '            '"SS_GetUserKey() " & SS_GetUserKey.ToString)
    '        End If
    '    Catch ex As Exception
    '        message_box.Text = ex.Message
    '    End Try
    'End Sub

    'Private Sub exit_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exit_btn.Click
    '    Try
    '        Application.Exit()
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub setkey_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles setkey_btn.Click
    '    Try
    '        'Antidebugging() 'Antidebugging or AntiTracing Protection

    '        'Dim ApplicationName As String
    '        'Dim ApplicationKey As String

    '        '' Set variable to allow key from you
    '        'ApplicationName = "IndicLabs RTf Converter"
    '        'ApplicationKey = "@mpul@1020I"

    '        'Call SS_R("Smart Solutions", "L86QOpXVtLpKmFJO994JdqUxkIkoYjH4gB5Y79te6DNE1Y/y7aUcZ/kAX28wBZNC66vEfdlZzjj4o4p7Fim3QQ==")
    '        'Call SetApplicationInfo(ApplicationName, ApplicationKey)
    '        'Call SS_Initialize()

    '        ' Set SerialShield for your application parameters and initialization

    '        Dim ResultID As Integer
    '        message_box.Text = ""
    '        ResultID = SSUser(name_box.Text, key_box.Text, "42009601")
    '        'result_box.Text = ("SS_IsUnlocked() " & SS_IsUnlocked.ToString & vbNewLine & _
    '        '"SS_TrialExpired() " & SS_TrialExpired.ToString & vbNewLine & _
    '        '"SS_TrialMode() " & SS_TrialMode.ToString & vbNewLine & _
    '        '"SS_GetUserName() " & SS_GetUserName.ToString & vbNewLine & _
    '        '"SS_GetUserKey() " & SS_GetUserKey.ToString & vbNewLine & _
    '        '"ResultID " & ResultID)

    '        'Retrieve User and Key informations
    '        name_box.Text = SS_GetUserName
    '        key_box.Text = SS_GetUserKey

    '        If ResultID = 4 Then
    '            message_box.Text = "Software Full Version Mode"
    '            ' MsgBox("Software Unlocked")
    '        End If

    '        If ResultID = 3 Then message_box.Text = ("Key incorrect!")

    '        If ResultID = 1 Then
    '            ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
    '            ' You can add your personal text with your language

    '            'Trial Mode = Days Expiration = 1
    '            If SS_TrialMode = 1 Then
    '                message_box.Text = (SS_LicenseInfo() & " day(s) left")
    '            End If

    '            'Trial Mode = DateExpiration = 2, return day remain before date expiration
    '            If SS_TrialMode = 2 Then
    '                message_box.Text = (SS_LicenseInfo() & " day(s) left before expiration")
    '            End If

    '            'Trial Mode = Run Count = 3
    '            If SS_TrialMode = 3 Then
    '                message_box.Text = (SS_LicenseInfo() & " run(s) left")
    '            End If

    '            'Trial Mode = FreeMode  = 4
    '            If SS_TrialMode = 4 Then
    '                message_box.Text = ("Free Mode")
    '            End If

    '        End If
    '        If SS_TrialMode = 1 Then
    '            'name_box.Text = SS_GetUserName
    '            'key_box.Text = SS_GetUserKey
    '            MessageBox.Show(SS_LicenseInfo() & " dayy(s) left")
    '            Application.Run(New Form1)
    '        End If
    '        Me.Close()
    '    Catch ex As Exception

    '    End Try
    'End Sub
End Class