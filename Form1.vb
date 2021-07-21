Imports System
Imports System.Windows.Forms
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
        Catch ex As Exception

        End Try
    End Sub
    Public Sub startup()

        ' SerialShield DLL Licensing Protection


        'Check if system clock has been moved back
        'If SS_TrialMode = 99 Then
        '    MsgBox("Your system clock has been moved back ! please restore to correct date to use software.")
        '    message_box.Text = "Clock moved back detected!"
        '    Exit Sub
        'End If

        'result_box.Text = ("SS_IsUnlocked() " & SS_IsUnlocked.ToString & vbNewLine & _
        '"SS_TrialExpired() " & SS_TrialExpired.ToString & vbNewLine & _
        '"SS_TrialMode() " & SS_TrialMode.ToString & vbNewLine & _
        '"SS_GetUserName() " & SS_GetUserName.ToString & vbNewLine & _
        '"SS_GetUserKey() " & SS_GetUserKey.ToString)


        '' Check if your software is unlocked mode
        'If SS_IsUnlocked = False Then

        '    ' Check if the trial period has been expired
        '    If SS_TrialExpired() = False Then

        '        ' SS_TrialMode is function to check which Trial Mode is used (Days = 1, Date = 2, Run = 3 Expiration)
        '        ' You can add your personal text with your language

        '        'Trial Mode = Days Expiration = 1
        '        If SS_TrialMode = 1 Then
        '            name_box.Text = SS_GetUserName
        '            key_box.Text = SS_GetUserKey
        '            message_box.Text = SS_LicenseInfo() & " day(s) left"
        '        End If

        '        'Trial Mode = DateExpiration = 2, return day remain before date expiration
        '        If SS_TrialMode = 2 Then
        '            name_box.Text = SS_GetUserName
        '            key_box.Text = SS_GetUserKey
        '            message_box.Text = SS_LicenseInfo() & " day(s) left before expiration"
        '        End If

        '        'Trial Mode = Run Count = 3
        '        If SS_TrialMode = 3 Then
        '            name_box.Text = SS_GetUserName
        '            key_box.Text = SS_GetUserKey
        '            message_box.Text = SS_LicenseInfo() & " run(s) left"
        '        End If

        '        'Trial Mode = FreeMode  = 4
        '        If SS_TrialMode = 4 Then
        '            name_box.Text = SS_GetUserName
        '            key_box.Text = SS_GetUserKey
        '            message_box.Text = "Free Mode"
        '        End If


        '    End If

        '    'Check if the Trial Period has been expired
        '    If SS_TrialExpired() = True Then
        '        MsgBox("Your trial period has been expired, please order now mysoftware name (e.g.)")
        '        message_box.Text = "Trial period has been expired !"
        '    End If

        'Else
        '    message_box.Text = "Registered Version"
        '    'Retrieve User and Key informations
        '    name_box.Text = SS_GetUserName
        '    key_box.Text = SS_GetUserKey

        'End If
    End Sub


    Private Sub exit_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exit_btn.Click
        Try
            Application.Exit()
        Catch ex As Exception

        End Try
    End Sub

End Class