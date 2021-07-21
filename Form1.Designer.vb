<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.exit_btn = New System.Windows.Forms.Button
        Me.setkey_btn = New System.Windows.Forms.Button
        Me.name_lbl = New System.Windows.Forms.Label
        Me.name_box = New System.Windows.Forms.TextBox
        Me.message_lbl = New System.Windows.Forms.Label
        Me.key_lbl = New System.Windows.Forms.Label
        Me.mac_id_lbl = New System.Windows.Forms.Label
        Me.message_box = New System.Windows.Forms.TextBox
        Me.key_box = New System.Windows.Forms.TextBox
        Me.mac_id_box = New System.Windows.Forms.TextBox
        Me.remove_btn = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'exit_btn
        '
        Me.exit_btn.Location = New System.Drawing.Point(198, 185)
        Me.exit_btn.Name = "exit_btn"
        Me.exit_btn.Size = New System.Drawing.Size(85, 23)
        Me.exit_btn.TabIndex = 21
        Me.exit_btn.Text = "Exit"
        '
        'setkey_btn
        '
        Me.setkey_btn.Location = New System.Drawing.Point(83, 185)
        Me.setkey_btn.Name = "setkey_btn"
        Me.setkey_btn.Size = New System.Drawing.Size(85, 23)
        Me.setkey_btn.TabIndex = 20
        Me.setkey_btn.Text = "Activate"
        '
        'name_lbl
        '
        Me.name_lbl.Location = New System.Drawing.Point(3, 73)
        Me.name_lbl.Name = "name_lbl"
        Me.name_lbl.Size = New System.Drawing.Size(64, 16)
        Me.name_lbl.TabIndex = 19
        Me.name_lbl.Text = "Serial:"
        '
        'name_box
        '
        Me.name_box.Location = New System.Drawing.Point(83, 73)
        Me.name_box.Name = "name_box"
        Me.name_box.Size = New System.Drawing.Size(296, 20)
        Me.name_box.TabIndex = 18
        '
        'message_lbl
        '
        Me.message_lbl.Location = New System.Drawing.Point(3, 153)
        Me.message_lbl.Name = "message_lbl"
        Me.message_lbl.Size = New System.Drawing.Size(64, 23)
        Me.message_lbl.TabIndex = 17
        Me.message_lbl.Text = "Status"
        '
        'key_lbl
        '
        Me.key_lbl.Location = New System.Drawing.Point(3, 105)
        Me.key_lbl.Name = "key_lbl"
        Me.key_lbl.Size = New System.Drawing.Size(64, 34)
        Me.key_lbl.TabIndex = 16
        Me.key_lbl.Text = "Activation Key:"
        '
        'mac_id_lbl
        '
        Me.mac_id_lbl.Location = New System.Drawing.Point(3, 41)
        Me.mac_id_lbl.Name = "mac_id_lbl"
        Me.mac_id_lbl.Size = New System.Drawing.Size(72, 23)
        Me.mac_id_lbl.TabIndex = 15
        Me.mac_id_lbl.Text = "Machine ID :"
        '
        'message_box
        '
        Me.message_box.Location = New System.Drawing.Point(83, 153)
        Me.message_box.Name = "message_box"
        Me.message_box.ReadOnly = True
        Me.message_box.Size = New System.Drawing.Size(296, 20)
        Me.message_box.TabIndex = 14
        '
        'key_box
        '
        Me.key_box.Location = New System.Drawing.Point(83, 105)
        Me.key_box.Name = "key_box"
        Me.key_box.Size = New System.Drawing.Size(296, 20)
        Me.key_box.TabIndex = 13
        '
        'mac_id_box
        '
        Me.mac_id_box.Location = New System.Drawing.Point(83, 41)
        Me.mac_id_box.Name = "mac_id_box"
        Me.mac_id_box.ReadOnly = True
        Me.mac_id_box.Size = New System.Drawing.Size(296, 20)
        Me.mac_id_box.TabIndex = 12
        '
        'remove_btn
        '
        Me.remove_btn.Location = New System.Drawing.Point(309, 185)
        Me.remove_btn.Name = "remove_btn"
        Me.remove_btn.Size = New System.Drawing.Size(85, 23)
        Me.remove_btn.TabIndex = 22
        Me.remove_btn.Text = "Remove Key"
        Me.remove_btn.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(406, 223)
        Me.Controls.Add(Me.remove_btn)
        Me.Controls.Add(Me.exit_btn)
        Me.Controls.Add(Me.setkey_btn)
        Me.Controls.Add(Me.name_lbl)
        Me.Controls.Add(Me.name_box)
        Me.Controls.Add(Me.message_lbl)
        Me.Controls.Add(Me.key_lbl)
        Me.Controls.Add(Me.mac_id_lbl)
        Me.Controls.Add(Me.message_box)
        Me.Controls.Add(Me.key_box)
        Me.Controls.Add(Me.mac_id_box)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "RTF Converter Activation"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents exit_btn As System.Windows.Forms.Button
    Friend WithEvents setkey_btn As System.Windows.Forms.Button
    Friend WithEvents name_lbl As System.Windows.Forms.Label
    Friend WithEvents name_box As System.Windows.Forms.TextBox
    Friend WithEvents message_lbl As System.Windows.Forms.Label
    Friend WithEvents key_lbl As System.Windows.Forms.Label
    Friend WithEvents mac_id_lbl As System.Windows.Forms.Label
    Friend WithEvents message_box As System.Windows.Forms.TextBox
    Friend WithEvents key_box As System.Windows.Forms.TextBox
    Friend WithEvents mac_id_box As System.Windows.Forms.TextBox
    Friend WithEvents remove_btn As System.Windows.Forms.Button
End Class
