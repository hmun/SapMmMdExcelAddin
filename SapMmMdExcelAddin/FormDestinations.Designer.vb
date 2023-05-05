<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormDestinations
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
        Me.ListBoxDest = New System.Windows.Forms.ListBox()
        Me.Button_OK = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBoxDest
        '
        Me.ListBoxDest.FormattingEnabled = True
        Me.ListBoxDest.Location = New System.Drawing.Point(12, 19)
        Me.ListBoxDest.Name = "ListBoxDest"
        Me.ListBoxDest.Size = New System.Drawing.Size(260, 186)
        Me.ListBoxDest.TabIndex = 0
        '
        'Button_OK
        '
        Me.Button_OK.Location = New System.Drawing.Point(12, 224)
        Me.Button_OK.Name = "Button_OK"
        Me.Button_OK.Size = New System.Drawing.Size(94, 26)
        Me.Button_OK.TabIndex = 1
        Me.Button_OK.Text = "OK"
        Me.Button_OK.UseVisualStyleBackColor = True
        '
        'FormDestinations
        '
        Me.AcceptButton = Me.Button_OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.Button_OK)
        Me.Controls.Add(Me.ListBoxDest)
        Me.Name = "FormDestinations"
        Me.Text = "SAP Destinations"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBoxDest As System.Windows.Forms.ListBox
    Friend WithEvents Button_OK As System.Windows.Forms.Button
End Class
