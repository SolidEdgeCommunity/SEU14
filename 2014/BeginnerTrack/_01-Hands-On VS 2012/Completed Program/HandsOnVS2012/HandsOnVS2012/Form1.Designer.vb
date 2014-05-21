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
        Me.btnStartSE = New System.Windows.Forms.Button()
        Me.chkboxNewSession = New System.Windows.Forms.CheckBox()
        Me.groupSETemplateOpt = New System.Windows.Forms.GroupBox()
        Me.radbuttonPart = New System.Windows.Forms.RadioButton()
        Me.radbuttonAssembly = New System.Windows.Forms.RadioButton()
        Me.groupSETemplateOpt.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStartSE
        '
        Me.btnStartSE.Location = New System.Drawing.Point(29, 140)
        Me.btnStartSE.Name = "btnStartSE"
        Me.btnStartSE.Size = New System.Drawing.Size(135, 50)
        Me.btnStartSE.TabIndex = 0
        Me.btnStartSE.Text = "Start Solid Edge"
        Me.btnStartSE.UseVisualStyleBackColor = True
        '
        'chkboxNewSession
        '
        Me.chkboxNewSession.AutoSize = True
        Me.chkboxNewSession.Location = New System.Drawing.Point(52, 92)
        Me.chkboxNewSession.Name = "chkboxNewSession"
        Me.chkboxNewSession.Size = New System.Drawing.Size(145, 21)
        Me.chkboxNewSession.TabIndex = 1
        Me.chkboxNewSession.Text = "Start New Session"
        Me.chkboxNewSession.UseVisualStyleBackColor = True
        '
        'groupSETemplateOpt
        '
        Me.groupSETemplateOpt.Controls.Add(Me.radbuttonPart)
        Me.groupSETemplateOpt.Controls.Add(Me.radbuttonAssembly)
        Me.groupSETemplateOpt.Location = New System.Drawing.Point(29, 12)
        Me.groupSETemplateOpt.Name = "groupSETemplateOpt"
        Me.groupSETemplateOpt.Size = New System.Drawing.Size(298, 61)
        Me.groupSETemplateOpt.TabIndex = 3
        Me.groupSETemplateOpt.TabStop = False
        Me.groupSETemplateOpt.Text = "SE Template Options"
        '
        'radbuttonPart
        '
        Me.radbuttonPart.AutoSize = True
        Me.radbuttonPart.Location = New System.Drawing.Point(211, 31)
        Me.radbuttonPart.Name = "radbuttonPart"
        Me.radbuttonPart.Size = New System.Drawing.Size(55, 21)
        Me.radbuttonPart.TabIndex = 1
        Me.radbuttonPart.TabStop = True
        Me.radbuttonPart.Text = "Part"
        Me.radbuttonPart.UseVisualStyleBackColor = True
        '
        'radbuttonAssembly
        '
        Me.radbuttonAssembly.AutoSize = True
        Me.radbuttonAssembly.Location = New System.Drawing.Point(23, 31)
        Me.radbuttonAssembly.Name = "radbuttonAssembly"
        Me.radbuttonAssembly.Size = New System.Drawing.Size(89, 21)
        Me.radbuttonAssembly.TabIndex = 0
        Me.radbuttonAssembly.TabStop = True
        Me.radbuttonAssembly.Text = "Assembly"
        Me.radbuttonAssembly.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(468, 238)
        Me.Controls.Add(Me.groupSETemplateOpt)
        Me.Controls.Add(Me.chkboxNewSession)
        Me.Controls.Add(Me.btnStartSE)
        Me.Name = "Form1"
        Me.Text = "My First Application"
        Me.groupSETemplateOpt.ResumeLayout(False)
        Me.groupSETemplateOpt.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnStartSE As System.Windows.Forms.Button
    Friend WithEvents chkboxNewSession As System.Windows.Forms.CheckBox
    Friend WithEvents groupSETemplateOpt As System.Windows.Forms.GroupBox
    Friend WithEvents radbuttonPart As System.Windows.Forms.RadioButton
    Friend WithEvents radbuttonAssembly As System.Windows.Forms.RadioButton

End Class
