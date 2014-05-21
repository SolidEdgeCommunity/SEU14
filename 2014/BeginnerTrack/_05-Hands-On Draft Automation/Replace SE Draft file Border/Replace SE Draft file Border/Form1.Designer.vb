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
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtStatusFile = New System.Windows.Forms.TextBox()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TxtStatus = New System.Windows.Forms.TextBox()
        Me.TxtCloseAfter = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RBShowSE = New System.Windows.Forms.RadioButton()
        Me.RBHideSE = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optAllFiles = New System.Windows.Forms.RadioButton()
        Me.optAllInDirectory = New System.Windows.Forms.RadioButton()
        Me.optSelected = New System.Windows.Forms.RadioButton()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.TxtFolderName = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TxtFileContainingNewBorder = New System.Windows.Forms.TextBox()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(622, 344)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 17)
        Me.Label4.TabIndex = 42
        Me.Label4.Text = "Version 1.00.00.01"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(9, 405)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(62, 17)
        Me.Label2.TabIndex = 41
        Me.Label2.Text = "Log File:"
        '
        'TxtStatusFile
        '
        Me.TxtStatusFile.Enabled = False
        Me.TxtStatusFile.Location = New System.Drawing.Point(13, 424)
        Me.TxtStatusFile.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtStatusFile.Name = "TxtStatusFile"
        Me.TxtStatusFile.Size = New System.Drawing.Size(735, 22)
        Me.TxtStatusFile.TabIndex = 40
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(512, 338)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 28)
        Me.Button4.TabIndex = 39
        Me.Button4.Text = "View LogFile"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(576, 450)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(67, 28)
        Me.Button3.TabIndex = 38
        Me.Button3.Text = "OK"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(649, 450)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 37
        Me.Button2.Text = "Cancel"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TxtStatus
        '
        Me.TxtStatus.Location = New System.Drawing.Point(13, 376)
        Me.TxtStatus.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtStatus.Name = "TxtStatus"
        Me.TxtStatus.Size = New System.Drawing.Size(735, 22)
        Me.TxtStatus.TabIndex = 36
        '
        'TxtCloseAfter
        '
        Me.TxtCloseAfter.Location = New System.Drawing.Point(429, 344)
        Me.TxtCloseAfter.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtCloseAfter.Name = "TxtCloseAfter"
        Me.TxtCloseAfter.Size = New System.Drawing.Size(55, 22)
        Me.TxtCloseAfter.TabIndex = 35
        Me.TxtCloseAfter.Text = "30"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 344)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(417, 17)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Close and restart Solid Edge after processing this number of files"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RBShowSE)
        Me.GroupBox2.Controls.Add(Me.RBHideSE)
        Me.GroupBox2.Location = New System.Drawing.Point(444, 234)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox2.Size = New System.Drawing.Size(307, 81)
        Me.GroupBox2.TabIndex = 33
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Solid Edge Visibility"
        Me.GroupBox2.Visible = False
        '
        'RBShowSE
        '
        Me.RBShowSE.AutoSize = True
        Me.RBShowSE.Checked = True
        Me.RBShowSE.Location = New System.Drawing.Point(9, 53)
        Me.RBShowSE.Margin = New System.Windows.Forms.Padding(4)
        Me.RBShowSE.Name = "RBShowSE"
        Me.RBShowSE.Size = New System.Drawing.Size(252, 21)
        Me.RBShowSE.TabIndex = 1
        Me.RBShowSE.TabStop = True
        Me.RBShowSE.Text = "Show Solid Edge during processing"
        Me.RBShowSE.UseVisualStyleBackColor = True
        '
        'RBHideSE
        '
        Me.RBHideSE.AutoSize = True
        Me.RBHideSE.Location = New System.Drawing.Point(9, 25)
        Me.RBHideSE.Margin = New System.Windows.Forms.Padding(4)
        Me.RBHideSE.Name = "RBHideSE"
        Me.RBHideSE.Size = New System.Drawing.Size(247, 21)
        Me.RBHideSE.TabIndex = 0
        Me.RBHideSE.Text = "Hide Solid Edge during processing"
        Me.RBHideSE.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optAllFiles)
        Me.GroupBox1.Controls.Add(Me.optAllInDirectory)
        Me.GroupBox1.Controls.Add(Me.optSelected)
        Me.GroupBox1.Location = New System.Drawing.Point(442, 115)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(308, 111)
        Me.GroupBox1.TabIndex = 32
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Define Scope"
        '
        'optAllFiles
        '
        Me.optAllFiles.AutoSize = True
        Me.optAllFiles.Checked = True
        Me.optAllFiles.Location = New System.Drawing.Point(8, 81)
        Me.optAllFiles.Margin = New System.Windows.Forms.Padding(4)
        Me.optAllFiles.Name = "optAllFiles"
        Me.optAllFiles.Size = New System.Drawing.Size(288, 21)
        Me.optAllFiles.TabIndex = 2
        Me.optAllFiles.TabStop = True
        Me.optAllFiles.Text = "All files in selected folder and sub-folders"
        Me.optAllFiles.UseVisualStyleBackColor = True
        '
        'optAllInDirectory
        '
        Me.optAllInDirectory.AutoSize = True
        Me.optAllInDirectory.Location = New System.Drawing.Point(8, 52)
        Me.optAllInDirectory.Margin = New System.Windows.Forms.Padding(4)
        Me.optAllInDirectory.Name = "optAllInDirectory"
        Me.optAllInDirectory.Size = New System.Drawing.Size(227, 21)
        Me.optAllInDirectory.TabIndex = 1
        Me.optAllInDirectory.Text = "All files in selected folder ONLY"
        Me.optAllInDirectory.UseVisualStyleBackColor = True
        '
        'optSelected
        '
        Me.optSelected.AutoSize = True
        Me.optSelected.Location = New System.Drawing.Point(8, 23)
        Me.optSelected.Margin = New System.Windows.Forms.Padding(4)
        Me.optSelected.Name = "optSelected"
        Me.optSelected.Size = New System.Drawing.Size(155, 21)
        Me.optSelected.TabIndex = 0
        Me.optSelected.Text = "Selected files ONLY"
        Me.optSelected.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(13, 115)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.ScrollAlwaysVisible = True
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBox1.Size = New System.Drawing.Size(420, 212)
        Me.ListBox1.TabIndex = 31
        '
        'TxtFolderName
        '
        Me.TxtFolderName.Location = New System.Drawing.Point(94, 83)
        Me.TxtFolderName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtFolderName.Name = "TxtFolderName"
        Me.TxtFolderName.Size = New System.Drawing.Size(655, 22)
        Me.TxtFolderName.TabIndex = 30
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(13, 80)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(73, 28)
        Me.Button1.TabIndex = 29
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 59)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(690, 17)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "Browse to select the top most folder containing Solid Edge draft files where the " & _
            "border needs to be replaced:"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(16, 31)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(75, 23)
        Me.Button5.TabIndex = 44
        Me.Button5.Text = "Browse"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(434, 17)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "Browse to select the Solid Edge draft file containing the new border:"
        '
        'TxtFileContainingNewBorder
        '
        Me.TxtFileContainingNewBorder.Location = New System.Drawing.Point(98, 31)
        Me.TxtFileContainingNewBorder.Name = "TxtFileContainingNewBorder"
        Me.TxtFileContainingNewBorder.Size = New System.Drawing.Size(650, 22)
        Me.TxtFileContainingNewBorder.TabIndex = 46
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(770, 484)
        Me.Controls.Add(Me.TxtFileContainingNewBorder)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TxtStatusFile)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TxtStatus)
        Me.Controls.Add(Me.TxtCloseAfter)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.TxtFolderName)
        Me.Controls.Add(Me.Button1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Replace draft border utility"
        Me.TopMost = True
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TxtStatusFile As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TxtStatus As System.Windows.Forms.TextBox
    Friend WithEvents TxtCloseAfter As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RBShowSE As System.Windows.Forms.RadioButton
    Friend WithEvents RBHideSE As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optAllFiles As System.Windows.Forms.RadioButton
    Friend WithEvents optAllInDirectory As System.Windows.Forms.RadioButton
    Friend WithEvents optSelected As System.Windows.Forms.RadioButton
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents TxtFolderName As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TxtFileContainingNewBorder As System.Windows.Forms.TextBox

End Class
