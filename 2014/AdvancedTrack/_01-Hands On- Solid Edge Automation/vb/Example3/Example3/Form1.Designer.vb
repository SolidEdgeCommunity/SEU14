Namespace Example3
	Partial Public Class Form1
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.button1 = New System.Windows.Forms.Button()
			Me.label1 = New System.Windows.Forms.Label()
			Me.backgroundWorker1 = New System.ComponentModel.BackgroundWorker()
			Me.SuspendLayout()
			' 
			' button1
			' 
			Me.button1.Location = New System.Drawing.Point(31, 59)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(254, 40)
			Me.button1.TabIndex = 0
			Me.button1.Text = "Execute task in separate AppDomain."
			Me.button1.UseVisualStyleBackColor = True
'			Me.button1.Click += New System.EventHandler(Me.button1_Click)
			' 
			' label1
			' 
			Me.label1.AutoSize = True
			Me.label1.Location = New System.Drawing.Point(64, 122)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(177, 13)
			Me.label1.TabIndex = 1
			Me.label1.Text = "Notice that the UI is NOT blocked..."
			Me.label1.Visible = False
			' 
			' backgroundWorker1
			' 
'			Me.backgroundWorker1.DoWork += New System.ComponentModel.DoWorkEventHandler(Me.backgroundWorker1_DoWork)
'			Me.backgroundWorker1.RunWorkerCompleted += New System.ComponentModel.RunWorkerCompletedEventHandler(Me.backgroundWorker1_RunWorkerCompleted)
			' 
			' Form1
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(323, 174)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.button1)
			Me.Name = "Form1"
			Me.Text = "Form1"
'			Me.Load += New System.EventHandler(Me.Form1_Load)
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub

		#End Region

		Private WithEvents button1 As System.Windows.Forms.Button
		Private label1 As System.Windows.Forms.Label
		Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker
	End Class
End Namespace

