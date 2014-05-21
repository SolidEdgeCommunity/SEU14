<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASMEdgebarCtrl
    Inherits SolidEdge.CommonUI.EdgeBarControl

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.btnToggleSketches = New System.Windows.Forms.Button()
		Me.btnToggleRefPlanes = New System.Windows.Forms.Button()
		Me.btnToggleCsys = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'btnToggleSketches
		'
		Me.btnToggleSketches.Image = Global.My.Resources.MyResources.btnSketch
		Me.btnToggleSketches.Location = New System.Drawing.Point(77, 1)
		Me.btnToggleSketches.Name = "btnToggleSketches"
		Me.btnToggleSketches.Size = New System.Drawing.Size(32, 32)
		Me.btnToggleSketches.TabIndex = 1
		Me.btnToggleSketches.UseVisualStyleBackColor = True
		'
		'btnToggleRefPlanes
		'
		Me.btnToggleRefPlanes.Image = Global.My.Resources.MyResources.btnRefplanes
		Me.btnToggleRefPlanes.Location = New System.Drawing.Point(39, 1)
		Me.btnToggleRefPlanes.Name = "btnToggleRefPlanes"
		Me.btnToggleRefPlanes.Size = New System.Drawing.Size(32, 32)
		Me.btnToggleRefPlanes.TabIndex = 1
		Me.btnToggleRefPlanes.UseVisualStyleBackColor = True
		'
		'btnToggleCsys
		'
		Me.btnToggleCsys.Image = Global.My.Resources.MyResources.btnCsys
		Me.btnToggleCsys.Location = New System.Drawing.Point(1, 1)
		Me.btnToggleCsys.Name = "btnToggleCsys"
		Me.btnToggleCsys.Size = New System.Drawing.Size(32, 32)
		Me.btnToggleCsys.TabIndex = 0
		Me.btnToggleCsys.UseVisualStyleBackColor = True
		'
		'ASMEdgebarCtrl
		'
		Me.BitmapID = 100
		Me.Controls.Add(Me.btnToggleSketches)
		Me.Controls.Add(Me.btnToggleRefPlanes)
		Me.Controls.Add(Me.btnToggleCsys)
		Me.Name = "ASMEdgebarCtrl"
		Me.Size = New System.Drawing.Size(545, 51)
		Me.ToolTip = "ASM Commands"
		Me.ResumeLayout(False)

End Sub
		Friend WithEvents btnToggleCsys As System.Windows.Forms.Button
	Friend WithEvents btnToggleRefPlanes As System.Windows.Forms.Button
	Friend WithEvents btnToggleSketches As System.Windows.Forms.Button

End Class
