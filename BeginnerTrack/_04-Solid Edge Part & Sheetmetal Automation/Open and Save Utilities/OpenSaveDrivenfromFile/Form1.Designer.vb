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
        Me.CBXRecompute = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TxtLogFileName = New System.Windows.Forms.TextBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnProcess = New System.Windows.Forms.Button()
        Me.TxtTextFilename = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.CBCreadtePDFFromDraft = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBoxCreatPreview = New System.Windows.Forms.CheckBox()
        Me.TextBoxCheckProperty = New System.Windows.Forms.TextBox()
        Me.CheckBoxAddHardWare = New System.Windows.Forms.CheckBox()
        Me.CheckBoxAlreadyImported = New System.Windows.Forms.CheckBox()
        Me.CheckBoxTurnOffGradientBackground = New System.Windows.Forms.CheckBox()
        Me.TextBoxFindReplaceLinkPropertyText = New System.Windows.Forms.TextBox()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ListBoxFindReplaceLinkedPropertyText = New System.Windows.Forms.ListBox()
        Me.CheckBoxReplaceCharacterInFilename = New System.Windows.Forms.CheckBox()
        Me.CheckBoxCalloutBOMFindandReplace = New System.Windows.Forms.CheckBox()
        Me.TextBoxNewCharacter = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBoxOldCharacter = New System.Windows.Forms.TextBox()
        Me.CheckBoxTurnOffCSs = New System.Windows.Forms.CheckBox()
        Me.TextBoxPropsToRemove = New System.Windows.Forms.TextBox()
        Me.CheckBoxRemoveProperties = New System.Windows.Forms.CheckBox()
        Me.CBCopyStyles = New System.Windows.Forms.CheckBox()
        Me.CBRsetFileStatusToAvailable = New System.Windows.Forms.CheckBox()
        Me.CBUpdateDrafts = New System.Windows.Forms.CheckBox()
        Me.CBFitAndShade = New System.Windows.Forms.CheckBox()
        Me.CBCheckAssemblyForCorruptLink = New System.Windows.Forms.CheckBox()
        Me.CheckBox1TimeFixCFGs = New System.Windows.Forms.CheckBox()
        Me.CBExtractPreview = New System.Windows.Forms.CheckBox()
        Me.CBUpDateAssemblyLinks = New System.Windows.Forms.CheckBox()
        Me.CBTurnOffDisplayNextHighestAssyOccProp = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtFilesFound = New System.Windows.Forms.TextBox()
        Me.lblFilesFound = New System.Windows.Forms.Label()
        Me.lblPaths = New System.Windows.Forms.Label()
        Me.lstPath = New System.Windows.Forms.ListBox()
        Me.lblFileList = New System.Windows.Forms.Label()
        Me.RBTraversFolders = New System.Windows.Forms.RadioButton()
        Me.RBListFormFile = New System.Windows.Forms.RadioButton()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.lstFiles = New System.Windows.Forms.ListBox()
        Me.GrpFolderOptions = New System.Windows.Forms.GroupBox()
        Me.optAllFiles = New System.Windows.Forms.RadioButton()
        Me.optAllInDirectory = New System.Windows.Forms.RadioButton()
        Me.GrpDocumentTypes = New System.Windows.Forms.GroupBox()
        Me.chkWeldment = New System.Windows.Forms.CheckBox()
        Me.chkDraft = New System.Windows.Forms.CheckBox()
        Me.chkAssembly = New System.Windows.Forms.CheckBox()
        Me.chkSheetmetal = New System.Windows.Forms.CheckBox()
        Me.chkPart = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CheckBoxResetBodyStyle = New System.Windows.Forms.CheckBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GrpFolderOptions.SuspendLayout()
        Me.GrpDocumentTypes.SuspendLayout()
        Me.SuspendLayout()
        '
        'CBXRecompute
        '
        Me.CBXRecompute.AutoSize = True
        Me.CBXRecompute.Location = New System.Drawing.Point(7, 191)
        Me.CBXRecompute.Margin = New System.Windows.Forms.Padding(4)
        Me.CBXRecompute.Name = "CBXRecompute"
        Me.CBXRecompute.Size = New System.Drawing.Size(264, 21)
        Me.CBXRecompute.TabIndex = 58
        Me.CBXRecompute.Text = "Recompute Part and Sheetmetal files"
        Me.CBXRecompute.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 661)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 17)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Label2"
        '
        'Button4
        '
        Me.Button4.Enabled = False
        Me.Button4.Location = New System.Drawing.Point(13, 690)
        Me.Button4.Margin = New System.Windows.Forms.Padding(4)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(80, 28)
        Me.Button4.TabIndex = 55
        Me.Button4.Text = "View Log"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TxtLogFileName
        '
        Me.TxtLogFileName.Location = New System.Drawing.Point(101, 694)
        Me.TxtLogFileName.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtLogFileName.Name = "TxtLogFileName"
        Me.TxtLogFileName.Size = New System.Drawing.Size(680, 22)
        Me.TxtLogFileName.TabIndex = 54
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(878, 688)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(67, 28)
        Me.Button3.TabIndex = 53
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnProcess
        '
        Me.btnProcess.Enabled = False
        Me.btnProcess.Location = New System.Drawing.Point(814, 688)
        Me.btnProcess.Margin = New System.Windows.Forms.Padding(4)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(56, 28)
        Me.btnProcess.TabIndex = 52
        Me.btnProcess.Text = "OK"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'TxtTextFilename
        '
        Me.TxtTextFilename.Enabled = False
        Me.TxtTextFilename.Location = New System.Drawing.Point(111, 44)
        Me.TxtTextFilename.Margin = New System.Windows.Forms.Padding(4)
        Me.TxtTextFilename.Name = "TxtTextFilename"
        Me.TxtTextFilename.Size = New System.Drawing.Size(815, 22)
        Me.TxtTextFilename.TabIndex = 51
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(35, 44)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(68, 22)
        Me.Button1.TabIndex = 50
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'CBCreadtePDFFromDraft
        '
        Me.CBCreadtePDFFromDraft.AutoSize = True
        Me.CBCreadtePDFFromDraft.Location = New System.Drawing.Point(566, 204)
        Me.CBCreadtePDFFromDraft.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBCreadtePDFFromDraft.Name = "CBCreadtePDFFromDraft"
        Me.CBCreadtePDFFromDraft.Size = New System.Drawing.Size(235, 21)
        Me.CBCreadtePDFFromDraft.TabIndex = 59
        Me.CBCreadtePDFFromDraft.Text = "Create PDF files for ONLY drafts"
        Me.CBCreadtePDFFromDraft.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBoxResetBodyStyle)
        Me.GroupBox1.Controls.Add(Me.CheckBoxCreatPreview)
        Me.GroupBox1.Controls.Add(Me.TextBoxCheckProperty)
        Me.GroupBox1.Controls.Add(Me.CheckBoxAddHardWare)
        Me.GroupBox1.Controls.Add(Me.CheckBoxAlreadyImported)
        Me.GroupBox1.Controls.Add(Me.CheckBoxTurnOffGradientBackground)
        Me.GroupBox1.Controls.Add(Me.TextBoxFindReplaceLinkPropertyText)
        Me.GroupBox1.Controls.Add(Me.Button6)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.ListBoxFindReplaceLinkedPropertyText)
        Me.GroupBox1.Controls.Add(Me.CheckBoxReplaceCharacterInFilename)
        Me.GroupBox1.Controls.Add(Me.CheckBoxCalloutBOMFindandReplace)
        Me.GroupBox1.Controls.Add(Me.TextBoxNewCharacter)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TextBoxOldCharacter)
        Me.GroupBox1.Controls.Add(Me.CheckBoxTurnOffCSs)
        Me.GroupBox1.Controls.Add(Me.TextBoxPropsToRemove)
        Me.GroupBox1.Controls.Add(Me.CheckBoxRemoveProperties)
        Me.GroupBox1.Controls.Add(Me.CBCreadtePDFFromDraft)
        Me.GroupBox1.Controls.Add(Me.CBCopyStyles)
        Me.GroupBox1.Controls.Add(Me.CBRsetFileStatusToAvailable)
        Me.GroupBox1.Controls.Add(Me.CBUpdateDrafts)
        Me.GroupBox1.Controls.Add(Me.CBFitAndShade)
        Me.GroupBox1.Controls.Add(Me.CBCheckAssemblyForCorruptLink)
        Me.GroupBox1.Controls.Add(Me.CBXRecompute)
        Me.GroupBox1.Controls.Add(Me.CheckBox1TimeFixCFGs)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 376)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox1.Size = New System.Drawing.Size(932, 257)
        Me.GroupBox1.TabIndex = 60
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Available Options during processing"
        '
        'CheckBoxCreatPreview
        '
        Me.CheckBoxCreatPreview.AutoSize = True
        Me.CheckBoxCreatPreview.Location = New System.Drawing.Point(303, 219)
        Me.CheckBoxCreatPreview.Name = "CheckBoxCreatPreview"
        Me.CheckBoxCreatPreview.Size = New System.Drawing.Size(125, 21)
        Me.CheckBoxCreatPreview.TabIndex = 85
        Me.CheckBoxCreatPreview.Text = "Create Preview"
        Me.CheckBoxCreatPreview.UseVisualStyleBackColor = True
        '
        'TextBoxCheckProperty
        '
        Me.TextBoxCheckProperty.Enabled = False
        Me.TextBoxCheckProperty.Location = New System.Drawing.Point(323, 191)
        Me.TextBoxCheckProperty.Name = "TextBoxCheckProperty"
        Me.TextBoxCheckProperty.Size = New System.Drawing.Size(216, 22)
        Me.TextBoxCheckProperty.TabIndex = 84
        Me.TextBoxCheckProperty.Text = "Category=Hardware"
        '
        'CheckBoxAddHardWare
        '
        Me.CheckBoxAddHardWare.AutoSize = True
        Me.CheckBoxAddHardWare.Location = New System.Drawing.Point(303, 173)
        Me.CheckBoxAddHardWare.Name = "CheckBoxAddHardWare"
        Me.CheckBoxAddHardWare.Size = New System.Drawing.Size(255, 21)
        Me.CheckBoxAddHardWare.TabIndex = 83
        Me.CheckBoxAddHardWare.Text = "Check the ""Hardware part"" option if:"
        Me.CheckBoxAddHardWare.UseVisualStyleBackColor = True
        '
        'CheckBoxAlreadyImported
        '
        Me.CheckBoxAlreadyImported.Location = New System.Drawing.Point(303, 120)
        Me.CheckBoxAlreadyImported.Name = "CheckBoxAlreadyImported"
        Me.CheckBoxAlreadyImported.Size = New System.Drawing.Size(243, 47)
        Me.CheckBoxAlreadyImported.TabIndex = 82
        Me.CheckBoxAlreadyImported.Text = "Check if file has already been imported to TC or SP"
        Me.CheckBoxAlreadyImported.UseVisualStyleBackColor = True
        '
        'CheckBoxTurnOffGradientBackground
        '
        Me.CheckBoxTurnOffGradientBackground.AutoSize = True
        Me.CheckBoxTurnOffGradientBackground.Location = New System.Drawing.Point(312, 68)
        Me.CheckBoxTurnOffGradientBackground.Name = "CheckBoxTurnOffGradientBackground"
        Me.CheckBoxTurnOffGradientBackground.Size = New System.Drawing.Size(215, 21)
        Me.CheckBoxTurnOffGradientBackground.TabIndex = 81
        Me.CheckBoxTurnOffGradientBackground.Text = "Turn off gradient background"
        Me.CheckBoxTurnOffGradientBackground.UseVisualStyleBackColor = True
        '
        'TextBoxFindReplaceLinkPropertyText
        '
        Me.TextBoxFindReplaceLinkPropertyText.Enabled = False
        Me.TextBoxFindReplaceLinkPropertyText.Location = New System.Drawing.Point(566, 125)
        Me.TextBoxFindReplaceLinkPropertyText.Name = "TextBoxFindReplaceLinkPropertyText"
        Me.TextBoxFindReplaceLinkPropertyText.Size = New System.Drawing.Size(351, 22)
        Me.TextBoxFindReplaceLinkPropertyText.TabIndex = 80
        Me.TextBoxFindReplaceLinkPropertyText.Text = "%{Title|G},%{Part.Title|G}"
        '
        'Button6
        '
        Me.Button6.Enabled = False
        Me.Button6.Location = New System.Drawing.Point(669, 151)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(145, 23)
        Me.Button6.TabIndex = 79
        Me.Button6.Text = "Remove Selected"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Enabled = False
        Me.Button2.Location = New System.Drawing.Point(566, 151)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(47, 23)
        Me.Button2.TabIndex = 78
        Me.Button2.Text = "Add"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'ListBoxFindReplaceLinkedPropertyText
        '
        Me.ListBoxFindReplaceLinkedPropertyText.Enabled = False
        Me.ListBoxFindReplaceLinkedPropertyText.FormattingEnabled = True
        Me.ListBoxFindReplaceLinkedPropertyText.ItemHeight = 16
        Me.ListBoxFindReplaceLinkedPropertyText.Location = New System.Drawing.Point(566, 47)
        Me.ListBoxFindReplaceLinkedPropertyText.Name = "ListBoxFindReplaceLinkedPropertyText"
        Me.ListBoxFindReplaceLinkedPropertyText.ScrollAlwaysVisible = True
        Me.ListBoxFindReplaceLinkedPropertyText.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBoxFindReplaceLinkedPropertyText.Size = New System.Drawing.Size(351, 68)
        Me.ListBoxFindReplaceLinkedPropertyText.TabIndex = 77
        '
        'CheckBoxReplaceCharacterInFilename
        '
        Me.CheckBoxReplaceCharacterInFilename.AutoSize = True
        Me.CheckBoxReplaceCharacterInFilename.Location = New System.Drawing.Point(7, 135)
        Me.CheckBoxReplaceCharacterInFilename.Name = "CheckBoxReplaceCharacterInFilename"
        Me.CheckBoxReplaceCharacterInFilename.Size = New System.Drawing.Size(224, 21)
        Me.CheckBoxReplaceCharacterInFilename.TabIndex = 71
        Me.CheckBoxReplaceCharacterInFilename.Text = "Replace Character in file name"
        Me.CheckBoxReplaceCharacterInFilename.UseVisualStyleBackColor = True
        '
        'CheckBoxCalloutBOMFindandReplace
        '
        Me.CheckBoxCalloutBOMFindandReplace.AutoSize = True
        Me.CheckBoxCalloutBOMFindandReplace.Location = New System.Drawing.Point(566, 20)
        Me.CheckBoxCalloutBOMFindandReplace.Name = "CheckBoxCalloutBOMFindandReplace"
        Me.CheckBoxCalloutBOMFindandReplace.Size = New System.Drawing.Size(263, 21)
        Me.CheckBoxCalloutBOMFindandReplace.TabIndex = 76
        Me.CheckBoxCalloutBOMFindandReplace.Text = "Callout/BOM column find and replace"
        Me.CheckBoxCalloutBOMFindandReplace.UseVisualStyleBackColor = True
        '
        'TextBoxNewCharacter
        '
        Me.TextBoxNewCharacter.Enabled = False
        Me.TextBoxNewCharacter.Location = New System.Drawing.Point(159, 162)
        Me.TextBoxNewCharacter.Name = "TextBoxNewCharacter"
        Me.TextBoxNewCharacter.Size = New System.Drawing.Size(84, 22)
        Me.TextBoxNewCharacter.TabIndex = 74
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(126, 166)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 17)
        Me.Label3.TabIndex = 73
        Me.Label3.Text = "with"
        '
        'TextBoxOldCharacter
        '
        Me.TextBoxOldCharacter.Enabled = False
        Me.TextBoxOldCharacter.Location = New System.Drawing.Point(35, 162)
        Me.TextBoxOldCharacter.Name = "TextBoxOldCharacter"
        Me.TextBoxOldCharacter.Size = New System.Drawing.Size(89, 22)
        Me.TextBoxOldCharacter.TabIndex = 72
        '
        'CheckBoxTurnOffCSs
        '
        Me.CheckBoxTurnOffCSs.AutoSize = True
        Me.CheckBoxTurnOffCSs.Enabled = False
        Me.CheckBoxTurnOffCSs.Location = New System.Drawing.Point(312, 46)
        Me.CheckBoxTurnOffCSs.Name = "CheckBoxTurnOffCSs"
        Me.CheckBoxTurnOffCSs.Size = New System.Drawing.Size(227, 21)
        Me.CheckBoxTurnOffCSs.TabIndex = 70
        Me.CheckBoxTurnOffCSs.Text = "Turn Off all coordinate systems"
        Me.CheckBoxTurnOffCSs.UseVisualStyleBackColor = True
        '
        'TextBoxPropsToRemove
        '
        Me.TextBoxPropsToRemove.Enabled = False
        Me.TextBoxPropsToRemove.Location = New System.Drawing.Point(35, 50)
        Me.TextBoxPropsToRemove.Name = "TextBoxPropsToRemove"
        Me.TextBoxPropsToRemove.Size = New System.Drawing.Size(217, 22)
        Me.TextBoxPropsToRemove.TabIndex = 69
        '
        'CheckBoxRemoveProperties
        '
        Me.CheckBoxRemoveProperties.AutoSize = True
        Me.CheckBoxRemoveProperties.Location = New System.Drawing.Point(7, 23)
        Me.CheckBoxRemoveProperties.Name = "CheckBoxRemoveProperties"
        Me.CheckBoxRemoveProperties.Size = New System.Drawing.Size(291, 21)
        Me.CheckBoxRemoveProperties.TabIndex = 68
        Me.CheckBoxRemoveProperties.Text = "Remove Properties (comma delimited list)"
        Me.CheckBoxRemoveProperties.UseVisualStyleBackColor = True
        '
        'CBCopyStyles
        '
        Me.CBCopyStyles.AutoSize = True
        Me.CBCopyStyles.Location = New System.Drawing.Point(303, 94)
        Me.CBCopyStyles.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBCopyStyles.Name = "CBCopyStyles"
        Me.CBCopyStyles.Size = New System.Drawing.Size(243, 21)
        Me.CBCopyStyles.TabIndex = 64
        Me.CBCopyStyles.Text = "Copy styles from OOTB templates"
        Me.CBCopyStyles.UseVisualStyleBackColor = True
        '
        'CBRsetFileStatusToAvailable
        '
        Me.CBRsetFileStatusToAvailable.AutoSize = True
        Me.CBRsetFileStatusToAvailable.Location = New System.Drawing.Point(7, 81)
        Me.CBRsetFileStatusToAvailable.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBRsetFileStatusToAvailable.Name = "CBRsetFileStatusToAvailable"
        Me.CBRsetFileStatusToAvailable.Size = New System.Drawing.Size(252, 21)
        Me.CBRsetFileStatusToAvailable.TabIndex = 63
        Me.CBRsetFileStatusToAvailable.Text = "Reset file status back to ""Available"""
        Me.CBRsetFileStatusToAvailable.UseVisualStyleBackColor = True
        '
        'CBUpdateDrafts
        '
        Me.CBUpdateDrafts.AutoSize = True
        Me.CBUpdateDrafts.Location = New System.Drawing.Point(566, 179)
        Me.CBUpdateDrafts.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBUpdateDrafts.Name = "CBUpdateDrafts"
        Me.CBUpdateDrafts.Size = New System.Drawing.Size(222, 21)
        Me.CBUpdateDrafts.TabIndex = 62
        Me.CBUpdateDrafts.Text = "Update SE draft drawing views"
        Me.CBUpdateDrafts.UseVisualStyleBackColor = True
        '
        'CBFitAndShade
        '
        Me.CBFitAndShade.AutoSize = True
        Me.CBFitAndShade.Location = New System.Drawing.Point(303, 23)
        Me.CBFitAndShade.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBFitAndShade.Name = "CBFitAndShade"
        Me.CBFitAndShade.Size = New System.Drawing.Size(214, 21)
        Me.CBFitAndShade.TabIndex = 61
        Me.CBFitAndShade.Text = "Fit View and shade (if 3D file)"
        Me.CBFitAndShade.UseVisualStyleBackColor = True
        '
        'CBCheckAssemblyForCorruptLink
        '
        Me.CBCheckAssemblyForCorruptLink.AutoSize = True
        Me.CBCheckAssemblyForCorruptLink.Location = New System.Drawing.Point(7, 107)
        Me.CBCheckAssemblyForCorruptLink.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBCheckAssemblyForCorruptLink.Name = "CBCheckAssemblyForCorruptLink"
        Me.CBCheckAssemblyForCorruptLink.Size = New System.Drawing.Size(275, 21)
        Me.CBCheckAssemblyForCorruptLink.TabIndex = 60
        Me.CBCheckAssemblyForCorruptLink.Text = "Check for invalid link in Solid Edge files"
        Me.CBCheckAssemblyForCorruptLink.UseVisualStyleBackColor = True
        '
        'CheckBox1TimeFixCFGs
        '
        Me.CheckBox1TimeFixCFGs.AutoSize = True
        Me.CheckBox1TimeFixCFGs.Location = New System.Drawing.Point(63, 133)
        Me.CheckBox1TimeFixCFGs.Name = "CheckBox1TimeFixCFGs"
        Me.CheckBox1TimeFixCFGs.Size = New System.Drawing.Size(167, 21)
        Me.CheckBox1TimeFixCFGs.TabIndex = 75
        Me.CheckBox1TimeFixCFGs.Text = "Temporary fix for cfgs"
        Me.CheckBox1TimeFixCFGs.UseVisualStyleBackColor = True
        Me.CheckBox1TimeFixCFGs.Visible = False
        '
        'CBExtractPreview
        '
        Me.CBExtractPreview.AutoSize = True
        Me.CBExtractPreview.Location = New System.Drawing.Point(952, 488)
        Me.CBExtractPreview.Margin = New System.Windows.Forms.Padding(4)
        Me.CBExtractPreview.Name = "CBExtractPreview"
        Me.CBExtractPreview.Size = New System.Drawing.Size(159, 21)
        Me.CBExtractPreview.TabIndex = 67
        Me.CBExtractPreview.Text = "Extract Preview BMP"
        Me.CBExtractPreview.UseVisualStyleBackColor = True
        '
        'CBUpDateAssemblyLinks
        '
        Me.CBUpDateAssemblyLinks.AutoSize = True
        Me.CBUpDateAssemblyLinks.Location = New System.Drawing.Point(952, 542)
        Me.CBUpDateAssemblyLinks.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBUpDateAssemblyLinks.Name = "CBUpDateAssemblyLinks"
        Me.CBUpDateAssemblyLinks.Size = New System.Drawing.Size(225, 21)
        Me.CBUpDateAssemblyLinks.TabIndex = 66
        Me.CBUpDateAssemblyLinks.Text = "Update external assembly links"
        Me.CBUpDateAssemblyLinks.UseVisualStyleBackColor = True
        '
        'CBTurnOffDisplayNextHighestAssyOccProp
        '
        Me.CBTurnOffDisplayNextHighestAssyOccProp.AutoSize = True
        Me.CBTurnOffDisplayNextHighestAssyOccProp.Location = New System.Drawing.Point(951, 456)
        Me.CBTurnOffDisplayNextHighestAssyOccProp.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.CBTurnOffDisplayNextHighestAssyOccProp.Name = "CBTurnOffDisplayNextHighestAssyOccProp"
        Me.CBTurnOffDisplayNextHighestAssyOccProp.Size = New System.Drawing.Size(291, 21)
        Me.CBTurnOffDisplayNextHighestAssyOccProp.TabIndex = 65
        Me.CBTurnOffDisplayNextHighestAssyOccProp.Text = "Turn off hidden parts in next highest assy"
        Me.CBTurnOffDisplayNextHighestAssyOccProp.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtFilesFound)
        Me.GroupBox2.Controls.Add(Me.lblFilesFound)
        Me.GroupBox2.Controls.Add(Me.lblPaths)
        Me.GroupBox2.Controls.Add(Me.lstPath)
        Me.GroupBox2.Controls.Add(Me.lblFileList)
        Me.GroupBox2.Controls.Add(Me.RBTraversFolders)
        Me.GroupBox2.Controls.Add(Me.RBListFormFile)
        Me.GroupBox2.Controls.Add(Me.Button5)
        Me.GroupBox2.Controls.Add(Me.lstFiles)
        Me.GroupBox2.Controls.Add(Me.GrpFolderOptions)
        Me.GroupBox2.Controls.Add(Me.GrpDocumentTypes)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.TxtTextFilename)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.GroupBox2.Size = New System.Drawing.Size(933, 359)
        Me.GroupBox2.TabIndex = 61
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "File Selection Options"
        '
        'txtFilesFound
        '
        Me.txtFilesFound.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtFilesFound.Location = New System.Drawing.Point(648, 104)
        Me.txtFilesFound.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFilesFound.Name = "txtFilesFound"
        Me.txtFilesFound.ReadOnly = True
        Me.txtFilesFound.Size = New System.Drawing.Size(97, 22)
        Me.txtFilesFound.TabIndex = 62
        '
        'lblFilesFound
        '
        Me.lblFilesFound.AutoSize = True
        Me.lblFilesFound.Location = New System.Drawing.Point(493, 107)
        Me.lblFilesFound.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFilesFound.Name = "lblFilesFound"
        Me.lblFilesFound.Size = New System.Drawing.Size(147, 17)
        Me.lblFilesFound.TabIndex = 61
        Me.lblFilesFound.Text = "Number of files found:"
        '
        'lblPaths
        '
        Me.lblPaths.AutoSize = True
        Me.lblPaths.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPaths.Location = New System.Drawing.Point(279, 109)
        Me.lblPaths.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPaths.Name = "lblPaths"
        Me.lblPaths.Size = New System.Drawing.Size(92, 17)
        Me.lblPaths.TabIndex = 60
        Me.lblPaths.Text = "File Path(s)"
        '
        'lstPath
        '
        Me.lstPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstPath.FormattingEnabled = True
        Me.lstPath.HorizontalScrollbar = True
        Me.lstPath.ItemHeight = 16
        Me.lstPath.Location = New System.Drawing.Point(283, 128)
        Me.lstPath.Margin = New System.Windows.Forms.Padding(4)
        Me.lstPath.Name = "lstPath"
        Me.lstPath.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstPath.Size = New System.Drawing.Size(636, 68)
        Me.lstPath.Sorted = True
        Me.lstPath.TabIndex = 59
        '
        'lblFileList
        '
        Me.lblFileList.AutoSize = True
        Me.lblFileList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFileList.Location = New System.Drawing.Point(279, 200)
        Me.lblFileList.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFileList.Name = "lblFileList"
        Me.lblFileList.Size = New System.Drawing.Size(164, 17)
        Me.lblFileList.TabIndex = 58
        Me.lblFileList.Text = "Files to be processed"
        '
        'RBTraversFolders
        '
        Me.RBTraversFolders.AutoSize = True
        Me.RBTraversFolders.Location = New System.Drawing.Point(8, 80)
        Me.RBTraversFolders.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RBTraversFolders.Name = "RBTraversFolders"
        Me.RBTraversFolders.Size = New System.Drawing.Size(416, 21)
        Me.RBTraversFolders.TabIndex = 56
        Me.RBTraversFolders.Text = "Traverse specified folder(s) to generate list of files to process"
        Me.RBTraversFolders.UseVisualStyleBackColor = True
        '
        'RBListFormFile
        '
        Me.RBListFormFile.AutoSize = True
        Me.RBListFormFile.Checked = True
        Me.RBListFormFile.Location = New System.Drawing.Point(8, 21)
        Me.RBListFormFile.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.RBListFormFile.Name = "RBListFormFile"
        Me.RBListFormFile.Size = New System.Drawing.Size(405, 21)
        Me.RBListFormFile.TabIndex = 55
        Me.RBListFormFile.TabStop = True
        Me.RBListFormFile.Text = "Select a text or CSV file containing the list of files to process"
        Me.RBListFormFile.UseVisualStyleBackColor = True
        '
        'Button5
        '
        Me.Button5.Enabled = False
        Me.Button5.Location = New System.Drawing.Point(19, 104)
        Me.Button5.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(240, 23)
        Me.Button5.TabIndex = 54
        Me.Button5.Text = "Browse to locate top level folder"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'lstFiles
        '
        Me.lstFiles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstFiles.ColumnWidth = 372
        Me.lstFiles.FormattingEnabled = True
        Me.lstFiles.HorizontalScrollbar = True
        Me.lstFiles.ItemHeight = 16
        Me.lstFiles.Location = New System.Drawing.Point(283, 221)
        Me.lstFiles.Margin = New System.Windows.Forms.Padding(4)
        Me.lstFiles.Name = "lstFiles"
        Me.lstFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstFiles.Size = New System.Drawing.Size(635, 132)
        Me.lstFiles.TabIndex = 53
        '
        'GrpFolderOptions
        '
        Me.GrpFolderOptions.Controls.Add(Me.optAllFiles)
        Me.GrpFolderOptions.Controls.Add(Me.optAllInDirectory)
        Me.GrpFolderOptions.Location = New System.Drawing.Point(12, 264)
        Me.GrpFolderOptions.Margin = New System.Windows.Forms.Padding(4)
        Me.GrpFolderOptions.Name = "GrpFolderOptions"
        Me.GrpFolderOptions.Padding = New System.Windows.Forms.Padding(4)
        Me.GrpFolderOptions.Size = New System.Drawing.Size(263, 91)
        Me.GrpFolderOptions.TabIndex = 4
        Me.GrpFolderOptions.TabStop = False
        Me.GrpFolderOptions.Text = "Documents to Process"
        '
        'optAllFiles
        '
        Me.optAllFiles.Location = New System.Drawing.Point(8, 41)
        Me.optAllFiles.Margin = New System.Windows.Forms.Padding(4)
        Me.optAllFiles.Name = "optAllFiles"
        Me.optAllFiles.Size = New System.Drawing.Size(239, 39)
        Me.optAllFiles.TabIndex = 2
        Me.optAllFiles.Text = "All files in selected directory and subdirectories"
        Me.optAllFiles.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.optAllFiles.UseVisualStyleBackColor = True
        '
        'optAllInDirectory
        '
        Me.optAllInDirectory.AutoSize = True
        Me.optAllInDirectory.Checked = True
        Me.optAllInDirectory.Location = New System.Drawing.Point(8, 23)
        Me.optAllInDirectory.Margin = New System.Windows.Forms.Padding(4)
        Me.optAllInDirectory.Name = "optAllInDirectory"
        Me.optAllInDirectory.Size = New System.Drawing.Size(204, 21)
        Me.optAllInDirectory.TabIndex = 0
        Me.optAllInDirectory.TabStop = True
        Me.optAllInDirectory.Text = "All files in selected directory"
        Me.optAllInDirectory.UseVisualStyleBackColor = True
        '
        'GrpDocumentTypes
        '
        Me.GrpDocumentTypes.Controls.Add(Me.chkWeldment)
        Me.GrpDocumentTypes.Controls.Add(Me.chkDraft)
        Me.GrpDocumentTypes.Controls.Add(Me.chkAssembly)
        Me.GrpDocumentTypes.Controls.Add(Me.chkSheetmetal)
        Me.GrpDocumentTypes.Controls.Add(Me.chkPart)
        Me.GrpDocumentTypes.Location = New System.Drawing.Point(12, 134)
        Me.GrpDocumentTypes.Margin = New System.Windows.Forms.Padding(4)
        Me.GrpDocumentTypes.Name = "GrpDocumentTypes"
        Me.GrpDocumentTypes.Padding = New System.Windows.Forms.Padding(4)
        Me.GrpDocumentTypes.Size = New System.Drawing.Size(263, 123)
        Me.GrpDocumentTypes.TabIndex = 52
        Me.GrpDocumentTypes.TabStop = False
        Me.GrpDocumentTypes.Text = "Document Types to Open"
        '
        'chkWeldment
        '
        Me.chkWeldment.AutoSize = True
        Me.chkWeldment.Location = New System.Drawing.Point(8, 95)
        Me.chkWeldment.Margin = New System.Windows.Forms.Padding(4)
        Me.chkWeldment.Name = "chkWeldment"
        Me.chkWeldment.Size = New System.Drawing.Size(168, 21)
        Me.chkWeldment.TabIndex = 7
        Me.chkWeldment.Text = "Weldment Documents"
        Me.chkWeldment.UseVisualStyleBackColor = True
        '
        'chkDraft
        '
        Me.chkDraft.AutoSize = True
        Me.chkDraft.Location = New System.Drawing.Point(9, 77)
        Me.chkDraft.Margin = New System.Windows.Forms.Padding(4)
        Me.chkDraft.Name = "chkDraft"
        Me.chkDraft.Size = New System.Drawing.Size(134, 21)
        Me.chkDraft.TabIndex = 5
        Me.chkDraft.Text = "Draft documents"
        Me.chkDraft.UseVisualStyleBackColor = True
        '
        'chkAssembly
        '
        Me.chkAssembly.AutoSize = True
        Me.chkAssembly.Location = New System.Drawing.Point(8, 59)
        Me.chkAssembly.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAssembly.Name = "chkAssembly"
        Me.chkAssembly.Size = New System.Drawing.Size(163, 21)
        Me.chkAssembly.TabIndex = 2
        Me.chkAssembly.Text = "Assembly documents"
        Me.chkAssembly.UseVisualStyleBackColor = True
        '
        'chkSheetmetal
        '
        Me.chkSheetmetal.AutoSize = True
        Me.chkSheetmetal.Location = New System.Drawing.Point(8, 41)
        Me.chkSheetmetal.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSheetmetal.Name = "chkSheetmetal"
        Me.chkSheetmetal.Size = New System.Drawing.Size(178, 21)
        Me.chkSheetmetal.TabIndex = 1
        Me.chkSheetmetal.Text = "Sheet Metal documents"
        Me.chkSheetmetal.UseVisualStyleBackColor = True
        '
        'chkPart
        '
        Me.chkPart.AutoSize = True
        Me.chkPart.Location = New System.Drawing.Point(8, 23)
        Me.chkPart.Margin = New System.Windows.Forms.Padding(4)
        Me.chkPart.Name = "chkPart"
        Me.chkPart.Size = New System.Drawing.Size(129, 21)
        Me.chkPart.TabIndex = 0
        Me.chkPart.Text = "Part documents"
        Me.chkPart.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 635)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 17)
        Me.Label1.TabIndex = 62
        Me.Label1.Text = "Label1"
        '
        'CheckBoxResetBodyStyle
        '
        Me.CheckBoxResetBodyStyle.AutoSize = True
        Me.CheckBoxResetBodyStyle.Location = New System.Drawing.Point(7, 219)
        Me.CheckBoxResetBodyStyle.Name = "CheckBoxResetBodyStyle"
        Me.CheckBoxResetBodyStyle.Size = New System.Drawing.Size(202, 21)
        Me.CheckBoxResetBodyStyle.TabIndex = 68
        Me.CheckBoxResetBodyStyle.Text = "Reset Body Style to ""None"""
        Me.CheckBoxResetBodyStyle.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(944, 734)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TxtLogFileName)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.CBUpDateAssemblyLinks)
        Me.Controls.Add(Me.CBTurnOffDisplayNextHighestAssyOccProp)
        Me.Controls.Add(Me.CBExtractPreview)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.Text = "Open and Save Utility"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GrpFolderOptions.ResumeLayout(False)
        Me.GrpFolderOptions.PerformLayout()
        Me.GrpDocumentTypes.ResumeLayout(False)
        Me.GrpDocumentTypes.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CBXRecompute As System.Windows.Forms.CheckBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TxtLogFileName As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents TxtTextFilename As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents CBCreadtePDFFromDraft As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CBCheckAssemblyForCorruptLink As System.Windows.Forms.CheckBox
    Friend WithEvents CBFitAndShade As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lstFiles As System.Windows.Forms.ListBox
    Friend WithEvents GrpFolderOptions As System.Windows.Forms.GroupBox
    Friend WithEvents optAllFiles As System.Windows.Forms.RadioButton
    Friend WithEvents optAllInDirectory As System.Windows.Forms.RadioButton
    Friend WithEvents GrpDocumentTypes As System.Windows.Forms.GroupBox
    Friend WithEvents chkWeldment As System.Windows.Forms.CheckBox
    Friend WithEvents chkDraft As System.Windows.Forms.CheckBox
    Friend WithEvents chkAssembly As System.Windows.Forms.CheckBox
    Friend WithEvents chkSheetmetal As System.Windows.Forms.CheckBox
    Friend WithEvents chkPart As System.Windows.Forms.CheckBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents RBTraversFolders As System.Windows.Forms.RadioButton
    Friend WithEvents RBListFormFile As System.Windows.Forms.RadioButton
    Friend WithEvents txtFilesFound As System.Windows.Forms.TextBox
    Friend WithEvents lblFilesFound As System.Windows.Forms.Label
    Friend WithEvents lblPaths As System.Windows.Forms.Label
    Friend WithEvents lstPath As System.Windows.Forms.ListBox
    Friend WithEvents lblFileList As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents CBUpdateDrafts As System.Windows.Forms.CheckBox
    Friend WithEvents CBRsetFileStatusToAvailable As System.Windows.Forms.CheckBox
    Friend WithEvents CBCopyStyles As System.Windows.Forms.CheckBox
    Friend WithEvents CBTurnOffDisplayNextHighestAssyOccProp As System.Windows.Forms.CheckBox
    Friend WithEvents CBUpDateAssemblyLinks As System.Windows.Forms.CheckBox
    Friend WithEvents CBExtractPreview As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxPropsToRemove As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxRemoveProperties As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxTurnOffCSs As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxNewCharacter As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBoxOldCharacter As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxReplaceCharacterInFilename As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1TimeFixCFGs As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCalloutBOMFindandReplace As System.Windows.Forms.CheckBox
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents ListBoxFindReplaceLinkedPropertyText As System.Windows.Forms.ListBox
    Friend WithEvents TextBoxFindReplaceLinkPropertyText As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxTurnOffGradientBackground As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxAlreadyImported As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxCheckProperty As System.Windows.Forms.TextBox
    Friend WithEvents CheckBoxAddHardWare As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxCreatPreview As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxResetBodyStyle As System.Windows.Forms.CheckBox

End Class
