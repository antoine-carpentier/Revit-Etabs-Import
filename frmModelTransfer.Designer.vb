'
' Created by SharpDevelop.
' User: michael
' Date: 3/27/2015
' Time: 2:58 PM
' 
' To change this template use Tools | Options | Coding | Edit Standard Headers.
'
Partial Class frmRenameSections
	Inherits System.Windows.Forms.Form
	
	''' <summary>
	''' Designer variable used to keep track of non-visual components.
	''' </summary>
	Private components As System.ComponentModel.IContainer
	
	''' <summary>
	''' Disposes resources used by the form.
	''' </summary>
	''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
	Protected Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If components IsNot Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub
	
	''' <summary>
	''' This method is required for Windows Forms designer support.
	''' Do not change the method contents inside the source code editor. The Forms designer might
	''' not be able to load this method if it was changed manually.
	''' </summary>
	Private Sub InitializeComponent()
		Me.btnOK = New System.Windows.Forms.Button()
		Me.btnCancel = New System.Windows.Forms.Button()
		Me.CSVLocation = New System.Windows.Forms.TextBox()
		Me.CSVBrowse = New System.Windows.Forms.Button()
		Me.CSVOpenFile = New System.Windows.Forms.OpenFileDialog()
		Me.label2 = New System.Windows.Forms.Label()
		Me.GridcheckBox = New System.Windows.Forms.CheckBox()
		Me.SlabcheckBox = New System.Windows.Forms.CheckBox()
		Me.WallcheckBox = New System.Windows.Forms.CheckBox()
		Me.BeamcheckBox = New System.Windows.Forms.CheckBox()
		Me.ColcheckBox = New System.Windows.Forms.CheckBox()
		Me.groupBox1 = New System.Windows.Forms.GroupBox()
		Me.LevelcheckBox = New System.Windows.Forms.CheckBox()
		Me.groupBox1.SuspendLayout
		Me.SuspendLayout
		'
		'btnOK
		'
		Me.btnOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
		Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnOK.Location = New System.Drawing.Point(318, 291)
		Me.btnOK.Name = "btnOK"
		Me.btnOK.Size = New System.Drawing.Size(75, 23)
		Me.btnOK.TabIndex = 3
		Me.btnOK.Text = "OK"
		Me.btnOK.UseCompatibleTextRendering = true
		Me.btnOK.UseVisualStyleBackColor = true
		'
		'btnCancel
		'
		Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right),System.Windows.Forms.AnchorStyles)
		Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnCancel.Location = New System.Drawing.Point(228, 291)
		Me.btnCancel.Name = "btnCancel"
		Me.btnCancel.Size = New System.Drawing.Size(75, 23)
		Me.btnCancel.TabIndex = 4
		Me.btnCancel.Text = "Cancel"
		Me.btnCancel.UseCompatibleTextRendering = true
		Me.btnCancel.UseVisualStyleBackColor = true
		'
		'CSVLocation
		'
		Me.CSVLocation.Location = New System.Drawing.Point(20, 67)
		Me.CSVLocation.Name = "CSVLocation"
		Me.CSVLocation.Size = New System.Drawing.Size(280, 20)
		Me.CSVLocation.TabIndex = 6
		'
		'CSVBrowse
		'
		Me.CSVBrowse.Location = New System.Drawing.Point(322, 65)
		Me.CSVBrowse.Name = "CSVBrowse"
		Me.CSVBrowse.Size = New System.Drawing.Size(75, 23)
		Me.CSVBrowse.TabIndex = 7
		Me.CSVBrowse.Text = "Browse"
		Me.CSVBrowse.UseCompatibleTextRendering = true
		Me.CSVBrowse.UseVisualStyleBackColor = true
		'
		'label2
		'
		Me.label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
		Me.label2.Location = New System.Drawing.Point(22, 32)
		Me.label2.Name = "label2"
		Me.label2.Size = New System.Drawing.Size(260, 23)
		Me.label2.TabIndex = 5
		Me.label2.Text = "Select a CSV File to Import:"
		Me.label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
		Me.label2.UseCompatibleTextRendering = true
		'
		'GridcheckBox
		'
		Me.GridcheckBox.Checked = true
		Me.GridcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.GridcheckBox.Location = New System.Drawing.Point(22, 33)
		Me.GridcheckBox.Name = "GridcheckBox"
		Me.GridcheckBox.Size = New System.Drawing.Size(68, 24)
		Me.GridcheckBox.TabIndex = 9
		Me.GridcheckBox.Text = "Grids"
		Me.GridcheckBox.UseCompatibleTextRendering = true
		Me.GridcheckBox.UseVisualStyleBackColor = true
		'
		'SlabcheckBox
		'
		Me.SlabcheckBox.Checked = true
		Me.SlabcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.SlabcheckBox.Location = New System.Drawing.Point(22, 81)
		Me.SlabcheckBox.Name = "SlabcheckBox"
		Me.SlabcheckBox.Size = New System.Drawing.Size(68, 24)
		Me.SlabcheckBox.TabIndex = 10
		Me.SlabcheckBox.Text = "Slabs"
		Me.SlabcheckBox.UseCompatibleTextRendering = true
		Me.SlabcheckBox.UseVisualStyleBackColor = true
		'
		'WallcheckBox
		'
		Me.WallcheckBox.Checked = true
		Me.WallcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.WallcheckBox.Location = New System.Drawing.Point(253, 33)
		Me.WallcheckBox.Name = "WallcheckBox"
		Me.WallcheckBox.Size = New System.Drawing.Size(90, 24)
		Me.WallcheckBox.TabIndex = 11
		Me.WallcheckBox.Text = "Walls"
		Me.WallcheckBox.UseCompatibleTextRendering = true
		Me.WallcheckBox.UseVisualStyleBackColor = true
		'
		'BeamcheckBox
		'
		Me.BeamcheckBox.Checked = true
		Me.BeamcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.BeamcheckBox.Location = New System.Drawing.Point(253, 81)
		Me.BeamcheckBox.Name = "BeamcheckBox"
		Me.BeamcheckBox.Size = New System.Drawing.Size(68, 24)
		Me.BeamcheckBox.TabIndex = 12
		Me.BeamcheckBox.Text = "Beams"
		Me.BeamcheckBox.UseCompatibleTextRendering = true
		Me.BeamcheckBox.UseVisualStyleBackColor = true
		'
		'ColcheckBox
		'
		Me.ColcheckBox.Checked = true
		Me.ColcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.ColcheckBox.Location = New System.Drawing.Point(130, 81)
		Me.ColcheckBox.Name = "ColcheckBox"
		Me.ColcheckBox.Size = New System.Drawing.Size(77, 24)
		Me.ColcheckBox.TabIndex = 14
		Me.ColcheckBox.Text = "Columns"
		Me.ColcheckBox.UseCompatibleTextRendering = true
		Me.ColcheckBox.UseVisualStyleBackColor = true
		'
		'groupBox1
		'
		Me.groupBox1.Controls.Add(Me.LevelcheckBox)
		Me.groupBox1.Controls.Add(Me.ColcheckBox)
		Me.groupBox1.Controls.Add(Me.GridcheckBox)
		Me.groupBox1.Controls.Add(Me.BeamcheckBox)
		Me.groupBox1.Controls.Add(Me.WallcheckBox)
		Me.groupBox1.Controls.Add(Me.SlabcheckBox)
		Me.groupBox1.Location = New System.Drawing.Point(26, 132)
		Me.groupBox1.Name = "groupBox1"
		Me.groupBox1.Size = New System.Drawing.Size(365, 126)
		Me.groupBox1.TabIndex = 15
		Me.groupBox1.TabStop = false
		Me.groupBox1.Text = "Import Options"
		Me.groupBox1.UseCompatibleTextRendering = true
		'
		'LevelcheckBox
		'
		Me.LevelcheckBox.Checked = true
		Me.LevelcheckBox.CheckState = System.Windows.Forms.CheckState.Checked
		Me.LevelcheckBox.Location = New System.Drawing.Point(132, 33)
		Me.LevelcheckBox.Name = "LevelcheckBox"
		Me.LevelcheckBox.Size = New System.Drawing.Size(68, 24)
		Me.LevelcheckBox.TabIndex = 14
		Me.LevelcheckBox.Text = "Levels"
		Me.LevelcheckBox.UseCompatibleTextRendering = true
		Me.LevelcheckBox.UseVisualStyleBackColor = true
		'
		'frmRenameSections
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(415, 335)
		Me.Controls.Add(Me.groupBox1)
		Me.Controls.Add(Me.CSVBrowse)
		Me.Controls.Add(Me.CSVLocation)
		Me.Controls.Add(Me.label2)
		Me.Controls.Add(Me.btnCancel)
		Me.Controls.Add(Me.btnOK)
		Me.Name = "frmRenameSections"
		Me.Text = "Import Model"
		Me.groupBox1.ResumeLayout(false)
		Me.ResumeLayout(false)
		Me.PerformLayout
	End Sub
	Private groupBox1 As System.Windows.Forms.GroupBox
	Private ColcheckBox As System.Windows.Forms.CheckBox
	Private LevelcheckBox As System.Windows.Forms.CheckBox
	Private BeamcheckBox As System.Windows.Forms.CheckBox
	Private WallcheckBox As System.Windows.Forms.CheckBox
	Private SlabcheckBox As System.Windows.Forms.CheckBox
	Private GridcheckBox As System.Windows.Forms.CheckBox
	Private CSVOpenFile As System.Windows.Forms.OpenFileDialog
	Private withevents CSVBrowse As System.Windows.Forms.Button
	Private CSVLocation As System.Windows.Forms.TextBox
	Private label2 As System.Windows.Forms.Label
	Private btnCancel As System.Windows.Forms.Button
	Private withevents btnOK As System.Windows.Forms.Button
End Class
