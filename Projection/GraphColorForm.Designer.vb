<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLineColor
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.lblYear1 = New System.Windows.Forms.Label()
        Me.lblYear2 = New System.Windows.Forms.Label()
        Me.lblYear3 = New System.Windows.Forms.Label()
        Me.lblYear4 = New System.Windows.Forms.Label()
        Me.lblYear5 = New System.Windows.Forms.Label()
        Me.lstColorName = New System.Windows.Forms.ListBox()
        Me.lblYear5RGB = New System.Windows.Forms.Label()
        Me.lblYear4RGB = New System.Windows.Forms.Label()
        Me.lblYear3RGB = New System.Windows.Forms.Label()
        Me.lblYear2RGB = New System.Windows.Forms.Label()
        Me.lblYear1RGB = New System.Windows.Forms.Label()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblYearTitle = New System.Windows.Forms.Label()
        Me.lblRGB = New System.Windows.Forms.Label()
        Me.lblSelection = New System.Windows.Forms.Label()
        Me.btnAbout = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tpColor = New System.Windows.Forms.TabPage()
        Me.tpCheck = New System.Windows.Forms.TabPage()
        Me.tpTriangle = New System.Windows.Forms.TabPage()
        Me.lblRsvTotal = New System.Windows.Forms.Label()
        Me.lblReserves = New System.Windows.Forms.Label()
        Me.btnPrior = New System.Windows.Forms.Button()
        Me.btnDefault = New System.Windows.Forms.Button()
        Me.btnSelected = New System.Windows.Forms.Button()
        Me.dgvTriangle = New System.Windows.Forms.DataGridView()
        Me.grpProjBase = New System.Windows.Forms.GroupBox()
        Me.rdoPaid = New System.Windows.Forms.RadioButton()
        Me.rdoIncurred = New System.Windows.Forms.RadioButton()
        Me.TabControl1.SuspendLayout()
        Me.tpColor.SuspendLayout()
        Me.tpTriangle.SuspendLayout()
        CType(Me.dgvTriangle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProjBase.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblYear1
        '
        Me.lblYear1.AutoSize = True
        Me.lblYear1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear1.Location = New System.Drawing.Point(23, 57)
        Me.lblYear1.Name = "lblYear1"
        Me.lblYear1.Size = New System.Drawing.Size(43, 15)
        Me.lblYear1.TabIndex = 0
        Me.lblYear1.Text = "Label1"
        '
        'lblYear2
        '
        Me.lblYear2.AutoSize = True
        Me.lblYear2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear2.Location = New System.Drawing.Point(23, 91)
        Me.lblYear2.Name = "lblYear2"
        Me.lblYear2.Size = New System.Drawing.Size(43, 15)
        Me.lblYear2.TabIndex = 1
        Me.lblYear2.Text = "Label2"
        '
        'lblYear3
        '
        Me.lblYear3.AutoSize = True
        Me.lblYear3.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear3.Location = New System.Drawing.Point(23, 126)
        Me.lblYear3.Name = "lblYear3"
        Me.lblYear3.Size = New System.Drawing.Size(43, 15)
        Me.lblYear3.TabIndex = 2
        Me.lblYear3.Text = "Label3"
        '
        'lblYear4
        '
        Me.lblYear4.AutoSize = True
        Me.lblYear4.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear4.Location = New System.Drawing.Point(23, 160)
        Me.lblYear4.Name = "lblYear4"
        Me.lblYear4.Size = New System.Drawing.Size(43, 15)
        Me.lblYear4.TabIndex = 3
        Me.lblYear4.Text = "Label4"
        '
        'lblYear5
        '
        Me.lblYear5.AutoSize = True
        Me.lblYear5.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear5.Location = New System.Drawing.Point(23, 196)
        Me.lblYear5.Name = "lblYear5"
        Me.lblYear5.Size = New System.Drawing.Size(43, 15)
        Me.lblYear5.TabIndex = 4
        Me.lblYear5.Text = "Label5"
        '
        'lstColorName
        '
        Me.lstColorName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstColorName.FormattingEnabled = True
        Me.lstColorName.ItemHeight = 15
        Me.lstColorName.Items.AddRange(New Object() {"seq-1", "seq-2", "seq-3", "seq-4", "seq-5", "seq-6", "seq-7", "seq-8", "seq-9", "seq-10", "seq-11", "seq-12", "seq-13", "seq-14", "seq-15", "seq-16", "seq-17", "seq-18", "div-1", "div-2", "div-3", "div-4", "div-5", "div-6", "div-7", "div-8", "div-9", "qual-1", "qual-2", "qual-3", "qual-4", "qual-5", "qual-6", "qual-7", "qual-8"})
        Me.lstColorName.Location = New System.Drawing.Point(249, 57)
        Me.lstColorName.Name = "lstColorName"
        Me.lstColorName.Size = New System.Drawing.Size(148, 154)
        Me.lstColorName.TabIndex = 5
        '
        'lblYear5RGB
        '
        Me.lblYear5RGB.AutoSize = True
        Me.lblYear5RGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear5RGB.Location = New System.Drawing.Point(123, 195)
        Me.lblYear5RGB.Name = "lblYear5RGB"
        Me.lblYear5RGB.Size = New System.Drawing.Size(43, 15)
        Me.lblYear5RGB.TabIndex = 10
        Me.lblYear5RGB.Text = "Label6"
        '
        'lblYear4RGB
        '
        Me.lblYear4RGB.AutoSize = True
        Me.lblYear4RGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear4RGB.Location = New System.Drawing.Point(123, 160)
        Me.lblYear4RGB.Name = "lblYear4RGB"
        Me.lblYear4RGB.Size = New System.Drawing.Size(43, 15)
        Me.lblYear4RGB.TabIndex = 9
        Me.lblYear4RGB.Text = "Label7"
        '
        'lblYear3RGB
        '
        Me.lblYear3RGB.AutoSize = True
        Me.lblYear3RGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear3RGB.Location = New System.Drawing.Point(123, 126)
        Me.lblYear3RGB.Name = "lblYear3RGB"
        Me.lblYear3RGB.Size = New System.Drawing.Size(43, 15)
        Me.lblYear3RGB.TabIndex = 8
        Me.lblYear3RGB.Text = "Label8"
        '
        'lblYear2RGB
        '
        Me.lblYear2RGB.AutoSize = True
        Me.lblYear2RGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear2RGB.Location = New System.Drawing.Point(123, 91)
        Me.lblYear2RGB.Name = "lblYear2RGB"
        Me.lblYear2RGB.Size = New System.Drawing.Size(43, 15)
        Me.lblYear2RGB.TabIndex = 7
        Me.lblYear2RGB.Text = "Label9"
        '
        'lblYear1RGB
        '
        Me.lblYear1RGB.AutoSize = True
        Me.lblYear1RGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear1RGB.Location = New System.Drawing.Point(123, 57)
        Me.lblYear1RGB.Name = "lblYear1RGB"
        Me.lblYear1RGB.Size = New System.Drawing.Size(50, 15)
        Me.lblYear1RGB.TabIndex = 6
        Me.lblYear1RGB.Text = "Label10"
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(166, 242)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(87, 27)
        Me.btnApply.TabIndex = 11
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(310, 242)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(87, 27)
        Me.btnCancel.TabIndex = 12
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblYearTitle
        '
        Me.lblYearTitle.AutoSize = True
        Me.lblYearTitle.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYearTitle.Location = New System.Drawing.Point(23, 27)
        Me.lblYearTitle.Name = "lblYearTitle"
        Me.lblYearTitle.Size = New System.Drawing.Size(31, 15)
        Me.lblYearTitle.TabIndex = 13
        Me.lblYearTitle.Text = "Year"
        '
        'lblRGB
        '
        Me.lblRGB.AutoSize = True
        Me.lblRGB.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRGB.Location = New System.Drawing.Point(123, 27)
        Me.lblRGB.Name = "lblRGB"
        Me.lblRGB.Size = New System.Drawing.Size(74, 15)
        Me.lblRGB.TabIndex = 14
        Me.lblRGB.Text = "RGB (R, G, B)"
        '
        'lblSelection
        '
        Me.lblSelection.AutoSize = True
        Me.lblSelection.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelection.Location = New System.Drawing.Point(246, 27)
        Me.lblSelection.Name = "lblSelection"
        Me.lblSelection.Size = New System.Drawing.Size(89, 15)
        Me.lblSelection.TabIndex = 15
        Me.lblSelection.Text = "Color Selection"
        '
        'btnAbout
        '
        Me.btnAbout.Location = New System.Drawing.Point(26, 242)
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.Size = New System.Drawing.Size(87, 27)
        Me.btnAbout.TabIndex = 16
        Me.btnAbout.Text = "About"
        Me.btnAbout.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.tpColor)
        Me.TabControl1.Controls.Add(Me.tpCheck)
        Me.TabControl1.Controls.Add(Me.tpTriangle)
        Me.TabControl1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(26, 24)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(507, 371)
        Me.TabControl1.TabIndex = 17
        '
        'tpColor
        '
        Me.tpColor.Controls.Add(Me.lblYear1)
        Me.tpColor.Controls.Add(Me.btnAbout)
        Me.tpColor.Controls.Add(Me.lblYear2)
        Me.tpColor.Controls.Add(Me.lblSelection)
        Me.tpColor.Controls.Add(Me.lblYear3)
        Me.tpColor.Controls.Add(Me.lblRGB)
        Me.tpColor.Controls.Add(Me.lblYear4)
        Me.tpColor.Controls.Add(Me.lblYearTitle)
        Me.tpColor.Controls.Add(Me.lblYear5)
        Me.tpColor.Controls.Add(Me.btnCancel)
        Me.tpColor.Controls.Add(Me.lstColorName)
        Me.tpColor.Controls.Add(Me.btnApply)
        Me.tpColor.Controls.Add(Me.lblYear1RGB)
        Me.tpColor.Controls.Add(Me.lblYear5RGB)
        Me.tpColor.Controls.Add(Me.lblYear2RGB)
        Me.tpColor.Controls.Add(Me.lblYear4RGB)
        Me.tpColor.Controls.Add(Me.lblYear3RGB)
        Me.tpColor.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpColor.Location = New System.Drawing.Point(4, 24)
        Me.tpColor.Name = "tpColor"
        Me.tpColor.Padding = New System.Windows.Forms.Padding(3)
        Me.tpColor.Size = New System.Drawing.Size(499, 343)
        Me.tpColor.TabIndex = 0
        Me.tpColor.Text = "Graph Line Color"
        Me.tpColor.UseVisualStyleBackColor = True
        '
        'tpCheck
        '
        Me.tpCheck.Location = New System.Drawing.Point(4, 24)
        Me.tpCheck.Name = "tpCheck"
        Me.tpCheck.Padding = New System.Windows.Forms.Padding(3)
        Me.tpCheck.Size = New System.Drawing.Size(499, 343)
        Me.tpCheck.TabIndex = 1
        Me.tpCheck.Text = "Validate Reserves"
        Me.tpCheck.UseVisualStyleBackColor = True
        '
        'tpTriangle
        '
        Me.tpTriangle.Controls.Add(Me.grpProjBase)
        Me.tpTriangle.Controls.Add(Me.lblRsvTotal)
        Me.tpTriangle.Controls.Add(Me.lblReserves)
        Me.tpTriangle.Controls.Add(Me.btnPrior)
        Me.tpTriangle.Controls.Add(Me.btnDefault)
        Me.tpTriangle.Controls.Add(Me.btnSelected)
        Me.tpTriangle.Controls.Add(Me.dgvTriangle)
        Me.tpTriangle.Location = New System.Drawing.Point(4, 24)
        Me.tpTriangle.Name = "tpTriangle"
        Me.tpTriangle.Padding = New System.Windows.Forms.Padding(3)
        Me.tpTriangle.Size = New System.Drawing.Size(499, 343)
        Me.tpTriangle.TabIndex = 2
        Me.tpTriangle.Text = "Complete Triangle"
        Me.tpTriangle.UseVisualStyleBackColor = True
        '
        'lblRsvTotal
        '
        Me.lblRsvTotal.AutoSize = True
        Me.lblRsvTotal.Location = New System.Drawing.Point(398, 235)
        Me.lblRsvTotal.Name = "lblRsvTotal"
        Me.lblRsvTotal.Size = New System.Drawing.Size(43, 15)
        Me.lblRsvTotal.TabIndex = 7
        Me.lblRsvTotal.Text = "Label2"
        '
        'lblReserves
        '
        Me.lblReserves.AutoSize = True
        Me.lblReserves.Location = New System.Drawing.Point(398, 200)
        Me.lblReserves.Name = "lblReserves"
        Me.lblReserves.Size = New System.Drawing.Size(85, 15)
        Me.lblReserves.TabIndex = 6
        Me.lblReserves.Text = "Total Reserves"
        '
        'btnPrior
        '
        Me.btnPrior.Location = New System.Drawing.Point(398, 154)
        Me.btnPrior.Name = "btnPrior"
        Me.btnPrior.Size = New System.Drawing.Size(75, 23)
        Me.btnPrior.TabIndex = 5
        Me.btnPrior.Text = "Prior"
        Me.btnPrior.UseVisualStyleBackColor = True
        '
        'btnDefault
        '
        Me.btnDefault.Location = New System.Drawing.Point(398, 111)
        Me.btnDefault.Name = "btnDefault"
        Me.btnDefault.Size = New System.Drawing.Size(75, 23)
        Me.btnDefault.TabIndex = 4
        Me.btnDefault.Text = "Default"
        Me.btnDefault.UseVisualStyleBackColor = True
        '
        'btnSelected
        '
        Me.btnSelected.Location = New System.Drawing.Point(398, 68)
        Me.btnSelected.Name = "btnSelected"
        Me.btnSelected.Size = New System.Drawing.Size(75, 23)
        Me.btnSelected.TabIndex = 3
        Me.btnSelected.Text = "Selected"
        Me.btnSelected.UseVisualStyleBackColor = True
        '
        'dgvTriangle
        '
        Me.dgvTriangle.AllowUserToAddRows = False
        Me.dgvTriangle.AllowUserToDeleteRows = False
        Me.dgvTriangle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTriangle.Location = New System.Drawing.Point(24, 68)
        Me.dgvTriangle.Name = "dgvTriangle"
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvTriangle.RowHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvTriangle.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.dgvTriangle.Size = New System.Drawing.Size(349, 245)
        Me.dgvTriangle.TabIndex = 0
        '
        'grpProjBase
        '
        Me.grpProjBase.Controls.Add(Me.rdoIncurred)
        Me.grpProjBase.Controls.Add(Me.rdoPaid)
        Me.grpProjBase.Location = New System.Drawing.Point(24, 7)
        Me.grpProjBase.Name = "grpProjBase"
        Me.grpProjBase.Size = New System.Drawing.Size(349, 46)
        Me.grpProjBase.TabIndex = 8
        Me.grpProjBase.TabStop = False
        '
        'rdoPaid
        '
        Me.rdoPaid.AutoSize = True
        Me.rdoPaid.Location = New System.Drawing.Point(5, 17)
        Me.rdoPaid.Name = "rdoPaid"
        Me.rdoPaid.Size = New System.Drawing.Size(48, 19)
        Me.rdoPaid.TabIndex = 0
        Me.rdoPaid.TabStop = True
        Me.rdoPaid.Text = "Paid"
        Me.rdoPaid.UseVisualStyleBackColor = True
        '
        'rdoIncurred
        '
        Me.rdoIncurred.AutoSize = True
        Me.rdoIncurred.Location = New System.Drawing.Point(140, 17)
        Me.rdoIncurred.Name = "rdoIncurred"
        Me.rdoIncurred.Size = New System.Drawing.Size(71, 19)
        Me.rdoIncurred.TabIndex = 1
        Me.rdoIncurred.TabStop = True
        Me.rdoIncurred.Text = "Incurred"
        Me.rdoIncurred.UseVisualStyleBackColor = True
        '
        'frmLineColor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(563, 424)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmLineColor"
        Me.Text = "Multi-Purpose Form"
        Me.TabControl1.ResumeLayout(False)
        Me.tpColor.ResumeLayout(False)
        Me.tpColor.PerformLayout()
        Me.tpTriangle.ResumeLayout(False)
        Me.tpTriangle.PerformLayout()
        CType(Me.dgvTriangle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProjBase.ResumeLayout(False)
        Me.grpProjBase.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblYear2 As System.Windows.Forms.Label
    Friend WithEvents lblYear3 As System.Windows.Forms.Label
    Friend WithEvents lblYear4 As System.Windows.Forms.Label
    Friend WithEvents lblYear5 As System.Windows.Forms.Label
    Friend WithEvents lstColorName As System.Windows.Forms.ListBox
    Friend WithEvents lblYear5RGB As System.Windows.Forms.Label
    Friend WithEvents lblYear4RGB As System.Windows.Forms.Label
    Friend WithEvents lblYear3RGB As System.Windows.Forms.Label
    Friend WithEvents lblYear2RGB As System.Windows.Forms.Label
    Friend WithEvents lblYear1RGB As System.Windows.Forms.Label
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblYearTitle As System.Windows.Forms.Label
    Friend WithEvents lblRGB As System.Windows.Forms.Label
    Friend WithEvents lblSelection As System.Windows.Forms.Label
    Friend WithEvents lblYear1 As System.Windows.Forms.Label
    Friend WithEvents btnAbout As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents tpColor As System.Windows.Forms.TabPage
    Friend WithEvents tpCheck As System.Windows.Forms.TabPage
    Friend WithEvents tpTriangle As System.Windows.Forms.TabPage
    Friend WithEvents lblRsvTotal As System.Windows.Forms.Label
    Friend WithEvents lblReserves As System.Windows.Forms.Label
    Friend WithEvents btnPrior As System.Windows.Forms.Button
    Friend WithEvents btnDefault As System.Windows.Forms.Button
    Friend WithEvents btnSelected As System.Windows.Forms.Button
    Friend WithEvents dgvTriangle As System.Windows.Forms.DataGridView
    Friend WithEvents grpProjBase As System.Windows.Forms.GroupBox
    Friend WithEvents rdoIncurred As System.Windows.Forms.RadioButton
    Friend WithEvents rdoPaid As System.Windows.Forms.RadioButton
End Class
