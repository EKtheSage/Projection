<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FactorSelection
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.HScrollBar1 = New System.Windows.Forms.HScrollBar()
        Me.HScrollBar2 = New System.Windows.Forms.HScrollBar()
        Me.HScrollBar3 = New System.Windows.Forms.HScrollBar()
        Me.txtBxSimpleAvg = New System.Windows.Forms.TextBox()
        Me.txtBxWtdAvg = New System.Windows.Forms.TextBox()
        Me.txtBxLstSqr = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(190, 364)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'HScrollBar1
        '
        Me.HScrollBar1.Location = New System.Drawing.Point(176, 61)
        Me.HScrollBar1.Name = "HScrollBar1"
        Me.HScrollBar1.Size = New System.Drawing.Size(80, 20)
        Me.HScrollBar1.TabIndex = 1
        '
        'HScrollBar2
        '
        Me.HScrollBar2.Location = New System.Drawing.Point(176, 87)
        Me.HScrollBar2.Name = "HScrollBar2"
        Me.HScrollBar2.Size = New System.Drawing.Size(80, 20)
        Me.HScrollBar2.TabIndex = 2
        '
        'HScrollBar3
        '
        Me.HScrollBar3.Location = New System.Drawing.Point(176, 113)
        Me.HScrollBar3.Name = "HScrollBar3"
        Me.HScrollBar3.Size = New System.Drawing.Size(80, 20)
        Me.HScrollBar3.TabIndex = 3
        '
        'txtBxSimpleAvg
        '
        Me.txtBxSimpleAvg.Location = New System.Drawing.Point(21, 61)
        Me.txtBxSimpleAvg.Multiline = True
        Me.txtBxSimpleAvg.Name = "txtBxSimpleAvg"
        Me.txtBxSimpleAvg.Size = New System.Drawing.Size(108, 20)
        Me.txtBxSimpleAvg.TabIndex = 4
        Me.txtBxSimpleAvg.Text = "Simple Average"
        '
        'txtBxWtdAvg
        '
        Me.txtBxWtdAvg.Location = New System.Drawing.Point(21, 87)
        Me.txtBxWtdAvg.Multiline = True
        Me.txtBxWtdAvg.Name = "txtBxWtdAvg"
        Me.txtBxWtdAvg.Size = New System.Drawing.Size(108, 20)
        Me.txtBxWtdAvg.TabIndex = 5
        Me.txtBxWtdAvg.Text = "Weighted Average"
        '
        'txtBxLstSqr
        '
        Me.txtBxLstSqr.Location = New System.Drawing.Point(21, 113)
        Me.txtBxLstSqr.Multiline = True
        Me.txtBxLstSqr.Name = "txtBxLstSqr"
        Me.txtBxLstSqr.Size = New System.Drawing.Size(108, 20)
        Me.txtBxLstSqr.TabIndex = 6
        Me.txtBxLstSqr.Text = "Least Squares"
        '
        'FactorSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.txtBxLstSqr)
        Me.Controls.Add(Me.txtBxWtdAvg)
        Me.Controls.Add(Me.txtBxSimpleAvg)
        Me.Controls.Add(Me.HScrollBar3)
        Me.Controls.Add(Me.HScrollBar2)
        Me.Controls.Add(Me.HScrollBar1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "FactorSelection"
        Me.Size = New System.Drawing.Size(297, 430)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents HScrollBar1 As System.Windows.Forms.HScrollBar
    Friend WithEvents HScrollBar2 As System.Windows.Forms.HScrollBar
    Friend WithEvents HScrollBar3 As System.Windows.Forms.HScrollBar
    Friend WithEvents txtBxSimpleAvg As System.Windows.Forms.TextBox
    Friend WithEvents txtBxWtdAvg As System.Windows.Forms.TextBox
    Friend WithEvents txtBxLstSqr As System.Windows.Forms.TextBox
End Class
