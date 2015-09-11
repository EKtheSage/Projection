Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices


<ComVisible(True)>
Public Class FactorSelectionPane
    Inherits UserControl

    Public theLabel As Label

    Public Sub New()

        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        theLabel = New Label()
        theLabel.Text = "Factor Selection"
        theLabel.Location = New Point(20, 20)
        theLabel.Size = New Size(200, 60)
        Controls.Add(theLabel)
    End Sub

    Private Sub InitializeComponent()
        HScrollBarSimAvg = New HScrollBar()
        HScrollBarWtdAvg = New HScrollBar()
        HScrollBarLstSqr = New HScrollBar()
        HScrollBarHiLo = New HScrollBar()
        HScrollBarSeasonal = New HScrollBar()
        SuspendLayout()
        '
        'HScrollBarSimAvg
        '
        HScrollBarSimAvg.Location = New Point(44, 66)
        HScrollBarSimAvg.Name = "HScrollBarSimAvg"
        HScrollBarSimAvg.Size = New Size(116, 17)
        HScrollBarSimAvg.TabIndex = 0
        HScrollBarSimAvg.Value = 50
        '
        'HScrollBarWtdAvg
        '
        HScrollBarWtdAvg.Location = New Point(44, 83)
        HScrollBarWtdAvg.Name = "HScrollBarWtdAvg"
        HScrollBarWtdAvg.Size = New Size(116, 17)
        HScrollBarWtdAvg.TabIndex = 1
        HScrollBarWtdAvg.Value = 50
        '
        'HScrollBarLstSqr
        '
        HScrollBarLstSqr.Location = New Point(44, 100)
        HScrollBarLstSqr.Name = "HScrollBarLstSqr"
        HScrollBarLstSqr.Size = New Size(116, 17)
        HScrollBarLstSqr.TabIndex = 2
        HScrollBarLstSqr.Value = 50
        '
        'HScrollBarHiLo
        '
        HScrollBarHiLo.Location = New Point(44, 117)
        HScrollBarHiLo.Name = "HScrollBarHiLo"
        HScrollBarHiLo.Size = New Size(116, 17)
        HScrollBarHiLo.TabIndex = 3
        HScrollBarHiLo.Value = 50
        '
        'HScrollBarSeasonal
        '
        HScrollBarSeasonal.Location = New Point(44, 134)
        HScrollBarSeasonal.Name = "HScrollBarSeasonal"
        HScrollBarSeasonal.Size = New Size(116, 17)
        HScrollBarSeasonal.TabIndex = 4
        HScrollBarSeasonal.Value = 50
        '
        'FactorSelectionPane
        '
        Controls.Add(HScrollBarSeasonal)
        Controls.Add(HScrollBarHiLo)
        Controls.Add(HScrollBarSimAvg)
        Controls.Add(HScrollBarWtdAvg)
        Controls.Add(HScrollBarLstSqr)
        Name = "FactorSelectionPane"
        Size = New Size(235, 301)
        ResumeLayout(False)

    End Sub

    Friend WithEvents HScrollBarSimAvg As HScrollBar
    Friend WithEvents HScrollBarWtdAvg As HScrollBar
    Friend WithEvents HScrollBarLstSqr As HScrollBar
    Friend WithEvents HScrollBarHiLo As HScrollBar
    Friend WithEvents HScrollBarSeasonal As HScrollBar


    Private Sub HScrollBarSimAvg_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBarSimAvg.Scroll
        Dim wkst As Excel.Worksheet = CType(Application.ActiveSheet, Excel.Worksheet)
        Dim rng As Excel.Range = wkst.Range("E364")
        rng.Value = HScrollBarSimAvg.Value
    End Sub


End Class
