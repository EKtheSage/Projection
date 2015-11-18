Imports ExcelDna.Integration.CustomUI
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration

<ComVisible(True)>
Public Class TestRibbon
    Inherits ExcelRibbon

    Private ctp As CustomTaskPane

    Public Shared Sub ctp_VisibleStateChange(ctp As CustomTaskPane)
        MsgBox("Visibility changed to " & ctp.Visible)
    End Sub

    Public Shared Sub ctp_DockPositionStateChange(ctp As CustomTaskPane)
        Dim ctrl As FactorSelectionPane
        ctrl = CType(ctp.ContentControl, FactorSelectionPane)
        ctrl.TheLabel.Text = "Moved to " & ctp.DockPosition.ToString()
    End Sub

    Public Sub OnShowCTP(ByVal control As IRibbonControl)
        'This sub needs to be modified when working with Excel 2013
        Dim theType As Type
        theType = GetType(FactorSelectionPane)
        If ctp Is Nothing Then
            ctp = CustomTaskPaneFactory.CreateCustomTaskPane(theType, "Factor Selection Pane")
            ctp.Visible = True
            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight
            ctp.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal
            AddHandler ctp.DockPositionStateChange, AddressOf ctp_DockPositionStateChange
            AddHandler ctp.VisibleStateChange, AddressOf ctp_VisibleStateChange
        Else
            ctp.Visible = True
        End If
    End Sub

    Public Sub OnDeleteCTP(ByVal control As IRibbonControl)
        If Not ctp Is Nothing Then
            ctp.Delete()
            ctp = Nothing
        End If
    End Sub

    Public Function customGetImage(control As IRibbonControl) As Object
        If control.Id = "finalizeQPage" Then
            Return My.Resources.star_struck
        ElseIf control.Id = "getTriangleList" Then
            Return My.Resources.step1
        ElseIf control.Id = "assignValueToControl" Then
            Return My.Resources.step2
        ElseIf control.Id = "getHistoryData" Then
            Return My.Resources.step3
        ElseIf control.Id = "setupProjection" Then
            Return My.Resources.step4
        ElseIf control.Id = "selected" Then
            Return My.Resources.laugh5
        ElseIf control.Id = "default" Then
            Return My.Resources.laugh4
        ElseIf control.Id = "prior" Then
            Return My.Resources.laugh3
        ElseIf control.Id = "showAllTriangles" Then
            Return My.Resources.cool
        ElseIf control.Id = "showMonthlyTriangles" Then
            Return My.Resources.laugh1
        ElseIf control.Id = "showQuarterlyTriangles" Then
            Return My.Resources.laugh2
        ElseIf control.Id = "showDefaultTriangle" Then
            Return My.Resources.angel
        ElseIf control.Id = "clear" Then
            Return My.Resources.xd
        Else
            Return 1
        End If
    End Function
End Class

Public Module Module1
    'Have a module that houses the macros for ribbons
    Public Sub clearCompleted()
        If evalGroup = "Monthly" Then
            clearTriangle(projBase, projBase & "_data")
        Else
            clearTriangle(projBase, projBase & "_qtrlydata")
        End If
    End Sub
    Public Sub useSelected()
        If evalGroup = "Monthly" Then
            completeTriangle(projBase, projBase & "_data", projBase & "_sel_ATA")
        Else
            completeTriangle(projBase, projBase & "_qtrlydata", projBase & "_sel_ATA_qtrly")
        End If
    End Sub
    Public Sub useDefault()
        If evalGroup = "Monthly" Then
            completeTriangle(projBase, projBase & "_data", projBase & "_default_ATA")
        Else
            completeTriangle(projBase, projBase & "_qtrlydata", projBase & "_default_ATA_qtrly")
        End If
    End Sub
    Public Sub usePrior()
        If evalGroup = "Monthly" Then
            completeTriangle(projBase, projBase & "_data", projBase & "_lastTime_ATA")
        Else
            completeTriangle(projBase, projBase & "_qtrlydata", projBase & "_lastTime_ATA_qtrly")
        End If
    End Sub

    Sub Button2()
        MsgBox("Greetings from Button2!")
    End Sub

    Public Sub testATA()
        'currently this produces values as hard-coded numbers, not sure if people agree with it
        'just like what we have on webtriangle, we only see numbers, but we trust the code behind
        'then we have to be 100% sure our code works, how do we test it works?

        Dim selection As ExcelReference = CType(XlCall.Excel(XlCall.xlfSelection), ExcelReference)
        Dim selectval As Object(,) = CType(selection.GetValue(), Object(,))
        Dim result As Double(,)

        result = ChainLadder.ATA(selectval)

        Dim sheet2 As ExcelReference = CType(XlCall.Excel(XlCall.xlSheetId, "sheet2"), ExcelReference)
        Dim target As ExcelReference = New ExcelReference(0, result.GetLength(0) - 1, 0, result.GetLength(1) - 1, sheet2.SheetId)

        target.SetValue(result)
    End Sub
End Module
