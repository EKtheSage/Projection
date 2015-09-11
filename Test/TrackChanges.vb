Imports ExcelDna.Integration
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class TrackChanges
    Implements IExcelAddIn

    Dim WithEvents Application As Application

    Public Sub AutoClose() Implements IExcelAddIn.AutoClose

    End Sub

    'to-do: route multiple SheetChange events to the same handler
    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen
        Application = CType(ExcelDnaUtil.Application, Application)
        AddHandler Application.SheetChange, AddressOf ATASheetChange
    End Sub

    Private Sub ATASheetChange(sh As Object, target As Range)
        Dim rngName As Name

        If CType(sh, Worksheet).Name = projBase Then
            'this will auto-update Age1 factor, EL, reserves to review template
            worksheetTriangleChange(target)
        End If

        If CType(sh, Worksheet).Name = "Exp Loss" Then
            worksheetExpLossChange(target)
        End If

        If CType(sh, Worksheet).Name = "Review Template" Then
            worksheetReviewTemplateChange(target)
        End If

        If CType(sh, Worksheet).Name = "Control" Then
            Try
                rngName = CType(target.Name, Name)
                If rngName.Name = "reset" Then
                    resetControl(CType(target.Value, String))
                ElseIf rngName.Name = "eval_group" Then
                    evalGroup = CType(target.Value, String)
                ElseIf rngName.Name = "proj_base" Then
                    projBase = CType(target.Value, String)
                ElseIf rngName.Name = "include_ss"
                    includeSS = CType(target.Value, String)
                End If
            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub worksheetTriangleChange(target As Range)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(projBase), Worksheet)
        Dim rng As Range
        Dim row As Integer
        Dim dt As Date = Now
        row = CType(wkstReviewTemplate.Cells(wkstReviewTemplate.Rows.Count, 25), Range).End(XlDirection.xlUp).Row + 1

        If evalGroup = "Monthly" Then
            rng = wkst.Range(projBase & "_sel_ATA").Offset(-1, 0).Resize(1, 1)
        Else
            rng = wkst.Range(projBase & "_sel_ATA_qtrly").Offset(-1, 0).Resize(1, 1)
        End If

        If Application.Intersect(target, rng) Is Nothing Then
            Exit Sub
        End If

        CType(wkstReviewTemplate.Cells(row, 25), Range).Value = projBase
        CType(wkstReviewTemplate.Cells(row, 26), Range).Value = target.Offset(1, 0).Value
        CType(wkstReviewTemplate.Cells(row, 27), Range).Value = wkstExpLoss.Range("$P$11").Value
        CType(wkstReviewTemplate.Cells(row, 28), Range).Value =
            sumRange(CType(CType(wkstSummary.Range("summary").Columns(33), Range).Value, Object(,)))
        CType(wkstReviewTemplate.Cells(row, 29), Range).Value = dt
        CType(wkstReviewTemplate.Cells(row, 30), Range).Value =
            CType(Application.ActiveWorkbook.BuiltinDocumentProperties, DocumentProperties)("Last Author").Value
    End Sub

    Private Sub worksheetReviewTemplateChange(target As Range)
        'check target is in final sel, if it is, get the index of the target
        'assign the value to the Paid/Incurred monthly/quarterly ATA.
        Dim selAddress As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        Dim colIndex As Integer
        Dim finalSel As Range = wkstReviewTemplate.Range("finalATASel")
        Dim lookup As Range = wkstExpLoss.Range("lookup_age")
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(projBase), Worksheet)
        Dim ATANamedRange As String

        If Application.Intersect(target, finalSel) Is Nothing And
           Application.Intersect(target, Application.Range("D21")) Is Nothing And
           Application.Intersect(target, Application.Range("D22")) Is Nothing And
           Application.Intersect(target, Application.Range("D23")) Is Nothing And
           Application.Intersect(target, Application.Range("F27")) Is Nothing Then
            Exit Sub
        End If

        If evalGroup = "Monthly" Then
            lookup.Value = 1
        Else
            lookup.Value = 3
        End If

        If Application.Intersect(target, finalSel) IsNot Nothing Then
            If evalGroup = "Monthly" Then
                ATANamedRange = projBase & "_sel_ATA"
            Else
                ATANamedRange = projBase & "_sel_ATA_qtrly"
            End If

            For i As Integer = 0 To finalSel.Rows.Count - 1
                selAddress.Add(CType(finalSel.Item(finalSel.Rows.Count - i, 1), Range).Address, i)
            Next

            colIndex = selAddress.Item(target.Address)

            'get the cell based on the column index and the selected ATA named range.
            wkst.Range(ATANamedRange).Offset(-1, colIndex).Resize(1, 1).Value = target.Value

        ElseIf Application.Intersect(target, Application.Range("D21")) IsNot Nothing Then 'sev trend
            lookup.Offset(2, 0).Value = target.Value
        ElseIf Application.Intersect(target, Application.Range("D22")) IsNot Nothing Then 'pp trend
            lookup.Offset(5, 0).Value = target.Value
        ElseIf Application.Intersect(target, Application.Range("D23")) IsNot Nothing Then 'lr trend
            lookup.Offset(8, 0).Value = target.Value
        ElseIf Application.Intersect(target, Application.Range("F27")) IsNot Nothing Then 'final exp loss
            lookup.Offset(10, 0).Value = target.Value
        End If

    End Sub

    Private Sub worksheetExpLossChange(target As Range)
        Dim expLoss As Range = wkstSummary.Range("exp_loss")
        Dim lookup As Range = wkstExpLoss.Range("lookup_age")
        Dim rowNum As Integer = expLoss.Rows.Count
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(projBase), Worksheet)
        Dim age1 As Object
        Dim dt As Date = Now
        Dim row As Integer
        Dim counter As Integer

        If Application.Intersect(target, wkstExpLoss.Range("P11")) Is Nothing Then Exit Sub

        If evalGroup = "Monthly" Then
            counter = 1
            age1 = wkst.Range(projBase & "_sel_ATA").Resize(1, 1).Value
        Else
            counter = 3
            age1 = wkst.Range(projBase & "_sel_ATA_qtrly").Resize(1, 1).Value
        End If
        CType(expLoss.Cells(rowNum - (CType(lookup.Value, Integer) / counter) + 1, 1), Range).Value = target.Value

        row = CType(wkstReviewTemplate.Cells(wkstReviewTemplate.Rows.Count, 25), Range).End(XlDirection.xlUp).Row + 1
        CType(wkstReviewTemplate.Cells(row, 25), Range).Value = projBase
        CType(wkstReviewTemplate.Cells(row, 26), Range).Value = age1
        CType(wkstReviewTemplate.Cells(row, 27), Range).Value = target.Value
        CType(wkstReviewTemplate.Cells(row, 28), Range).Value =
            sumRange(CType(CType(wkstSummary.Range("summary").Columns(33), Range).Value, Object(,)))
        CType(wkstReviewTemplate.Cells(row, 29), Range).Value = dt
        CType(wkstReviewTemplate.Cells(row, 30), Range).Value =
            CType(Application.ActiveWorkbook.BuiltinDocumentProperties, DocumentProperties)("Last Author").Value

    End Sub
    Private Sub resetControl(reset As String)
        If reset = "Yes" Then
            For Each pvtTbl As PivotTable In CType(wkstControl.PivotTables, PivotTables)
                pvtTbl.ClearAllFilters()
            Next
        End If
    End Sub

    Private Sub resizeNamedRangeAndCreateValidation(rngName As String, selectedItm As String)
        'takes in parameter to decide which of the 4 strings to not be executed
        Dim rng As Range
        Dim listAsNames As String
        Dim triangleList As PivotTable = CType(wkstData.PivotTables("PT_TriangleList"), PivotTable)
        Dim pvtFld, pvtFld2 As PivotField
        Dim pvtItm As PivotItem



        For Each pvtFld In CType(triangleList.RowFields, PivotFields)
            If pvtFld.Name = rngName Then
                'first reset the visibility of each item to be true
                For Each pvtItm In CType(pvtFld.PivotItems, PivotItems)
                    pvtItm.Visible = True
                Next
                'then make all non-selected items false
                For Each pvtItm In CType(pvtFld.PivotItems, PivotItems)
                    If pvtItm.Name <> selectedItm Then
                        pvtItm.Visible = False
                    End If
                Next
                'then adjust other fields so data-validation are adjusted
                For Each pvtFld2 In CType(triangleList.RowFields, PivotFields)
                    If pvtFld2.Name <> "coverage" And pvtFld2.Name <> "lob" Then
                        listAsNames = getList(pvtFld2.Name & "List")
                        rng = wkstControl.Range(pvtFld2.Name)
                        rng.Validation.Delete()
                        rng.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                                           XlFormatConditionOperator.xlBetween, listAsNames)
                    End If
                Next
            End If
        Next

    End Sub

    Private Function getList(name As String) As String
        'a named range with un-contiguous ranges, join the values, remove dups, return a comma delimited string
        Dim rng As Range
        Dim strList As List(Of String) = New List(Of String)
        Dim strDelim As String

        rng = wkstData.Range(name)

        For Each c As Range In rng
            strList.Add(CType(c.Value, String))
        Next

        strDelim = String.Join(", ", strList.Distinct().ToList)
        Return strDelim
    End Function
End Class
