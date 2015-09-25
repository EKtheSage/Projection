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
        AddHandler Application.SheetChange, AddressOf WorksheetChange
    End Sub

    Private Sub WorksheetChange(sh As Object, target As Range)
        Dim rngName As Name

        If CType(sh, Worksheet).Name = "Exp Loss" Then
            worksheetExpLossChange(target)
        End If

        If CType(sh, Worksheet).Name = "Review Template" Then
            worksheetReviewTemplateChange(target)
        End If

        If CType(sh, Worksheet).Name = "Control" Then
            Try
                rngName = CType(target.Name, Name)
                If rngName.Name = "eval_group" Then
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
           Application.Intersect(target, wkstReviewTemplate.Range("RT_SevTrnd")) Is Nothing And
           Application.Intersect(target, wkstReviewTemplate.Range("RT_PPTrnd")) Is Nothing And
           Application.Intersect(target, wkstReviewTemplate.Range("RT_LRTrnd")) Is Nothing And
           Application.Intersect(target, wkstReviewTemplate.Range("RT_ExpLossAge1")) Is Nothing Then
            Exit Sub
        End If

        'assigns the age cell to 1 or 3 first before changing the trend values below
        If evalGroup = "Monthly" Then
            lookup.Value = 1
            ATANamedRange = projBase & "_sel_ATA"
        Else
            lookup.Value = 3
            ATANamedRange = projBase & "_sel_ATA_qtrly"
        End If

        'if target is in F10:F15 and active worksheet is Review Template - then change the ATA factors in Paid/Incurred
        If Application.Intersect(target, finalSel) IsNot Nothing And CType(Application.ActiveWorkbook.ActiveSheet, Worksheet).Name = "Review Template" Then

            'add range's address to selAddress
            For i As Integer = 0 To finalSel.Rows.Count - 1
                selAddress.Add(CType(finalSel.Item(finalSel.Rows.Count - i, 1), Range).Address, i)
            Next
            colIndex = selAddress.Item(target.Address)
            'get the cell based on the column index and the selected ATA named range.
            wkst.Range(ATANamedRange).Offset(-1, colIndex).Resize(1, 1).Value = target.Value

            'add track changes here?
        ElseIf Application.Intersect(target, wkstReviewTemplate.Range("RT_SevTrnd")) IsNot Nothing And
                CType(Application.ActiveWorkbook.ActiveSheet, Worksheet).Name = "Review Template" Then 'sev trend

            lookup.Offset(2, 0).Value = target.Value

        ElseIf Application.Intersect(target, wkstReviewTemplate.Range("RT_PPTrnd")) IsNot Nothing And
                CType(Application.ActiveWorkbook.ActiveSheet, Worksheet).Name = "Review Template" Then 'pp trend

            lookup.Offset(5, 0).Value = target.Value

        ElseIf Application.Intersect(target, wkstReviewTemplate.Range("RT_LRTrnd")) IsNot Nothing And
                CType(Application.ActiveWorkbook.ActiveSheet, Worksheet).Name = "Review Template" Then 'lr trend

            lookup.Offset(8, 0).Value = target.Value

        ElseIf Application.Intersect(target, wkstReviewTemplate.Range("RT_ExpLossAge1")) IsNot Nothing And
                CType(Application.activeworkbook.ActiveSheet, Worksheet).Name = "Review Template" Then 'final exp loss

            lookup.Offset(10, 0).Value = target.Value
            'add track changes here?
        End If

    End Sub

    Private Sub worksheetExpLossChange(target As Range)
        Dim expLoss As Range = wkstSummary.Range("exp_loss")
        Dim lookup As Range = wkstExpLoss.Range("lookup_age")
        Dim rowNum As Integer = expLoss.Rows.Count
        Dim row As Integer
        Dim counter As Integer

        If Application.Intersect(target, wkstExpLoss.Range("P11")) Is Nothing And
            Application.Intersect(target, lookup) Is Nothing Then
            Exit Sub
        End If

        If evalGroup = "Monthly" Then
            counter = 1
        Else
            counter = 3
        End If

        'get the exp loss from Summary tab - when we change the age cell
        If Application.Intersect(target, lookup) IsNot Nothing Then
            row = rowNum - CType(CType(lookup.Value, Integer) / counter, Integer) + 1
            wkstExpLoss.Range("P11").Value = CType(expLoss.Cells(row, 1), Range).Value
            Exit Sub
        End If

        'change the exp loss in Summary tab to the target value
        CType(expLoss.Cells(rowNum - (CType(lookup.Value, Integer) / counter) + 1, 1), Range).Value = target.Value

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
