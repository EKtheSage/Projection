'Future add-ons to the module:
' 1. Apply different ATA factors based on number of working days (Good to have)
' 2. Conditional formatting on the triangle, show heatmap ??
' 3. Reserve Ranges ??
' 4. Initialize triangle objects for calculation behind the scene ??

Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core

Module Globals
    Public ReadOnly Property Application As Application
        Get
            Application = CType(ExcelDnaUtil.Application, Application)
        End Get
    End Property
End Module

Public Module ProjectionFormat
    Private Const monthRowNum As Integer = 180
    Private Const quarterRowNum As Integer = 60

    'declare variables to be used throughout the project
    Public wkstControl As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Control"), Worksheet)
    Public wkstConstants As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Constants"), Worksheet)
    Public wkstCount As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Count"), Worksheet)
    Public wkstPaid As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Paid"), Worksheet)
    Public wkstIncurred As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Incurred"), Worksheet)
    Public wkstExpLoss As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Exp Loss"), Worksheet)
    Public wkstSummary As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Summary"), Worksheet)
    Public wkstQPage As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Q Page"), Worksheet)
    Public wkstReviewTemplate As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Review Template"), Worksheet)
    Public wkstData As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Data"), Worksheet)
    Public wkstIBNRCnt As Worksheet = CType(Application.ActiveWorkbook.Worksheets("GU IBNR Count"), Worksheet)
    Public wkstClsdAvg As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Closed Avg"), Worksheet)
    Public wkstClsMod As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Closure Model"), Worksheet)
    Public evalGroup As String = CType(wkstControl.Range("eval_group").Value, String)
    Public projBase As String = CType(wkstControl.Range("proj_base").Value, String)
    Public includeSS As String = CType(wkstControl.Range("include_ss").Value, String)
    Public coverageField As String = CType(wkstControl.Range("coverage").Value, String)
    Public priorRes As Double = CType(CType(wkstData.ListObjects("BI_ind_res25").DataBodyRange.Cells(1, 2), Range).Value, Double)

    Enum namedRanges
        'this enum will allow us to do range resize, will need to change numbers 
        'whenever summary worksheet columns change
        accident_date = 1
        age = 2
        ep = 3
        ee = 4
        avg_prem = 5
        ult_counts = 6
        freq = 7
        cur_paid = 8
        percent_paid = 9
        ult_paid = 10
        cur_incurred = 11
        percent_incurred = 12
        ult_incurred = 13
        exp_loss = 14
        bf = 15
        gb = 16
        letter = 17
        preIC_ultloss = 18
        IC_ultloss = 19
        wtd_ultloss = 21
        sel_ultloss = 22
        IC_lr = 24
        IC_sev = 25
        IC_pp = 27
        preIC_res_spr = 29
        preIC_res = 31
        preIC_sev = 40
        preIC_pp = 41
        preIC_lr = 42
    End Enum

    Enum namedRangesTriangle
        'this enum will allow us to do range resize on the triangle worksheets
        _CurAmt = 3
        _Cap = 4
        _Exclusion = 5
        _prior_ATU = 6
        _default_ATU = 7
        _sel_ATU = 8
    End Enum

    Public Sub setup()

        'turn off calculation until everything is setup
        Application.Calculation = XlCalculation.xlCalculationManual

        monthToQuarter("Count")
        makeTriangleSheets("Count")

        monthToQuarter("Paid")
        makeTriangleSheets("Paid")

        monthToQuarter("Incurred")
        makeTriangleSheets("Incurred")


        'wait...if the ata data will be either the actual data or ata data, do we really have to do this part?
        'yes we do, because we need the macro to produce quarterly data triangle
        monthToQuarterAlt("Incurred")
        monthToQuarterAlt("Paid")

        showDefaultTriangleView()
        summary()
        expLoss()
        QPageFormat()
        reviewTemplate()

        'calculate all sheets, then turn calculation back to automatic
        Application.Calculate()
        Application.Calculation = XlCalculation.xlCalculationAutomatic

        wkstExpLoss.Activate()
        graphsUpdate("Exp Loss")
        graphsUpdate("Review Template")

        'remove closure model monthly spread first. need to think about quarterly spread...
        wkstClsMod.Range("clos_mod_spr_monthly").ClearContents()
        CType(Application.Worksheets(projBase), Worksheet).Activate()

    End Sub

    Public Sub clearTriangle(wkstName As String, name As String)
        Application.ScreenUpdating = False

        'First get the body of triangle
        Dim rng As String = wkstName & "!" & name
        Dim wkst As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        Dim dataVal As Object(,) = CType(wkst.GetValue, Object(,))
        Dim counter As Integer

        If evalGroup = "Monthly" Then
            counter = 179
        Else
            counter = 59
        End If

        For i As Integer = 1 To dataVal.GetUpperBound(0)
            For j As Integer = 1 + counter - i To dataVal.GetUpperBound(1)
                dataVal(i, j) = Nothing
            Next
        Next

        'Go back to data range
        wkst = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        wkst.SetValue(dataVal)

        Application.ScreenUpdating = True
    End Sub
    Public Sub completeTriangle(wkstName As String, name As String, ataName As String)

        Application.ScreenUpdating = False
        Dim counter, num As Integer

        'First get the body of triangle
        Dim rng As String = wkstName & "!" & name
        Dim wkst As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        Dim dataVal As Object(,) = CType(wkst.GetValue, Object(,))

        rng = wkstName & "!" & ataName
        wkst = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        Dim ataVal As Object(,) = CType(wkst.GetValue, Object(,))

        If dataVal.GetUpperBound(1) = 179 Then
            num = 1
            counter = 179
        Else
            num = 3
            counter = 59
        End If
        'Calculate ATA * triangle
        For i As Integer = 1 To dataVal.GetUpperBound(0)
            For j As Integer = 1 + counter - i To dataVal.GetUpperBound(1)
                dataVal(i, j) = CType(dataVal(i, j - 1), Double) * CType(ataVal(0, j - 1), Double)
            Next
        Next

        'Go back to data range
        rng = wkstName & "!" & name
        wkst = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        wkst.SetValue(dataVal)

        Application.ScreenUpdating = True
    End Sub

    Public Sub monthToQuarterAlt(data As String)
        Application.Calculation = XlCalculation.xlCalculationManual
        Dim dataRng As String = "'ATA Data'!" & data & "_data_alt"

        Dim sht2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        Dim selectVal As Object(,) = CType(sht2.GetValue(), Object(,))
        Dim qtrTri As Double(,) = quarterTriangle(selectVal)

        dataRng = "'ATA Data'!" & data & "_qtrlydata_alt"
        Dim target2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        target2.SetValue(qtrTri)

        'also assigns new formula to the ATA block - ATA data equals actual data if ATA selections are the same, no need to have if/else
        'actually no need to do this step
        'Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(data), Worksheet)
        'wkst.Range(data & "_ATA").FormulaArray = "=ATA(" & data & "_data_alt)"
        'wkst.Range(data & "_qtrlyATA").FormulaArray = "=ATA(" & data & "_qtrlydata_alt)"

    End Sub
    Public Sub monthToQuarter(shtName As String)
        Application.Calculation = XlCalculation.xlCalculationManual
        Dim dataRng As String
        dataRng = shtName & "!" & shtName & "_data"

        Dim sht2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        Dim selectVal As Object(,) = CType(sht2.GetValue(), Object(,))
        Dim qtrTri As Double(,)

        qtrTri = quarterTriangle(selectVal)

        dataRng = shtName & "!" & shtName & "_qtrlydata"
        Dim target2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        target2.SetValue(qtrTri)

    End Sub

    Public Sub showAllTriangles()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet)
        If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
            wkst.Rows.EntireRow.Hidden = False
            wkst.Columns.EntireColumn.Hidden = False
        End If
    End Sub

    Public Sub showMonthlyTriangles()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet)
        Dim hideRows As Range
        Application.ScreenUpdating = False
        If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
            CType(wkst.Rows("1:379"), Range).EntireRow.Hidden = False
            hideRows = Application.Union(CType(wkst.Rows("2:156"), Range),
                                         CType(wkst.Rows("184:337"), Range),
                                         CType(wkst.Rows("380:519"), Range))
            For Each a As Range In hideRows.Areas
                a.EntireRow.Hidden = True
            Next
        End If
        Application.ScreenUpdating = True
    End Sub

    Public Sub showQuarterlyTriangles()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet)
        Dim hideRows As Range
        Application.ScreenUpdating = False
        If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
            CType(wkst.Rows("380:519"), Range).EntireRow.Hidden = False
            hideRows = Application.Union(CType(wkst.Rows("381:415"), Range),
                                        CType(wkst.Rows("443:476"), Range),
                                        CType(wkst.Range("1:379"), Range))
            For Each a As Range In hideRows.Areas
                a.EntireRow.Hidden = True
            Next
        End If
        Application.ScreenUpdating = True
    End Sub

    Public Sub showDefaultTriangleView()
        Dim wkst As Worksheet
        Dim month, quarter As Boolean
        Dim hideRows As Range

        Application.Calculation = XlCalculation.xlCalculationManual

        For Each wkst In Application.ActiveWorkbook.Worksheets
            If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
                If evalGroup = "Monthly" Then
                    month = False
                    quarter = True
                    hideRows = Application.Union(CType(wkst.Rows("2:156"), Range), CType(wkst.Rows("184:337"), Range))
                Else
                    month = True
                    quarter = False
                    hideRows = Application.Union(CType(wkst.Rows("381:415"), Range), CType(wkst.Rows("443:476"), Range))
                End If

                CType(wkst.Rows("1:379"), Range).EntireRow.Hidden = month
                CType(wkst.Rows("380:519"), Range).EntireRow.Hidden = quarter

                For Each area In hideRows.Areas
                    hideRows.EntireRow.Hidden = True
                Next
            End If
        Next

        Application.Calculate()
        Application.Calculation = XlCalculation.xlCalculationAutomatic
    End Sub

    Public Sub makeTriangleSheets(shtName As String)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(shtName), Worksheet)
        Dim rng, rng2 As Range
        Dim dateRng, dataRng, lastTime, defaultATA, selATA As String
        Dim rowNum, counter As Integer
        Dim nameOfRange As Name
        Dim namedRangeValues As Array = System.Enum.GetValues(GetType(namedRangesTriangle))

        Application.Calculation = XlCalculation.xlCalculationManual

        rng = wkst.Range(shtName & "_Summary")
        nameOfRange = CType(rng.Name, Name)
        rng.ClearContents()
        rng.Offset(1, 0).ClearContents() 'remove total row

        If evalGroup = "Monthly" Then
            dateRng = "=accident_date_mthly"
            dataRng = shtName & "_data"
            lastTime = "=ATU(" & shtName & "_lastTime_ATA)"
            defaultATA = "= ATU(" & shtName & "_default_ATA)"
            selATA = "=ATU(" & shtName & "_sel_ATA)"
            rowNum = monthRowNum
            counter = 1
            'adding accident date comments above age cells
            For i = 1 To 179
                Dim temp As String
                temp = CType(CType(wkstCount.Range("accident_date_mthly").Cells(181 - i, 1), Range).Value, String)
                With CType(wkst.Cells(364, i + 1), Range)
                    .ClearComments()
                    .AddComment("")
                    .Comment.Text(Text:=temp)
                    .Comment.Visible = False
                    .Comment.Shape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle
                    .Comment.Shape.Fill.PresetGradient(MsoGradientStyle.msoGradientHorizontal,
                                                       1,
                                                       MsoPresetGradientType.msoGradientCalmWater)
                    .Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
                    .Comment.Shape.TextFrame.Characters.Font.Bold = True
                    .Comment.Shape.TextFrame.Characters.Font.Size = 12
                End With
            Next
        Else
            dateRng = "=accident_date_qtrly"
            dataRng = shtName & "_qtrlydata"
            lastTime = "=ATU(" & shtName & "_lastTime_ATA_qtrly)"
            defaultATA = "=ATU(" & shtName & "_default_ATA_qtrly)"
            selATA = "=ATU(" & shtName & "_sel_ATA_qtrly)"
            rowNum = quarterRowNum
            counter = 3

            For i = 1 To 59
                Dim temp As String
                temp = CType(CType(wkstCount.Range("accident_date_qtrly").Cells(61 - i, 1), Range).Value, String)
                With CType(wkst.Cells(503, i + 1), Range)
                    .ClearComments()
                    .AddComment("")
                    .Comment.Text(Text:=temp)
                    .Comment.Visible = False
                    .Comment.Shape.AutoShapeType = MsoAutoShapeType.msoShapeRoundedRectangle
                    .Comment.Shape.Fill.PresetGradient(MsoGradientStyle.msoGradientHorizontal,
                                                       1,
                                                       MsoPresetGradientType.msoGradientCalmWater)
                    .Comment.Shape.TextFrame.Characters.Font.Name = "Calibri"
                    .Comment.Shape.TextFrame.Characters.Font.Bold = True
                    .Comment.Shape.TextFrame.Characters.Font.Size = 12
                End With
            Next
        End If
        'remove selections
        wkst.Range(shtName & "_sel_ATA").Offset(-1, 0).ClearContents()
        wkst.Range(shtName & "_sel_ATA_qtrly").Offset(-1, 0).ClearContents()

        'resize named ranges here
        For Each name As namedRangesTriangle In namedRangeValues
            nameOfRange = CType(wkst.Range(shtName & name.ToString()).Name, Name)
            rng2 = wkst.Range(shtName & name.ToString()).Resize(rowNum)
            With nameOfRange
                .Name = shtName & name.ToString()
                .RefersTo = rng2
            End With
        Next

        'For Count worksheet, there is one more named range called Count_GUIBNR
        'Also, resize the summary to monthly or quarterly based on rowNum
        If shtName = "Count" Then
            nameOfRange = CType(wkst.Range("Count_GUIBNR").Name, Name)
            rng2 = wkst.Range(shtName & "_GUIBNR").Resize(rowNum)
            With nameOfRange
                .Name = shtName & "_GUIBNR"
                .RefersTo = rng2
            End With

            CType(rng.Columns(9), Range).Offset(0, 1).ClearContents()
            CType(CType(rng.Columns(9), Range).Offset(0, 1).Cells(rowNum, 1), Range).Offset(1, 0).ClearContents()
            rng = rng.Resize(rowNum, 9)
        Else
            rng = rng.Resize(rowNum, 11)
        End If

        CType(rng.Columns(1), Range).FormulaArray = dateRng

        For i As Integer = 1 To rowNum
            CType(rng.Cells(i, 2), Range).Value =
                DateDiff("m", CType(rng.Cells(i, 1), Range).Value, wkstControl.Range("CurrentEvalDate").Value) + counter
            CType(rng.Cells(i, 3), Range).Value = CType(wkst.Range(dataRng).Cells(i, rowNum - i + 1), Range).Value
        Next

        'add formula for total at bottom of range
        CType(CType(rng.Columns(3), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
            "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_curAmt))"
        CType(CType(rng.Columns(4), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
            "=SUM(INDEX(" & shtName & "_Summary,,column_summary_cap))"
        CType(CType(rng.Columns(5), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
            "=SUM(INDEX(" & shtName & "_Summary,,column_summary_exclusion))"

        If shtName = "Count" Then
            CType(rng.Columns(6), Range).FormulaArray = defaultATA
            CType(rng.Columns(7), Range).FormulaArray = selATA
            CType(rng.Columns(8), Range).Formula =
                "=(" & shtName & "_CurAmt-" & shtName & "_Cap-" & shtName & "_Exclusion)*" & shtName & "_default_ATU+" & shtName & "_Cap"
            CType(CType(rng.Columns(8), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
                "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_defaultUlt))"
            CType(rng.Columns(9), Range).Formula =
                "=IF(SUM(Count_GUIBNR) = 0, (Count_CurAmt-Count_Cap-Count_Exclusion)*Count_sel_ATU+Count_Cap," &
                    "Count_CurAmt-Count_Exclusion+Count_GUIBNR)"
            CType(CType(rng.Columns(9), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
            "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_selUlt))"

            If evalGroup = "Monthly" Then
                CType(rng.Columns(9), Range).Offset(0, 1).Formula = "=IFERROR(VLOOKUP($A521,tbl_IBNRCount,2,0),0)"
            Else
                CType(rng.Columns(9), Range).Offset(0, 1).Formula =
                    "=IFERROR(INDEX(tbl_IBNRCount,MATCH($A521,tbl_IBNRCount[date],0), " &
                    "MATCH(""IBNRCnt2"",tbl_IBNRCount[#Headers],0)),0)"
            End If
            CType(CType(rng.Columns(9), Range).Offset(0, 1).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
                "=SUM(Count_GUIBNR)"
        Else
            CType(rng.Columns(6), Range).FormulaArray = lastTime
            CType(rng.Columns(7), Range).FormulaArray = defaultATA
            CType(rng.Columns(8), Range).FormulaArray = selATA
            CType(rng.Columns(9), Range).Formula =
            "=(" & shtName & "_CurAmt-" & shtName & "_Cap-" & shtName & "_Exclusion)*" & shtName & "_prior_ATU+" & shtName & "_Cap"
            CType(CType(rng.Columns(9), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
                "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_priorUlt))"

            CType(rng.Columns(10), Range).Formula =
                "=(" & shtName & "_CurAmt-" & shtName & "_Cap-" & shtName & "_Exclusion)*" & shtName & "_default_ATU+" & shtName & "_Cap"
            CType(CType(rng.Columns(10), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
                "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_defaultUlt))"
            CType(rng.Columns(11), Range).Formula =
                "=(" & shtName & "_CurAmt-" & shtName & "_Cap-" & shtName & "_Exclusion)*" & shtName & "_sel_ATU+" & shtName & "_Cap"
            CType(CType(rng.Columns(11), Range).Cells(rowNum, 1), Range).Offset(1, 0).Formula =
            "=SUM(INDEX(" & shtName & "_Summary,,column_" & shtName & "_summary_selUlt))"
        End If

        'Needs to assign nameOfRange to a range's name first
        nameOfRange = CType(wkst.Range(shtName & "_Summary").Name, Name)
        With nameOfRange
            .Name = shtName & "_Summary"
            .RefersTo = rng
        End With

    End Sub

    Public Sub expLoss()
        Application.ScreenUpdating = False
        Dim rng As Range
        Dim counter As Integer
        Dim tblSev As ListObject = wkstExpLoss.ListObjects("tbl_sev")

        If evalGroup = "Monthly" Then
            wkstExpLoss.Range("lookup_age").Value = 1
            counter = 12
        Else
            wkstExpLoss.Range("lookup_age").Value = 3
            counter = 4
        End If

        'default formula for trended off of values.
        wkstExpLoss.Range("O3").Formula = "=Index(summary, MATCH(lookup_age + 12, age, 0), column_preIC_sev)"
        wkstExpLoss.Range("O6").Formula = "=Index(summary, MATCH(lookup_age + 12, age, 0), column_preIC_pp)"
        wkstExpLoss.Range("O9").Formula = "=Index(summary, MATCH(lookup_age + 12, age, 0), column_preIC_lr)"

        For Each tbl As ListObject In wkstExpLoss.ListObjects
            rng = tbl.DataBodyRange.Resize(1) 'stores rng as the first row of the table
            tbl.DataBodyRange.Offset(1, 0).ClearContents() 'delete contents in all but first row
            tbl.Resize(rng) 'weird stuff..need to resize to 1 row for the table first
            tbl.Resize(rng.Resize(counter)) 'then resize to mthly/qtrly table
        Next

        For c As Integer = 1 To counter
            If counter = 12 Then
                CType(tblSev.ListColumns(1).Range.Cells(c), Range).Value = c
            Else
                If Month(CType(wkstControl.Range("CurrentEvalDate").Value, Date)) Mod 3 = 0 Then
                    '3 6 9 12
                    CType(tblSev.ListColumns(1).Range.Cells(c), Range).Value = 3 * c
                ElseIf Month(CType(wkstControl.Range("CurrentEvalDate").Value, Date)) Mod 3 = 1 Then
                    '1 4 7 10
                    CType(tblSev.ListColumns(1).Range.Cells(c), Range).Value = 3 * c - 2
                Else
                    '2 5 8 11
                    CType(tblSev.ListColumns(1).Range.Cells(c), Range).Value = 3 * c - 1
                End If
            End If
        Next
        Application.ScreenUpdating = True
    End Sub

    Public Sub graphsUpdate(wkstName As String)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(wkstName), Worksheet)

        For Each chartObj As ChartObject In CType(wkst.ChartObjects, ChartObjects)
            'charobject.chart.name is tricky -- it will show the activesheet's name 
            'in front of the chart name
            Dim sheetName As String = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet).Name
            Dim chartName As String = chartObj.Chart.Name.Substring(Len(sheetName) + 1)
            Dim tbl As ListObject = wkstExpLoss.ListObjects("tbl_" & chartName)
            Dim min, max, majorUnit As Double

            min = Double.MaxValue
            max = Double.MinValue

            'exclude month/quarter column
            For Each c As Range In tbl.DataBodyRange.Offset(0, 1).Resize(, 5)
                If CType(c.Value, Double) > 0 Then
                    min = Math.Min(min, CType(c.Value, Double))
                    max = Math.Max(max, CType(c.Value, Double))
                End If
            Next

            If max > 1000 Then
                max = UDFRoundUp(max, 100)
                min = UDFRoundDown(min, 100)
            ElseIf max < 1000 And max > 100 Then
                max = UDFRoundUp(max, 10)
                min = UDFRoundDown(min, 10)
            ElseIf max < 100 And max > 10 Then
                max = UDFRoundUp(max, 1)
                min = UDFRoundDown(min, 1)
            ElseIf max < 10 And max > 2 Then
                max = UDFRoundUp(max, 0.1)
                min = UDFRoundDown(min, 0.1)
            ElseIf max < 2 And max > 0 Then
                max = UDFRoundUp(max, 0.01)
                min = UDFRoundDown(min, 0.01)
            End If

            If min < 0 Then
                min = 0
            End If
            majorUnit = (max - min) / 10

            With CType(chartObj.Chart.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue), Axis)
                .MaximumScale = max
                .MinimumScale = min
                .MajorUnit = majorUnit
            End With
        Next
    End Sub
    Public Sub summary()
        'The column numbers will need to change whenever the Summary named range is changed
        Dim rng As Range
        Dim nameOfRange As Name
        Dim rowNum, offsetRows As Integer
        Dim namedRangeValues As Array = System.Enum.GetValues(GetType(namedRanges))

        rng = wkstSummary.Range("summary")
        rng.ClearContents()

        'this block resizes the named ranges On the summary tab, so the formulas below can work
        For Each name As namedRanges In namedRangeValues
            If evalGroup = "Monthly" Then
                rowNum = monthRowNum
            Else
                rowNum = quarterRowNum
            End If

            nameOfRange = CType(wkstSummary.Range(name.ToString()).Name, Name)
            rng = wkstSummary.Range(name.ToString()).Resize(rowNum)
            With nameOfRange
                .Name = name.ToString()
                .RefersTo = rng
            End With
        Next

        If evalGroup = "Monthly" Then
            rowNum = monthRowNum
            offsetRows = 12
        Else
            rowNum = quarterRowNum
            offsetRows = 4
        End If

        rng = wkstSummary.Range("summary")
        nameOfRange = CType(wkstSummary.Range("summary").Name, Name)
        rng = rng.Resize(rowNum)
        CType(rng.Columns(1), Range).Formula = "=Count!A521"
        CType(rng.Columns(2), Range).Formula = "=Count!B521"

        If evalGroup = "Monthly" Then
            CType(rng.Columns(3), Range).Formula = "=IFERROR(VLOOKUP(accident_date,tbl_epee,column_ep,0)/1000,0)"
            CType(rng.Columns(4), Range).Formula = "=IFERROR(VLOOKUP(accident_date,tbl_epee,column_ee,0)/365,0)"
            CType(rng.Columns(20), Range).Formula =
                "=If(SUM(clos_mod_spr_monthly)=0, 0, INDEX(clos_mod_ult_monthly,MATCH($D2,age,0),1))"
        Else
            CType(rng.Columns(3), Range).Formula = "=IFERROR(VLOOKUP(accident_date,tbl_epee_qtrly,column_ep,0)/1000,0)"
            CType(rng.Columns(4), Range).Formula = "=IFERROR(VLOOKUP(accident_date,tbl_epee_qtrly,column_ee,0)/365,0)"
            CType(rng.Columns(20), Range).Formula = "=IFERROR(INDEX(tbl_clsmod,MATCH(age,age,0),tbl_closmod_column)/1000,0)"
        End If

        CType(rng.Columns(5), Range).Formula = "=IFERROR(ep/ee*1000,0)"
        CType(rng.Columns(6), Range).Formula = "=VLOOKUP(accident_date,Count_Summary,column_count_summary_selULT,0)"
        CType(rng.Columns(7), Range).Formula = "=IFERROR(ult_counts/ee*1000,0)"
        CType(rng.Columns(8), Range).FormulaArray = "=Paid_CurAmt-Paid_Cap-Paid_Exclusion"
        CType(rng.Columns(9), Range).Formula = "=IFERROR(1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_selATU,0),0)"
        CType(rng.Columns(10), Range).FormulaArray = "=IFERROR(cur_paid/percent_paid,0)+Paid_Cap"
        CType(rng.Columns(11), Range).FormulaArray = "=Incurred_CurAmt-Incurred_Cap-Incurred_Exclusion"
        CType(rng.Columns(12), Range).Formula = "=IFERROR(1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_selATU,0),0)"
        CType(rng.Columns(13), Range).FormulaArray = "=IFERROR(cur_incurred/percent_incurred,0)+Incurred_Cap"

        'remove age 1 exp loss formula
        CType(rng.Columns(14), Range).Formula =
            "=IFERROR(INDEX(tbl_expLoss[exp_loss], MATCH(accident_date,tbl_expLoss[accident_date],0),1),0)"
        CType(rng.Columns(14), Range).End(XlDirection.xlDown).ClearContents()

        CType(rng.Columns(15), Range).Formula =
            "=ultLoss(""E"",proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss,0)"
        CType(rng.Columns(16), Range).Formula =
            "=ultLoss(""S"",proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss,0)"

        'letter selection column needs to be updated based on Paid/Incurred, A or H and B or G
        If projBase = "Paid" Then
            CType(rng.Columns(17), Range).Formula =
                "=If(percent_paid>(1/1.07), If(ult_paid>=cur_incurred, ""A"", ""H""), ""E"")"
        Else
            CType(rng.Columns(17), Range).Formula =
                "=If(percent_incurred>(1/1.07), ""B"", ""E"")"
        End If

        CType(rng.Columns(18), Range).Formula =
            "=ultLoss(letter,proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred," &
            "ult_incurred,exp_loss,IFERROR(VLOOKUP(accident_date,tbl_expLoss,5,0),0))"
        'age 1 exp loss doesn't use prior loss
        CType(rng.Columns(18), Range).End(XlDirection.xlDown).Formula =
            "=ultLoss(letter,proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss, 0)"

        CType(rng.Columns(19), Range).Formula = "=preIC_ultloss+volatility"

        CType(rng.Columns(21), Range).Formula = "=clos_mod*clos_mod_weight+(1-clos_mod_weight)*IC_ultloss"

        'BI needs special formula
        If coverageField = "BI" Then
            'any risk projections use this formula
            If CType(wkstControl.Range("risk").Value, String) <> "ALL" Then
                CType(rng.Columns(22), Range).Formula =
                "=If(YEAR(accident_date)<2007,ult_incurred,MAX(INDEX(cur_incurred,ROW()-1,1),INDEX(wtd_ultloss,ROW()-1,1)))"
            Else
                'otherwise use this formula (e.g., company projections)
                CType(rng.Columns(22), Range).Formula =
                "=MAX(INDEX(cur_incurred,ROW()-1,1),INDEX(wtd_ultloss,ROW()-1,1))"
            End If
        Else
            CType(rng.Columns(22), Range).Formula = "=IC_ultloss"
        End If

        'Development over prior, age 1 is blank
        CType(rng.Columns(23), Range).Formula =
            "=IFERROR(sel_ultloss-INDEX(tbl_expLoss[sel_ult_loss], MATCH(accident_date,tbl_expLoss[accident_date],0),1),0)"
        CType(rng.Columns(23), Range).End(XlDirection.xlDown).ClearContents()

        CType(rng.Columns(24), Range).Formula = "=IFERROR(sel_ultloss/ep,0)"
        CType(rng.Columns(25), Range).Formula = "=IFERROR(sel_ultloss/ult_counts*1000,0)"
        CType(rng.Columns(27), Range).Formula = "=IFERROR(sel_ultloss/ee*1000,0)"
        CType(rng.Columns(29), Range).Formula = "=If(age<=IC_spr_age,preIC_res/SUMIFS(preIC_res,age, ""<=""&IC_spr_age), 0)"
        CType(rng.Columns(30), Range).Formula = "=sel_volatility*preIC_res_spr"
        CType(rng.Columns(31), Range).Formula = "=preIC_ultloss-cur_paid"
        CType(rng.Columns(32), Range).Formula = "=cur_incurred-cur_paid"
        CType(rng.Columns(33), Range).Formula = "=sel_ultloss-cur_incurred"
        CType(rng.Columns(34), Range).Formula = "=sel_ultloss-cur_paid"
        CType(rng.Columns(40), Range).Formula = "=IFERROR(preIC_ultloss/ult_counts*1000,0)"
        CType(rng.Columns(41), Range).Formula = "=IFERROR(preIC_ultloss/ee*1000,0)"
        CType(rng.Columns(42), Range).Formula = "=IFERROR(preIC_ultloss/ep,0)"
        CType(rng.Columns(26), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(IC_sev, eval_group)"
        CType(rng.Columns(28), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(IC_pp, eval_group)"
        CType(rng.Columns(35), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(IC_lr, eval_group)"
        CType(rng.Columns(36), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(freq, eval_group)"
        CType(rng.Columns(37), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(avg_prem, eval_group)"
        CType(rng.Columns(38), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(ee, eval_group)"
        CType(rng.Columns(39), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(ep, eval_group)"
        CType(rng.Columns(43), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(preIC_sev, eval_group)"
        CType(rng.Columns(44), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(preIC_pp, eval_group)"
        CType(rng.Columns(45), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray =
            "=getTrend(preIC_lr, eval_group)"
        With nameOfRange
            .Name = "summary"
            .RefersTo = rng
        End With

    End Sub

    Public Sub QPageFormat()
        Dim dataRng As Range = wkstQPage.ListObjects("tbl_QPage").Range
        Dim endColumn As Integer = CType(dataRng.Columns(dataRng.Columns.Count), Range).Column

        For i As Integer = 1 To endColumn
            Dim c As Range = CType(wkstQPage.Cells(8, i), Range)
            c.EntireColumn.Hidden = False
            If CType(c.Value, String) = "(hide)" Then
                c.EntireColumn.Hidden = True
            End If
        Next

    End Sub
    Public Sub reviewTemplate()
        Dim rowCount As Integer = wkstReviewTemplate.Range("Z1").End(XlDirection.xlDown).Row
        'clear last time's track changes
        wkstReviewTemplate.Range("Z1:AE1").Offset(1, 0).Resize(rowCount - 1, 6).ClearContents()

        wkstReviewTemplate.Range("RT_selATA").ClearContents()

        'refill letter selection formulas
        wkstReviewTemplate.Range("RT_letterSel").Formula = "=INDEX(letter,ROWS(letter)-$H21,1)"

        'update 7-ult ATU
        CType(wkstReviewTemplate.Cells(9, 3), Range).Formula =
                "=INDEX(" & projBase & "_Summary,ROWS(" & projBase & "_Summary)-$A9,column_" & projBase & "_summary_priorATU)"
        CType(wkstReviewTemplate.Cells(9, 4), Range).Formula =
                "=INDEX(" & projBase & "_Summary,ROWS(" & projBase & "_Summary)-$A9,column_" & projBase & "_summary_defaultATU)"
        CType(wkstReviewTemplate.Cells(9, 6), Range).Formula =
                "=INDEX(" & projBase & "_Summary,ROWS(" & projBase & "_Summary)-$A9,column_" & projBase & "_summary_selATU)"
        'update 1-6 ATA
        If evalGroup = "Monthly" Then
            wkstReviewTemplate.Range("RT_priorATA").Formula = "=INDEX(" & projBase & "_lastTime_ATA,,$A10+1)"
            wkstReviewTemplate.Range("RT_defaultATA").Formula = "=INDEX(" & projBase & "_default_ATA,,$A10+1)"
        Else
            wkstReviewTemplate.Range("RT_priorATA").Formula = "=INDEX(" & projBase & "_lastTime_ATA_qtrly,,$A10+1)"
            wkstReviewTemplate.Range("RT_defaultATA").Formula = "=INDEX(" & projBase & "_default_ATA_qtrly,,$A10+1)"
        End If
    End Sub

    'this part updates prior ATA, default ATA in the reviewTemplate
    Public Sub reviewTemplate2()
        Dim summary As Range = wkstSummary.Range("summary")
        Dim reserves As Double
        If projBase = "Paid" Then
            'change paid ATU to prior ATU first, get the reserves using prior sel
            CType(summary.Columns(9), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_priorATU,0),0)"
            reserves = sumRange(CType(CType(summary.Columns(34), Range).Value, Object(,))) + priorRes
            wkstReviewTemplate.Range("C16").Value = reserves

            'change paid ATU to default ATU, get the reserves with default sel
            CType(summary.Columns(9), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_defaultATU,0),0)"
            reserves = sumRange(CType(CType(summary.Columns(34), Range).Value, Object(,))) + priorRes
            wkstReviewTemplate.Range("D16").Value = reserves

            'finally change paid ATU back to selected ATU
            CType(summary.Columns(9), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_selATU,0),0)"
        Else
            CType(summary.Columns(12), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_priorATU,0),0)"
            reserves = sumRange(CType(CType(summary.Columns(34), Range).Value, Object(,)))
            wkstReviewTemplate.Range("C16").Value = reserves + priorRes
            CType(summary.Columns(12), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_defaultATU,0),0)"
            reserves = sumRange(CType(CType(summary.Columns(34), Range).Value, Object(,)))
            wkstReviewTemplate.Range("D16").Value = reserves + priorRes
            CType(summary.Columns(12), Range).Formula =
                "=IFERROR(1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_selATU,0),0)"
        End If
    End Sub

    'Before going into PRP, update both initial and final ATA columns, so flag = 0
    'During PRP, we're just going to update the final ATA column, so make flag = 1
    Public Sub finalizeATA(flag As Integer)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(projBase), Worksheet)
        Dim row As Integer
        Dim dt As Date = Now

        'bring in the prior and default reserves in
        reviewTemplate2()

        Dim rng As Range
        Dim selATA As Object(,)

        If evalGroup = "Monthly" Then
            rng = wkst.Range(wkst.Name & "_sel_ATA").Resize(1, 6)
        Else
            rng = wkst.Range(wkst.Name & "_sel_ATA_qtrly").Resize(1, 6)
        End If

        'this part just converts a row of cells into a column of cells
        selATA = New Object(5, 0) {}

        For i As Integer = 0 To rng.Columns.Count - 1
            selATA(i, 0) = CType(rng.Cells(1, rng.Columns.Count - i), Range).Value
        Next

        'assigns the newly created column of cells to review template
        CType(wkstReviewTemplate.Cells(9, 5), Range).Value = CType(wkstReviewTemplate.Cells(9, 6), Range).Value
        For i As Integer = 0 To 5
            'only update the initial ATA column when flag is set to 0
            If flag = 0 Then
                CType(wkstReviewTemplate.Cells(10 + i, 5), Range).Value = selATA(i, 0)
                CType(wkstReviewTemplate.Cells(16, 5), Range).Value = CType(wkstReviewTemplate.Cells(16, 6), Range).Value
            End If
            CType(wkstReviewTemplate.Cells(10 + i, 6), Range).Value = selATA(i, 0)
        Next

        'track changes on the review template
        row = CType(wkstReviewTemplate.Cells(wkstReviewTemplate.Rows.Count, 26), Range).End(XlDirection.xlUp).Row + 1
        CType(wkstReviewTemplate.Cells(row, 26), Range).Value = projBase
        CType(wkstReviewTemplate.Cells(row, 27), Range).Value = rng.Resize(1, 1).Value
        CType(wkstReviewTemplate.Cells(row, 28), Range).Value = wkstExpLoss.Range("$P$11").Value
        CType(wkstReviewTemplate.Cells(row, 29), Range).Value =
                sumRange(CType(CType(wkstSummary.Range("summary").Columns(34), Range).Value, Object(,)))
        CType(wkstReviewTemplate.Cells(row, 30), Range).Value = dt
        CType(wkstReviewTemplate.Cells(row, 31), Range).Value =
            CType(Application.ActiveWorkbook.BuiltinDocumentProperties, DocumentProperties)("Last Author").Value
    End Sub

    Public Sub finalizeExpLoss(flag As Integer)
        Dim counter As Integer
        Dim wkst2 As Worksheet = CType(Application.ActiveWorkbook.Worksheets(projBase), Worksheet)
        Dim age1 As Object
        Dim row As Integer
        Dim dt As Date = Now


        'bring in the prior and default reserves in
        reviewTemplate2()

        If evalGroup = "Monthly" Then
            counter = 1
            age1 = wkst2.Range(projBase & "_sel_ATA").Resize(1, 1).Value
        Else
            counter = 3
            age1 = wkst2.Range(projBase & "_sel_ATA_qtrly").Resize(1, 1).Value
        End If

        If (CType(wkstExpLoss.Range("lookup_age").Value, Integer) = 1 And evalGroup = "Monthly") Or
            (CType(wkstExpLoss.Range("lookup_age").Value, Integer) = 3 And evalGroup = "Quarterly") Then
            wkstReviewTemplate.Range("RT_SevTrnd").Value = wkstExpLoss.Range("P3").Value
            wkstReviewTemplate.Range("RT_PPTrnd").Value = wkstExpLoss.Range("P6").Value
            wkstReviewTemplate.Range("RT_LRTrnd").Value = wkstExpLoss.Range("P9").Value
            wkstReviewTemplate.Range("RT_ExpLossAge1").Value = wkstExpLoss.Range("P11").Value

            'when flag is 0 which means it's before PRP, update the initial Exp Loss
            'when flag is 1, don't update
            If flag = 0 Then
                wkstReviewTemplate.Range("E27").Value = wkstExpLoss.Range("P11").Value
                wkstReviewTemplate.Range("E28").Value = wkstReviewTemplate.Range("F28").Value
            End If

            'track changes on Review Template
            row = CType(wkstReviewTemplate.Cells(wkstReviewTemplate.Rows.Count, 26), Range).End(XlDirection.xlUp).Row + 1
            CType(wkstReviewTemplate.Cells(row, 26), Range).Value = projBase
            CType(wkstReviewTemplate.Cells(row, 27), Range).Value = age1
            CType(wkstReviewTemplate.Cells(row, 28), Range).Value = wkstExpLoss.Range("P11").Value
            CType(wkstReviewTemplate.Cells(row, 29), Range).Value =
                sumRange(CType(CType(wkstSummary.Range("summary").Columns(34), Range).Value, Object(,)))
            CType(wkstReviewTemplate.Cells(row, 30), Range).Value = dt
            CType(wkstReviewTemplate.Cells(row, 31), Range).Value =
                CType(Application.ActiveWorkbook.BuiltinDocumentProperties, DocumentProperties)("Last Author").Value
        Else
            MsgBox("Can only bring Age " & counter & " expected loss to review template!")
            Exit Sub
        End If
    End Sub

    Public Sub finalizeGraphs()
        graphsUpdate("Exp Loss")
        graphsUpdate("Review Template")
    End Sub

    Public Sub inputReviewTemplateBeforePRP()
        finalizeATA(0)
        finalizeExpLoss(0)
    End Sub

    Public Sub inputReviewTemplateDuringPRP()
        finalizeATA(1)
        finalizeExpLoss(1)
    End Sub

    Public Sub adjustGraphLineColors()
        Dim colorForm As frmLineColor = New frmLineColor
        colorForm.Show()
    End Sub

    Private Function sumRange(ByVal rngToSum As Object(,)) As Double
        Dim out As Double = 0
        For i As Integer = 1 To rngToSum.GetUpperBound(0)
            out = out + CType(rngToSum(i, 1), Double)
        Next
        Return out
    End Function

    'A simple function to convert dates to quarterly dates (not very elegant...)
    Public Function convertDate(ByVal month As Integer, ByVal year As Integer, ByVal currentDate As Date) As Date
        Dim newDate As Date
        If evalGroup = "Monthly" Then
            newDate = New Date(year, month, Date.DaysInMonth(year, month))
        Else
            If currentDate.Month Mod 3 = 0 Then
                newDate = New Date(year, month * 3, Date.DaysInMonth(year, month * 3))
            ElseIf currentDate.Month Mod 3 = 1 Then
                newDate = New Date(year, month * 3 - 2, Date.DaysInMonth(year, month * 3 - 2))
            Else
                newDate = New Date(year, month * 3 - 1, Date.DaysInMonth(year, month * 3 - 1))
            End If
        End If

        Return newDate
    End Function

    'Converts a monthly triangle to a quarterly triangle
    Private Function quarterTriangle(ByVal monthTriangle As Object(,)) As Double(,)
        Dim out(Convert.ToInt32(monthTriangle.GetUpperBound(0) / 3) - 1,
                Convert.ToInt32(monthTriangle.GetUpperBound(1) / 3) - 1) As Double

        For i As Integer = 0 To out.GetUpperBound(0)
            For j As Integer = 0 To out.GetUpperBound(1) - i
                out(i, j) = CType(monthTriangle(3 * i + 2, 3 * j), Double) +
                            CType(monthTriangle(3 * i + 1, 3 * j + 1), Double) +
                            CType(monthTriangle(3 * i, 3 * j + 2), Double)
            Next
        Next
        Return out
    End Function

    'Round numbers up to the specified multiple
    Private Function UDFRoundUp(num As Double, multiple As Double) As Double
        If (multiple = 0) Then
            Return 0
        End If
        Return Math.Ceiling(num / multiple) * multiple
    End Function

    'Round numbers down to the specified multiple
    Private Function UDFRoundDown(num As Double, multiple As Double) As Double
        If (multiple = 0) Then
            Return 0
        End If
        Return Math.Floor(num / multiple) * multiple
    End Function

End Module
