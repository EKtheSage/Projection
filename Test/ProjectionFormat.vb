'Future add-ons to the module:
' 1. Apply different ATA factors based on number of working days (Good to have)
' 2. Conditional formatting on the triangle, show heatmap ??
' 3. Project payments as lower half of rectangle ??
' 4. Reserve Ranges ??
' 5. Initialize triangle objects for calculation behind the scene ??

Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel

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
    Public wkstReviewTemplate As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Review Template"), Worksheet)
    Public wkstData As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Data"), Worksheet)
    Public wkstIBNRCnt As Worksheet = CType(Application.ActiveWorkbook.Worksheets("GU IBNR Count"), Worksheet)
    Public wkstClsdAvg As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Closed Avg"), Worksheet)
    Public wkstIC As Worksheet = CType(Application.ActiveWorkbook.Worksheets("IC Volatility"), Worksheet)
    Public evalGroup As String = CType(wkstControl.Range("eval_group").Value, String)
    Public projBase As String = CType(wkstControl.Range("proj_base").Value, String)
    Public includeSS As String = CType(wkstControl.Range("include_ss").Value, String)

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
        sel_ultloss = 22
        IC_lr = 23
        IC_sev = 24
        IC_pp = 26
        preIC_res_spr = 28
        preIC_res = 30
        preIC_sev = 39
        preIC_pp = 40
        preIC_lr = 41
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

        monthToQuarterAlt("Incurred")
        monthToQuarterAlt("Paid")

        showDefaultTriangleView()
        expLoss()
        summary()
        reviewTemplate()

        'calculate all sheets, then turn calculation back to automatic
        Application.Calculate()
        Application.Calculation = XlCalculation.xlCalculationAutomatic

        wkstExpLoss.Activate()
        graphsUpdate("Exp Loss")
        graphsUpdate("Review Template")

    End Sub

    Public Sub monthToQuarterAlt(data As String)
        Application.Calculation = XlCalculation.xlCalculationManual
        Dim dataRng As String = "'Alt Data'!" & data & "_data_alt"

        Dim sht2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        Dim selectVal As Object(,) = CType(sht2.GetValue(), Object(,))
        Dim qtrTri As Double(,) = quarterTriangle(selectVal)

        dataRng = "'Alt Data'!" & data & "_qtrlydata_alt"
        Dim target2 As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, dataRng), ExcelReference)
        target2.SetValue(qtrTri)

        'also assigns new formula to the ATA block - Alt data equals actual data if ATA selections are the same, no need to have if/else
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(data), Worksheet)
        wkst.Range(data & "_ATA").FormulaArray = "=ATA(" & data & "_data_alt)"
        wkst.Range(data & "_qtrlyATA").FormulaArray = "=ATA(" & data & "_qtrlydata_alt)"

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
            hideRows = Application.Union(CType(wkst.Rows("2:157"), Range),
                                         CType(wkst.Rows("184:338"), Range),
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
            hideRows = Application.Union(CType(wkst.Rows("381:416"), Range),
                                        CType(wkst.Rows("443:477"), Range),
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
        'Dim accDateCol As Range
        'Dim fieldInfoArray() As Integer = New Integer() {1, XlColumnDataType.xlYMDFormat}

        Application.Calculation = XlCalculation.xlCalculationManual

        For Each wkst In Application.ActiveWorkbook.Worksheets
            If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
                If evalGroup = "Monthly" Then
                    month = False
                    quarter = True
                    hideRows = Application.Union(CType(wkst.Rows("2:157"), Range), CType(wkst.Rows("184:338"), Range))
                Else
                    month = True
                    quarter = False
                    hideRows = Application.Union(CType(wkst.Rows("381:416"), Range), CType(wkst.Rows("443:477"), Range))
                End If
                'CType(wkst.Columns("Z:FX"), Range).EntireColumn.Hidden = True

                'accDateCol = wkst.Range(wkst.Name & "_data").Offset(0, -1).Resize(180, 1)
                'accDateCol.TextToColumns(Destination:=accDateCol, FieldInfo:=fieldInfoArray)
                'wkst.Activate()
                'MsgBox("does it work?")


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
        Dim rng As Range
        Dim dateRng, dataRng, lastTime, defaultATA, selATA As String
        Dim rowNum, counter As Integer
        Dim nameOfRange As Name

        Application.Calculation = XlCalculation.xlCalculationManual

        rng = wkst.Range(shtName & "_Summary")
        nameOfRange = CType(rng.Name, Name)
        rng.ClearContents()
        CType(rng.Columns(10), Range).Offset(0, 1).ClearContents()

        If evalGroup = "Monthly" Then
            dateRng = "=accident_date_mthly"
            dataRng = shtName & "_data"
            lastTime = "=ATU(" & shtName & "_lastTime_ATA)"
            defaultATA = "= ATU(" & shtName & "_default_ATA)"
            selATA = "=ATU(" & shtName & "_sel_ATA)"
            rowNum = monthRowNum
            counter = 1
        Else
            dateRng = "=accident_date_qtrly"
            dataRng = shtName & "_qtrlydata"
            lastTime = "=ATU(" & shtName & "_lastTime_ATA_qtrly)"
            defaultATA = "=ATU(" & shtName & "_default_ATA_qtrly)"
            selATA = "=ATU(" & shtName & "_sel_ATA_qtrly)"
            rowNum = quarterRowNum
            counter = 3
        End If

        rng = rng.Resize(rowNum, 10)
        CType(rng.Columns(1), Range).FormulaArray = dateRng

        For i As Integer = 1 To rowNum
            CType(rng.Cells(i, 2), Range).Value =
                DateDiff("m", CType(rng.Cells(i, 1), Range).Value, wkstControl.Range("CurrentEvalDate").Value) + counter
            CType(rng.Cells(i, 3), Range).Value = CType(wkst.Range(dataRng).Cells(i, rowNum - i + 1), Range).Value
        Next

        CType(rng.Columns(5), Range).FormulaArray = lastTime
        CType(rng.Columns(6), Range).FormulaArray = defaultATA
        CType(rng.Columns(7), Range).FormulaArray = selATA
        CType(rng.Columns(8), Range).Formula = "=$C521*(E521-$D521)+$D521"
        CType(rng.Columns(9), Range).Formula = "=$C521*(F521-$D521)+$D521"

        If shtName = "Count" Then
            CType(rng.Columns(10), Range).Formula = "=If($K521="""",$C521*(G521-$D521)+$D521,$K521+$C521)"
            CType(rng.Columns(10), Range).Offset(0, 1).Formula = "=IFERROR(VLOOKUP($A521,tbl_IBNRCount,2,0),0)"
        Else
            CType(rng.Columns(10), Range).Formula = "=$C521*(G521-$D521)+$D521"
        End If

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

        For Each tbl As ListObject In wkstExpLoss.ListObjects
            rng = tbl.DataBodyRange.Resize(1) 'stores rng as the first row of the table
            tbl.DataBodyRange.Offset(1, 0).ClearContents() 'delete contents in all but first row
            tbl.Resize(rng) 'weird stuff..need to resize to 1 row for the table first
            tbl.Resize(rng.Resize(counter)) 'then resize to mthly/qtrly table
        Next

        For c As Integer = 1 To counter
            CType(tblSev.ListColumns(1).Range.Cells(c), Range).Value = c
        Next
        Application.ScreenUpdating = True
    End Sub

    Public Sub graphsUpdate(wkstName As String)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(wkstName), Worksheet)
        For Each chartObj As ChartObject In CType(wkst.ChartObjects, ChartObjects)
            Dim sheetName As String = "Exp Loss"
            Dim chartName As String = chartObj.Chart.Name.Substring(Len(sheetName) + 1)
            Dim tbl As ListObject = wkstExpLoss.ListObjects("tbl_" & chartName)
            Dim min, max, majorUnit As Double

            min = Double.MaxValue
            max = Double.MinValue

            For Each c As Range In tbl.DataBodyRange.Offset(0, 1).Resize(, 5)
                If CType(c.Value, Double) > 0 Then
                    min = Math.Min(min, CType(c.Value, Double))
                    max = Math.Max(max, CType(c.Value, Double))
                End If
            Next

            If max > 1000 Then
                max = RoundUp(max, 1000)
                min = RoundUp(min, 1000) - 2000
            ElseIf max < 1000 And max > 100 Then
                max = RoundUp(max, 100)
                min = RoundUp(min, 100) - 200
            ElseIf max < 100 And max > 10 Then
                max = RoundUp(max, 5)
                min = RoundUp(min, 5) - 10
            ElseIf max < 10 And max > 1 Then
                max = Math.Round(max, 2)
                min = Math.Round(min, 2) - 0.5
            ElseIf max < 1 And max > 0 Then
                max = Math.Round(max, 2)
                min = Math.Round(min, 2) - 0.125
            End If

            If min < 0 Then
                min = 0
            End If
            majorUnit = (max - min) / 10

            With CType(chartObj.Chart.Axes(XlAxisType.xlValue), Axis)
                .MaximumScale = max
                .MinimumScale = min
                .MajorUnit = majorUnit
            End With
        Next
    End Sub
    Public Sub summary()
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
        CType(rng.Columns(3), Range).Formula = "=VLOOKUP(accident_date,tbl_epee,column_ep,0)"
        CType(rng.Columns(4), Range).Formula = "=VLOOKUP(accident_date,tbl_epee,column_ee,0)"
        CType(rng.Columns(5), Range).Formula = "=ep/ee"
        CType(rng.Columns(6), Range).Formula = "=VLOOKUP(accident_date,Count_Summary,column_count_summary_selULT,0)"
        CType(rng.Columns(7), Range).Formula = "=ult_counts/ee*1000"
        CType(rng.Columns(8), Range).Formula = "=VLOOKUP(accident_date,Paid_Summary,column_paid_summary_curAmt,0)"
        CType(rng.Columns(9), Range).Formula = "=1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_selATU,0)"
        CType(rng.Columns(10), Range).Formula = "=cur_paid/percent_paid"
        CType(rng.Columns(11), Range).Formula = "=VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_curAmt,0)"
        CType(rng.Columns(12), Range).Formula = "=1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_selATU,0)"
        CType(rng.Columns(13), Range).Formula = "=cur_incurred/percent_incurred"

        CType(rng.Columns(14), Range).Formula = "=VLOOKUP(accident_date, ExpLoss,2,0)"
        'remove age 1 exp loss formula
        CType(rng.Columns(14), Range).End(XlDirection.xlDown).ClearContents()

        CType(rng.Columns(15), Range).Formula =
            "=ultLoss(""E"",proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss,0)"
        CType(rng.Columns(16), Range).Formula =
            "=ultLoss(""S"",proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss,0)"

        'letter selection column needs to be updated based on Paid/Incurred
        If projBase = "Paid" Then
            CType(rng.Columns(17), Range).Formula = "=IF(percent_paid>0.935, ""A"", ""E"")"
        Else
            CType(rng.Columns(17), Range).Formula = "=IF(percent_incurred>0.935, ""B"", ""E"")"
        End If

        CType(rng.Columns(18), Range).Formula =
            "=ultLoss(letter,proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss,VLOOKUP(accident_date,ExpLoss,5,0))"
        'age 1 exp loss doesn't use prior loss
        CType(rng.Columns(18), Range).End(XlDirection.xlDown).Formula =
            "=ultLoss(letter,proj_base,cur_paid,percent_paid,ult_paid,cur_incurred,percent_incurred,ult_incurred,exp_loss, 0)"

        CType(rng.Columns(19), Range).Formula = "=preIC_ultloss+volatility"
        'rng.Columns(20).Formula
        CType(rng.Columns(21), Range).Formula = "=clos_mod*clos_mod_weight+(1-clos_mod_weight)*IC_ultloss"
        CType(rng.Columns(22), Range).Formula = "=IC_ultloss"
        CType(rng.Columns(23), Range).Formula = "=sel_ultloss/ep*1000"
        CType(rng.Columns(24), Range).Formula = "=sel_ultloss/ult_counts*1000"
        CType(rng.Columns(26), Range).Formula = "=sel_ultloss/ee*1000"
        CType(rng.Columns(28), Range).Formula = "=IF(age<IC_spr_age,preIC_res/SUMIFS(preIC_res,age, ""<""&IC_spr_age), 0)"
        CType(rng.Columns(29), Range).Formula = "=sel_volatility*preIC_res_spr"
        CType(rng.Columns(30), Range).Formula = "=preIC_ultloss-cur_paid"
        CType(rng.Columns(31), Range).Formula = "=cur_incurred-cur_paid"
        CType(rng.Columns(32), Range).Formula = "=sel_ultloss-cur_incurred"
        CType(rng.Columns(33), Range).Formula = "=sel_ultloss-cur_paid"
        CType(rng.Columns(39), Range).Formula = "=preIC_ultloss/ult_counts*1000"
        CType(rng.Columns(40), Range).Formula = "=preIC_ultloss/ee*1000"
        CType(rng.Columns(41), Range).Formula = "=preIC_ultloss/ep*1000"
        CType(rng.Columns(25), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(IC_sev, eval_group)"
        CType(rng.Columns(27), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(IC_pp, eval_group)"
        CType(rng.Columns(34), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(IC_lr, eval_group)"
        CType(rng.Columns(35), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(freq, eval_group)"
        CType(rng.Columns(36), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(avg_prem, eval_group)"
        CType(rng.Columns(37), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(ee, eval_group)"
        CType(rng.Columns(38), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(ep, eval_group)"
        CType(rng.Columns(42), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(preIC_sev, eval_group)"
        CType(rng.Columns(43), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(preIC_pp, eval_group)"
        CType(rng.Columns(44), Range).Offset(offsetRows, 0).Resize(rng.Rows.Count - offsetRows).FormulaArray = "=getTrend(preIC_lr, eval_group)"
        With nameOfRange
            .Name = "summary"
            .RefersTo = rng
        End With

    End Sub

    Public Sub reviewTemplate()
        Dim summary As Range = wkstSummary.Range("summary")
        Dim reserves As Double

        Dim rowCount As Integer = wkstReviewTemplate.Range("Y1").End(XlDirection.xlDown).Row
        'clear last time's track changes
        wkstReviewTemplate.Range("Y1:AD1").Offset(1, 0).Resize(rowCount - 1, 6).ClearContents()

        wkstReviewTemplate.Range("RT_selATA").ClearContents()
        If evalGroup = "Monthly" Then
            wkstReviewTemplate.Range("RT_priorATA").Formula = "=INDEX(" & projBase & "_lastTime_ATA,,$A10+1)"
            wkstReviewTemplate.Range("RT_defaultATA").Formula = "=INDEX(" & projBase & "_default_ATA,,$A10+1)"
        Else
            wkstReviewTemplate.Range("RT_priorATA").Formula = "=INDEX(" & projBase & "_lastTime_ATA_qtrly,,$A10+1)"
            wkstReviewTemplate.Range("RT_defaultATA").Formula = "=INDEX(" & projBase & "_default_ATA_qtrly,,$A10+1)"
        End If

        If projBase = "Paid" Then
            'change paid ATU to prior first, get the reserves using prior sel
            CType(summary.Columns(9), Range).Formula =
                "=1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_priorATU,0)"
            reserves = sumRange(CType(CType(summary.Columns(33), Range).Value, Object(,)))
            wkstReviewTemplate.Range("C16").Value = reserves

            'change paid ATU to default ATU, get the reserves with default sel
            CType(summary.Columns(9), Range).Formula =
                "=1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_defaultATU,0)"
            reserves = sumRange(CType(CType(summary.Columns(33), Range).Value, Object(,)))
            wkstReviewTemplate.Range("D16").Value = reserves

            'finally change paid ATU to selected ATU
            CType(summary.Columns(9), Range).Formula =
                "=1/VLOOKUP(accident_date,Paid_Summary,column_paid_summary_selATU,0)"
        Else
            CType(summary.Columns(12), Range).Formula =
                "=1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_priorATU,0)"
            reserves = sumRange(CType(CType(summary.Columns(33), Range).Value, Object(,)))
            wkstReviewTemplate.Range("C16").Value = reserves
            CType(summary.Columns(12), Range).Formula =
                "=1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_defaultATU,0)"
            reserves = sumRange(CType(CType(summary.Columns(33), Range).Value, Object(,)))
            wkstReviewTemplate.Range("D16").Value = reserves
            CType(summary.Columns(12), Range).Formula =
                "=1/VLOOKUP(accident_date,Incurred_Summary,column_incurred_summary_selATU,0)"
        End If

    End Sub

    Public Sub finalizeATA()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet)
        Dim rng As Range
        Dim selATA As Object(,)

        If wkst.Name = projBase Then
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
                CType(wkstReviewTemplate.Cells(10 + i, 5), Range).Value = selATA(i, 0)
                CType(wkstReviewTemplate.Cells(10 + i, 6), Range).Value = selATA(i, 0)
            Next
            CType(wkstReviewTemplate.Cells(16, 5), Range).Value = CType(wkstReviewTemplate.Cells(16, 6), Range).Value
        ElseIf wkst.Name <> projBase And (wkst.Name = "Paid" Or wkst.Name = "Incurred") Then
            MsgBox("Your projection does not use " & wkst.Name & " ATA factors for estimating the reserves!")
        Else
            MsgBox("You are on the wrong tab!")
        End If
    End Sub

    Public Sub finalizeExpLoss()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.ActiveSheet, Worksheet)
        Dim counter As Integer

        If wkst.Name <> "Exp Loss" Then
            Exit Sub
        End If

        If (CType(wkst.Range("lookup_age").Value, Integer) = 1 And evalGroup = "Monthly") Or
            (CType(wkst.Range("lookup_age").Value, Integer) = 3 And evalGroup = "Quarterly") Then
            wkstReviewTemplate.Range("RT_SevTrnd").Value = wkst.Range("P3").Value
            wkstReviewTemplate.Range("RT_PPTrnd").Value = wkst.Range("P6").Value
            wkstReviewTemplate.Range("RT_LRTrnd").Value = wkst.Range("P9").Value
            wkstReviewTemplate.Range("E27").Value = wkst.Range("P11").Value
            wkstReviewTemplate.Range("RT_ExpLossAge1").Value = wkst.Range("P11").Value
            wkstReviewTemplate.Range("E28").Value = wkstReviewTemplate.Range("F28").Value
        Else
            If evalGroup = "Monthly" Then
                counter = 1
            Else
                counter = 3
            End If
            MsgBox("Can only bring Age " & counter & " expected loss to review template!")
            Exit Sub
        End If
    End Sub

    Public Sub finalizeGraphs()
        graphsUpdate("Exp Loss")
        graphsUpdate("Review Template")
    End Sub

    Public Function sumRange(ByVal rngToSum As Object(,)) As Double
        Dim out As Double = 0
        For i As Integer = 1 To rngToSum.GetUpperBound(0)
            out = out + CType(rngToSum(i, 1), Double)
        Next
        Return out
    End Function

    Public Sub highlightBorderTriangle()
        Dim wkst As Worksheet
        Dim rng As Range

        For Each wkst In Application.ActiveWorkbook.Worksheets
            If wkst.Name = "Count" Or wkst.Name = "Paid" Or wkst.Name = "Incurred" Then
                rng = wkst.Range(wkst.Name & "_data")
                For i As Integer = 1 To rng.Rows.Count
                    For j As Integer = 1 To rng.Columns.Count

                    Next
                Next
            End If
        Next
    End Sub
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
    Private Function RoundUp(num As Double, multiple As Integer) As Integer

        If (multiple = 0) Then
            Return 0
        End If
        Dim add As Integer
        add = CType(multiple / Math.Abs(multiple), Integer)
        Return CType(((num + multiple - add) / multiple), Integer) * multiple
    End Function

End Module
