Imports ExcelDna.Integration
Imports ExcelDna.Integration.XlCall
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Excel

'One day could think about having one function to evaluate all different algorithms, the structures are really similar
'Module means the class is static, only one instance is needed. Any functions inside will not need to have shared keyword
Public Module ChainLadder
    'ExcelReference + XlCall
    'In a UDF, use XlCall(C API); in a macro, use COM interface(Office Interop)
    'Calculate ATU, convert to %, to me it seems easier to understand.
    'One-Dimension array are rows in Excel
    'ATA Algorithms with number of points, populate the range of estimes
    Enum ATAType
        WeightedAverage = 1
        SimpleAverage = 2
        LeastSquare = 3
        HighLow = 4
        Seasonal = 5
    End Enum

    Public Function ultLoss(ByVal letter As String, ByVal projBase As String,
                            ByVal curPaid As Double, ByVal percPaid As Double, ByVal ultPaid As Double,
                            ByVal curInc As Double, ByVal percInc As Double, ByVal ultInc As Double,
                            ByVal expLoss As Double, ByVal priorSel As Double) As Double
        Dim bf, gb As Double
        If projBase = "Incurred" Then
            bf = curInc + (1 - percInc) * expLoss
            gb = curInc + (1 - percInc) * bf
        ElseIf projBase = "Paid" Then
            bf = curPaid + (1 - percPaid) * expLoss
            gb = curPaid + (1 - percPaid) * bf
        End If

        Select Case letter
            Case "A"
                ultLoss = Decimal.Round(CType(ultPaid, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "B"
                ultLoss = Decimal.Round(CType(ultInc, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "C"
                ultLoss = Decimal.Round(CType((2 * ultInc + ultPaid) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "D"
                ultLoss = Decimal.Round(CType((ultInc + 2 * ultPaid) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "E"
                ultLoss = Decimal.Round(CType(bf, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "F"
                ultLoss = Decimal.Round(CType(priorSel, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "G"
                ultLoss = Decimal.Round(CType((ultPaid + ultInc) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "H"
                ultLoss = Decimal.Round(CType(curInc, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "I"
                ultLoss = Decimal.Round(CType((priorSel + ultPaid) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "J"
                ultLoss = Decimal.Round(CType((priorSel + ultInc) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "K"
                ultLoss = Decimal.Round(CType((priorSel + 2 * ultInc) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "L"
                ultLoss = Decimal.Round(CType((priorSel + bf) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "M"
                ultLoss = Decimal.Round(CType((priorSel + 2 * bf) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "P"
                ultLoss = Decimal.Round(CType((2 * priorSel + bf) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "Q"
                ultLoss = Decimal.Round(CType((3 * priorSel + bf) / 4, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "S"
                ultLoss = Decimal.Round(CType(gb, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "U"
                ultLoss = Decimal.Round(CType(expLoss, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "V"
                ultLoss = Decimal.Round(CType((expLoss + bf) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "W"
                ultLoss = Decimal.Round(CType((ultInc + bf) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "X"
                ultLoss = Decimal.Round(CType((ultPaid + bf) / 2, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "Y"
                ultLoss = Decimal.Round(CType((2 * priorSel + ultInc) / 3, Decimal), 6, MidpointRounding.AwayFromZero)
            Case "Z"
                ultLoss = Decimal.Round(CType((9 * ultInc + ultPaid) / 10, Decimal), 6, MidpointRounding.AwayFromZero)
        End Select

        Return ultLoss
    End Function

    'Trend function based on monthly or quarterly file
    Public Function getTrend(ByVal data() As Object, group As String) As Double(,)
        Dim indicator As Integer
        Dim out(,) As Double

        If group = "Monthly" Then
            indicator = 12
        ElseIf group = "Quarterly" Then
            indicator = 4
        End If

        out = New Double(data.GetUpperBound(0) - indicator, 0) {}
        For i As Integer = 0 To out.GetUpperBound(0)
            out(i, 0) = CType(data(i + indicator), Double) / CType(data(i), Double) - 1
        Next
        Return out
    End Function
    'Takes in a triangle range, convert it to array, then do ATA calculation
    Public Function ATA(ByVal triangle(,) As Object) As Double(,)
        'first need to create an ATA array, which has dimensions of the triangle -1
        Dim out(triangle.GetUpperBound(0) - 1, triangle.GetUpperBound(1) - 1) As Double

        For i As Integer = 0 To out.GetUpperBound(0)
            For j As Integer = 0 To out.GetUpperBound(1) - i
                'cannot allow division by 0

                If CType(triangle(i, j + 1), Double) = 0 Or CType(triangle(i, j), Double) = 0 Or
                    (CType(triangle(i, j), Double) > 0 And CType(triangle(i, j), Double) < 0.00001) Then
                    out(i, j) = 1
                Else
                    out(i, j) = Decimal.Round(
                        CType(CType(triangle(i, j + 1), Double) / CType(triangle(i, j), Double), Decimal),
                        4, MidpointRounding.AwayFromZero)
                End If
            Next
        Next
        Return out
    End Function

    'Takes in an ATA range, convert it to ATU
    Public Function ATU(ByVal ATA() As Object) As Double(,)
        Dim out(ATA.GetUpperBound(0) + 1, 0) As Double
        Dim out2(ATA.GetUpperBound(0) + 1, 0) As Double
        out(out.GetUpperBound(0), 0) = 1
        For i As Integer = out.GetUpperBound(0) To 1 Step -1
            out(i - 1, 0) = Decimal.Round(CType(out(i, 0) * CType(ATA(i - 1), Double), Decimal),
                                       4, MidpointRounding.AwayFromZero)
        Next

        'need to reverse the ATU...not good
        For i As Integer = 0 To out.GetUpperBound(0)
            out2(i, 0) = out(out.GetUpperBound(0) - i, 0)
        Next
        Return out2
    End Function

    'split string into type and numPt
    Public Function parseAlgorithm(ByVal algorithm As String, ByVal index As Integer) As Integer
        Dim stringList As Dictionary(Of Integer, String) = New Dictionary(Of Integer, String)
        stringList.Add(1, Regex.Replace(algorithm, "[\d]", ""))
        stringList.Add(2, Regex.Replace(algorithm, "[^\d]", ""))
        Select Case index
            Case 1
                Select Case stringList.Item(index)
                    Case "W"
                        Return ATAType.WeightedAverage
                    Case "A"
                        Return ATAType.SimpleAverage
                    Case "LS"
                        Return ATAType.LeastSquare
                    Case "H"
                        Return ATAType.HighLow
                    Case Else
                        Return ATAType.Seasonal
                End Select
            Case Else
                Return Convert.ToInt32(stringList.Item(index))
        End Select
    End Function

    'this function returns an array of ATA factors for a specified algorithm and number of points
    Public Function ATAAlg(ByVal algType As Integer, ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()
        Dim out(triangle.GetUpperBound(1) - 1) As Double

        Select Case algType
            Case ATAType.WeightedAverage
                out = WtdAvg(numPt, triangle)
            Case ATAType.SimpleAverage
                out = SimAvg(numPt, triangle)
            Case ATAType.LeastSquare
                out = LstSqr(numPt, triangle)
            Case ATAType.HighLow
                out = HighLow(numPt, triangle)
            Case ATAType.Seasonal
                out = seasonal(numPt, triangle)
            Case Else
        End Select

        Return out
    End Function

    'this function returns the ATA factor at specified algorithm, number of points, and age
    Public Function AtaWhichAge(ByVal algorithm As String, ByVal age As Integer,
                                ByVal triangle(,) As Object) As Double

        Dim alg As String = Regex.Replace(algorithm, "[\d]", "")
        Dim algType As Integer
        Select Case alg
            Case "W"
                algType = ATAType.WeightedAverage
            Case "A"
                algType = ATAType.SimpleAverage
            Case "LS"
                algType = ATAType.LeastSquare
            Case "H"
                algType = ATAType.HighLow
            Case "S"
                algType = ATAType.Seasonal
        End Select
        Dim numPt As Integer = Integer.Parse(Regex.Replace(algorithm, "[^\d]", ""))

        Return ATAAlg(algType, numPt, triangle)(age - 1)
    End Function

    Public Function WtdAvg(ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()

        'Create a 1-Dimension ATA array, length equal to number of columns in the triangle - 1
        Dim out(triangle.GetUpperBound(1) - 1) As Double
        Dim thisAge As Double
        Dim nextAge As Double
        Dim counter As Integer

        For j As Integer = 0 To triangle.GetUpperBound(1) - 1
            thisAge = 0
            nextAge = 0

            'let 0 = using all available points
            If numPt = 0 Then
                numPt = triangle.GetUpperBound(0) - j
            End If

            'if number of rows are less than specified points, use all available points
            If triangle.GetUpperBound(0) - j < numPt Then
                counter = numPt - (triangle.GetUpperBound(0) - j)
            Else
                counter = 0
            End If

            For i As Integer = triangle.GetUpperBound(0) - j - numPt + counter To triangle.GetUpperBound(0) - j - 1
                thisAge = thisAge + CType(triangle(i, j), Double)
                nextAge = nextAge + CType(triangle(i, j + 1), Double)
            Next

            'force ATA to 1 if it's less than 1
            If nextAge < thisAge Then
                out(j) = 1
            Else
                out(j) = Decimal.Round(CType(nextAge / thisAge, Decimal), 4, MidpointRounding.AwayFromZero)
            End If
        Next
        Return out

    End Function

    Public Function SimAvg(ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()

        Dim ATATri(triangle.GetUpperBound(0) - 1, triangle.GetUpperBound(1) - 1) As Double
        Dim out(triangle.GetUpperBound(1) - 1) As Double
        Dim ATASum As Double
        Dim counter As Integer
        ATATri = ATA(triangle)

        For j As Integer = 0 To ATATri.GetUpperBound(1)
            ATASum = 0
            'let 0 = using all available points
            If numPt = 0 Then
                numPt = triangle.GetUpperBound(0) - j
            End If

            'if number of rows are less than specified points, use all available points
            If triangle.GetUpperBound(0) - j < numPt Then
                counter = numPt - (triangle.GetUpperBound(0) - j)
            Else
                counter = 0
            End If

            For i As Integer = ATATri.GetUpperBound(0) - j - numPt + 1 + counter To ATATri.GetUpperBound(0) - j
                ATASum = ATASum + ATATri(i, j)
            Next

            ATASum = Decimal.Round(CType(ATASum, Decimal), 4, MidpointRounding.AwayFromZero)

            'force ATA to 1 if it's less than 1
            If ATASum < (numPt - counter) Then
                out(j) = 1
            Else
                out(j) = Decimal.Round(CType(ATASum / (numPt - counter), Decimal), 4, MidpointRounding.AwayFromZero)
            End If
        Next
        Return out
    End Function

    Public Function LstSqr(ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()
        'y = b x + a
        'b = (avg(xy) - avg(x)avg(y)) / (avg(x^2) - avg(x)^2)
        'b = (n * sum(xy) - sum(x)*sum(y)) / (n*sum(xx) - sum(x)sum(x))
        'a = avg(y) - b * avg(x)

        Dim x As Double, y As Double, a As Double, b As Double
        Dim xy As Double, xSqr As Double, nextAge As Double
        Dim counter As Integer
        Dim out(triangle.GetUpperBound(1) - 1) As Double

        For j As Integer = 0 To triangle.GetUpperBound(1) - 1
            a = 0
            b = 0
            x = 0
            y = 0
            xy = 0
            xSqr = 0
            nextAge = 0

            If numPt = 0 Then
                numPt = triangle.GetUpperBound(0) - j
            End If

            If triangle.GetUpperBound(0) - j < numPt Then
                counter = numPt - (triangle.GetUpperBound(0) - j)
            Else
                counter = 0
            End If

            For i As Integer = triangle.GetUpperBound(0) - j - numPt + counter To triangle.GetUpperBound(0) - j - 1
                x = x + CType(triangle(i, j), Double)
                y = y + CType(triangle(i, j + 1), Double)
                xy = xy + CType(triangle(i, j), Double) * CType(triangle(i, j + 1), Double)
                xSqr = xSqr + CType(triangle(i, j), Double) ^ 2
            Next
            b = ((numPt - counter) * xy - x * y) / ((numPt - counter) * xSqr - x ^ 2)
            a = (y / (numPt - counter)) - (b * (x / (numPt - counter)))
            nextAge = b * CType(triangle(triangle.GetUpperBound(0) - j, j), Double) + a

            'boundary condition - set factor to 1
            If triangle.GetUpperBound(0) - j - 1 < 2 Then
                out(j) = 1
            ElseIf nextAge < CType(triangle(triangle.GetUpperBound(0) - j, j), Double)
                out(j) = 1
            Else
                out(j) = Decimal.Round(
                    CType(nextAge / CType(triangle(triangle.GetUpperBound(0) - j, j), Double), Decimal),
                    4, MidpointRounding.AwayFromZero)
            End If
        Next
        Return out
    End Function

    Public Function HighLow(ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()
        Dim out(triangle.GetUpperBound(1) - 1) As Double
        Dim ATATri(triangle.GetUpperBound(0) - 1, triangle.GetUpperBound(1) - 1) As Double
        Dim ATAMax As Double
        Dim ATAMin As Double
        Dim ATASum As Double
        Dim counter As Integer
        ATATri = ATA(triangle)

        For j As Integer = 0 To ATATri.GetUpperBound(1)
            ATASum = 0
            ATAMax = Double.MinValue
            ATAMin = Double.MaxValue
            'let 0 = using all available points
            If numPt = 0 Then
                numPt = triangle.GetUpperBound(0) - j
            End If

            'if number of rows are less than specified points, use all available points
            If triangle.GetUpperBound(0) - j < numPt Then
                counter = numPt - (triangle.GetUpperBound(0) - j)
            Else
                counter = 0
            End If

            For i As Integer = ATATri.GetUpperBound(0) - j - numPt + 1 + counter To ATATri.GetUpperBound(0) - j
                ATASum = ATASum + ATATri(i, j)
                ATAMax = Math.Max(ATAMax, ATATri(i, j))
                ATAMin = Math.Min(ATAMin, ATATri(i, j))
            Next

            ATASum = Decimal.Round(CType(ATASum, Decimal), 4, MidpointRounding.AwayFromZero)
            If ATASum < (numPt - counter - 2) Then
                out(j) = 1
            Else
                out(j) = Decimal.Round(
                    CType((ATASum - ATAMax - ATAMin) / (numPt - counter - 2), Decimal),
                    4, MidpointRounding.AwayFromZero)
            End If

            'manually assign last 2 points to be 1
            If j > ATATri.GetUpperBound(1) - 2 Then
                out(j) = 1
            End If
        Next
        Return out
    End Function

    Public Function seasonal(ByVal numPt As Integer, ByVal triangle(,) As Object) As Double()
        Dim out(triangle.GetUpperBound(1) - 1) As Double

        Return out
    End Function

    Public Function completeRectangle(ByVal triangle(,) As Object, ByVal ataFactors() As Object) As Double(,)
        Dim out(triangle.GetUpperBound(0), triangle.GetUpperBound(1)) As Double

        For i As Integer = 0 To triangle.GetUpperBound(0)
            For j As Integer = 0 To triangle.GetUpperBound(1)

            Next
        Next
        Return out
    End Function

    Private Function ReferenceToRange(ByVal xlRef As ExcelReference) As Object
        Dim strAddress As String = CType(Excel(XlCall.xlfReftext, xlRef, True), String)

        ReferenceToRange = CType(ExcelDnaUtil.Application, Application).Range(strAddress)
    End Function
End Module
