Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration

'The goal of Triangle Class is to store triangle values in an object. 
'Do calculation in .net, bring result back to excel, so it would be faster.
Public Class Triangle
    Private name As String
    Private data As Double(,)
    Private ATA As Double(,)
    Private ATU As Double(,)

    Public Sub New(ByVal inputData As Object(,), ByVal inputName As String)
        'this part is tricky, inputData is coming from excel (1-based array)
        'data is a .net array (0-based array)
        data = New Double(inputData.GetUpperBound(0) - 1, inputData.GetUpperBound(1) - 1) {}
        For i As Integer = 0 To data.GetUpperBound(0)
            For j As Integer = 0 To data.GetUpperBound(1)
                data(i, j) = CType(inputData(i + 1, j + 1), Double)
            Next
        Next

        ATA = New Double(data.GetUpperBound(0) - 1, data.GetUpperBound(1) - 1) {}
        For i As Integer = 0 To ATA.GetUpperBound(0)
            For j As Integer = 0 To ATA.GetUpperBound(1) - i
                ATA(i, j) = Math.Round(data(i, j + 1) / data(i, j),
                                       4, MidpointRounding.AwayFromZero)
            Next
        Next

        name = inputName
    End Sub

    Public Function getATU(ByVal inputData As Object()) As Double(,)
        ATU = New Double(inputData.GetUpperBound(1) + 1, 0) {}

        Return ATU
    End Function

    Public Function getData() As Double(,)
        Return data
    End Function

    Public Function getATA() As Double(,)
        Return ATA
    End Function

    Public Function getName() As String
        Return name
    End Function


End Class

Public Module test
    Public Sub test1()
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets("Sheet1"), Worksheet)
        Dim rng As Range = wkst.Range("A1:J10")

        Dim test As Triangle = New Triangle(CType(rng.Value, Object(,)), "Paid")

        rng.Offset(20, 0).Resize(9, 9).Value = test.getATA()
        CType(wkst.Cells(45, 1), Range).Value = test.getName()
    End Sub

    Public Sub testing()
        'ProjectionFormat.summary()
        'ProjectionFormat.showDefaultTriangleView()
        'ProjectionFormat.expLoss()
        'ProjectionFormat.testButton()
        'ProjectionFormat.defaultTriangleView("Count")
        'ProjectionFormat.reviewTemplate()
        'PullData.getInitialTriangleList()
        'PullData.getTrianglesFromSqlSvr()
        'PullData.testConvert()
    End Sub
End Module
