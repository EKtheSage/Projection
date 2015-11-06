Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel


'Licensed under the Apache License, Version 2.0 (the "License"); you may Not use this file except In compliance With the License.	
'You may obtain a copy Of the License at http://www.apache.org/licenses/LICENSE-2.0	
'Unless required by applicable law Or agreed To In writing, software distributed under the License Is distributed On an "AS IS" BASIS,	
'WITHOUT WARRANTIES Or CONDITIONS Of ANY KIND, either express Or implied.	
'See the License For the specific language governing permissions And limitations under the License.	


Public Class frmLineColor
    Inherits Form

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        System.Windows.Forms.Application.EnableVisualStyles()

        ' Add any initialization after the InitializeComponent() call.
        For i = 1 To 5
            Dim myLabel As System.Windows.Forms.Label = CType(Controls("lblYear" & i), System.Windows.Forms.Label)
            myLabel.Text = (Year(CType(wkstControl.Range("CurrentEvalDate").Value, Date)) + i - 5).ToString
            myLabel.Height = 20
        Next

        lstColorName.SelectedIndex = 0
        AddHandler lstColorName.SelectedIndexChanged, AddressOf lstColorNameIndexChanged

    End Sub


    Private Sub lstColorNameIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles lstColorName.SelectedIndexChanged
        Dim colorTable As Data.DataTable = getTable()
        Dim selColorName As String
        Dim lblYear, lblRGB As System.Windows.Forms.Label

        If lstColorName.SelectedIndex >= 0 Then
            selColorName = lstColorName.SelectedItem.ToString()
        Else
            selColorName = lstColorName.Items(0).ToString
        End If

        Dim query =
            From cTbl In colorTable.AsEnumerable
            Where cTbl.Field(Of String)("colorName") = selColorName
            Select New With
                {.level = cTbl.Item("level"), .red = cTbl.Item("red"), .green = cTbl.Item("green"), .blue = cTbl.Item("blue")}

        For Each item In query
            lblRGB = CType(Controls("lblYear" & CType(item.level, Integer) & "RGB"), System.Windows.Forms.Label)
            lblYear = CType(Controls("lblYear" & CType(item.level, Integer)), System.Windows.Forms.Label)
            lblRGB.Text = CType(item.red, Integer) & "," & CType(item.green, Integer) & "," & CType(item.blue, Integer)
            lblYear.BackColor = Color.FromArgb(CType(item.red, Integer), CType(item.green, Integer), CType(item.blue, Integer))
        Next

    End Sub

    Private Function getTable() As Data.DataTable
        Dim table As New Data.DataTable
        table.Columns.Add("colorName", GetType(String))
        table.Columns.Add("level", GetType(Integer))
        table.Columns.Add("red", GetType(Integer))
        table.Columns.Add("green", GetType(Integer))
        table.Columns.Add("blue", GetType(Integer))

        table.Rows.Add("seq-1", 1, 255, 255, 204)
        table.Rows.Add("seq-1", 2, 194, 230, 153)
        table.Rows.Add("seq-1", 3, 120, 198, 121)
        table.Rows.Add("seq-1", 4, 49, 163, 84)
        table.Rows.Add("seq-1", 5, 0, 104, 55)
        table.Rows.Add("seq-2", 1, 255, 255, 204)
        table.Rows.Add("seq-2", 2, 161, 218, 180)
        table.Rows.Add("seq-2", 3, 65, 182, 196)
        table.Rows.Add("seq-2", 4, 44, 127, 184)
        table.Rows.Add("seq-2", 5, 37, 52, 148)
        table.Rows.Add("seq-3", 1, 240, 249, 232)
        table.Rows.Add("seq-3", 2, 186, 228, 188)
        table.Rows.Add("seq-3", 3, 123, 204, 196)
        table.Rows.Add("seq-3", 4, 67, 162, 202)
        table.Rows.Add("seq-3", 5, 8, 104, 172)
        table.Rows.Add("seq-4", 1, 237, 248, 251)
        table.Rows.Add("seq-4", 2, 178, 226, 226)
        table.Rows.Add("seq-4", 3, 102, 194, 164)
        table.Rows.Add("seq-4", 4, 44, 162, 95)
        table.Rows.Add("seq-4", 5, 0, 109, 44)
        table.Rows.Add("seq-5", 1, 246, 239, 247)
        table.Rows.Add("seq-5", 2, 189, 201, 225)
        table.Rows.Add("seq-5", 3, 103, 169, 207)
        table.Rows.Add("seq-5", 4, 28, 144, 153)
        table.Rows.Add("seq-5", 5, 1, 108, 89)
        table.Rows.Add("seq-6", 1, 241, 238, 246)
        table.Rows.Add("seq-6", 2, 189, 201, 225)
        table.Rows.Add("seq-6", 3, 116, 169, 207)
        table.Rows.Add("seq-6", 4, 43, 140, 190)
        table.Rows.Add("seq-6", 5, 4, 90, 141)
        table.Rows.Add("seq-7", 1, 237, 248, 251)
        table.Rows.Add("seq-7", 2, 179, 205, 227)
        table.Rows.Add("seq-7", 3, 140, 150, 198)
        table.Rows.Add("seq-7", 4, 136, 86, 167)
        table.Rows.Add("seq-7", 5, 129, 15, 124)
        table.Rows.Add("seq-8", 1, 254, 235, 226)
        table.Rows.Add("seq-8", 2, 251, 180, 185)
        table.Rows.Add("seq-8", 3, 247, 104, 161)
        table.Rows.Add("seq-8", 4, 197, 27, 138)
        table.Rows.Add("seq-8", 5, 122, 1, 119)
        table.Rows.Add("seq-9", 1, 241, 238, 246)
        table.Rows.Add("seq-9", 2, 215, 181, 216)
        table.Rows.Add("seq-9", 3, 223, 101, 176)
        table.Rows.Add("seq-9", 4, 221, 28, 119)
        table.Rows.Add("seq-9", 5, 152, 0, 67)
        table.Rows.Add("seq-10", 1, 254, 240, 217)
        table.Rows.Add("seq-10", 2, 253, 204, 138)
        table.Rows.Add("seq-10", 3, 252, 141, 89)
        table.Rows.Add("seq-10", 4, 227, 74, 51)
        table.Rows.Add("seq-10", 5, 179, 0, 0)
        table.Rows.Add("seq-11", 1, 255, 255, 178)
        table.Rows.Add("seq-11", 2, 254, 204, 92)
        table.Rows.Add("seq-11", 3, 253, 141, 60)
        table.Rows.Add("seq-11", 4, 240, 59, 32)
        table.Rows.Add("seq-11", 5, 189, 0, 38)
        table.Rows.Add("seq-12", 1, 255, 255, 212)
        table.Rows.Add("seq-12", 2, 254, 217, 142)
        table.Rows.Add("seq-12", 3, 254, 153, 41)
        table.Rows.Add("seq-12", 4, 217, 95, 14)
        table.Rows.Add("seq-12", 5, 153, 52, 4)
        table.Rows.Add("seq-13", 1, 242, 240, 247)
        table.Rows.Add("seq-13", 2, 203, 201, 226)
        table.Rows.Add("seq-13", 3, 158, 154, 200)
        table.Rows.Add("seq-13", 4, 117, 107, 177)
        table.Rows.Add("seq-13", 5, 84, 39, 143)
        table.Rows.Add("seq-14", 1, 239, 243, 255)
        table.Rows.Add("seq-14", 2, 189, 215, 231)
        table.Rows.Add("seq-14", 3, 107, 174, 214)
        table.Rows.Add("seq-14", 4, 49, 130, 189)
        table.Rows.Add("seq-14", 5, 8, 81, 156)
        table.Rows.Add("seq-15", 1, 237, 248, 233)
        table.Rows.Add("seq-15", 2, 186, 228, 179)
        table.Rows.Add("seq-15", 3, 116, 196, 118)
        table.Rows.Add("seq-15", 4, 49, 163, 84)
        table.Rows.Add("seq-15", 5, 0, 109, 44)
        table.Rows.Add("seq-16", 1, 254, 237, 222)
        table.Rows.Add("seq-16", 2, 253, 190, 133)
        table.Rows.Add("seq-16", 3, 253, 141, 60)
        table.Rows.Add("seq-16", 4, 230, 85, 13)
        table.Rows.Add("seq-16", 5, 166, 54, 3)
        table.Rows.Add("seq-17", 1, 254, 229, 217)
        table.Rows.Add("seq-17", 2, 252, 174, 145)
        table.Rows.Add("seq-17", 3, 251, 106, 74)
        table.Rows.Add("seq-17", 4, 222, 45, 38)
        table.Rows.Add("seq-17", 5, 165, 15, 21)
        table.Rows.Add("seq-18", 1, 247, 247, 247)
        table.Rows.Add("seq-18", 2, 204, 204, 204)
        table.Rows.Add("seq-18", 3, 150, 150, 150)
        table.Rows.Add("seq-18", 4, 99, 99, 99)
        table.Rows.Add("seq-18", 5, 37, 37, 37)
        table.Rows.Add("div-1", 1, 230, 97, 1)
        table.Rows.Add("div-1", 2, 253, 184, 99)
        table.Rows.Add("div-1", 3, 247, 247, 247)
        table.Rows.Add("div-1", 4, 178, 171, 210)
        table.Rows.Add("div-1", 5, 94, 60, 153)
        table.Rows.Add("div-2", 1, 166, 97, 26)
        table.Rows.Add("div-2", 2, 223, 194, 125)
        table.Rows.Add("div-2", 3, 245, 245, 245)
        table.Rows.Add("div-2", 4, 128, 205, 193)
        table.Rows.Add("div-2", 5, 1, 133, 113)
        table.Rows.Add("div-3", 1, 123, 50, 148)
        table.Rows.Add("div-3", 2, 194, 165, 207)
        table.Rows.Add("div-3", 3, 247, 247, 247)
        table.Rows.Add("div-3", 4, 166, 219, 160)
        table.Rows.Add("div-3", 5, 0, 136, 55)
        table.Rows.Add("div-4", 1, 208, 28, 139)
        table.Rows.Add("div-4", 2, 241, 182, 218)
        table.Rows.Add("div-4", 3, 247, 247, 247)
        table.Rows.Add("div-4", 4, 184, 225, 134)
        table.Rows.Add("div-4", 5, 77, 172, 38)
        table.Rows.Add("div-5", 1, 202, 0, 32)
        table.Rows.Add("div-5", 2, 244, 165, 130)
        table.Rows.Add("div-5", 3, 247, 247, 247)
        table.Rows.Add("div-5", 4, 146, 197, 222)
        table.Rows.Add("div-5", 5, 5, 113, 176)
        table.Rows.Add("div-6", 1, 202, 0, 32)
        table.Rows.Add("div-6", 2, 244, 165, 130)
        table.Rows.Add("div-6", 3, 255, 255, 255)
        table.Rows.Add("div-6", 4, 186, 186, 186)
        table.Rows.Add("div-6", 5, 64, 64, 64)
        table.Rows.Add("div-7", 1, 215, 25, 28)
        table.Rows.Add("div-7", 2, 253, 174, 97)
        table.Rows.Add("div-7", 3, 255, 255, 191)
        table.Rows.Add("div-7", 4, 171, 217, 233)
        table.Rows.Add("div-7", 5, 44, 123, 182)
        table.Rows.Add("div-8", 1, 215, 25, 28)
        table.Rows.Add("div-8", 2, 253, 174, 97)
        table.Rows.Add("div-8", 3, 255, 255, 191)
        table.Rows.Add("div-8", 4, 171, 221, 164)
        table.Rows.Add("div-8", 5, 43, 131, 186)
        table.Rows.Add("div-9", 1, 215, 25, 28)
        table.Rows.Add("div-9", 2, 253, 174, 97)
        table.Rows.Add("div-9", 3, 255, 255, 191)
        table.Rows.Add("div-9", 4, 166, 217, 106)
        table.Rows.Add("div-9", 5, 26, 150, 65)
        table.Rows.Add("qual-1", 1, 141, 211, 199)
        table.Rows.Add("qual-1", 2, 255, 255, 179)
        table.Rows.Add("qual-1", 3, 190, 186, 218)
        table.Rows.Add("qual-1", 4, 251, 128, 114)
        table.Rows.Add("qual-1", 5, 128, 177, 211)
        table.Rows.Add("qual-2", 1, 251, 180, 174)
        table.Rows.Add("qual-2", 2, 179, 205, 227)
        table.Rows.Add("qual-2", 3, 204, 235, 197)
        table.Rows.Add("qual-2", 4, 222, 203, 228)
        table.Rows.Add("qual-2", 5, 254, 217, 166)
        table.Rows.Add("qual-3", 1, 228, 26, 28)
        table.Rows.Add("qual-3", 2, 55, 126, 184)
        table.Rows.Add("qual-3", 3, 77, 175, 74)
        table.Rows.Add("qual-3", 4, 152, 78, 163)
        table.Rows.Add("qual-3", 5, 255, 127, 0)
        table.Rows.Add("qual-4", 1, 179, 226, 205)
        table.Rows.Add("qual-4", 2, 253, 205, 172)
        table.Rows.Add("qual-4", 3, 203, 213, 232)
        table.Rows.Add("qual-4", 4, 244, 202, 228)
        table.Rows.Add("qual-4", 5, 230, 245, 201)
        table.Rows.Add("qual-5", 1, 102, 194, 165)
        table.Rows.Add("qual-5", 2, 252, 141, 98)
        table.Rows.Add("qual-5", 3, 141, 160, 203)
        table.Rows.Add("qual-5", 4, 231, 138, 195)
        table.Rows.Add("qual-5", 5, 166, 216, 84)
        table.Rows.Add("qual-6", 1, 27, 158, 119)
        table.Rows.Add("qual-6", 2, 217, 95, 2)
        table.Rows.Add("qual-6", 3, 117, 112, 179)
        table.Rows.Add("qual-6", 4, 231, 41, 138)
        table.Rows.Add("qual-6", 5, 102, 166, 30)
        table.Rows.Add("qual-7", 1, 166, 206, 227)
        table.Rows.Add("qual-7", 2, 31, 120, 180)
        table.Rows.Add("qual-7", 3, 178, 223, 138)
        table.Rows.Add("qual-7", 4, 51, 160, 44)
        table.Rows.Add("qual-7", 5, 251, 154, 153)
        table.Rows.Add("qual-8", 1, 127, 201, 127)
        table.Rows.Add("qual-8", 2, 190, 174, 212)
        table.Rows.Add("qual-8", 3, 253, 192, 134)
        table.Rows.Add("qual-8", 4, 255, 255, 153)
        table.Rows.Add("qual-8", 5, 56, 108, 176)

        Return table
    End Function
    Private Sub frmLineColor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        updateColor("Exp Loss")
        updateColor("Review Template")
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Hide()
    End Sub

    Private Sub updateColor(wkstName As String)
        Dim wkst As Worksheet = CType(Application.ActiveWorkbook.Worksheets(wkstName), Worksheet)
        Dim red, green, blue As Integer
        For Each chartObj As ChartObject In CType(wkst.ChartObjects, ChartObjects)
            MsgBox(chartObj.Chart.Name)
            For i = 1 To 5
                With CType(chartObj.Chart.SeriesCollection(i), Series)
                    red = Integer.Parse(Controls("lblYear" & i & "RGB").Text.Split(New Char() {","c})(0))
                    green = Integer.Parse(Controls("lblYear" & i & "RGB").Text.Split(New Char() {","c})(1))
                    blue = Integer.Parse(Controls("lblYear" & i & "RGB").Text.Split(New Char() {","c})(2))

                    'note the tricky part here, the interop has a bug that takes in argument for RGB as Blue-Green-Red
                    'instead of the expected argument Red-Green-Blue
                    .Format.Fill.ForeColor.RGB = Color.FromArgb(blue, green, red).ToArgb
                    .Format.Line.ForeColor.RGB = Color.FromArgb(blue, green, red).ToArgb
                End With
            Next
        Next
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        MsgBox("This is an implementation of ColorBrewer. Credit goes to Cynthia Brewer who created the color themes. " &
               "Visit colorbrewer2.org for more information. The color options " &
               "allow for clearer identification between different data groups." & vbCrLf & vbCrLf &
               "1. Sequential (seq-n) schemes are suited to ordered data that progress from low to high. " &
               "Lightness steps dominate the look of these schemes, with light colors for low data " &
               "values to dark colors for high data values." & vbCrLf & vbCrLf &
               "2. Diverging (div-n) schemes put equal emphasis on mid-range critical values and " &
               "extremes at both ends of the data range. The critical class Or break in the " &
               "middle of the legend Is emphasized with light colors And low And high extremes " &
               "are emphasized with dark colors that have contrasting hues." & vbCrLf & vbCrLf &
               "3.Qualitative (qual-n) schemes Do Not imply magnitude differences between legend classes, " &
               "And hues are used to create the primary visual differences between classes. " &
               "Qualitative schemes are best suited to representing nominal Or categorical data.", Title:="Color Guide")
    End Sub
End Class

Public Class colorRGB
    Public Property level() As Integer
    Public Property red() As Integer
    Public Property green() As Integer
    Public Property blue() As Integer
End Class