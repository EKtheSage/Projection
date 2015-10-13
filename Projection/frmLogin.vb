Imports System.Windows.Forms

Public Class frmLogin
    Inherits Form

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If txtBxPw.Text = "" Or txtBxUser.Text = "" Then
            MsgBox("You need to enter both user name and password!")
        Else
            Hide()
            'testFormLogin()
            getGUIBNRCount()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Hide()
        txtBxPw.Clear()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class

Public Module playForm
    'still trying to figure out how to set cursor on the username textbox...
    Public frmLogin As frmLogin = New frmLogin

    Public Sub initializeForm()
        frmLogin.txtBxPw.Clear()
        frmLogin.Show()
        frmLogin.txtBxUser.Focus()
    End Sub
    Public Sub testFormLogin()
        MsgBox(frmLogin.txtBxUser.Text)
        MsgBox(frmLogin.txtBxPw.Text)
    End Sub
End Module