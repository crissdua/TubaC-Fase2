
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports TubacBarCodeProduction.Conexion
Public Class Login
    Public con As New Conexion
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = String.Empty Or TextBox2.Text = String.Empty Then
            MessageBox.Show("Verifique Nombre y contraseña")
        Else
            con.MakeConnectionSAP(TextBox1.Text, TextBox2.Text)
            Dim user As String
            user = TextBox1.Text
            If con.Connected = True Then
                'MessageBox.Show(con.Connected.ToString)
                Me.Hide()
                Dim frms As New MDI_F(user)
                frms.Show()
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Dim frm As New cParametros
        frm.Show()
    End Sub

    Private Sub EnterClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode.Equals(Keys.Enter) Then
            Button1_Click(1, e)
        End If
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Call EnterClick(sender, e)
    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        Call EnterClick(sender, e)
    End Sub
End Class