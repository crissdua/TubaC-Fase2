Public Class MDI_F

    Private Const CP_NOCLOSE_BUTTON As Integer = &H200
    Private user As String
    Private v As Integer

    Public Sub New(ByVal user As String)
        MyBase.New()
        InitializeComponent()
        '  Note which form has called this one
        ToolStripStatusLabel1.Text = user
    End Sub
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property
    Private Sub Fase1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Fase1ToolStripMenuItem.Click
        Dim frm As New FrmFase1
        frm.MdiParent = Me
        frm.Show()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Dim result As Integer = MessageBox.Show("Desea salir del modulo?", "Atencion", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            MessageBox.Show("Puede continuar")
        ElseIf result = DialogResult.Yes Then
            MessageBox.Show("Finalizando modulo")
            Application.Exit()
            Me.Close()
        End If
    End Sub

    Private Sub Fase2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Fase2ToolStripMenuItem.Click
        Dim frm As New FrmFase2
        frm.MdiParent = Me
        frm.Show()
    End Sub

    Private Sub MDI_F_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class