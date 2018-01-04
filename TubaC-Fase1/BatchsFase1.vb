Public Class BatchsFase1
    Public batchsnum As Double
    Dim cantidadR As Double
    Dim itemcode As String
    Dim oCompany As SAPbobsCOM.Company = Login.con.oCompany
    Dim Connected As Boolean = Login.con.Connected
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Dim GoodsReceiptPO = FrmFase1.GoodsReceiptPO
    Dim objectCode As Integer
    Private Const CP_NOCLOSE_BUTTON As Integer = &H200

    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim myCp As CreateParams = MyBase.CreateParams
            myCp.ClassStyle = myCp.ClassStyle Or CP_NOCLOSE_BUTTON
            Return myCp
        End Get
    End Property

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Convert.ToDouble(Label6.Text) = 0 Then
            batchs()
        Else
            MessageBox.Show("Verifique que la cantidad requerida concuerde con el total consumido")
        End If
    End Sub
    Friend Sub load(item As String, cant As Double)
        DGV.Rows.Clear()
        itemcode = item
        Label4.Text = itemcode
        cantidadR = cant
        Label2.Text = cant
    End Sub
    Private Sub Batchs_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Public Function batchs()
        Dim oReturn As Integer = -1
        Dim line As Integer
        Dim i As Integer = 0
        While i < DGV.Rows.Count - 1
            GoodsReceiptPO.Lines.BatchNumbers.SetCurrentLine(i)
            GoodsReceiptPO.Lines.BatchNumbers.BatchNumber = "000-00" & DGV.Rows(i).Cells(0).Value.ToString
            GoodsReceiptPO.Lines.BatchNumbers.Quantity = Convert.ToDouble(DGV.Rows(i).Cells(1).Value)
            GoodsReceiptPO.Lines.BatchNumbers.AddmisionDate = "2017-12-29"
            GoodsReceiptPO.Lines.BatchNumbers.Add()
            i = i + 1
        End While
        Me.Hide()

    End Function



    Private Sub DGV_KeyUp(sender As Object, e As KeyEventArgs) Handles DGV.KeyUp
        Dim suma As Double
        Dim queda As Double
        For Each row As DataGridViewRow In DGV.Rows
            suma += Val(row.Cells(1).Value)
        Next
        queda = Convert.ToDouble(Label2.Text) - suma
        Label6.Text = queda
        Label6.Refresh()

    End Sub
End Class