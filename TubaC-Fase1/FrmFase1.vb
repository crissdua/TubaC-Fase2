Imports System.Windows.Forms
Imports System.IO
Imports System.Data.SqlClient
Public Class FrmFase1

    Dim objectCode As String
    Dim itemcode As String
    Dim oCompany As SAPbobsCOM.Company = Login.con.oCompany
    Dim Connected As Boolean = Login.con.Connected
    Dim connectionString As String = Conexion.ObtenerConexion.ConnectionString
    Public Shared PO As SAPbobsCOM.Documents
    Public Shared GoodsReceiptPO As SAPbobsCOM.Documents
    'se crea la lsita de batch a ingresar
    'Dim listadebach As New List(Of BatchsubENC)
    'Dim ItemcodeGrid As String = String.Empty
    Public Shared SQL_Conexion As SqlConnection = New SqlConnection()

    Private Sub FrmFase1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Select()
        PO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
        GoodsReceiptPO = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
        cargaORDER()

        'GR_from_PO()
    End Sub
    Public Function cargaORDER()
        Dim con As SqlConnection = New SqlConnection(connectionString)
        con.Open()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OPOR T0 WHERE T0.DocType ='I' and  T0.CANCELED = 'N' and  T0.DocStatus ='O'", con)
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.Close()
    End Function



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim con As SqlConnection = New SqlConnection(connectionString)
        con.Open()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.DocNum FROM OPOR T0 WHERE T0.DocType ='I' and  T0.CANCELED = 'N' and  T0.DocStatus ='O' and T0.DocNum LIKE '" + TextBox2.Text + "%' ORDER BY T0.DocNum", con)
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV.DataSource = DT_dat
        con.Close()
    End Sub

    Private Sub GR_from_PO()
        Dim iResult As Integer = -1
        Dim iResult2 As Integer = -1
        Dim sResult As String = String.Empty
        Dim sOutput As String = String.Empty
        Dim sql As String
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim sql2 As String
        Dim oRecordSet2 As SAPbobsCOM.Recordset
        Dim docentry As String
        Dim frms As New BatchsFase1
        Try





            '------------------OBTIENE DOCENTRY------------
            sql = ("SELECT T0.DocEntry FROM OPOR T0 WHERE T0.DocNum = '" + txtOrder.Text + "'")
            oRecordSet = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(sql)
            If oRecordSet.RecordCount > 0 Then
                docentry = oRecordSet.Fields.Item(0).Value
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()

            'Dim PO As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            'Dim GoodsReceiptPO As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)

            '----------------------------------------------
            PO.GetByKey(docentry)
            '----------------------------------------------

            GoodsReceiptPO.CardCode = PO.CardCode
            GoodsReceiptPO.CardName = PO.CardName
            '----------- LINEAS -----------------------------

            Dim itemcode As String
            Dim quantity As Double
            Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
            Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
            If existe = False Then
                DGV2.Columns.Add(i)
                i.HeaderText = "CHK"
                i.Name = "CHK"
                i.Width = 32
                i.DisplayIndex = 0
            End If
            Dim result As Integer = MessageBox.Show("Desea Ingresar la Orden?", "Atencion", MessageBoxButtons.YesNoCancel)
            If result = DialogResult.Cancel Then
                MessageBox.Show("Cancelado")
            ElseIf result = DialogResult.No Then
                MessageBox.Show("No se realizara la orden")
            ElseIf result = DialogResult.Yes Then
                For Each row As DataGridViewRow In DGV2.Rows
                    Dim chk As DataGridViewCheckBoxCell = row.Cells("CHK")
                    If chk.Value IsNot Nothing AndAlso chk.Value = True Then
                        PO.Lines.SetCurrentLine(DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString)
                        PO.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                        'GoodsReceiptPO.Lines.ItemCode = DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString
                        'GoodsReceiptPO.Lines.ItemDescription = PO.Lines.ItemDescription
                        GoodsReceiptPO.Lines.Quantity = DGV2.Rows(chk.RowIndex).Cells.Item(2).Value.ToString
                        GoodsReceiptPO.Lines.BaseEntry = PO.DocEntry
                        GoodsReceiptPO.Lines.BaseLine = DGV2.Rows(chk.RowIndex).Cells.Item(3).Value.ToString
                        GoodsReceiptPO.Lines.BaseType = Convert.ToInt32(PO.DocObjectCodeEx)
                        frms.load(DGV2.Rows(chk.RowIndex).Cells.Item(1).Value.ToString,
                                  DGV2.Rows(chk.RowIndex).Cells.Item(2).Value.ToString)
                        frms.ShowDialog()
                        'batchs()
                        GoodsReceiptPO.Lines.Add()
                    End If
                Next

            End If

            '-------------------------------BATCH--------------------------------------

            '---------------------------------------------------------------------------

            ' GoodsReceiptPO.Comments = PO.DocEntry
            iResult = GoodsReceiptPO.Add()
            If iResult <> 0 Then
                MessageBox.Show(oCompany.GetLastErrorDescription)
            Else
                PO.Comments = PO.DocEntry
                iResult2 = PO.Update()
                If iResult2 <> 0 Then
                    MessageBox.Show(oCompany.GetLastErrorDescription)
                End If
                MessageBox.Show("listo")
            End If
        Catch ex As Exception
            MsgBox("Error: " + ex.Message.ToString)
        End Try
    End Sub

    Private Sub DGV_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        txtOrder.Text = DGV(0, DGV.CurrentCell.RowIndex).Value.ToString()
    End Sub

    Private Sub txtOrder_TextChanged(sender As Object, e As EventArgs) Handles txtOrder.TextChanged
        'listadebach = Nothing
        'listadebach = New List(Of BatchsubENC)
        Dim i As DataGridViewCheckBoxColumn = New DataGridViewCheckBoxColumn()
        Dim existe As Boolean = DGV2.Columns.Cast(Of DataGridViewColumn).Any(Function(x) x.Name = "CHK")
        If existe = False Then
            DGV2.Columns.Add(i)
            i.HeaderText = "CHK"
            i.Name = "CHK"
            i.Width = 32
            i.DisplayIndex = 0
        End If
        Dim con As SqlConnection = New SqlConnection(connectionString)
        con.Open()
        Dim SQL_da As SqlDataAdapter = New SqlDataAdapter("SELECT T0.[ItemCode], T0.[Quantity], isnull(T0.LineNum,0) FROM POR1 T0 WHERE T0.[LineStatus] = 'O' and T0.[DocEntry] like '" + txtOrder.Text + "%'", con)
        Dim DT_dat As System.Data.DataTable = New System.Data.DataTable()
        SQL_da.Fill(DT_dat)
        DGV2.DataSource = DT_dat
        'Se crea la variable de tipo entero para recorrer el grid que se lleno con la consulta sql que se obtiene de la tabla
        'POR1 
        'Dim icontador As Integer = 0
        'While icontador <= DGV2.Rows.Count - 1
        '    Dim item As New BatchsubENC
        '    item._ITEMCODE = DGV2.Rows(icontador).Cells(1).Value
        '    item._CATNIDADTOTAL = DGV2.Rows(icontador).Cells(2).Value
        '    listadebach.Add(item)

        '    icontador = icontador + 1
        'End While


        con.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        GR_from_PO()
        DGV.DataSource = Nothing
        DGV2.DataSource = Nothing
        TextBox2.Clear()
        txtOrder.Clear()
        'Dim listadebach2 = From liBatch In listadebach Where liBatch._ITEMCODE = ItemcodeGrid
        '                   Take 1 Select liBatch
        'Dim icontador As Integer = 0
        'listadebach2.First()._batchsubdet = New List(Of BatchSubDET)
        'While icontador <= DGV3.Rows.Count - 1
        '    Dim item As New BatchSubDET
        '    item._Nombre = DGV3.Rows(icontador).Cells(0).Value
        '    item._Cantidad = DGV3.Rows(icontador).Cells(1).Value
        '    listadebach2.First()._batchsubdet.Add(item)
        '    icontador = icontador + 1
        'End While



    End Sub

    Private Sub btnFinalizar_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As Integer = MessageBox.Show("Desea limpiar el objeto?", "Atencion", MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Cancel Then
            MessageBox.Show("Cancelado")
        ElseIf result = DialogResult.No Then
            MessageBox.Show("Puede continuar!")
        ElseIf result = DialogResult.Yes Then
            TextBox2.Clear()
            txtOrder.Clear()
            PO = Nothing
            GoodsReceiptPO = Nothing
            DGV.DataSource = Nothing
            DGV2.DataSource = Nothing
            MessageBox.Show("Inicie un objeto nuevo")
        End If
    End Sub

    'Private Sub DGV2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DGV2.CellContentClick
    '    'Dim listadebach2 = From liBatch In listadebach Where liBatch._ITEMCODE = DGV2.Rows(1).Cells(1).Value
    '    '                   Select liBatch
    '    ItemcodeGrid = DGV2(1, DGV2.CurrentCell.RowIndex).Value.ToString()
    '    If Not listadebach Is Nothing Then
    '        If listadebach.Count > 0 Then
    '            Dim listadebach2 = From liBatch In listadebach Where liBatch._ITEMCODE = ItemcodeGrid
    '                               Take 1 Select liBatch
    '            If Not listadebach2.FirstOrDefault()._batchsubdet Is Nothing Then
    '                If listadebach2.FirstOrDefault()._batchsubdet.Count > 0 Then
    '                    Dim icontador As Integer = 0
    '                    While icontador < listadebach2.FirstOrDefault()._batchsubdet.Count - 1

    '                        DGV3.Rows.Add(listadebach2.FirstOrDefault()._batchsubdet(icontador)._Nombre,
    '                                      listadebach2.FirstOrDefault()._batchsubdet(icontador)._Cantidad)

    '                        icontador = icontador + 1
    '                    End While
    '                End If
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Sub btnFinalizar_Click(sender As Object, e As EventArgs) Handles btnFinalizar.Click
    '    For Each item In listadebach
    '        item._ITEMCODE
    '        item._CATNIDADTOTAL
    '        For Each itemsub In item._batchsubdet
    '            itemsub._Nombre
    '            itemsub._Cantidad
    '            item._ITEMCODE
    '        Next

    '    Next

    'End Sub

End Class