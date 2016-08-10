Public Class CustomerGridForm

    Private Sub CustomerGridForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Displays all customer records

        DataGrid1.DataSource = arrCustomers

    End Sub
End Class