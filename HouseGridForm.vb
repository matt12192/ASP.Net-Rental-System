Public Class HouseGridForm

    Private Sub HouseGridForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Displays all compiled data related to Houses

        DataGrid1.DataSource = arrHouses

    End Sub
End Class