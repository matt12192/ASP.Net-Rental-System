Public Class Menu

    Private Sub Menu_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Loads the menu form
        'Initialise Public_Module variables

        intCustomers = 0

        intHouses = 0
       
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'Used to display customer Form

        Customer_Form.Show()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        'Used to display House Form

        House_Form.Show()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        'Exits the application

        Application.Exit()
    End Sub
End Class