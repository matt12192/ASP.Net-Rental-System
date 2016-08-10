Public Class Customer_Find

    Private Sub Customer_Find_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Allows user to search records of specifically stored customers and display them on screen


        For loopCounter = 0 To intCustomers - 1

            ComboBox1.Items.Add(arrCustomers(loopCounter).CustomerID.ToString)

        Next

        ComboBox1.SelectedIndex = -1

        ComboBox1.Text = "Please Select"

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        CustomerID = ComboBox1.SelectedItem

        Customer_Form.Find_Customer()

    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Used to close form when finished using

        Me.Close()

    End Sub
End Class