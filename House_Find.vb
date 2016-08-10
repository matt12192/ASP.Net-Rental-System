Public Class House_Find

    Private Sub House_Find_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Allows user to search records of specifically stored houses and display them on screen

        For loopCounter = 0 To intHouses - 1

            ComboBox1.Items.Add(arrHouses(loopCounter).HouseReference.ToString)

        Next

        ComboBox1.SelectedIndex = -1

        ComboBox1.Text = "Please Select"

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

        HouseReference = ComboBox1.SelectedItem

        House_Form.Find_House()

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Used to close form when finished using

        Me.Close()

    End Sub
End Class