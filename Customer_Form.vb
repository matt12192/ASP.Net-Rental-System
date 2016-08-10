Public Class Customer_Form

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'Customer details error code
        If TextBox1.Text = "" Then
            ErrorProvider1.SetError(TextBox1, "Required Field")
            TextBox1.Select()
            Exit Sub

        ElseIf TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Required Field")
            TextBox2.Select()
            Exit Sub

        ElseIf TextBox3.Text = "" Then
            ErrorProvider1.SetError(TextBox3, "Required Field")
            TextBox3.Select()
            Exit Sub

        ElseIf TextBox4.Text = "" Then
            ErrorProvider1.SetError(TextBox4, "Required Field")
            TextBox4.Select()
            Exit Sub

        ElseIf TextBox5.Text = "" Then
            ErrorProvider1.SetError(TextBox5, "Required Field")
            TextBox5.Select()
            Exit Sub

        ElseIf TextBox6.Text = "" Then
            ErrorProvider1.SetError(TextBox6, "Required Field")
            TextBox6.Select()
            Exit Sub

        ElseIf TextBox7.Text = "" Then
            ErrorProvider1.SetError(TextBox7, "Required Field")
            TextBox7.Select()
            Exit Sub

        ElseIf TextBox8.Text = "" Then
            ErrorProvider1.SetError(TextBox8, "Required Field")
            TextBox8.Select()
            Exit Sub

        ElseIf TextBox9.Text = "" Then
            ErrorProvider1.SetError(TextBox8, "Required Field")
            TextBox8.Select()
            Exit Sub

        End If

        'Creates an Instance of the Customer Class

        'Assigns customer Instance values to each of the objects

        arrCustomers(intCustomers) = New Customer(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, _
                                                  TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text)


        'Increase the customer counter by 1 so that next time a customer is added to the arrCustomer they will be assigned to the correct position in the array


        intCustomers = intCustomers + 1

        'Clear all Textboxes

        Clear_TextBoxes()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        'Displays customer Grid form
        CustomerGridForm.Show()

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        'Displays customer find form

        Customer_Find.Show()

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        Clear_TextBoxes()

    End Sub

    Private Sub Clear_TextBoxes()

        'Loop through all Visual Control names inside GroupBox1

        For intControlCounter = 0 To GroupBox1.Controls.Count - 1

            'If the Name of the Visual Control in GroupBox1 is longer than 6 characters 

            If GroupBox1.Controls(intControlCounter).Name.Length > 6 Then

                'Check the Name of the Visual Control to see if it is a TextBox Control or not

                If GroupBox1.Controls(intControlCounter).Name.Substring(0, 7) = "TextBox" Then

                    'If the Visual Control is a TextBox then set its contents to no text

                    GroupBox1.Controls(intControlCounter).Text = ""

                End If

            End If
        Next

    End Sub

    Public Sub Find_Customer()

        'Loop through each row in the arrCustomers

        For loopCounter = 0 To intCustomers - 1

            'If the CustomerID for the row matches the CustomerID variable then all data will display in textboxes

            If arrCustomers(loopCounter).CustomerID = CustomerID Then

                TextBox1.Text = arrCustomers(loopCounter).CustomerID.ToString
                TextBox2.Text = arrCustomers(loopCounter).FirstName.ToString
                TextBox3.Text = arrCustomers(loopCounter).Surname.ToString
                TextBox4.Text = arrCustomers(loopCounter).Address.ToString
                TextBox5.Text = arrCustomers(loopCounter).Town.ToString
                TextBox6.Text = arrCustomers(loopCounter).County.ToString
                TextBox7.Text = arrCustomers(loopCounter).PostCode.ToString
                TextBox8.Text = arrCustomers(loopCounter).Phone.ToString
                TextBox9.Text = arrCustomers(loopCounter).Email.ToString

            End If

        Next

    End Sub



End Class