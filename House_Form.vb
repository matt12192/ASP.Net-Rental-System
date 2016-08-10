Public Class House_Form

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'House details error code
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
            TextBox9.Select()
            Exit Sub

        ElseIf TextBox10.Text = "" Then
            ErrorProvider1.SetError(TextBox10, "Required Field")
            TextBox10.Select()
            Exit Sub

        ElseIf TextBox11.Text = "" Then
            ErrorProvider1.SetError(TextBox11, "Required Field")
            TextBox11.Select()
            Exit Sub

        ElseIf TextBox12.Text = "" Then
            ErrorProvider1.SetError(TextBox12, "Required Field")
            TextBox12.Select()
            Exit Sub

        ElseIf TextBox13.Text = "" Then
            ErrorProvider1.SetError(TextBox13, "Required Field")
            TextBox13.Select()
            Exit Sub

        End If

        'Creates an Instance of the House Class

        'Assigns House Instance values to each of the objects

        arrHouses(intHouses) = New House(TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, TextBox5.Text, _
                                         TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox9.Text, TextBox10.Text, TextBox11.Text, TextBox12.Text, TextBox13.Text)


        'Code below allocates the correct row to each user by adding 1 to the House counter each time a new user adds data. 
        intHouses = intHouses + 1

        'Clear all Textboxes

        Clear_TextBoxes()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        'Displays house Grid form

        HouseGridForm.Show()

    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

        'Displays House find form

        House_Find.Show()

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

    Public Sub Find_House()

        'Loop through each row in the arrCustomers

        For loopCounter = 0 To intHouses - 1

            'If the HouseReference for the row matches the HouseReference variable then all data will display in textboxes

            If arrHouses(loopCounter).HouseReference = HouseReference Then

                TextBox1.Text = arrHouses(loopCounter).HouseReference.ToString
                TextBox2.Text = arrHouses(loopCounter).CustomerID.ToString
                TextBox3.Text = arrHouses(loopCounter).Address.ToString
                TextBox4.Text = arrHouses(loopCounter).Town.ToString
                TextBox5.Text = arrHouses(loopCounter).County.ToString
                TextBox6.Text = arrHouses(loopCounter).PostCode.ToString
                TextBox7.Text = arrHouses(loopCounter).Type.ToString
                TextBox8.Text = arrHouses(loopCounter).NoOfBedrooms.ToString
                TextBox9.Text = arrHouses(loopCounter).NoOfReceiptionRooms.ToString
                TextBox10.Text = arrHouses(loopCounter).Garage.ToString
                TextBox11.Text = arrHouses(loopCounter).RentalCosts.ToString
                TextBox12.Text = arrHouses(loopCounter).LeaseLength.ToString
                TextBox13.Text = arrHouses(loopCounter).Status.ToString

            End If

        Next

    End Sub



End Class