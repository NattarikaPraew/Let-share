Public Class Form1
    Dim menuName As String
    Dim menuPrice As Double
    Dim pplName As String
    Dim value As Integer
    Dim totalPrice As Integer

    Dim ppl As New Dictionary(Of String, Integer)

    Private Sub tbAddName_Click(sender As System.Object, e As System.EventArgs) Handles btnAddName.Click
        Dim isAdded As Boolean = False

        If tbName.Text = "" Then
            MsgBox("Please add some name")
        Else
            pplName = tbName.Text

            For i As Integer = 0 To lbNames.Items.Count - 1
                If lbNames.Items(i).ToString = pplName Then
                    MsgBox("This name already has been added")
                    isAdded = True
                    Exit For
                End If

            Next

            If Not isAdded Then
                lbNames.Items.Add(pplName)

                ppl.Add(pplName, value)
                tbName.Clear()

                lbPeople.Text = lbNames.Items.Count
            End If
        End If
    End Sub

    Private Sub btnToLeft_Click(sender As System.Object, e As System.EventArgs) Handles btnToLeft.Click

        Try
            'ขวามาซ้าย
            lbSelectedName.Items.Add(lbNames.SelectedItem)

        Catch ex As Exception
            If lbNames.Items.Count = 0 Then
                MsgBox("Please add some name")
            End If
        End Try


    End Sub

    Private Sub btnToRight_Click(sender As System.Object, e As System.EventArgs) Handles btnToRight.Click
        'เอาซ้ายออก
        lbSelectedName.Items.Remove(lbSelectedName.SelectedItem)
    End Sub

    Private Sub btnAddMenu_Click_1(sender As System.Object, e As System.EventArgs) Handles btnAddMenu.Click
        Dim pplNumber As Integer
        Dim priceForEach As Integer
        Dim name As String

        If (tbMenuName.Text = "" And tbMenuPrice.Text = "") Then
            MessageBox.Show("Please add menu name and price")
        ElseIf tbMenuName.Text = "" Then
            MessageBox.Show("Please add menu name")
        ElseIf tbMenuPrice.Text = "" Then
            MessageBox.Show("Please add price")
        ElseIf lbSelectedName.Items.Count = 0 Then
            MessageBox.Show("Please select some name")
        Else
            pplNumber = lbSelectedName.Items.Count

            menuName = tbMenuName.Text
            menuPrice = Val(tbMenuPrice.Text)
            totalPrice += menuPrice

            priceForEach = menuPrice / pplNumber

            tbMenuAndName.AppendText("--------------------------------------" & ControlChars.NewLine)
            tbMenuAndName.AppendText(menuName & Chr(9) & Chr(9) & menuPrice & Chr(9) & priceForEach & ControlChars.NewLine)

            TextBox1.Clear()

            tbMenuAndName.AppendText("Payer list" & ControlChars.NewLine)
            For i As Integer = 0 To pplNumber - 1
                name = lbSelectedName.Items(i).ToString
                tbMenuAndName.AppendText(name & " " & vbCrLf)
                ppl(name) += priceForEach
            Next


            tbMenuName.Clear()
            tbMenuPrice.Clear()


            For Each pair As KeyValuePair(Of String, Integer) In ppl
                TextBox1.AppendText(pair.Key & Chr(9) & Chr(9) & pair.Value & ControlChars.NewLine)
            Next
        End If

        lbTotalPrice.Text = totalPrice & " ฿"

    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click

        Dim fileData As IO.StreamWriter
        Dim result As DialogResult = SaveFileDialog1.ShowDialog()
        'Test result

        SaveFileDialog1.Filter = "Text File (*.txt)|Rich Text File (*.rtf)"

        If result = DialogResult.OK Then
            Dim fileName As String
            fileName = SaveFileDialog1.FileName
            fileData = IO.File.AppendText(fileName)
            'write to a string
            Dim strMenuAndName, strPay As String
            strMenuAndName = tbMenuAndName.Text
            strPay = TextBox1.Text

            fileData.WriteLine("Menu & Price")
            fileData.WriteLine(strMenuAndName)
            fileData.WriteLine("Payer")
            fileData.WriteLine(strPay)

            fileData.Close()
        End If

    End Sub

End Class
