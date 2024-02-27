Private Sub CommandButton1_Click()
    partName = TextBox1.Value
    partCount = TextBox2.Value
    partPrice = TextBox3.Value
    partTotalPrice = CDbl(partCount) * CDbl(partPrice)
    With UserForm1
        With .ListBox1
            rowscount = .ListCount
            .AddItem ""
            .List(rowscount, 0) = partName
            .List(rowscount, 1) = partCount
            .List(rowscount, 2) = partPrice
            .List(rowscount, 3) = partTotalPrice
        End With
    End With
    TextBox1.Value = Empty
    TextBox2.Value = Empty
    TextBox3.Value = Empty
    With UserForm1
        .Label6.Caption = CDbl(.Label6.Caption) + partTotalPrice
    End With
End Sub
