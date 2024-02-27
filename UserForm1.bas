Public carAll, orderNumber, currentDate As Variant


Private Sub UserForm_Initialize()
    With UserForm1
        orderNumber = Int((999 * Rnd) + 1)
        currentDate = Date
        .TextBox1.Value = orderNumber
        .TextBox2.Value = currentDate
        .Label6.Caption = 0
    End With
    With ComboBox1
        rngg = Range("Сотрудники[ФИО]").Value
        .List = rngg
    End With
    Dim carBrands As New Collection
    With ComboBox2
        carAll = Sheets("Справочник").ListObjects("ТС").DataBodyRange
        On Error Resume Next
        For i = LBound(carAll) To UBound(carAll)
            carBrands.Add carAll(i, 1), carAll(i, 1)
        Next i
        On Error GoTo 0
        For Each e In carBrands
            .AddItem e
        Next e
    End With
    'Set lb1 = ListBox1
    With ListBox1
        .ColumnCount = 4
        .ColumnWidths = "130;50;90;90"
        .ColumnHeads = False
    End With
End Sub

Private Sub ComboBox2_DropButtonClick()
    With ComboBox3
        .Clear
        selectedValue = ComboBox2.Value
        For i = LBound(carAll) To UBound(carAll)
            If carAll(i, 1) = selectedValue Then .AddItem carAll(i, 2)
        Next i
    End With
End Sub

Private Sub CommandButton1_Click()
    UserForm2.Show
End Sub

Private Sub CommandButton3_Click()
    With Sheets("Заказы")
        lastrow = .Cells(Rows.Count, 1).End(xlUp).Row
        lastcolumn = .Cells(1, Columns.Count).End(xlToLeft).Column
        worker = UserForm1.ComboBox1.Value
        For i = 0 To UserForm1.ListBox1.ListCount - 1
            .Cells(lastrow + 1 + i, 1) = orderNumber
            .Cells(lastrow + 1 + i, 2) = currentDate
            .Cells(lastrow + 1 + i, 3) = worker
            .Cells(lastrow + 1 + i, 4) = UserForm1.ComboBox2.Value
            .Cells(lastrow + 1 + i, 5) = UserForm1.ComboBox3.Value
            .Cells(lastrow + 1 + i, 6) = UserForm1.ListBox1.List(i, 0)
            .Cells(lastrow + 1 + i, 7) = UserForm1.ListBox1.List(i, 1)
            .Cells(lastrow + 1 + i, 8) = UserForm1.ListBox1.List(i, 2)
            .Cells(lastrow + 1 + i, 9) = UserForm1.ListBox1.List(i, 3)
            .Range(.Cells(lastrow + 1, 1), .Cells(lastrow + UserForm1.ListBox1.ListCount, lastcolumn)).Borders.LineStyle = xlContinuous
        Next i
    End With
End Sub
