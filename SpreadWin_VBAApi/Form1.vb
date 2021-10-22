Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' バージョンの表示
        Me.Text = $"SPREAD ver.{FpSpread1.ProductVersion}"

        ' コンボボックスの設定
        ComboBox1.Items.AddRange({"SetValueA", "SetStyleA", "CopyCellA", "ClearCellA", "SelectCellA", "CopySheetA"})
        ComboBox2.Items.AddRange({"SetValueB", "SetStyleB", "CopyCellB", "ClearCellB", "SelectCellB", "CopySheetB"})

        ' SPREADの初期設定
        FpSpread1.TabStripRatio = 0.7
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' シートの初期化
        FpSpread1.Sheets.Clear()
        Dim sheet = New FarPoint.Win.Spread.SheetView()
        sheet.Columns.Default.Width = 72
        FpSpread1.Sheets.Add(sheet)

        ' コンボボックスの初期化
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
    End Sub

    ' 従来の方法による操作
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                Call SetValueA()
            Case 1
                Call SetStyleA()
            Case 2
                Call CopyCellA()
            Case 3
                Call ClearCellA()
            Case 4
                Call SelectCellA()
            Case 5
                Call CopySheetA()
        End Select
    End Sub

    ' VBA互換APIを使う方法による操作
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Select Case ComboBox2.SelectedIndex
            Case 0
                Call SetValueB()
            Case 1
                Call SetStyleB()
            Case 2
                Call CopyCellB()
            Case 3
                Call ClearCellB()
            Case 4
                Call SelectCellB()
            Case 5
                Call CopySheetB()
        End Select
    End Sub

    ' 値の設定：VBA互換APIを使う方法
    Private Sub SetValueA()
        With FpSpread1.Sheets(0).AsWorksheet()
            ' 文字列でセル範囲を指定する方法
            .Range("A1:B2").Value = "ABC"

            ' インデックス番号でセル範囲を指定する方法
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 3, 3)}).Value = "XYZ"
        End With
    End Sub

    ' 値の設定：従来の方法
    Private Sub SetValueB()
        With FpSpread1.Sheets(0)
            ' 文字列でセル範囲を指定する方法
            .Cells("A1:B2").Value = "ABC"

            ' インデックス番号でセル範囲を指定する方法
            .Cells(2, 2, 3, 3).Value = "XYZ"
        End With
    End Sub

    ' スタイルの設定：VBA互換APIを使う方法
    Private Sub SetStyleA()
        With FpSpread1.Sheets(0).AsWorksheet()
            ' 文字列でセル範囲を指定する方法
            .Range("A1:B2").Interior.Color = GrapeCity.Spreadsheet.Color.FromArgb(Color.Lavender.ToArgb())
            .Range("A1:B2").Font.Color = GrapeCity.Spreadsheet.Color.FromArgb(Color.Blue.ToArgb())

            ' インデックス番号でセル範囲を指定する方法
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 3, 3)}).Interior.ColorIndex = (3 - 1)
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 3, 3)}).Font.Color = GrapeCity.Spreadsheet.Color.FromIndexedColor(2 - 1)
        End With
    End Sub

    ' スタイルの設定：従来の方法
    Private Sub SetStyleB()
        With FpSpread1.Sheets(0)
            ' 文字列でセル範囲を指定する方法
            .Cells("A1:B2").BackColor = Color.Lavender
            .Cells("A1:B2").ForeColor = Color.Blue

            ' インデックス番号でセル範囲を指定する方法
            .Cells(2, 2, 3, 3).BackColor = Color.Red
            .Cells(2, 2, 3, 3).ForeColor = Color.White
        End With
    End Sub

    ' セル範囲のコピー：VBA互換APIを使う方法
    Private Sub CopyCellA()
        With FpSpread1.Sheets(0).AsWorksheet()
            ' 文字列でセル範囲を指定する方法
            .Range("A1:B2").Copy(destination:= .Range("A3"))

            ' インデックス番号でセル範囲を指定する方法
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 3, 3)}).Copy(destination:= .Range({New GrapeCity.Spreadsheet.Reference(4, 2, 4, 2)}))
        End With
    End Sub

    ' セル範囲のコピー：従来の方法
    Private Sub CopyCellB()
        With FpSpread1.Sheets(0)
            ' 文字列でセル範囲を指定する方法
            .Cells("A3:B4").Value = .Cells("A1:B2").Value
            .Cells("A3:B4").BackColor = .Cells("A1:B2").BackColor
            .Cells("A3:B4").ForeColor = .Cells("A1:B2").ForeColor

            ' インデックス番号でセル範囲を指定する方法
            .CopyRange(2, 2, 4, 2, 2, 2, False)
        End With
    End Sub

    ' セル範囲のクリア：VBA互換APIを使う方法
    Private Sub ClearCellA()
        With FpSpread1.Sheets("Sheet1").AsWorksheet()
            ' 文字列でセル範囲を指定する方法
            .Range("A1:B4").Clear()

            ' インデックス番号でセル範囲を指定する方法
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 5, 3)}).Clear()
        End With
    End Sub

    ' セル範囲のクリア：従来の方法
    Private Sub ClearCellB()
        With FpSpread1.Sheets("Sheet1")
            ' 文字列でセル範囲を指定する方法
            .Cells("A1:B4").ResetValue()
            .Cells("A1:B4").ResetBackColor()
            .Cells("A1:B4").ResetForeColor()

            ' インデックス番号でセル範囲を指定する方法
            .ClearRange(2, 2, 4, 2, False)
        End With
    End Sub

    ' セル範囲の選択：VBA互換APIを使う方法
    Private Sub SelectCellA()
        With FpSpread1.Sheets("Sheet1").AsWorksheet()
            '' 文字列でセル範囲を指定する方法
            '.Range("C3:D4").Select()

            ' インデックス番号でセル範囲を指定する方法
            .Range({New GrapeCity.Spreadsheet.Reference(2, 2, 3, 3)}).Select()
        End With
    End Sub

    ' セル範囲の選択：従来の方法
    Private Sub SelectCellB()
        With FpSpread1.Sheets("Sheet1")
            ' 文字列でセル範囲を指定する方法
            ' 該当する機能はありません

            ' インデックス番号でセル範囲を指定する方法
            '.SetActiveCell(2, 2) ' <== 明示的なアクティブ枠の設定
            .AddSelection(2, 2, 2, 2)
        End With
    End Sub

    ' シートのコピー：VBA互換APIを使う方法
    Private Sub CopySheetA()
        With FpSpread1.Sheets("Sheet1").AsWorksheet()
            ' ユニークなシート名が自動的に設定されます
            .Copy(position:=0)
        End With
    End Sub

    ' シートのコピー：従来の方法
    Private Sub CopySheetB()
        With FpSpread1.Sheets("Sheet1")
            ' ユニークなシート名を明示的に設定しないと例外が発生します
            Dim sheet = DirectCast(.Clone(), FarPoint.Win.Spread.SheetView)
            sheet.SheetName = "Sheet2"
            FpSpread1.Sheets.Insert(0, sheet)
            'FpSpread1.ActiveSheetIndex = 0 ' アクティブシートの指定
        End With
    End Sub

End Class
