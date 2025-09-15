Sub AggregateData_UserSelected_Ver4()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim key As String
    Dim cell As Range
    Dim outputRow As Long
    Dim colCount As Long
    Dim dictKey As Variant
    Dim rowIndex As Long
    Dim firstRow As Object ' 最初の行を記録するためのDictionary
    Dim selectedWorkbook As Workbook
    Dim selectedFileName As Variant
    Dim colIndex As Long
    Dim infoFileName As Variant
    Dim infoWorkbook As Workbook
    Dim infoWs As Worksheet
    Dim infoLastRow As Long
    Dim infoDict As Object

    ' 処理するブックを選択
    selectedFileName = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "Select the workbook cells to be processed")
    If selectedFileName = False Then Exit Sub ' ユーザーがキャンセルした場合終了
    
    ' 選択したブックを開く
    Set selectedWorkbook = Workbooks.Open(selectedFileName)
    
    ' 処理するシートを選択
    On Error Resume Next
    Set ws = Application.InputBox("Select the workbook cells to be processed." & vbCrLf & "※The cells to be selected can be anywhere.", Type:=8).Parent
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No sheet selected. Exiting.", vbExclamation
        Exit Sub
    End If

    ' 追加情報ファイルを冒頭で選択
    infoFileName = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , _
                    "Select the workbook containing additional information")
    If infoFileName <> False Then
        Set infoWorkbook = Workbooks.Open(infoFileName)
        Set infoWs = infoWorkbook.Sheets(1)
        infoLastRow = infoWs.Cells(infoWs.Rows.Count, "A").End(xlUp).Row

        Set infoDict = CreateObject("Scripting.Dictionary")
        For Each cell In infoWs.Range("A2:A" & infoLastRow)
            key = CStr(cell.Value)
            If Not infoDict.Exists(key) Then
                infoDict.Add key, infoWs.Cells(cell.Row, "T").Value
            End If
        Next cell

        infoWorkbook.Close SaveChanges:=False
    End If
    
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.EnableEvents = False
    
    ' 新しいシートを作成
    Set newWs = selectedWorkbook.Sheets.Add
    newWs.Name = "Aggregated Data"
    
    ' 元のシートの列数を取得
    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' 元のすべての列のヘッダーを新しいシートにコピー
    ws.Rows(1).Copy Destination:=newWs.Rows(1)
    
    ' Dictionaryを初期化
    Set dict = CreateObject("Scripting.Dictionary")
    Set firstRow = CreateObject("Scripting.Dictionary")
    
    ' データ範囲の最後の行を取得
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' データをループして条件に一致するものを集計
    For Each cell In ws.Range("R2:R" & lastRow)
        If cell.Value = 100 Or cell.Value = 200 Or (cell.Value = 650 And Right(CStr(ws.Cells(cell.Row, "B").Value), 3) = "-SK") Then
            key = CStr(ws.Cells(cell.Row, "B").Value)
            
            If Not dict.exists(key) Then
                dict.Add key, 0
                firstRow.Add key, cell.Row ' 最初の行番号を記録
            End If
            ' G列の数量を加算
            dict(key) = dict(key) + ws.Cells(cell.Row, "G").Value
        End If
    Next cell
    
    ' 集計結果を新しいシートに書き込み
    outputRow = 2
    For Each dictKey In dict.keys
        rowIndex = firstRow(dictKey)
        For colIndex = 1 To colCount
            newWs.Cells(outputRow, colIndex).Value = ws.Cells(rowIndex, colIndex).Value
            
            If colIndex = 7 Then ' G列の合計値を挿入
                newWs.Cells(outputRow, colIndex).Value = dict(dictKey)
            End If
        Next colIndex
        outputRow = outputRow + 1
    Next dictKey
    
    ' ★ここから列操作の順序を大幅に変更★
    
    ' 1. 不要な列を削除 (これによってC,D列が現在の位置に固定される)
    ' newWs.Range("A:A,F:F,I:T").Delete Shift:=xlToLeft
    ' 上記の削除範囲にI:Tが含まれるため、元シートのI列以降のデータが削除される。
    ' その後でE:FをC:Dに挿入すると位置がずれる可能性があるため、列の指定を正確にする
    
    ' まず、元のシートのG列に相当する集計値が入っている新しいシートのG列を考慮
    ' 元のI列以降（I:T）の削除
    newWs.Range("I:T").Delete Shift:=xlToLeft ' 元のI列からT列までを削除 (G列は集計値が入っているため残す)

    ' 元のF列の削除
    newWs.Columns("F").Delete Shift:=xlToLeft
    
    ' 元のA列の削除
    newWs.Columns("A").Delete Shift:=xlToLeft
    
    ' これにより、画像左側の結果に近い状態になるはず
    ' ただし、元のG列（集計値が入る列）はそのままの位置に保持される
    
    ' 2. E,F列をC,D列に移動 (現在の列位置に基づいて)
    '    ※A,B,C,D,E,F... の状態で、EとFをCとDの間に挿入
    '    つまり、元のC,Dが右にずれて、元E,Fが新しいC,Dになる
    newWs.Columns("E:F").Cut
    newWs.Columns("C:D").Insert Shift:=xlToRight

    ' ★ここまでの処理で、左の画像のレイアウトになったと仮定して、C,D列をG,H列にコピーする★
    ' この時点でのC列とD列は、元々のシートのE列とF列だったデータ
    
    ' 3. 現在のC列とD列をG列とH列にコピーする
    '    ※G,H列にデータがないことを想定 (もし既存データがあれば上書き)
    newWs.Columns("C:D").Copy Destination:=newWs.Columns("G:H")

    ' 4. ヘッダー名の設定とセルの色付け (現在のC1とG1の位置に適用)
    newWs.Range("C1").Value = "Qty.Actual"
    newWs.Range("C1").Interior.Color = RGB(255, 255, 0) ' 黄色
    
    newWs.Range("G1").Value = "Qty.BOM Data"
    newWs.Range("G1").Interior.Color = RGB(255, 165, 0) ' オレンジ

    ' 4b. 追加情報列の挿入 (I列) と冒頭で選択したファイルからの値の取得
    newWs.Columns("I").Insert
    newWs.Range("I1").Value = "Note"
    newWs.Range("I1").Interior.Color = RGB(0, 255, 0) ' 緑

    If Not infoDict Is Nothing Then
        For rowIndex = 2 To outputRow - 1
            key = CStr(newWs.Cells(rowIndex, "A").Value)
            If infoDict.Exists(key) Then
                newWs.Cells(rowIndex, "I").Value = infoDict(key)
            End If
        Next rowIndex
    End If

    ' 5. オートフィルターとソート
    newWs.Range("A1").AutoFilter
    ' ソートキーがRange("E1")なので、この時点でのE列のヘッダーがソート対象になる
    newWs.Range("A1").Sort key1:=newWs.Range("E1"), order1:=xlAscending, Header:=True

    ' 6. 列幅の自動調整 (A:I列までで十分か、必要に応じて調整)
    newWs.Columns("A:I").EntireColumn.AutoFit
    
    ' 7. 1行目でウインドウ枠の固定をする
    newWs.Activate ' 新しいシートをアクティブにする
    With ActiveWindow
        .FreezePanes = False ' 既存の固定を解除
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True ' 1行目を固定
    End With
    
    MsgBox "Aggregation has been completed！", vbInformation

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.EnableEvents = True

End Sub
