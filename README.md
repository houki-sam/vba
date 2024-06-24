'--------------------------------------------------------
' 指定フォルダ内の全てのブックから任意の文字列を検索する
'--------------------------------------------------------
' https://excel.syogyoumujou.com/vba/find_allbooks.html
'--------------------------------------------------------
Sub searchAllBooksForAnyString() 'メイン
    '--------------------------------
    ' 検索する文字列を配列として設定
    '--------------------------------
    Dim varArray As Variant
    varArray = Array("富山", "神奈川") '検索文字列

    '--------------------------------
    ' フォルダの選択
    '--------------------------------
    Dim strFolderPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then strFolderPath = .SelectedItems(1)
    End With
    If Len(strFolderPath) = 0 Then Exit Sub

    '--------------------------------
    ' フォルダの存在確認
    '--------------------------------
    If Dir(strFolderPath, vbDirectory) = "" Then
        MsgBox "対象のフォルダが見つかりません", vbExclamation, "終了します"
        Exit Sub
    End If

    '--------------------------------
    ' フォルダ内ブックを検索
    '--------------------------------
    Dim strFileName As String
    strFolderPath = strFolderPath & Application.PathSeparator 'フォルダパスに区切り文字追加
    strFileName = Dir(strFolderPath & "*.xls?")               'フォルダからExcelブックを検索
    If strFileName = "" Then                                  'ブックのパスを取得できなければ終了
        MsgBox "指定フォルダ内にExcelブックが見つかりません", vbExclamation, "終了します"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False                        '画面更新無効
    Application.EnableEvents = False                          'イベント無効

    '--------------------------------
    ' 新規ブック追加・見出し設定
    '--------------------------------
    Dim shtWrite As Worksheet '書き込みシート
    Set shtWrite = Workbooks.Add.Worksheets(1)
    shtWrite.Range("A4:C4").Value = Array("検索値", "ブック名", "シート名")
    shtWrite.Range("1:1,4:4").Interior.Color = RGB(217, 225, 242)
    
    '--------------------------------
    ' ブック内から文字列を検索
    '--------------------------------
    Dim bokTarget As Workbook
    Dim shtTarget As Worksheet
    Dim rngTarget As Range
    Dim varWhat   As Variant
    Dim lngCount  As Long
On Error Resume Next
    Do
        'フォルダ内のブックを開く
        Set bokTarget = Workbooks.Open(strFolderPath & strFileName)
        
        '--------------------------------
        'ブックの各シートで検索を実行
        '--------------------------------
        For Each shtTarget In bokTarget.Worksheets
            For Each varWhat In varArray
                '対象シートの全てのセルから任意の文字列を検索
                Set rngTarget = findTargetCell(shtTarget.Cells, varWhat)
                
                '検索に一致するセルが存在する場合は新規ブックに情報を書き込み
                If Not rngTarget Is Nothing Then
                    With shtWrite.Cells(5 + lngCount, "A").Resize(1, 4)
                        .Value = Array(varWhat, bokTarget.Name, shtTarget.Name, rngTarget.Address(0, 0))
                        lngCount = lngCount + 1
                    End With
                End If
            Next
        Next
        
        '開いたブックを保存せずに閉じる
        bokTarget.Close SaveChanges:=False
        
        strFileName = Dir()     '次のExcelブックを検索
    Loop Until strFileName = "" 'ブックが見つからなければループから抜ける
    
    strFileName = Dir("")
On Error GoTo 0

    '--------------------------------
    ' 検索値が見つからなければ終了
    '--------------------------------
    If lngCount = 0 Then
        shtWrite.Parent.Close SaveChanges:=False
        MsgBox "検索値は見つかりませんでした", vbInformation
        GoTo LBL_FINALLY
    End If

    '--------------------------------
    ' 新規ブックレイアウト調整
    '--------------------------------
    shtWrite.Columns(4).TextToColumns Destination:=Range("D1"), Comma:=True   'セルアドレスをコンマで分割
    shtWrite.UsedRange.EntireColumn.AutoFit                                   '列幅を自動調整する
    shtWrite.Range("A1").Value = "フォルダ"                                   'タイトルを設定
    shtWrite.Range("A2").Value = Left$(strFolderPath, Len(strFolderPath) - 1) 'フォルダパス設定
    shtWrite.Range("D4").Value = "セルアドレス"
    
LBL_FINALLY:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'-------------------------------------------------------------------------------------
' 対象セル範囲から任意の文字列を検索するプロシージャ
'-------------------------------------------------------------------------------------
'［引数］
'   rngTarget      ：対象セル範囲
'   What           ：検索する文字列
'   LookIn         ：情報種類　値：xlValues[既定]　数式：xlFormulas　コメント文：xlComments
'   LookAt         ：一致の種類    部分一致：xlPart[既定]　全体一致：xlWhole
'   SearchOrder    ：検索方法   1行ごと検索：xlByRows[既定]　1列ごと検索：xlByColumns
'   SearchDirection：検索順  一致する次の値：xlNext[既定]　一致する前の値：xlPrevious
'   MatchCase      ：大文字・小文字の区別  区別する：True　区別しない：False[既定]
'   MatchByte      ：全角・半角の区別      区別する：True　区別しない：False[既定]
'［戻り値］
'   検索値のセルの集合　検索値がない場合はNothing
'［作成日］2023.12.19　［更新日］2023.12.22
' https://excel.syogyoumujou.com/vba/find_allbooks.html
'-------------------------------------------------------------------------------------
Function findTargetCell(ByRef rngTarget As Range, _
                        ByVal What As String, _
                        Optional ByVal LookIn As XlFindLookIn = xlValues, _
                        Optional ByVal LookAt As XlLookAt = xlPart, _
                        Optional ByVal SearchOrder As XlSearchOrder = xlByRows, _
                        Optional ByVal SearchDirection As XlSearchDirection = xlNext, _
                        Optional ByVal MatchCase As Boolean = False, _
                        Optional ByVal MatchByte As Boolean = False) As Range
    '検索実行
    Dim rngFind As Range
    Set rngFind = rngTarget.Find(What, , LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte)
    
    '検索に一致のセルがない場合は抜ける
    If rngFind Is Nothing Then Exit Function
    
    Dim strAddress As String
    Dim rngUnion   As Range
    strAddress = rngFind.Address                  '最初に検索一致したセルのアドレスを取得
    Set rngUnion = rngFind
    Do
        Set rngUnion = Union(rngUnion, rngFind)   'セルを集合
        Set rngFind = rngTarget.FindNext(rngFind) '次の一致セルを検索
        If rngFind Is Nothing Then Exit Do
    Loop Until strAddress = rngFind.Address       'セルアドレスが最初のセルと同じ場合はループを抜ける
    
    Set findTargetCell = rngUnion
    
End Function
