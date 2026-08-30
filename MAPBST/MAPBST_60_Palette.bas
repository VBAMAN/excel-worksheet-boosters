Attribute VB_Name = "MAPBST_60_Palette"
'==============================================================================
' RunPalette
'------------------------------------------------------------------------------
' パレット生成を実行します。
'
' アクティブシートで使用されている背景色を収集し、
' パレットシートへ一覧を作成します。
'
' 出力内容
'   ・ビットマップ番号
'   ・背景色
'   ・RGBコード
'
' ※ 色数が設定値を超える場合は確認ダイアログを表示します。
'==============================================================================
Public Sub RunPalette()

    Dim wsSource As Worksheet
    Dim wsPalette As Worksheet

    '----------------------------------------------------------
    ' 初期設定
    '----------------------------------------------------------
    MAPBST_Initialize

    Set wsSource = ActiveSheet
    Set wsPalette = GetPaletteSheet()

    '----------------------------------------------------------
    ' パレット作成
    '----------------------------------------------------------
    CreatePalette wsSource, wsPalette

End Sub

'==============================================================================
' CreatePalette
'------------------------------------------------------------------------------
' 指定したシートで使用されている背景色を収集し、
' パレットシートへ一覧を出力します。
'
' 出力内容
'   A列 : ビットマップ番号（元シートの値）
'   B列 : 背景色
'   C列 : RGBコード
'
' ※ 白色(RGB255,255,255)は対象外です。
'==============================================================================
Public Sub CreatePalette( _
    ByVal wsSource As Worksheet, _
    ByVal wsPalette As Worksheet)

    Dim SourceRange As Range
    Dim Cell As Range

    Dim ColorDict As Object
    Dim ValueDict As Object

    Dim ColorCode As Long
    Dim CellValue As Variant

    Set ColorDict = CreateObject("Scripting.Dictionary")
    Set ValueDict = CreateObject("Scripting.Dictionary")

    '----------------------------------------------------------
    ' パレットシート初期化
    '----------------------------------------------------------
    wsPalette.Cells.ClearContents
    
    wsPalette.Cells.NumberFormat = "@"
    wsPalette.Columns("A:C").ColumnWidth = 16

    Set SourceRange = wsSource.UsedRange

    '----------------------------------------------------------
    ' 使用色を収集
    '----------------------------------------------------------
    For Each Cell In SourceRange

        ColorCode = Cell.Interior.Color
        CellValue = Cell.Value

        ' 白セルは除外
        If ColorCode <> RGB(255, 255, 255) Then

            If Not ColorDict.Exists(ColorCode) Then
                ColorDict.Add ColorCode, True
            End If

            ' 最初に見つかった番号だけ保持
            If Not ValueDict.Exists(ColorCode) Then

                If Trim(CellValue & "") <> "" Then
                    ValueDict.Add ColorCode, CellValue
                End If

            End If

        End If

    Next Cell

    '----------------------------------------------------------
    ' 色数確認
    '----------------------------------------------------------
    If ColorDict.Count >= PALETTE_CONFIRM_LIMIT Then

        If MsgBox( _
            "使用色は " & ColorDict.Count & " 色あります。" & vbCrLf & _
            "パレットを作成しますか？", _
            vbYesNo + vbQuestion, _
            "Palette") = vbNo Then

            Exit Sub

        End If

    End If

    '----------------------------------------------------------
    ' パレット出力
    '----------------------------------------------------------
    WritePalette wsPalette, ColorDict, ValueDict

    MsgBox ColorDict.Count & " 色のパレットを作成しました。", vbInformation

End Sub

'==============================================================================
' BuildPaletteDictionary
'------------------------------------------------------------------------------
' ワークシートで使用されている背景色を収集します。
'
' ColorDict
'   Key   : 背景色(Color)
'   Item  : True
'
' ValueDict
'   Key   : 背景色(Color)
'   Item  : 最初に見つかったセルの値
'
' ※ 白色(RGB255,255,255)は対象外です。
'==============================================================================
Private Sub BuildPaletteDictionary( _
    ByVal ws As Worksheet, _
    ByRef ColorDict As Object, _
    ByRef ValueDict As Object)

    Dim SourceRange As Range
    Dim Cell As Range

    Dim ColorCode As Long
    Dim CellValue As Variant

    Set SourceRange = ws.UsedRange

    For Each Cell In SourceRange

        ColorCode = Cell.Interior.Color
        CellValue = Cell.Value

        '----------------------------------------------------------
        ' 白セルは対象外
        '----------------------------------------------------------
        If ColorCode <> RGB(255, 255, 255) Then

            ' 使用色を登録
            If Not ColorDict.Exists(ColorCode) Then

                ColorDict.Add ColorCode, True

            End If

            ' 最初に見つかった番号を保持
            If Not ValueDict.Exists(ColorCode) Then

                If Len(Trim(CStr(CellValue))) > 0 Then

                    ValueDict.Add ColorCode, CellValue

                End If

            End If

        End If

    Next Cell

End Sub

'==============================================================================
' WritePalette
'------------------------------------------------------------------------------
' パレット情報をシートへ出力します。
'
' 出力形式
'   A列 : ビットマップ番号
'   B列 : 背景色
'   C列 : RGBコード
'
' 引数
'   wsPalette : パレット出力先シート
'   ColorDict : 使用色一覧
'   ValueDict : ビットマップ番号一覧
'==============================================================================
Private Sub WritePalette( _
    ByVal wsPalette As Worksheet, _
    ByVal ColorDict As Object, _
    ByVal ValueDict As Object)

    Dim ColorCode As Variant

    Dim r As Long
    Dim g As Long
    Dim b As Long

    Dim RowNo As Long

    '----------------------------------------------------------
    ' 見出し
    '----------------------------------------------------------
    wsPalette.Range("A1").Value = "Bitmap"
    wsPalette.Range("B1").Value = "Color"
    wsPalette.Range("C1").Value = "RGB"

    RowNo = 2

    '----------------------------------------------------------
    ' パレット出力
    '----------------------------------------------------------
    For Each ColorCode In ColorDict.Keys

        ' ビットマップ番号
        If ValueDict.Exists(ColorCode) Then

            wsPalette.Cells(RowNo, 1).Value = ValueDict(ColorCode)

        End If

        ' 背景色
        wsPalette.Cells(RowNo, 2).Interior.Color = ColorCode

        ' RGBコード
        r = ColorCode Mod 256
        g = (ColorCode \ 256) Mod 256
        b = (ColorCode \ 65536) Mod 256

        wsPalette.Cells(RowNo, 3).Value = r & "," & g & "," & b

        RowNo = RowNo + 1

    Next ColorCode

    '----------------------------------------------------------
    ' 列幅調整
    '----------------------------------------------------------
    wsPalette.Columns("A:C").AutoFit

End Sub

'==============================================================================
' GetPaletteSheet
'------------------------------------------------------------------------------
' パレットシートを取得します。
'
' シートが存在しない場合は自動的に作成します。
'==============================================================================
Public Function GetPaletteSheet() As Worksheet

    On Error Resume Next
    Set GetPaletteSheet = ThisWorkbook.Worksheets(PALETTE_SHEET_NAME)
    On Error GoTo 0

    If GetPaletteSheet Is Nothing Then

        Set GetPaletteSheet = ThisWorkbook.Worksheets.Add( _
            After:=Worksheets(Worksheets.Count))

        GetPaletteSheet.Name = PALETTE_SHEET_NAME

    End If
    
End Function
