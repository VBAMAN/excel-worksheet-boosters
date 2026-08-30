Attribute VB_Name = "MAPBST_40_BlockFill"
Option Explicit

'==============================================================================
' MAPBST_40_BlockFill
'------------------------------------------------------------------------------
' Excel Map Booster
'
' 2×2 の正方形ブロックへ4種類の番号を配置します。
'
' 配置例
'
'     80 81
'     82 83
'
' 正方形の検索は MAPBST_30_SquareCounter が担当します。
' このモジュールは番号を書き込む処理のみ行います。
'==============================================================================

'------------------------------------------------------------------------------
' BlockFill 実行
'------------------------------------------------------------------------------
Public Sub RunBlockFill()

    Dim ws As Worksheet

    Dim MapTop As Long
    Dim MapBottom As Long

    Dim StartRow As Long
    Dim StartCol As Long

    Dim r As Long
    Dim c As Long

    Set ws = ActiveSheet

    ' 初期化
    MAPBST_Initialize

    '==============================================================
    ' マップ単位で処理
    '==============================================================
    For MapTop = 1 To MAP_LAST_ROW Step MAP_HEIGHT

        MapBottom = MapTop + MAP_HEIGHT - 1

        '----------------------------------------------------------
        ' このマップの2×2グリッドの揃い方を取得
        '----------------------------------------------------------
        If Not FindMapAlignment(ws, MapTop, StartRow, StartCol) Then

            ' 正方形が1つも無いマップはスキップ
            GoTo NextMap

        End If

        '----------------------------------------------------------
        ' このマップのみ処理
        '----------------------------------------------------------
        For r = MapTop + StartRow - 1 To MapBottom Step 2

            For c = StartCol To MAP_WIDTH Step 2

                If IsSquare(ws, r, c) Then

                    WriteSquare ws, r, c

                End If

            Next c

        Next r

NextMap:

    Next MapTop

    MsgBox "Block Fill が完了しました。", vbInformation

End Sub

'==============================================================================
' FindBaseSquare
'------------------------------------------------------------------------------
' マップ内で最初に見つかった2×2正方形を取得します。
'
' この正方形を基準として、
' 以降の2セル単位の探索を行います。
'
' 戻り値
'   True  : 基準正方形を発見
'   False : 見つからなかった
'==============================================================================
Public Function FindBaseSquare( _
    ByVal ws As Worksheet, _
    ByRef BaseRow As Long, _
    ByRef BaseCol As Long) As Boolean

    Dim r As Long
    Dim c As Long

    For r = 1 To MAP_HEIGHT - 1

        For c = 1 To MAP_WIDTH - 1

            If IsSquare(ws, r, c) Then

                BaseRow = r
                BaseCol = c

                FindBaseSquare = True

                Exit Function

            End If

        Next c

    Next r

End Function

'==============================================================================
' FindMapAlignment
'------------------------------------------------------------------------------
' 指定したマップ内で最初に見つかった 2×2 正方形から、
' マップ内の 2 セルグリッドの開始位置（奇数／偶数）を取得します。
'==============================================================================
Public Function FindMapAlignment( _
    ByVal ws As Worksheet, _
    ByVal MapTop As Long, _
    ByRef StartRow As Long, _
    ByRef StartCol As Long) As Boolean

    Dim r As Long
    Dim c As Long

    Dim LocalRow As Long
    Dim LocalCol As Long

    For r = MapTop To MapTop + MAP_HEIGHT - 2

        For c = 1 To MAP_WIDTH - 1

            If IsSquare(ws, r, c) Then

                '------------------------------
                ' マップ内での相対位置
                '------------------------------
                LocalRow = r - MapTop + 1
                LocalCol = c

                If (LocalRow Mod 2) = 0 Then
                    StartRow = 2
                Else
                    StartRow = 1
                End If

                If (LocalCol Mod 2) = 0 Then
                    StartCol = 2
                Else
                    StartCol = 1
                End If

                FindMapAlignment = True
                Exit Function

            End If

        Next c

    Next r

End Function

'==============================================================================
' IsSquare
'------------------------------------------------------------------------------
' 指定した位置が 2×2 の正方形か判定します。
'
' 判定方法
'   ・左上を基準とした 2×2 のセルを取得
'   ・4セルすべてが対象色であること
'
' 引数
'   ws  : 対象ワークシート
'   Row : 左上セルの行番号
'   Col : 左上セルの列番号
'
' 戻り値
'   True  : 2×2 の正方形
'   False : 正方形ではない
'==============================================================================
Public Function IsSquare( _
    ByVal ws As Worksheet, _
    ByVal Row As Long, _
    ByVal Col As Long) As Boolean

    Dim targetRange As Range

    ' シート範囲外を防ぐ
    If Row >= ws.Rows.Count Then Exit Function
    If Col >= ws.Columns.Count Then Exit Function

    ' 2×2 の範囲を取得
    Set targetRange = ws.Range( _
        ws.Cells(Row, Col), _
        ws.Cells(Row + 1, Col + 1))

    ' 4セルすべてが対象色なら True
    IsSquare = IsAllSame(targetRange)

End Function

'==============================================================================
' IsAllSame
'------------------------------------------------------------------------------
' 指定したセル範囲が、すべて対象色で構成されているか判定します。
'
' 本ツールでは 2×2 の正方形判定に使用しますが、
' 任意のセル範囲にも利用できます。
'
' 引数
'   TargetRange : 判定するセル範囲
'
' 戻り値
'   True  : 全セルが対象色
'   False : 対象色以外のセルが含まれる
'==============================================================================
Private Function IsAllSame( _
    ByVal targetRange As Range) As Boolean

    Dim Cell As Range

    ' すべてのセルを確認
    For Each Cell In targetRange.Cells

        If Cell.Interior.Color <> SQUARE_TARGET_COLOR Then

            IsAllSame = False
            Exit Function

        End If

    Next Cell

    ' 全セル一致
    IsAllSame = True

End Function

'------------------------------------------------------------------------------
' 2×2 ブロックへ番号を書き込む
'------------------------------------------------------------------------------
Private Sub WriteSquare( _
    ws As Worksheet, _
    ByVal Row As Long, _
    ByVal Col As Long)

    ws.Cells(Row, Col).Value = SQUARE_NUMBERS(0)
    ws.Cells(Row, Col + 1).Value = SQUARE_NUMBERS(1)

    ws.Cells(Row + 1, Col).Value = SQUARE_NUMBERS(2)
    ws.Cells(Row + 1, Col + 1).Value = SQUARE_NUMBERS(3)

End Sub
