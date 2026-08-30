Attribute VB_Name = "MAPBST_50_DominoPaint"
Option Explicit

'==============================================================================
' MAPBST_50_DominoPaint
'------------------------------------------------------------------------------
' Excel Map Booster
'
' Domino Paint
'
' アクティブセルと同じ背景色で上下左右に連結されたセルへ、
' アクティブセルの値を一括入力します。
'
' 昔のペイントソフトの「塗りつぶし(Flood Fill)」と同じ動作です。
'==============================================================================

Public Sub RunDominoPaint()

    MAPBST_Initialize

    FloodFillSearch ActiveSheet, ActiveCell

End Sub

'==============================================================================
' FloodFillSearch
'------------------------------------------------------------------------------
' 開始セルと同じ背景色で上下左右に連結されたセルを探索し、
' アクティブセルの値を一括入力します。
'
' ・探索には幅優先探索（Breadth First Search）を使用
' ・上下左右のみを連結対象とします
'------------------------------------------------------------------------------
Private Sub FloodFillSearch( _
    ByVal ws As Worksheet, _
    ByVal StartCell As Range)

    Dim TargetColor As Long
    Dim TargetValue As Variant

    Dim Visited As Object
    Dim Queue As Collection

    Dim CurrentCell As Range
    Dim NextCell As Range

    Dim dr As Variant
    Dim dc As Variant

    Dim NewRow As Long
    Dim NewCol As Long

    Dim Key As String

    Dim Count As Long
    Dim i As Long

    '----------------------------------------------------------
    ' 初期設定
    '----------------------------------------------------------
    TargetColor = StartCell.Interior.Color
    TargetValue = StartCell.Value

    Set Visited = CreateObject("Scripting.Dictionary")
    Set Queue = New Collection

    Queue.Add StartCell
    Visited.Add StartCell.Address(False, False), True

    ' 上下左右
    dr = Array(-1, 1, 0, 0)
    dc = Array(0, 0, -1, 1)

    '----------------------------------------------------------
    ' 幅優先探索（Flood Fill）
    '----------------------------------------------------------
    Do While Queue.Count > 0

        Set CurrentCell = Queue(1)
        Queue.Remove 1

        Count = Count + 1

        For i = 0 To 3

            NewRow = CurrentCell.Row + dr(i)
            NewCol = CurrentCell.Column + dc(i)

            If NewRow >= 1 And NewCol >= 1 Then

                Set NextCell = ws.Cells(NewRow, NewCol)

                Key = NextCell.Address(False, False)

                If Not Visited.Exists(Key) Then

                    If NextCell.Interior.Color = TargetColor Then

                        Visited.Add Key, True
                        Queue.Add NextCell

                    End If

                End If

            End If

        Next i

    Loop

    '----------------------------------------------------------
    ' 大量セル確認
    '----------------------------------------------------------
    If Count >= DOMINO_CONFIRM_LIMIT Then

        If MsgBox( _
            Count & " 個のセルが対象です。" & vbCrLf & _
            "値を書き込みますか？", _
            vbYesNo + vbQuestion, _
            "Domino Paint") = vbNo Then

            Exit Sub

        End If

    End If

    '----------------------------------------------------------
    ' 一括書き込み
    '----------------------------------------------------------
    WriteDominoValue ws, Visited, TargetValue
    
    MsgBox Count & " 個のセルへ値を入力しました。", vbInformation

End Sub

'==============================================================================
' WriteDominoValue
'------------------------------------------------------------------------------
' FloodFillSearch で取得したセル一覧へ値を一括書き込みします。
'==============================================================================
Private Sub WriteDominoValue( _
    ByVal ws As Worksheet, _
    ByVal TargetCells As Object, _
    ByVal Value As Variant)

    Dim Key As Variant

    For Each Key In TargetCells.Keys

        ws.Range(Key).Value = Value

    Next Key

End Sub

