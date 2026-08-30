Attribute VB_Name = "MAPBST_80_Reset"
'==============================================================================
' RunReset
'------------------------------------------------------------------------------
' アクティブセルと同じ背景色を持つセルのビットマップ番号を削除します。
'
' ・背景色は保持します。
' ・セルの値のみ削除します。
'==============================================================================
Public Sub RunReset()

    Dim ws As Worksheet
    Dim TargetColor As Long

    Set ws = ActiveSheet

    TargetColor = ActiveCell.Interior.Color

    If MsgBox( _
        "同じ背景色のビットマップ番号を削除しますか？", _
        vbYesNo + vbQuestion, _
        "Reset") = vbNo Then

        Exit Sub

    End If

    ResetBitmapNumbers ws, TargetColor

End Sub

'==============================================================================
' ResetBitmapNumbers
'------------------------------------------------------------------------------
' 指定した背景色を持つセルの値を削除します。
'------------------------------------------------------------------------------
Private Sub ResetBitmapNumbers( _
    ByVal ws As Worksheet, _
    ByVal TargetColor As Long)

    Dim Cell As Range
    Dim Count As Long

    For Each Cell In ws.UsedRange

        If Cell.Interior.Color = TargetColor Then

            Cell.ClearContents

            Count = Count + 1

        End If

    Next Cell

    MsgBox Count & " 個のセルをクリアしました。", vbInformation

End Sub

