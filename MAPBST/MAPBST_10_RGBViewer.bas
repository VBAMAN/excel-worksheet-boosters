Attribute VB_Name = "MAPBST_10_RGBViewer"
Option Explicit

'==============================================================================
' MAPBST_10_RGBViewer
'------------------------------------------------------------------------------
' Excel Map Booster
'
' アクティブセルの背景色を RGB 値として表示します。
'
' 【機能】
' ・アクティブセルの背景色を取得
' ・RGB表示用シェイプが存在しない場合は自動作成
' ・表示位置を更新
'
' シェイプは MAPBST_00_Config の設定に従って表示されます。
'==============================================================================

'------------------------------------------------------------------------------
' RGB表示
'------------------------------------------------------------------------------
Public Sub ShowActiveCellColor()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim RGB_STRING As String

    Set ws = ActiveSheet

    ' アクティブセルのRGB文字列を取得
    RGB_STRING = GetActiveCellColorRGB()

    ' シェイプ取得（存在しなければ作成）
    Set shp = GetRGBShape(ws)

    ' 表示位置更新
    With shp

        ' 表示列
        .Left = ws.Cells(1, RGB_SHAPE_COLUMN).Left

        ' アクティブセルの高さへ移動
        .Top = ActiveCell.Top

        ' 表示文字列
        .TextFrame2.TextRange.Text = RGB_STRING
        .TextFrame2.TextRange.Font.Size = 12

    End With

End Sub

'------------------------------------------------------------------------------
' アクティブセルのRGBコード取得
'
' 戻り値
'   RGB(216,216,216)
'------------------------------------------------------------------------------
Private Function GetActiveCellColorRGB() As String

    Dim colorValue As Long

    Dim RedValue As Long
    Dim GreenValue As Long
    Dim BlueValue As Long

    colorValue = ActiveCell.Interior.Color

    RedValue = colorValue Mod 256
    GreenValue = (colorValue \ 256) Mod 256
    BlueValue = (colorValue \ 65536) Mod 256

    GetActiveCellColorRGB = _
        "RGB (" & _
        RedValue & "," & _
        GreenValue & "," & _
        BlueValue & ")"

End Function

'------------------------------------------------------------------------------
' RGB表示シェイプ取得
'
' シェイプが存在しない場合は自動生成します。
'------------------------------------------------------------------------------
Private Function GetRGBShape(ws As Worksheet) As Shape

    On Error Resume Next
    Set GetRGBShape = ws.Shapes(RGB_SHAPE_NAME)
    On Error GoTo 0

    If GetRGBShape Is Nothing Then

        Set GetRGBShape = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            ws.Cells(1, RGB_SHAPE_COLUMN).Left, _
            ws.Cells(1, RGB_SHAPE_COLUMN).Top, _
            120, _
            30)

        With GetRGBShape

            .Name = RGB_SHAPE_NAME

            .Fill.ForeColor.RGB = RGB(27, 27, 27)

            .Line.Visible = msoFalse

            .TextFrame2.TextRange.Font.Size = 16

        End With

    End If

End Function

