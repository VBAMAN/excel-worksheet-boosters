Attribute VB_Name = "MAPBST_70_MapRestore"
'==============================================================================
' RunMapRestore
'------------------------------------------------------------------------------
' Paletteシートを参照し、背景色に対応するビットマップ番号を
' アクティブシートへ復元します。
'
' 用途
'   ・誤って削除した番号の復元
'   ・同じ配色の別マップへ番号を適用
'==============================================================================
Public Sub RunMapRestore()

    Dim wsMap As Worksheet
    Dim wsPalette As Worksheet

    MAPBST_Initialize
    
    If Not SheetExists(PALETTE_SHEET_NAME) Then

        MsgBox _
            "Paletteシートが見つかりません。" & vbCrLf & _
            "先にパレットを作成してください。", _
            vbExclamation

        Exit Sub

    End If

    Set wsPalette = Worksheets(PALETTE_SHEET_NAME)

    RestoreBitmapNumbers wsMap, wsPalette

End Sub

'==============================================================================
' SheetExists
'------------------------------------------------------------------------------
' 指定したシートが存在するか判定します。
'==============================================================================
Public Function SheetExists(ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Worksheets(sheetName)

    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function

'==============================================================================
' RestoreBitmapNumbers
'------------------------------------------------------------------------------
' Paletteを参照し、マップへビットマップ番号を書き込みます。
'==============================================================================
Private Sub RestoreBitmapNumbers( _
    ByVal wsMap As Worksheet, _
    ByVal wsPalette As Worksheet)

    Dim BitmapDict As Object

    Set BitmapDict = CreateObject("Scripting.Dictionary")

    BuildBitmapDictionary wsPalette, BitmapDict

    WriteBitmapNumbers wsMap, BitmapDict

End Sub

'==============================================================================
' BuildBitmapDictionary
'------------------------------------------------------------------------------
' Paletteシートから
'
'     背景色 → ビットマップ番号
'
' のDictionaryを作成します。
'==============================================================================
Private Sub BuildBitmapDictionary( _
    ByVal wsPalette As Worksheet, _
    ByRef BitmapDict As Object)

    Dim lastRow As Long
    Dim r As Long

    Dim ColorCode As Long
    Dim BitmapNo As Variant

    lastRow = wsPalette.Cells(wsPalette.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow

        ColorCode = wsPalette.Cells(r, 2).Interior.Color
        BitmapNo = wsPalette.Cells(r, 1).Value

        If Len(Trim(CStr(BitmapNo))) > 0 Then

            BitmapDict(ColorCode) = BitmapNo

        End If

    Next r

End Sub

'==============================================================================
' WriteBitmapNumbers
'------------------------------------------------------------------------------
' 背景色を参照し、対応するビットマップ番号を書き込みます。
'==============================================================================
Private Sub WriteBitmapNumbers( _
    ByVal wsMap As Worksheet, _
    ByVal BitmapDict As Object)

    Dim Cell As Range
    Dim ColorCode As Long

    For Each Cell In wsMap.UsedRange

        ColorCode = Cell.Interior.Color

        If BitmapDict.Exists(ColorCode) Then

            Cell.Value = BitmapDict(ColorCode)

        End If

    Next Cell

    MsgBox "ビットマップ番号を復元しました。", vbInformation

End Sub

