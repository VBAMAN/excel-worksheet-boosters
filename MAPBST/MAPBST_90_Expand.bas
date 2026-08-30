Attribute VB_Name = "MAPBST_90_Expand"
'==============================================================================
' RunExpandMap
'------------------------------------------------------------------------------
' マップサイズを変更した新しいシートを作成します。
'
' 元シートは変更しません。
'==============================================================================
Public Sub RunExpandMap()

    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet

    MAPBST_Initialize

    Set wsSource = ActiveSheet
    
    Dim ColumnWidths As Variant
    ColumnWidths = GetColumnWidths(wsSource)

    Set wsTarget = CreateExpandedSheet(wsSource)

    ExpandMapLayout wsSource, wsTarget
    
    RestoreColumnWidths wsTarget, ColumnWidths

    MsgBox "マップサイズ変更が完了しました。" & vbCrLf & _
           wsTarget.Name, vbInformation

End Sub

'==============================================================================
' CreateExpandedSheet
'------------------------------------------------------------------------------
' 展開先シートを作成します。
'==============================================================================
Private Function CreateExpandedSheet( _
    ByVal wsSource As Worksheet) As Worksheet

    Dim NewName As String

    NewName = wsSource.Name & "_Expand"

    '同名シート削除防止
    On Error Resume Next
    Application.DisplayAlerts = False

    Worksheets(NewName).Delete

    Application.DisplayAlerts = True
    On Error GoTo 0


    Set CreateExpandedSheet = _
        Worksheets.Add(After:=Worksheets(Worksheets.Count))


    CreateExpandedSheet.Name = NewName

End Function

'==============================================================================
' ExpandMapLayout
'------------------------------------------------------------------------------
' 旧サイズ単位でマップを読み込み、
' 新サイズ単位へ配置します。
'==============================================================================
Private Sub ExpandMapLayout( _
    ByVal wsSource As Worksheet, _
    ByVal wsTarget As Worksheet)

    Dim SourceRow As Long
    Dim SourceCol As Long

    Dim TargetRow As Long
    Dim TargetCol As Long


    TargetRow = 1


    For SourceRow = 1 _
        To wsSource.UsedRange.Rows.Count _
        Step OLD_MAP_HEIGHT


        TargetCol = 1


        For SourceCol = 1 _
            To wsSource.UsedRange.Columns.Count _
            Step OLD_MAP_WIDTH


            CopyMapBlock _
                wsSource, _
                wsTarget, _
                SourceRow, _
                SourceCol, _
                TargetRow, _
                TargetCol


            TargetCol = TargetCol + NEW_MAP_WIDTH

        Next SourceCol


        TargetRow = TargetRow + NEW_MAP_HEIGHT

    Next SourceRow


End Sub

'==============================================================================
' CopyMapBlock
'------------------------------------------------------------------------------
' 1マップ分をコピーします。
'==============================================================================
Private Sub CopyMapBlock( _
    ByVal wsSource As Worksheet, _
    ByVal wsTarget As Worksheet, _
    ByVal SourceRow As Long, _
    ByVal SourceCol As Long, _
    ByVal TargetRow As Long, _
    ByVal TargetCol As Long)


    Dim r As Long
    Dim c As Long


    For r = 0 To OLD_MAP_HEIGHT - 1

        For c = 0 To OLD_MAP_WIDTH - 1


            wsSource.Cells(SourceRow + r, SourceCol + c).Copy _
                Destination:= _
                wsTarget.Cells(TargetRow + r, TargetCol + c)


        Next c

    Next r


End Sub

'==============================================================================
' GetColumnWidths
'------------------------------------------------------------------------------
' 元シートの列幅を保存します。
'==============================================================================
Private Function GetColumnWidths( _
    ByVal ws As Worksheet) As Variant

    Dim Widths() As Double
    Dim c As Long

    ReDim Widths(1 To ws.UsedRange.Columns.Count)

    For c = 1 To ws.UsedRange.Columns.Count

        Widths(c) = ws.Columns(c).ColumnWidth

    Next c

    GetColumnWidths = Widths

End Function

'==============================================================================
' RestoreColumnWidths
'------------------------------------------------------------------------------
' 保存した列幅を復元します。
'==============================================================================
Private Sub RestoreColumnWidths( _
    ByVal ws As Worksheet, _
    ByVal Widths As Variant)

    Dim c As Long

    For c = LBound(Widths) To UBound(Widths)

        ws.Columns(c).ColumnWidth = Widths(c)

    Next c

End Sub
