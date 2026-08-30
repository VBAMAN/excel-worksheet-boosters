Attribute VB_Name = "SHTBST_10_Store"
Option Explicit

'==============================================================================
' SHTBST_10_Store
'------------------------------------------------------------------------------
' Excel Worksheet Boosters
' SHTBST - Sheet Store Booster
'
' 指定したRangeのValueを、Store ID単位で保管します。
'
' このモジュールは、保存対象となるデータの意味を知りません。
'
' Course、Map、Route、NPCなどのアプリケーション固有情報は扱わず、
' 「指定されたRangeの状態を保管する」ことだけを担当します。
'==============================================================================


'==============================================================================
' StoreRange
'------------------------------------------------------------------------------
' 指定したRangeのValueを保管します。
'
' 使用例：
'
'   StoreRange Sheet1.Range("E5:W20"), 1001
'
' Store IDは保管先シートの列番号として使用します。
'
' 同じIDがすでに存在する場合は、上書き確認を行います。
'==============================================================================

Public Sub StoreRange( _
    ByVal Target As Range, _
    ByVal StoreID As Long)

    Dim storeSheet As Worksheet
    Dim targetData As Variant

    Dim targetRows As Long
    Dim targetColumns As Long

    Dim storeColumn As Long
    Dim answer As VbMsgBoxResult

    '--------------------------------------------------------------------------
    ' Store IDチェック
    '--------------------------------------------------------------------------
    
    If StoreID < 1 Then
        
        MsgBox _
            "Store IDには1上の値を指定してください。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' Rangeチェック
    '--------------------------------------------------------------------------
    
    If Target Is Nothing Then
        
        MsgBox _
            "保存対象のRangeが指定されていません。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 保管先シート取得
    '--------------------------------------------------------------------------
    
    Set storeSheet = GetStoreSheet()
    
    If storeSheet Is Nothing Then Exit Sub


    '--------------------------------------------------------------------------
    ' Store ID = 列番号
    '--------------------------------------------------------------------------
    
    storeColumn = StoreID


    '--------------------------------------------------------------------------
    ' 上書き確認
    '--------------------------------------------------------------------------
    
    If storeSheet.Cells(SHTBST_ROW_ID, storeColumn).Value <> "" Then
        
        answer = MsgBox( _
            "Store ID " & StoreID & _
            " は既に存在します。" & vbCrLf & _
            "上書きしますか？", _
            vbYesNo + vbQuestion)
        
        If answer = vbNo Then Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 保存対象データ取得
    '--------------------------------------------------------------------------
    
    targetData = Target.Value
    
    targetRows = Target.Rows.Count
    targetColumns = Target.Columns.Count


    '--------------------------------------------------------------------------
    ' 既存データ削除
    '--------------------------------------------------------------------------
    
    ClearStoredData _
        storeSheet, _
        storeColumn


    '--------------------------------------------------------------------------
    ' メタデータ保存
    '--------------------------------------------------------------------------
    
    With storeSheet
        
        ' Store ID
        .Cells(SHTBST_ROW_ID, storeColumn).Value = StoreID
        
        ' Sheet Name
        .Cells(SHTBST_ROW_SHEET, storeColumn).Value = _
            Target.Worksheet.Name
        
        ' Range Address
        .Cells(SHTBST_ROW_RANGE, storeColumn).Value = _
            Target.Address(False, False)
        
        ' Rows
        .Cells(SHTBST_ROW_ROWS, storeColumn).Value = _
            targetRows
        
        ' Columns
        .Cells(SHTBST_ROW_COLUMNS, storeColumn).Value = _
            targetColumns
        
    End With


    '--------------------------------------------------------------------------
    ' データ保存
    '--------------------------------------------------------------------------
    
    WriteStoredData _
        storeSheet, _
        storeColumn, _
        targetData, _
        targetRows, _
        targetColumns


    '--------------------------------------------------------------------------
    ' 完了
    '--------------------------------------------------------------------------
    
    MsgBox _
        "Store ID " & StoreID & " を保存しました。" & vbCrLf & _
        "Sheet : " & Target.Worksheet.Name & vbCrLf & _
        "Range : " & Target.Address(False, False), _
        vbInformation

End Sub


'==============================================================================
' WriteStoredData
'------------------------------------------------------------------------------
' RangeのValueを1列へ展開して保管します。
'
' Range.Valueが2次元配列になることを利用します。
'==============================================================================

Private Sub WriteStoredData( _
    ByVal storeSheet As Worksheet, _
    ByVal storeColumn As Long, _
    ByVal targetData As Variant, _
    ByVal targetRows As Long, _
    ByVal targetColumns As Long)

    Dim r As Long
    Dim c As Long
    Dim position As Long

    position = SHTBST_ROW_DATA_START


    '--------------------------------------------------------------------------
    ' 1セルの場合
    '--------------------------------------------------------------------------
    
    If targetRows = 1 And targetColumns = 1 Then
        
        storeSheet.Cells(position, storeColumn).Value = targetData
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 2次元データを1列へ展開
    '--------------------------------------------------------------------------
    
    For r = 1 To targetRows
        
        For c = 1 To targetColumns
            
            storeSheet.Cells(position, storeColumn).Value = _
                targetData(r, c)
            
            position = position + 1
            
        Next c
        
    Next r

End Sub


'==============================================================================
' ClearStoredData
'------------------------------------------------------------------------------
' 指定したStore IDの既存データを削除します。
'==============================================================================

Private Sub ClearStoredData( _
    ByVal storeSheet As Worksheet, _
    ByVal storeColumn As Long)

    Dim lastRow As Long

    lastRow = storeSheet.Cells( _
        storeSheet.Rows.Count, _
        storeColumn).End(xlUp).Row


    If lastRow >= SHTBST_ROW_ID Then
        
        storeSheet.Range( _
            storeSheet.Cells(SHTBST_ROW_ID, storeColumn), _
            storeSheet.Cells(lastRow, storeColumn) _
        ).ClearContents
        
    End If

End Sub


'==============================================================================
' GetStoreSheet
'------------------------------------------------------------------------------
' 保管用ワークシートを取得します。
'
' 存在しない場合は自動作成します。
'==============================================================================

Private Function GetStoreSheet() As Worksheet

    Dim wb As Workbook
    Dim ws As Worksheet


    Set wb = ThisWorkbook


    '--------------------------------------------------------------------------
    ' 既存シート検索
    '--------------------------------------------------------------------------
    
    On Error Resume Next
    
    Set ws = wb.Worksheets(SHTBST_STORE_SHEET_NAME)
    
    On Error GoTo 0


    '--------------------------------------------------------------------------
    ' 存在しない場合は作成
    '--------------------------------------------------------------------------
    
    If ws Is Nothing Then
        
        On Error GoTo CreateError
        
        Set ws = wb.Worksheets.Add( _
            After:=wb.Worksheets(wb.Worksheets.Count))
        
        ws.Name = SHTBST_STORE_SHEET_NAME
        
    End If


    Set GetStoreSheet = ws
    
    Exit Function


'------------------------------------------------------------------------------
' 作成エラー
'------------------------------------------------------------------------------

CreateError:

    MsgBox _
        "保管用シート """ & _
        SHTBST_STORE_SHEET_NAME & _
        """ を作成できませんでした。", _
        vbExclamation

    Set GetStoreSheet = Nothing

End Function

