Attribute VB_Name = "SHTBST_20_Restore"
Option Explicit

'==============================================================================
' SHTBST_20_Restore
'------------------------------------------------------------------------------
' Excel Worksheet Boosters
' SHTBST - Sheet Store Booster
'
' SHTBST_10_Store で保管したワークシートデータを復元します。
'
' Restore対象：
'   ・セルのValue
'
' 保存されているSheet NameとRange Addressを使用して、
' Store時と同じワークシート上のRangeへデータを戻します。
'==============================================================================


'==============================================================================
' RestoreRange
'------------------------------------------------------------------------------
' 指定したStore IDのデータを復元します。
'
' 使用例：
'
'   RestoreRange 1001
'
' Store IDは保管先シートの列番号として使用します。
'==============================================================================

Public Sub RestoreRange(ByVal StoreID As Long)

    Dim storeSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim targetRange As Range

    Dim sheetName As String
    Dim rangeAddress As String

    Dim rowsCount As Long
    Dim columnsCount As Long

    Dim data() As Variant

    Dim r As Long
    Dim c As Long
    Dim pos As Long

    Dim answer As VbMsgBoxResult


    '--------------------------------------------------------------------------
    ' Store IDチェック
    '--------------------------------------------------------------------------
    
    If StoreID < 1 Then
        
        MsgBox _
            "Store IDには1以上の値を指定してください。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 保管先シート取得
    '--------------------------------------------------------------------------
    
    Set storeSheet = GetStoreSheet()
    
    If storeSheet Is Nothing Then Exit Sub


    '--------------------------------------------------------------------------
    ' Store ID存在チェック
    '--------------------------------------------------------------------------
    
    If storeSheet.Cells(SHTBST_ROW_ID, StoreID).Value = "" Then
        
        MsgBox _
            "Store ID " & StoreID & _
            " の保管データがありません。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 保存情報取得
    '--------------------------------------------------------------------------
    
    sheetName = CStr( _
        storeSheet.Cells(SHTBST_ROW_SHEET, StoreID).Value)
    
    rangeAddress = CStr( _
        storeSheet.Cells(SHTBST_ROW_RANGE, StoreID).Value)
    
    rowsCount = CLng( _
        storeSheet.Cells(SHTBST_ROW_ROWS, StoreID).Value)
    
    columnsCount = CLng( _
        storeSheet.Cells(SHTBST_ROW_COLUMNS, StoreID).Value)


    '--------------------------------------------------------------------------
    ' 対象ワークシート取得
    '--------------------------------------------------------------------------
    
    On Error Resume Next
    
    Set targetSheet = _
        storeSheet.Parent.Worksheets(sheetName)
    
    On Error GoTo 0


    If targetSheet Is Nothing Then
        
        MsgBox _
            "保存時のワークシート """ & _
            sheetName & _
            """ が見つかりません。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' 対象Range取得
    '--------------------------------------------------------------------------
    
    On Error Resume Next
    
    Set targetRange = _
        targetSheet.Range(rangeAddress)
    
    On Error GoTo 0


    If targetRange Is Nothing Then
        
        MsgBox _
            "保存時のRange """ & _
            rangeAddress & _
            """ を取得できません。", _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' サイズチェック
    '--------------------------------------------------------------------------
    
    If targetRange.Rows.Count <> rowsCount _
       Or targetRange.Columns.Count <> columnsCount Then
        
        MsgBox _
            "保存時と現在のRangeサイズが一致しません。" & vbCrLf & _
            vbCrLf & _
            "保存時 : " & rowsCount & " × " & columnsCount & vbCrLf & _
            "現在   : " & targetRange.Rows.Count & _
            " × " & targetRange.Columns.Count, _
            vbExclamation
        
        Exit Sub
        
    End If


    '--------------------------------------------------------------------------
    ' Restore確認
    '--------------------------------------------------------------------------
    
    answer = MsgBox( _
        "Store ID " & StoreID & " のデータを復元します。" & vbCrLf & _
        vbCrLf & _
        "Sheet : " & sheetName & vbCrLf & _
        "Range : " & rangeAddress & vbCrLf & _
        vbCrLf & _
        "現在のデータは上書きされます。", _
        vbYesNo + vbQuestion)


    If answer = vbNo Then Exit Sub


    '--------------------------------------------------------------------------
    ' 1セルの場合
    '--------------------------------------------------------------------------
    
    If rowsCount = 1 And columnsCount = 1 Then
        
        targetRange.Value = _
            storeSheet.Cells(SHTBST_ROW_DATA_START, StoreID).Value
        
    Else
        
        '----------------------------------------------------------------------
        ' 保管データを2次元配列へ復元
        '----------------------------------------------------------------------
        
        ReDim data(1 To rowsCount, 1 To columnsCount)
        
        pos = SHTBST_ROW_DATA_START
        
        For r = 1 To rowsCount
            
            For c = 1 To columnsCount
                
                data(r, c) = _
                    storeSheet.Cells(pos, StoreID).Value
                
                pos = pos + 1
                
            Next c
            
        Next r
        
        
        '----------------------------------------------------------------------
        ' Rangeへ一括復元
        '----------------------------------------------------------------------
        
        targetRange.Value = data
        
    End If


    '--------------------------------------------------------------------------
    ' 完了
    '--------------------------------------------------------------------------
    
    MsgBox _
        "Store ID " & StoreID & " の復元が完了しました。" & vbCrLf & _
        "Sheet : " & sheetName & vbCrLf & _
        "Range : " & rangeAddress, _
        vbInformation

End Sub


'==============================================================================
' GetStoreSheet
'------------------------------------------------------------------------------
' 保管用ワークシートを取得します。
'
' SHTBST_10_Storeと同じ処理を使用します。
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
    ' 保管シートが存在しない場合
    '--------------------------------------------------------------------------
    
    If ws Is Nothing Then
        
        MsgBox _
            "保管用シート """ & _
            SHTBST_STORE_SHEET_NAME & _
            """ がありません。", _
            vbExclamation
        
        Set GetStoreSheet = Nothing
        
        Exit Function
        
    End If


    Set GetStoreSheet = ws

End Function

