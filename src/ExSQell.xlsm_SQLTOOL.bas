'========================================================
' ExSQell.xlsm - Copyright (C) 2016 nmrmsys
' sqltools.xls - Copyright (C) 2005 M.Nomura
'========================================================

Option Explicit

'定数定義
Const ccAppTitle = "ExSQell"
Const ccCfgShtName = "設定"
Const cnCfgMaxRow = 30
Const ccLibShtName = "ライブラリ"
Const ccQryShtName = "実行結果"
Const ccTblShtName = "表定義"
Const ccIdxShtName = "索引定義"
Const ccSrcShtName = "ソース"
Const ccWrkShtName = "作業用"
Const ccMntShtPref = "+"
Const ccQpsShtPref = "&"
Const ccMaxSqlWidth = 60
Const ccPriKeyPref = "PK_"
Const ccPriKeySuff = ""
Const ccWfOUShtName = "WfOU"

Const ccPURPLE = 13
Const ccBLUE = 5
Const ccRED = 3
Const ccWHITE = 2
Const ccGREY = 15
Const ccYELLOW = 36
Const ccPINK = 38

Const ccFT_NONE = 0
Const ccFT_DATE = 1
Const ccFT_TEXT = 2
Const ccFT_ELSE = 3

Const ccTPL_SEL = 0
Const ccTPL_INS = 1
Const ccTPL_UPD = 2
Const ccTPL_DEL = 3
Const ccTPL_CRE = 4
Const ccTPL_CLR = 5

Const ccMaxShFldLen = 2000

'変数定義
Public coSTBook As Workbook
Public coPrevSheet
Public coBalloon
Public cnLibMode As Integer
Public cnTplMode As Integer
Public csTplCell As String
Public coTplRs
Public coConObjs As New Prop

Public coWSH
Public coFSO
Public csShCmd
Public csUname

Const cnCurDirRowS = 29
Const cnCurDirRowE = 34

'メッセージ定義
Const ccUpdateDisclaim = "更新機能を使用するには「更新SQLの実行」設定を確認して下さい" & vbCrLf _
                        & "なお本ツールは未だ開発中の為、本番環境への更新はお控え下さい"


Sub Auto_Open()
    'マクロ存在ブックを退避
    Set coSTBook = ThisWorkbook
    'マクロ起動ショートキー登録
    With Application
        Call .OnKey("^{q}", "ST_Query")
        Call .OnKey("+^{q}", "ST_Query")
        Call .OnKey("%^{q}", "ST_Query")
        Call .OnKey("^{w}", "ST_Which")
        Call .OnKey("+^{w}", "ST_Which")
        Call .OnKey("%^{w}", "ST_Which")
        Call .OnKey("+%^{w}", "ST_Which")
        Call .OnKey("^{e}", "ST_ExtInfo")
        Call .OnKey("+^{e}", "ST_ExtInfo")
        Call .OnKey("^{p}", "ST_Plus")
        Call .OnKey("+^{p}", "ST_Plus")
        Call .OnKey("%^{p}", "ST_Plus")
        Call .OnKey("^{m}", "ST_Maintenance")
        Call .OnKey("+^{m}", "ST_Maintenance")
        Call .OnKey("%^{m}", "ST_Maintenance")
        Call .OnKey("+%^{m}", "ST_Maintenance")
        Call .OnKey("^{l}", "ST_Library")
        Call .OnKey("+^{l}", "ST_Library")
        Call .OnKey("%^{l}", "ST_Library")
        Call .OnKey("+%^{l}", "ST_Library")
        Call .OnKey("^{t}", "ST_Template")
        Call .OnKey("+^{t}", "ST_Template")
        Call .OnKey("%^{t}", "ST_Template")
        Call .OnKey("^{k}", "ST_Knowledge")
        Call .OnKey("%^{k}", "ST_Knowledge")
        
        Call .OnKey("^{DELETE}", "ST_ShtMntDel")
        Call .OnKey("+^{DELETE}", "ST_ShtDel")
        Call .OnKey("+^{TAB}", "ST_ShtCFG")
        Call .OnKey("+^{BS}", "ST_ShtBack")
        Call .OnKey("+^{:}", "ST_ShtClear")
        Call .OnKey("+^{INSERT}", "ST_ShtNew")
        Call .OnKey("+^{^}", "ST_ShtHideList")
        Call .OnKey("+^{-}", "ST_ShtHide")
        Call .OnKey("+^{v}", "ST_ShtPaste")
        Call .OnKey("+^{c}", "ST_ShtCopy")
        Call .OnKey("+%^{c}", "ST_ShtCopy")
        Call .OnKey("+^{ }", "ST_ShtEdit")
        Call .OnKey("+^{RETURN}", "ST_ShtEdit")
        Call .OnKey("+%^{RETURN}", "ST_CommitMenu")
        Call .OnKey("+^{u}", "ST_ShtUniqChk")
    End With
End Sub

Function GetSafeSelection() As Range
'    Dim nStaRow As Long, nEndRow As Long, nStaCol As Long, nEndCol As Long
'    If Selection.Rows.Count > ActiveSheet.UsedRange.Rows.Count Then
'        nStaRow = 1
'        nEndRow = ActiveSheet.UsedRange.Rows.Count
'    Else
'        nStaRow = Selection.Row
'        nEndRow = Selection.Row + Selection.Rows.Count - 1
'    End If
'    If Selection.Columns.Count > ActiveSheet.UsedRange.Columns.Count Then
'        nStaCol = 1
'        nEndCol = ActiveSheet.UsedRange.Columns.Count
'    Else
'        nStaCol = Selection.Column
'        nEndCol = Selection.Column + Selection.Columns.Count - 1
'    End If
'    Set GetSafeSelection = Range(Cells(nStaRow, nStaCol), Cells(nEndRow, nEndCol))
    Dim nMaxRows As Long, nMaxCols As Long, sMaxCols As String, sRange As String, oArea
    '使用範囲の最大行列取得
    With ActiveSheet.UsedRange
        nMaxRows = .Row - 1 + .Rows.Count
        nMaxCols = .Column - 1 + .Columns.Count
        sMaxCols = Cells(1, nMaxCols).Address(False, False)
        sMaxCols = Left(sMaxCols, Len(sMaxCols) - 1)
    End With
    '行全選択かどうか
'    MsgBox ActiveSheet.UsedRange.Address
    If Selection.Rows.Count > nMaxRows Then
        '列全選択かどうか
        If Selection.Columns.Count > nMaxCols Then
            'シート選択
            Set GetSafeSelection = ActiveSheet.UsedRange
        Else
            '列選択
            For Each oArea In Selection.Areas
                If sRange <> "" Then sRange = sRange & ","
                sRange = sRange & Replace(oArea.Address, ":", "1:") & nMaxRows
            Next
            Set GetSafeSelection = Range(sRange)
        End If
    Else
        '列全選択かどうか
        If Selection.Columns.Count > nMaxCols Then
            '行選択
            For Each oArea In Selection.Areas
                If sRange <> "" Then sRange = sRange & ","
                sRange = sRange & "$A" & Mid(Replace(oArea.Address, ":$", ":$" & sMaxCols), 2)
            Next
            Set GetSafeSelection = Range(sRange)
        Else
            '範囲選択
            Set GetSafeSelection = Selection
        End If
    End If
End Function

Sub ST_ShtCopy()
    Dim oDataObj As New DataObject, oSel As Range, oRow As Range, oCol As Range
    Dim s As String, l As String, bAlt As Boolean, sRowSep As String, sColSep As String
    If GetKeyState(VK_MENU) < 0 Then bAlt = True
    Set oSel = GetSafeSelection
    If bAlt Then
        If oSel.Columns.Count = 1 Then
            sRowSep = InputBox("行区切り文字の指定", ccAppTitle)
            sColSep = ""
        Else
            sRowSep = ""
            sColSep = InputBox("列区切り文字の指定", ccAppTitle)
        End If
    Else
        If oSel.Columns.Count = 1 Then
            sRowSep = ","
            sColSep = ""
        Else
            sRowSep = ""
            sColSep = ","
        End If
    End If
    For Each oRow In oSel.Rows
        If oRow.RowHeight > 0 Then
            If s <> "" Then
                s = s & sRowSep & vbCrLf
            End If
            l = ""
            For Each oCol In oRow.Columns
                If oCol.ColumnWidth > 0 Then
                    If l <> "" Then
                        l = l & sColSep
                    End If
                    l = l & oCol.Value
                End If
            Next
            s = s & l
        End If
    Next
    s = s & vbCrLf
    With oDataObj
        Call .SetText(s, 1)
        .PutInClipboard
    End With
End Sub

Sub ST_ShtPaste()
    Dim oDataObj As New DataObject
    With oDataObj
        .GetFromClipboard
        If .GetFormat(1) Then
            Call SetSql(Replace(.GetText, vbTab, "  "), ActiveCell)
        End If
    End With
End Sub

Sub ST_ShtEdit()
    Dim sSQL As String, nRow As Long, nCol As Long
    nRow = ActiveCell.Row
    nCol = ActiveCell.Column
    If ActiveCell <> "" Then
        Do Until nRow = 1 Or ActiveSheet.Cells(nRow, nCol) = ""
            nRow = nRow - 1
        Loop
        If ActiveSheet.Cells(nRow, nCol) = "" Then
            nRow = nRow + 1
        End If
        Do Until ActiveSheet.Cells(nRow, nCol) = ""
            sSQL = sSQL & ActiveSheet.Cells(nRow, nCol) & vbCrLf
            nRow = nRow + 1
        Loop
    End If
    If ST_WIN.ShowModal2(sSQL, "展開", "EDIT") Then
        Call SetSql(sSQL)
    End If
End Sub

Sub ST_ShtUniqChk()
    Dim nMaxRows As Long, nMaxCols As Long, nChkCol As Long, sFormula As String, oCol, oFC
    '使用範囲の最大行列取得
    With ActiveSheet.UsedRange
        nMaxRows = .Row - 1 + .Rows.Count
        nMaxCols = .Column - 1 + .Columns.Count
    End With
    If Selection.Rows.Count > nMaxRows Then
        nChkCol = nMaxCols + 1
        For Each oCol In Range(Cells(1, 1), Cells(1, nMaxCols))
            If Not (Range("A1") = "MAINTEN" And oCol.Column <= 2) Then
                If oCol = "重複CHK" Or oCol = "" Then
                    nChkCol = oCol.Column
                    Exit For
                End If
            End If
        Next
        For Each oCol In GetSafeSelection.Columns
            If oCol.Rows(1) <> "重複CHK" And oCol.Rows(1) <> "重複KEY" Then
                If sFormula <> "" Then
                    sFormula = sFormula & "&"
                End If
                sFormula = sFormula & "RC" & oCol.Column
            End If
        Next
        Application.ScreenUpdating = False
        ActiveSheet.EnableCalculation = False
        Application.EnableEvents = False
        With Cells(1, nChkCol)
            .ColumnWidth = 10.5
            .Value = "重複CHK"
        End With
        With Range(Cells(2, nChkCol), Cells(nMaxRows, nChkCol))
            .NumberFormat = ""
            .FormulaR1C1 = "=IF(COUNTIF(R2C" & nChkCol + 1 & ":R" & nMaxRows & "C" & nChkCol + 1 & _
                            ",RC" & nChkCol + 1 & ")>1,""重複レコード"","""")"
        End With
        With Cells(1, nChkCol + 1)
            .ColumnWidth = 10.5
            .Value = "重複KEY"
        End With
        With Range(Cells(2, nChkCol + 1), Cells(nMaxRows, nChkCol + 1))
            .NumberFormat = ""
            .FormulaR1C1 = "=T(" & sFormula & ")"
        End With
        Application.EnableEvents = True
        ActiveSheet.EnableCalculation = True
        If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
        Cells.AutoFilter
        Application.ScreenUpdating = True
    Else
        For Each oCol In Range(Cells(1, 1), Cells(1, nMaxCols))
            If oCol = "重複CHK" Then
                nChkCol = oCol.Column
                Exit For
            End If
        Next
        If nChkCol = 0 Then
            Call MsgBoxST("重複チェックする列を列選択してください")
        Else
            Application.ScreenUpdating = False
            If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
            Columns(nChkCol).Delete
            Columns(nChkCol).Delete
            Range(Cells(1, 1), Cells(nMaxRows, nChkCol - 1)).AutoFilter
            Application.ScreenUpdating = True
        End If
    End If
End Sub

Sub ST_ShtHide()
    Dim n, s
    For Each s In Sheets
        If s.Visible Then
            n = n + 1
        End If
    Next
    If n <> 1 And ActiveSheet.Name <> ccCfgShtName Then
        ActiveSheet.Visible = False
    End If
End Sub

Sub ST_ShtHideList()
    Call Application.Dialogs(xlDialogWorkbookUnhide).Show
End Sub

Sub ST_ShtMntDel()
    With ActiveSheet
        If .Range("A1") = "MAINTEN" Then
            If UCase(Cells(ActiveCell.Row, 2)) = "D" Then
                Cells(ActiveCell.Row, 2) = ""
            Else
                Cells(ActiveCell.Row, 2) = "D"
            End If
        End If
    End With
End Sub

Sub ST_ShtClear()
    With ActiveSheet
        If .Range("A1") = "QPARAMS" Then
            .Range(.Cells(3, 7), .Cells(.UsedRange.Rows.Count, 8)) = ""
        Else
            .Select
            If .FilterMode Then
                .ShowAllData
            End If
        End If
    End With
End Sub

Sub ST_ShtDel()
    If Sheets.Count > 1 And GetCfgSheet.Name <> ActiveSheet.Name Then
        ActiveSheet.Delete
        Application.StatusBar = ""
        On Error Resume Next
        ActiveSheet.Worksheet_Activate
        On Error GoTo 0
    End If
End Sub

Sub ST_ShtNew()
    Call GetNewSheet(ccWrkShtName, True)
    Cells.NumberFormatLocal = "@"
    Cells.Font.Name = "ＭＳ ゴシック"
End Sub

Sub ST_ShtCFG()
    With GetCfgSheet()
        .Parent.Activate
        .Select
    End With
End Sub

Sub ST_ShtBack()
    Dim oActiveSheet
    On Error Resume Next
    Set oActiveSheet = ActiveSheet
    coPrevSheet.Select
    Set coPrevSheet = oActiveSheet
    On Error GoTo 0
End Sub

'クエリーを実行
Sub ST_Query(Optional argSql, Optional argSqlTitle)
    Dim sConStr As String
    Dim sSQL As String
    Dim sSQLType As String
    Dim sTitle As String, sObjName As String
    Dim sShtName As String, sShtMode As String
    Dim oRs, oFld, oFlds, oOpts, oSht
    Dim nFldType As Integer, nEffRecCnt As Long
    'ExSQell暫定ロジック
    Dim sCmd As String
    Dim nDefExecMode As Integer
    
    nDefExecMode = GetDefExecMode '2 Shell実行モード
    
    'シート種別の判別
    sShtMode = ActiveSheet.Range("A1")
    Select Case sShtMode
    Case "MAINTEN"
        'メンテナンス反映
        Call UpdateMaintenanceSheet
    Case "NAME" 'ソース定義
        Call ST_Maintenance 'コンパイル
        Exit Sub
    Case "TABLE_NAME", "INDEX_NAME" '表定義、索引定義
        Exit Sub
    Case Else
        '通常クエリ実行、抽出条件入力
        
        'ExSQell暫定ロジック
        If nDefExecMode = 2 Or Left(ActiveCell.Value, 2) = "$ " Then
            
            'Shell実行モードで先頭 $ が無い場合は補完
            If nDefExecMode = 2 And Left(ActiveCell.Value, 1) <> "$" Then
                ActiveCell.Value = "$ " & ActiveCell.Value
            End If
            
            sCmd = Mid(ActiveCell.Value, 3)
            
            sSQL = ""
            sSQLType = "SELECT"
            sShtName = ""
            Set oOpts = New Prop
        Else
            sConStr = GetConStr
            
            If IsMissing(argSql) Then
                '現在セル位置からＳＱＬ文、タイプ、シート名の取得
                sSQL = GetSql(sSQLType, sShtName)
                If sSQL = "" Then
                    Exit Sub
                End If
            Else
                'ＳＱＬ文を指定して実行された場合
                sSQL = argSql
                sSQLType = "SELECT"
                sShtName = argSqlTitle
            End If
        End If
        
        'ALT同時押し以外の場合は抽出条件指定
        If Not GetKeyState(VK_MENU) < 0 And Not GetKeyState(VK_SHIFT) < 0 And sSQLType = "TABLE" Then
            Call MakeQpsSheet(Trim(ActiveCell))
            Exit Sub
        End If
        
        Select Case sSQLType
        Case "SELECT", "TABLE", "RELOAD"
        
            'ExSQell暫定ロジック
            If sCmd <> "" Then
                'シェルコマンドの実行
                If Not ExecShell(sCmd, oRs, oFlds) Then
                    Exit Sub
                End If
            Else
                'クエリの実行
                If Not ExecQuery(sConStr, sSQL, sSQLType, oRs, oFlds, oOpts) Then
                    Exit Sub
                End If
            End If
        
            
            '現在セル位置からのシート名が空白の場合はデフォルトのシート名
            sShtName = IIf(sShtName = "", ccQryShtName, sShtName)
            
            'クエリ結果シートの作成
            If GetCfgVal("行番号表示") <> 3 Then 'しない以外
                With oFlds
                    Set oFld = New Prop
                    oFld.Item("Title") = "No."
                    Select Case GetCfgVal("行番号表示")
                    Case 1 '値
                        oFld.Item("特殊列") = "行番号"
                    Case 2 '書式
                        oFld.Item("特殊列") = "行連番"
                    End Select
                    Set .Item(0, False) = oFld '先頭に追加
                End With
            End If
            With oOpts
                .Item("タイトル") = "フィルタ"
                .Item("行タイトル") = "$1:$1"
                If GetCfgVal("行番号表示") <> 3 Then 'しない以外
                    .Item("枠固定セル") = "B2"
                    .Item("列タイトル") = "$A:$A"
                End If
                .Item("スタイル") = "罫線描画、項目型色付"
                Select Case GetCfgVal("縞模様表示")
                Case 1 '条件付き書式で表示
                    .Item("スタイル") = .Item("スタイル") & "、縞模様書式"
                Case 2 'セル背景色で表示
                    .Item("スタイル") = .Item("スタイル") & "、縞模様セル"
                End Select
                Select Case GetCfgVal("空白セル強調")
                Case 1 '条件付き書式で強調
                    .Item("スタイル") = .Item("スタイル") & "、空白セル強調"
                End Select
                Select Case GetCfgVal("列幅のサイズ調整")
                Case 1 '列全体の文字幅
                    .Item("スタイル") = .Item("スタイル") & "、自動列幅列全体"
                Case 3 'データの文字幅
                    .Item("スタイル") = .Item("スタイル") & "、自動列幅データ"
                End Select
                .Item("コメント") = "" & IIf(GetCfgVal("件数、秒数のコメント") = 1, "、件数、秒数", "")
            End With
            
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            If sShtMode = "QPARAMS" Then
                If ActiveSheet.Range("I2") = "" Then
                    Select Case GetCfgVal("抽出条件シート")
                    Case 1 '抽出時に削除
                        ActiveSheet.Delete
                    Case 2 '抽出時に非表示
                        ActiveSheet.Visible = False
                    End Select
                End If
            End If
            If sShtMode <> "QPARAMS" And sSQLType <> "TABLE" And sSQLType <> "RELOAD" Then
                Set oSht = GetNewSheet(sShtName)
            Else
                Set oSht = GetNewSheet(sShtName, False, True)
            End If
            Call MakeList(oSht, oRs, oFlds, oOpts)
            With oSht
                Call SetCmtVal(.Range("A1"), , "接続先: " & sConStr & vbLf & sSQL)
            End With
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
        
        Case "DML", "DDL", "DCL", "BLOCK"
            Dim arLines, nLineNo As Integer, nTotalCnt As Integer, bChk As Boolean, bRet As Boolean
            Dim sTotalCnt As String
            Select Case GetCfgVal("更新SQLの実行")
            Case 2, 3 '確認して実行
                If MsgBoxST("接続先：" & GetConDisp(sConStr) & vbCrLf & vbCrLf _
                    & "更新SQLを実行しますか？", vbYesNo + vbDefaultButton2) = vbNo _
                Then
                    Exit Sub
                End If
                If GetCfgVal("更新SQLの実行") = 2 Then
                    bChk = True
                End If
            Case 4 '問答無用で実行
                '確認無し
            Case Else '実行しない
                Call MsgBoxST(ccUpdateDisclaim)
                Exit Sub
            End Select
            'DB接続、SQL実行
            arLines = Split(sSQL, ";" & vbCrLf)
            nTotalCnt = UBound(arLines) + 1
            sTotalCnt = "/" & Format(UBound(arLines) + 1, "000")
            For nLineNo = 1 To UBound(arLines) + 1
                If nTotalCnt = 1 Then
                    Application.StatusBar = "更新SQL実行中..."
                    bRet = True
                Else
                    If bChk Then
                        bRet = ST_WIN.ShowModal2(arLines(nLineNo - 1), "実行")
                    Else
                        bRet = True
                    End If
                    Application.StatusBar = "更新SQL実行中... " & _
                        Format(nLineNo, "000") & sTotalCnt
                End If
                If bRet Then
                    If Not ExecSql(sConStr, arLines(nLineNo - 1), argOpts:=oOpts) Then
                        Application.StatusBar = ""
                        Exit Sub
                    End If
                    nEffRecCnt = oOpts.Item("更新件数")
                    If nTotalCnt = 1 And nEffRecCnt <> 0 Then
                        MsgBoxST ("実行が完了しました" & vbCrLf _
                                & "影響レコード件数：" & nEffRecCnt & "件")
                    Else
                        Application.StatusBar = "影響レコード件数：" & nEffRecCnt & "件"
                    End If
                End If
            Next
            If nTotalCnt <> 1 Then
                MsgBoxST ("更新SQLバッチ実行が完了しました")
            End If
            Application.StatusBar = ""
        Case "FUNCTION", "PROCEDURE"
            '☆ADO対応要
            sObjName = ActiveCell
            If MsgBoxST(sSQLType & " " & sObjName & "を実行しますか？", vbYesNo + vbDefaultButton2) = vbYes _
            Then
                Dim sSpSql As String, sSpMsg As String, sRetType As String, oPrms As New Prop, oPrm
                If sSQLType = "FUNCTION" Then
                    sSpSql = "BEGIN :RESULT := " & sObjName & "("
                Else
                    sSpSql = "BEGIN " & sObjName & "("
                End If
                sSQL = GetLibSql("引数情報取得")
                sSQL = SetSqlParam(sSQL, "オブジェクト名", sObjName)
                If Not ExecSql(sConStr, sSQL, oRs) Then
                    Exit Sub
                End If
                Do Until oRs.EOF
                    If oRs("POSITION") = 0 Then
                        sRetType = oRs("DATA_TYPE").Value
                    Else
                        '実行SQLにプレースホルダ追加
                        If Right(sSpSql, 1) <> "(" Then
                            sSpSql = sSpSql & ", "
                        End If
                        sSpSql = sSpSql & ":" & oRs("ARGUMENT_NAME")
                        'パラメータプロパティ追加
                        Set oPrm = New Prop
                        oPrm.Item("Name") = oRs("ARGUMENT_NAME").Value
                        oPrm.Item("InOut") = "[" & oRs("IN_OUT").Value & "]"
                        oPrm.Item("Type") = oRs("DATA_TYPE").Value
                        '入力パラメータはダイアログ表示
                        If InStr(oRs("IN_OUT").Value, "IN") <> 0 Then
                            oPrm.Item("Value") = InputBox( _
                                oRs("ARGUMENT_NAME").Value & "   " & oRs("DATA_TYPE").Value, ccAppTitle)
                        End If
                        Set oPrms.Item() = oPrm
                    End If
                    oRs.MoveNext
                Loop
                sSpSql = sSpSql & "); END;"
                Set oOpts = New Prop
                Set oOpts.Item("パラメータ") = oPrms
                oOpts.Item("戻り値") = sRetType
                If Not ExecSql(sConStr, sSpSql, argOpts:=oOpts) Then
'                    Exit Sub
                End If
                sSpMsg = sSpMsg & sSQLType & " " & sObjName & vbLf & vbLf
                For Each oPrm In oPrms.Items
                    sSpMsg = sSpMsg & RPAD(oPrm.Item("Name"), 18) & RPAD(oPrm.Item("Type"), 8) & _
                        RPAD(oPrm.Item("InOut"), 5) & " =  " & oPrm.Item("Value") & vbLf
                Next
                If sSQLType = "FUNCTION" Then
                    sSpMsg = sSpMsg & vbLf & "戻り値 " & sRetType & " = " & oOpts.Item("戻り値") & vbLf
                End If
                Call ST_WIN.ShowModal2(sSpMsg)
            End If
        End Select
    End Select
End Sub

Sub ST_CommitMenu()
    Dim sConKey, nConMode, nCmtMode
    
    nConMode = GetCfgVal("接続モード")
    nCmtMode = GetCfgVal("コミットモード")
    
    If nCmtMode <> 1 Then
        Dim oPop As CommandBar
        Dim oMnu, oItm
        On Error Resume Next
        Set oPop = Application.CommandBars(ccAppTitle)
        On Error GoTo 0
        If Not oPop Is Nothing Then
            oPop.Delete
        End If
        Set oPop = Application.CommandBars.Add(Name:=ccAppTitle, Position:=msoBarPopup, Temporary:=True)
        With coConObjs
            For Each sConKey In .Keys
                If Left(sConKey, 1) = nConMode Then
                    Set oMnu = oPop.Controls.Add(msoControlPopup)
                    oMnu.Caption = .Item(sConKey).Item("Disp")
                    Set oItm = oMnu.Controls.Add(msoControlButton)
                    oItm.Caption = "コミット"
                    oItm.OnAction = "ST_CommitMenu_OnAction"
                    oItm.Tag = sConKey
                    oItm.Parameter = "C"
                    Set oItm = oMnu.Controls.Add(msoControlButton)
                    oItm.Caption = "ロールバック"
                    oItm.OnAction = "ST_CommitMenu_OnAction"
                    oItm.Tag = sConKey
                    oItm.Parameter = "R"
                    Set oItm = oMnu.Controls.Add(msoControlButton)
                    oItm.Caption = "切断"
                    oItm.OnAction = "ST_CommitMenu_OnAction"
                    oItm.Tag = sConKey
                    oItm.Parameter = "D"
                End If
            Next
        End With
        If oPop.Controls.Count > 0 Then
            oPop.ShowPopup
        End If
    End If
End Sub

Sub ST_CommitMenu_OnAction()
    Dim oConItem, oCon, sMsg
    
    With Application.CommandBars.ActionControl
        Set oConItem = coConObjs.Item(.Tag)
        Set oCon = oConItem.Item("Obj")
        
        Select Case .Parameter
        Case "C"
            sMsg = "コミット"
            oCon.CommitTrans
            oCon.BeginTrans
            oConItem.Item("Disp") = Replace(oConItem.Item("Disp"), "*", "")
        Case "R"
            sMsg = "ロールバック"
            If Left(.Tag, 1) = "1" Then
                oCon.RollBack
            Else
                oCon.RollBackTrans
            End If
            oCon.BeginTrans
            oConItem.Item("Disp") = Replace(oConItem.Item("Disp"), "*", "")
        Case "D"
            sMsg = "切断"
            Set oCon = Nothing
            Call coConObjs.Remove(.Tag)
        End Select
        
        Call MsgBoxST(sMsg & "が完了しました")
        
    End With
End Sub

'メンテナンスシートの更新
Private Function UpdateMaintenanceSheet()
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nCol As Long, nMinCol As Long, nMaxCol As Long
    Dim sConStr As String, sTblName As String, sSQL As String, sFldSQL As String
    Dim bFixFlg As Boolean
    Dim oRs, oFlds, oOpts, arLines
    Dim nFldType As Integer
    With ActiveSheet
        nMinRow = 2
        nMinCol = 3
        nMaxRow = .UsedRange.Rows.Count
        nMaxCol = .UsedRange.Columns.Count
        arLines = Split(GetCmtVal(ActiveSheet.Range("A1")), vbLf)
        sConStr = arLines(0)
        sTblName = Replace(.Name, ccMntShtPref, "")
        Select Case GetCfgVal("更新SQLの実行")
        Case 2, 3 '確認して実行
            If MsgBoxST("接続先：" & GetConDisp(sConStr) & vbCrLf & vbCrLf _
                & "更新SQLを実行しますか？", vbYesNo + vbDefaultButton2) = vbNo _
            Then
                Exit Function
            End If
        Case 4 '問答無用で実行
            '確認無し
        Case Else '実行しない
            Call MsgBoxST(ccUpdateDisclaim)
            Exit Function
        End Select
        'フィールド情報取得用
        If Not ExecSql(sConStr, "SELECT * FROM " & sTblName & " WHERE 0=1", oRs, oFlds) Then
            Exit Function
        End If
        For nRow = nMinRow To nMaxRow
            sSQL = ""
            Select Case UCase(.Cells(nRow, nMinCol - 1))
            Case "I"
                Application.StatusBar = "追加：" & Format(nRow, "00000") & "行目"
                sSQL = sSQL & "INSERT INTO " & sTblName & "(" & vbCrLf
                sFldSQL = ""
                For nCol = nMinCol To nMaxCol
                    If .Cells(nMinRow - 1, nCol).ColumnWidth > 0 Then
                        If GetFldType(oRs, .Cells(nMinRow - 1, nCol)) <> ccFT_NONE Then
                            If sFldSQL = "" Then
                                sFldSQL = sFldSQL & "  "
                            Else
                                sFldSQL = sFldSQL & ", "
                            End If
                            sFldSQL = sFldSQL & .Cells(nMinRow - 1, nCol)
                        End If
                    End If
                Next
                sSQL = sSQL & sFldSQL & vbCrLf
                sSQL = sSQL & ") VALUES(" & vbCrLf
                sFldSQL = ""
                For nCol = nMinCol To nMaxCol
                    If .Cells(nMinRow - 1, nCol).ColumnWidth > 0 Then
                        nFldType = GetFldType(oRs, .Cells(nMinRow - 1, nCol))
                        If nFldType <> ccFT_NONE Then
                            If sFldSQL = "" Then
                                sFldSQL = sFldSQL & "  "
                            Else
                                sFldSQL = sFldSQL & ", "
                            End If
                            Select Case nFldType
                            Case ccFT_DATE
                                sFldSQL = sFldSQL & "TO_DATE('" & .Cells(nRow, nCol) & "','YYYY/MM/DD HH24:MI:SS')"
                            Case ccFT_TEXT
                                sFldSQL = sFldSQL & "'" & .Cells(nRow, nCol) & "'"
                            Case ccFT_ELSE
                                If .Cells(nRow, nCol) <> "" Then
                                    sFldSQL = sFldSQL & .Cells(nRow, nCol)
                                Else
                                    sFldSQL = sFldSQL & "NULL"
                                End If
                            End Select
                        End If
                    End If
                Next
                sSQL = sSQL & sFldSQL & vbCrLf
                sSQL = sSQL & ")" & vbCrLf
            Case "U"
                Application.StatusBar = "更新：" & Format(nRow, "00000") & "行目"
                sSQL = sSQL & "UPDATE " & sTblName & " SET" & vbCrLf
                sFldSQL = ""
                For nCol = nMinCol To nMaxCol
                    If .Cells(nMinRow - 1, nCol).ColumnWidth > 0 Then
                        nFldType = GetFldType(oRs, .Cells(nMinRow - 1, nCol))
                        If nFldType <> ccFT_NONE Then
                            If sFldSQL = "" Then
                                sFldSQL = sFldSQL & "  "
                            Else
                                sFldSQL = sFldSQL & ", "
                            End If
                            sFldSQL = sFldSQL & .Cells(nMinRow - 1, nCol) & " = "
                            Select Case nFldType
                            Case ccFT_DATE
                                sFldSQL = sFldSQL & "TO_DATE('" & .Cells(nRow, nCol) & "','YYYY/MM/DD HH24:MI:SS')"
                            Case ccFT_TEXT
                                sFldSQL = sFldSQL & "'" & .Cells(nRow, nCol) & "'"
                            Case ccFT_ELSE
                                If .Cells(nRow, nCol) <> "" Then
                                    sFldSQL = sFldSQL & .Cells(nRow, nCol)
                                Else
                                    sFldSQL = sFldSQL & "NULL"
                                End If
                            End Select
                        End If
                    End If
                Next
                sSQL = sSQL & sFldSQL & vbCrLf
                sSQL = sSQL & "WHERE ROWID = '" & .Cells(nRow, 1) & "'" & vbCrLf
            Case "D"
                Application.StatusBar = "削除：" & Format(nRow, "00000") & "行目"
                sSQL = sSQL & "DELETE FROM " & sTblName & vbCrLf
                sSQL = sSQL & "WHERE ROWID = '" & .Cells(nRow, 1) & "'" & vbCrLf
            End Select
            If sSQL <> "" Then
                bFixFlg = True
                'DB接続、SQL実行
                If Not ExecSql(sConStr, sSQL, argOpts:=oOpts) Then
                    Exit Function
                End If
            End If
        Next
        Application.StatusBar = ""
    End With
    If bFixFlg And GetCfgVal("メンテシートＤＢ反映後") = 1 Then '自動で再抽出
        Call ST_Maintenance
    End If
End Function

'オブジェクト一覧/定義の表示
Sub ST_Which()
    Dim sWhich1 As String, sWhich2 As String, sWhich3 As String
    Dim sWhich As String, sWBpat As String, sWList As String
    Dim oRs, oFlds, oOpts As New Prop
    Dim sObjType As String, sObjType1 As String, sObjType2 As String
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nCol As Long, nMinCol As Long, nMaxCol As Long
    Dim nFldIdx As Integer
    
    'ExSQell暫定ロジック
    Dim sCmd As String
    Dim nDefExecMode As Integer
    
    nDefExecMode = GetDefExecMode '2 Shell実行モード
    
    With ActiveCell
        sWhich1 = Trim(.Value)
        sWhich2 = Trim(.Offset(0, 1).Value)
        sWhich3 = Trim(.Offset(0, 2).Value)
        If sWhich1 = "" And sWhich2 = "" And sWhich3 = "" Then
            Exit Sub
        End If
        If sWhich2 <> "" Or sWhich3 <> "" Then
            sWhich = sWhich1
            sWBpat = sWhich1 & "%" & "," & sWhich2 & "," & sWhich3
            sWList = sWhich1 & "," & sWhich2 & "," & sWhich3
        Else
            sWhich = sWhich1
            sWBpat = sWhich1 & "%"
            sWList = sWhich1
        End If
        nMinRow = .Row
        nMinCol = .Column
    End With
    
    'ExSQell暫定ロジック
    If nDefExecMode = 2 Or Left(ActiveCell.Value, 2) = "$ " Then
        
        'Shell実行モードで先頭 $ が無い場合は補完
        If nDefExecMode = 2 And Left(ActiveCell.Value, 1) <> "$" Then
            ActiveCell.Value = "$ " & ActiveCell.Value
        End If
    
        sCmd = Mid(ActiveCell.Value, 3)
        
        'MsgBox "ExSQell: " & sCmd
        
        nRow = nMinRow + 1
        Do Until Cells(nRow, nMinCol).Value = ""
            nRow = nRow + 1
        Loop
        nMaxRow = nRow
        nCol = nMinCol + 1
        Do Until Cells(nMinRow + 1, nCol).Value = ""
            nCol = nCol + 1
        Loop
        nMaxCol = nCol
        
        Range(Cells(nMinRow + 1, nMinCol), Cells(nMaxRow, nMaxCol)).Value = ""
        If Not ExecShell(sCmd, oRs, oFlds) Then
            Exit Sub
        End If
        Application.ScreenUpdating = False
        Call MakeList(ActiveCell.Offset(1, 0), oRs, oFlds, oOpts)
        Cells(nMinRow, nMinCol).Select
        Application.ScreenUpdating = True
        
        Exit Sub
    End If
    
    'まずは完全マッチで検索
    sObjType1 = GetDBObjectType(sWList)
    
    'SHIFT同時押しの場合は部分マッチの検索をキャンセル
    If InStr(sObjType1, "ADO/ODBC") = 0 Then
        If Not GetKeyState(VK_SHIFT) < 0 Then
            sObjType2 = GetDBObjectType(sWBpat)
        End If
    End If
    
    'ALT同時押しの場合は拡張一覧項目を表示
    If GetKeyState(VK_MENU) < 0 Then
        oOpts.Item("拡張一覧") = "on"
    End If
    
    '完全 部分 結果
    'Ｎ   Ｎ   Ｎ
    'Ｎ   Ｏ   Ｌ
    'Ｎ   Ｌ   Ｌ
    'Ｏ   Ｌ   Ｌ
    'Ｏ   Ｏ   Ｏ
    If sObjType1 = "" And sObjType2 = "" Then
        sObjType = ""
    ElseIf sObjType1 <> "LIST" And sObjType2 <> "LIST" Then
        If sObjType1 <> "" Then
            sObjType = sObjType1
        Else
            sObjType = "LIST"
        End If
    Else
        sObjType = "LIST"
    End If
    
    Select Case sObjType
    Case "LIST"
        nRow = nMinRow + 1
        Do Until Cells(nRow, nMinCol).Value = ""
            nRow = nRow + 1
        Loop
        nMaxRow = nRow
        nMaxCol = nMinCol + 2
        
        Range(Cells(nMinRow + 1, nMinCol), Cells(nMaxRow, nMaxCol)).Value = ""
        If Not GetDBObjectList(sWList, oRs, oFlds, oOpts) Then
            Exit Sub
        End If
        If oOpts.Item("拡張一覧") <> "" Then
            For nFldIdx = 4 To oFlds.Count
                Cells(nMinRow, nMinCol + nFldIdx - 1) = oFlds.Item(nFldIdx).Item("Title")
            Next
            oOpts.Item("行再描画") = "on"
        End If
        Application.ScreenUpdating = False
        Call MakeList(ActiveCell.Offset(1, 0), oRs, oFlds, oOpts)
        Cells(nMinRow, nMinCol).Select
        Application.ScreenUpdating = True
    Case "TABLE", "VIEW"
        If Not GetDBObjectInfo(sWhich, sObjType, oRs, oFlds, oOpts) Then
            Exit Sub
        End If
        Application.ScreenUpdating = False
        With GetNewSheet(ccTblShtName, False)
            On Error Resume Next
            nRow = Application.WorksheetFunction.Match(sWhich, .Columns(1), 0)
            On Error GoTo 0
            If nRow <> 0 Then
                If .Cells(nRow, 1) = sWhich Then
                    .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
                    Exit Sub
                End If
            End If
            nRow = 1
            Do Until .Cells(nRow, 2).Value = ""
                nRow = nRow + 1
            Loop
            With oOpts
                If nRow = 1 Then
                    .Item("タイトル") = "フィルタ"
                    .Item("枠固定セル") = "A2"
                    .Item("行タイトル") = "$1:$1"
                End If
                .Item("スタイル") = "罫線描画、縞模様書式"
            End With
            .Range("B:B").ColumnWidth = 25
            Call MakeList(.Cells(nRow, 1), oRs, oFlds, oOpts)
            .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
            .Range("B:B").ColumnWidth = 0
#If VBA7 Then
            Application.PrintCommunication = False
#End If
            With .PageSetup
                .LeftHeader = ""
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.393700787401575)
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
            End With
#If VBA7 Then
            Application.PrintCommunication = True
#End If
        End With
        Application.ScreenUpdating = True
    Case "INDEX"
        If Not GetDBObjectInfo(sWhich, sObjType, oRs, oFlds, oOpts) Then
            Exit Sub
        End If
        With GetNewSheet(ccIdxShtName, False)
            On Error Resume Next
            nRow = Application.WorksheetFunction.Match(sWhich, .Columns(1), 0)
            On Error GoTo 0
            If nRow <> 0 Then
                If .Cells(nRow, 1) = sWhich Then
                    .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
                    Exit Sub
                End If
            End If
            nRow = 1
            Do Until .Cells(nRow, 2).Value = ""
                nRow = nRow + 1
            Loop
            With oOpts
                If nRow = 1 Then
                    .Item("タイトル") = "フィルタ"
                    .Item("枠固定セル") = "A2"
                    .Item("行タイトル") = "$1:$1"
                End If
                .Item("スタイル") = "罫線描画、縞模様書式"
            End With
            Call MakeList(.Cells(nRow, 1), oRs, oFlds, oOpts)
            .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
#If VBA7 Then
            Application.PrintCommunication = False
#End If
            With .PageSetup
                .LeftHeader = ""
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.393700787401575)
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
            End With
#If VBA7 Then
            Application.PrintCommunication = True
#End If
        End With
    Case "PROCEDURE", "FUNCTION", "PACKAGE"
        If Not GetDBObjectInfo(sWhich, sObjType, oRs, oFlds, oOpts) Then
            Exit Sub
        End If
        With GetNewSheet(ccSrcShtName, False)
            On Error Resume Next
            nRow = Application.WorksheetFunction.Match(sWhich, .Columns(1), 0)
            On Error GoTo 0
            If nRow <> 0 Then
                If .Cells(nRow, 1) = sWhich Then
                    .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
                    Exit Sub
                End If
            End If
            nRow = 1
            Do Until .Cells(nRow, 2).Value = ""
                nRow = nRow + 1
            Loop
            With oFlds
                .Item(1).Item("列幅") = 10
                .Item(2).Item("列幅") = 12
                .Item(3).Item("列幅") = 7
                .Item(4).Item("列幅") = 2
                .Item(5).Item("列幅") = 85
            End With
            With oOpts
                If nRow = 1 Then
                    .Item("タイトル") = "フィルタ"
                    .Item("枠固定セル") = "A2"
                    .Item("行タイトル") = "$1:$1"
                End If
            End With
            Call MakeList(.Cells(nRow, 1), oRs, oFlds, oOpts)
            .Cells.AutoFilter Field:=1, Criteria1:="=" & sWhich, Operator:=xlAnd
#If VBA7 Then
            Application.PrintCommunication = False
#End If
            With .PageSetup
                .LeftHeader = ""
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.393700787401575)
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
            End With
#If VBA7 Then
            Application.PrintCommunication = True
#End If
        End With
    Case ""
    Case Else
        MsgBoxST sObjType
    End Select
    
    
End Sub

'拡張情報取得
Sub ST_ExtInfo()
    Dim sWhich As String
    Dim sObjType As String
    Dim nMinRow As Long
    Dim nMinCol As Long
    Dim oExtInfo
    Dim sExtName As String
    Dim sExtText As String
    Dim i As Integer
    sWhich = ActiveCell.Value
    nMinRow = ActiveCell.Row
    nMinCol = ActiveCell.Column

    sObjType = GetDBObjectType(sWhich)
    
    If GetCfgVal("拡張情報：状況") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "状況"
    End If
    If GetCfgVal("拡張情報：更新日時") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "更新日時"
    End If
    If GetCfgVal("拡張情報：シノニム") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "シノニム"
    End If
    If GetCfgVal("拡張情報：件数") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "件数"
    End If
    If GetCfgVal("拡張情報：行数") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "行数"
    End If
    If GetCfgVal("拡張情報：サイズ") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "サイズ"
    End If
    
    If sObjType <> "" And sObjType <> "LIST" Then
        Set oExtInfo = GetDBObjectExtInfo(sWhich, sObjType, sExtName)
        For i = 1 To oExtInfo.Count
            sExtText = sExtText & IIf(sExtText = "", "", vbLf) & oExtInfo.Item(i)
        Next
        If Not GetKeyState(VK_SHIFT) < 0 Then
            Call MsgBalloon(sWhich, sExtText)
        Else
            Call SetCmtVal(ActiveCell, , sExtText)
        End If
    End If
End Sub

'データメンテナンス
Sub ST_Maintenance()
    Dim sConStr As String, sTblName As String, sObjType As String, sObjName As String
    Dim sSQL As String, sErrSQL As String, sErrText As String, nErrLine As Integer
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nMinCol As Long
    Dim oRs, oFld, oFlds, oOpts, oSht As Worksheet
    Dim arLines, nLineNo As Integer
    
    Select Case GetCfgVal("更新SQLの実行")
    Case 2, 3, 4 '確認して実行、問答無用で実行
        '確認無し
    Case Else '実行しない
        Call MsgBoxST(ccUpdateDisclaim)
        Exit Sub
    End Select
    
    sConStr = GetConStr
    Select Case ActiveSheet.Range("A1")
    Case "QPARAMS" '抽出条件入力
        sTblName = Replace(ActiveSheet.Name, ccQpsShtPref, "")
        '入力抽出条件からＳＱＬ文を生成
        sSQL = GetSql("MAINTEN")
        If sSQL = "" Then
            Exit Sub
        End If
    Case "MAINTEN" 'メンテナンス
        'シート上のコメントから接続文字列、ＳＱＬ文を取得
        arLines = Split(GetCmtVal(ActiveSheet.Range("A1")), vbLf)
        sTblName = Replace(ActiveSheet.Name, ccMntShtPref, "")
        For nLineNo = 1 To UBound(arLines)
            sSQL = sSQL & arLines(nLineNo)
        Next
    Case "NAME" 'ソース定義
        With ActiveSheet
            'フィルタ選択されているか
            If .FilterMode Then
                If MsgBoxST("ストアドファンクション/プロシージャをコンパイルしますか？", vbYesNo + vbDefaultButton2) = vbYes _
                Then '☆ADO対応要
                    'コンパイルソースの取得
                    nMinRow = .Columns(1).SpecialCells(xlCellTypeVisible).Areas(2).Row
                    If .Cells(nMinRow, 1) = "" Then nMinRow = 2
                    sObjName = .Cells(nMinRow, 1)
                    nRow = nMinRow
                    Do Until .Cells(nRow, 1) <> sObjName
                        If Not .Cells(nRow, 3).Comment Is Nothing Then
                            .Cells(nRow, 3).Comment.Delete
                        End If
                        .Range(.Cells(nRow, 1), .Cells(nRow, 5)).Font.ColorIndex = xlAutomatic
                        .Cells(nRow, 4) = ""
                        If .Cells(nRow, 3) = "1" Then
                            sSQL = sSQL & "CREATE OR REPLACE "
                        Else
                            .Cells(nRow, 3) = .Cells(nRow - 1, 3) + 1
                        End If
                        sSQL = sSQL & .Cells(nRow, 5) & vbLf
                        nRow = nRow + 1
                    Loop
                    'ＤＤＬ発行
                    If Not ExecSql(sConStr, sSQL) Then
                        Exit Sub
                    End If
                    'エラー行を色付け
                    sErrSQL = GetLibSql("エラー情報取得")
                    sErrSQL = SetSqlParam(sErrSQL, "オブジェクト名", sObjName)
                    If Not ExecSql(sConStr, sErrSQL, oRs) Then
                        Exit Sub
                    End If
                    If Not oRs.EOF Then
                        nErrLine = oRs("LINE")
                    End If
                    Do Until oRs.EOF
                        If oRs("LINE") <> nErrLine Then
                            .Range(.Cells(nMinRow + nErrLine - 1, 1), _
                                    .Cells(nMinRow + nErrLine - 1, 5)).Font.ColorIndex = ccRED
                            Call SetCmtVal(.Cells(nMinRow + nErrLine - 1, 3), , sErrText)
                            .Cells(nMinRow + nErrLine - 1, 4) = "E"
                            nErrLine = oRs("LINE")
                            sErrText = ""
                        End If
                        If sErrText <> "" Then sErrText = sErrText & vbLf
                        sErrText = sErrText & oRs("TEXT")
                        oRs.MoveNext
                    Loop
                    If nErrLine <> 0 Then
                        .Range(.Cells(nMinRow + nErrLine - 1, 1), _
                                .Cells(nMinRow + nErrLine - 1, 5)).Font.ColorIndex = ccRED
                        Call SetCmtVal(.Cells(nMinRow + nErrLine - 1, 3), , sErrText)
                        .Cells(nMinRow + nErrLine - 1, 4) = "E"
                    End If
                End If
            Else
                Call MsgBoxST("コンパイルするソースをフィルタ選択してください")
            End If
        End With
        Exit Sub
    Case "TABLE_NAME", "INDEX_NAME" '表定義、索引定義
        Exit Sub
    Case Else '通常のシート時
        '表引き可能オブジェクトかどうか
        sTblName = Trim(ActiveCell.Value)
        sObjType = GetDBObjectType(sTblName)
        Select Case sObjType
        Case "TABLE"
        Case "FUNCTION", "PROCEDURE"
            If MsgBoxST("ストアドファンクション/プロシージャを再コンパイルしますか？", vbYesNo + vbDefaultButton2) = vbYes _
            Then
                arLines = Split(GetSql, vbCrLf)
                With GetSafeSelection
                    nMinRow = .Row
                    nMinCol = .Column
                End With
                For nLineNo = 0 To UBound(arLines)
                    sObjType = GetDBObjectType(arLines(nLineNo))
                    With Cells(nMinRow + nLineNo, nMinCol)
                        If Not .Comment Is Nothing Then
                            .Comment.Delete
                        End If
                        .Font.ColorIndex = xlAutomatic
                        If sObjType = "FUNCTION" Or sObjType = "PROCEDURE" Then
                            If Not ExecSql(sConStr, "ALTER " & sObjType & " " & arLines(nLineNo) & " COMPILE") Then
                                Call SetCmtVal(.Offset(0, 0), , "データベースエラーが発生")
                                .Font.ColorIndex = ccRED
                            Else
                                sErrSQL = GetLibSql("エラー情報取得")
                                sErrSQL = SetSqlParam(sErrSQL, "オブジェクト名", arLines(nLineNo))
                                If Not ExecSql(sConStr, sErrSQL, oRs) Then
                                    Exit Sub
                                End If
                                If Not oRs.EOF Then
                                    sErrText = ""
                                    Do Until oRs.EOF
                                        If sErrText <> "" Then sErrText = sErrText & vbLf
                                        sErrText = sErrText & Format(oRs("LINE"), "000") & ": " & oRs("TEXT")
                                        oRs.MoveNext
                                    Loop
                                    Call SetCmtVal(.Offset(0, 0), , sErrText)
                                    .Font.ColorIndex = ccRED
                                End If
                            End If
                        Else
                            Call SetCmtVal(.Offset(0, 0), , "オブジェクトが存在しません")
                            .Font.ColorIndex = ccRED
                        End If
                    End With
                Next
            End If
            Exit Sub
        Case Else
'            Call MsgBoxST("テーブルを指定して下さい")
            Exit Sub
        End Select
    
        'ALT同時押し以外の場合は抽出条件指定
        If Not GetKeyState(VK_MENU) < 0 Then
            Call MakeQpsSheet(sTblName)
            Exit Sub
        End If
        
        nMinRow = ActiveCell.Row
        nMinCol = ActiveCell.Column
    
        sSQL = GetTemplateSql(sTblName, ccTPL_SEL, "noselect, nowhere")
        sSQL = Replace(sSQL, "SELECT ", "SELECT ROWID MAINTEN, " & sTblName & ".")
    End Select
    
    'クエリの実行
    If Not ExecQuery(sConStr, sSQL, "MAINTEN", oRs, oFlds, oOpts) Then
        Exit Sub
    End If
        
    'クエリ結果シートの作成
    With oFlds
        Set oFld = New Prop
        oFld.Item("Title") = ""
        oFld.Item("特殊列") = "空白列"
        Set .Item(1, False) = oFld
        'ROWID列
        .Item(1).Item("列幅") = 0
        '追加/更新/削除マーク列
        .Item(2).Item("列幅") = 2.5
        .Item(2).Item("横位置") = "中央"
    End With
    With oOpts
        .Item("タイトル") = "フィルタ"
        .Item("枠固定セル") = "C2"
        .Item("行タイトル") = "$1:$1"
        .Item("列タイトル") = "$A:$B"
        .Item("スタイル") = "罫線描画、項目型色付"
        Select Case GetCfgVal("縞模様表示")
        Case 1, 2 '条件付き書式で表示、セル背景色で表示
            .Item("スタイル") = .Item("スタイル") & "、縞模様書式"
        End Select
        Select Case GetCfgVal("空白セル強調")
        Case 1 '条件付き書式で強調
            .Item("スタイル") = .Item("スタイル") & "、空白セル強調"
        End Select
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If ActiveSheet.Range("A1") = "QPARAMS" Then
        Select Case GetCfgVal("抽出条件シート")
        Case 1 '抽出時に削除
            ActiveSheet.Delete
        Case 2 '抽出時に非表示
            ActiveSheet.Visible = False
        End Select
    End If
    Set oSht = GetNewSheet(ccMntShtPref & sTblName, False, True)
    Call MakeList(oSht, oRs, oFlds, oOpts)
    With GetShtCodeModule(oSht)
        .InsertLines .CountOfLines + 1, "Public Sub Worksheet_Activate()"
        .InsertLines .CountOfLines + 1, "    Application.StatusBar = """ & GetConDisp(sConStr) & """"
        .InsertLines .CountOfLines + 1, "End Sub"
        .InsertLines .CountOfLines + 1, ""
        .InsertLines .CountOfLines + 1, "Private Sub Worksheet_Deactivate()"
        .InsertLines .CountOfLines + 1, "    Application.StatusBar = """""
        .InsertLines .CountOfLines + 1, "End Sub"
        .InsertLines .CountOfLines + 1, ""
        .InsertLines .CountOfLines + 1, "Private Sub Worksheet_Change(ByVal Target As Excel.Range)"
        .InsertLines .CountOfLines + 1, "    Application.EnableEvents = False"
        .InsertLines .CountOfLines + 1, "    If Target.Column > 2 Then"
        .InsertLines .CountOfLines + 1, "        If Cells(Target.Row, 2) = """" Then"
        .InsertLines .CountOfLines + 1, "            If Cells(Target.Row, 1) <> """" Then"
        .InsertLines .CountOfLines + 1, "                Cells(Target.Row, 2) = ""U"""
        .InsertLines .CountOfLines + 1, "            Else"
        .InsertLines .CountOfLines + 1, "                Cells(Target.Row, 2) = ""I"""
        .InsertLines .CountOfLines + 1, "            End If"
        .InsertLines .CountOfLines + 1, "        End If"
        .InsertLines .CountOfLines + 1, "    ElseIf Target.Column < 2 Then"
        .InsertLines .CountOfLines + 1, "        Dim nRow As Long"
        .InsertLines .CountOfLines + 1, "        For nRow = Target.Row To Target.Row + Target.Rows.Count - 1"
        .InsertLines .CountOfLines + 1, "            ActiveSheet.Cells(nRow, 2) = ""I"""
        .InsertLines .CountOfLines + 1, "        Next"
        .InsertLines .CountOfLines + 1, "    End If"
        .InsertLines .CountOfLines + 1, "    Application.EnableEvents = True"
        .InsertLines .CountOfLines + 1, "End Sub"
        .InsertLines .CountOfLines + 1, ""
    End With
    Call SetCmtVal(ActiveSheet.Range("A1"), , sConStr & vbLf & sSQL)
    Call SetCmtVal(ActiveSheet.Range("B1"), , "I:追加 U:更新 D:削除")
    Application.StatusBar = GetConDisp(sConStr)
    'キー項目の太字表示、データタイプ文字色
    Call GetDBObjectInfo(sTblName, "TABLE", oRs, oFlds, oOpts)
    With ActiveSheet
        Dim nCol As Long
        Dim nMaxCol As Long
        nMaxCol = .UsedRange.Columns.Count
        Do Until oRs.EOF
            If oRs("KEY").Value <> "" Then
                For nCol = 3 To nMaxCol
                    If .Cells(1, nCol) = oRs("COLUMN_NAME").Value Then
                        .Cells(1, nCol).Font.Bold = True
                        Exit For
                    End If
                Next
            End If
            Select Case oRs("DATA_TYPE").Value
            Case "DATE"
                For nCol = 3 To nMaxCol
                    If .Cells(1, nCol) = oRs("COLUMN_NAME").Value Then
                        .Cells(1, nCol).Font.ColorIndex = ccPURPLE
                        Exit For
                    End If
                Next
            Case "VARCHAR2", "CHAR"
            Case Else
                For nCol = 3 To nMaxCol
                    If .Cells(1, nCol) = oRs("COLUMN_NAME").Value Then
                        .Cells(1, nCol).Font.ColorIndex = ccBLUE
                        Exit For
                    End If
                Next
            End Select
            oRs.MoveNext
        Loop
    End With
End Sub

Private Function GetShtCodeModule(argSht As Worksheet) As CodeModule
    Dim oComp
    For Each oComp In argSht.Parent.VBProject.VBComponents
        If oComp.Name = argSht.CodeName Then
            Set GetShtCodeModule = oComp.CodeModule
        End If
    Next
End Function

'抽出条件シート生成
Private Function MakeQpsSheet(argObjName)
    Dim sSQL As String, sShtName As String, sConStr As String
    Dim oRs, oFlds, oOpts As New Prop
    Dim nRow As Long
    Dim nMaxRow As Long
    sSQL = GetLibSql("抽出条件シート生成")
    
    sSQL = SetSqlParam(sSQL, "オブジェクト名", argObjName)
    
    sConStr = GetConStr
    If Not ExecSql(sConStr, sSQL, oRs, oFlds, oOpts) Then
        Exit Function
    End If
    Application.ScreenUpdating = False
    sShtName = ccQpsShtPref & argObjName
    With GetNewSheet(sShtName, False, False)
        .Columns(1).ColumnWidth = 10
        Call SetCmtVal(.Range("A1"), , sConStr)
        .Columns(1).ColumnWidth = 0
        .Range("D1") = GetConDisp(sConStr)
        If .Range("A1") = "QPARAMS" Then
            Exit Function
        End If
        .Range("A1") = "QPARAMS"
        Call SetCmtVal(.Range("B1"), , _
            "項目行の位置がクエリの列位置になります。非表示行は表示/ソート/条件が全て無効" & vbLf & _
            "空行を挿入して条件のみを入力すると直前の項目の条件とのOR条件となります")
        Call SetCmtVal(.Range("B2"), , _
            "項目名が太字はキー項目、文字色が黒はVARCHAR, CHAR、紫はDATE、青はそれ以外")
        Call SetCmtVal(.Range("C2"), , _
            "↓カラムコメントを表示、項目名称オプションをONにすると結果の列タイトルが名称になります")
        Call SetCmtVal(.Range("F1"), , _
            "　文字列は前方一致、演算子は<>!=が使用可能" & vbLf & _
            "↓カンマ、改行区切りの文字列はIN展開されます")
        Call SetCmtVal(.Range("F2"), , _
            "↓NULL指定は先頭のISを除いて指定します")
        Call SetCmtVal(.Range("G1"), , _
            "　最初のWHERE,ANDを除いたSQL文を記述" & vbLf & _
            "↓先頭 - - でコメントアウトされます")
        With .Range("B1")
            .Value = "項目条件指定"
            .Font.Bold = True
        End With
        With .Range("B1:G1")
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("H1")
            .Value = "追加SQL文指定"
            .Font.Bold = True
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("I1:K2")
            .HorizontalAlignment = xlHAlignCenter
            With .Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End With
        With .Range("I1")
            .Value = "ロック"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("J1")
            .Value = "項目名称"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("K1")
            .Value = "SQL表示"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With oFlds
            .Item(1).Item("列幅") = 0
            .Item(2).Item("列幅") = 3
            .Item(3).Item("列幅") = 20
            .Item(4).Item("列幅") = 20
            .Item(5).Item("列幅") = 5
            .Item(5).Item("横位置") = "中央"
            .Item(6).Item("列幅") = 5
            .Item(6).Item("横位置") = "中央"
            .Item(7).Item("列幅") = 23
            .Item(8).Item("列幅") = 55
        End With
        With oOpts
            .Item("タイトル") = "-"
            .Item("スタイル") = "罫線描画"
        End With
        Call MakeList(.Cells(2, 1), oRs, oFlds, oOpts)
        nMaxRow = .UsedRange.Rows.Count
        .Range(.Cells(1, 1), .Cells(nMaxRow, 7)).VerticalAlignment = xlVAlignTop
        For nRow = 3 To nMaxRow
            If .Cells(nRow, 6) <> "" Then
                .Range(.Cells(nRow, 2), .Cells(nRow, 4)).Font.Bold = True
            End If
            Select Case .Cells(nRow, 1)
            Case "DATE"
                .Range(.Cells(nRow, 2), .Cells(nRow, 4)).Font.ColorIndex = ccPURPLE
            Case "VARCHAR2", "CHAR"
            Case Else
                .Range(.Cells(nRow, 2), .Cells(nRow, 4)).Font.ColorIndex = ccBLUE
            End Select
        Next
        '.Columns(4).ColumnWidth = 0
        With .Range(.Cells(3, 6), .Cells(nMaxRow, 6)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="昇順,降順"
        End With
#If VBA7 Then
        Application.PrintCommunication = False
#End If
        With .PageSetup
            .PrintArea = "$B:$H"
            .LeftHeader = Replace(sShtName, "&", "")
            .LeftMargin = Application.InchesToPoints(0.393700787401575)
            .RightMargin = Application.InchesToPoints(0.393700787401575)
            .TopMargin = Application.InchesToPoints(0.78740157480315)
            .BottomMargin = Application.InchesToPoints(0.78740157480315)
            .Zoom = False
            .FitToPagesTall = False
            .FitToPagesWide = 1
        End With
#If VBA7 Then
        Application.PrintCommunication = True
#End If
        .Range("H2") = ""
        .Range("A3").Select
        ActiveWindow.FreezePanes = True
        .Range("G3").Select
    End With
    Application.ScreenUpdating = True
    
End Function

'テンプレート生成
Sub ST_Template()
    Dim sSQL As String, sTpl As String
    If GetKeyState(VK_MENU) < 0 Then
        'テンプレート生成
        Const cNonChkWords = "SELECT INSERT UPDATE DELETE CREATE"
        Dim sSetCell As String, sStartWord As String, arWords
        sSetCell = ActiveCell.Offset(1, 0)
        arWords = Split(Trim(sSetCell), " ")
        sStartWord = arWords(0)
        If sSetCell = "" Or InStr(cNonChkWords, sStartWord) <> 0 Then
    '        If ActiveCell <> csTplCell Or sSetCell = "" Then
            If ActiveCell <> csTplCell Then
                cnTplMode = ccTPL_SEL
                csTplCell = ActiveCell
                Set coTplRs = Nothing
            End If
            sTpl = GetTemplateSql(ActiveCell, cnTplMode, argTplRs:=coTplRs)
            If sTpl <> "" Then
                If GetCfgVal("SQLの抽出, 整形, 生成") = 2 Then
                    Call SetSql(sTpl, ActiveCell.Offset(1, 0))
                Else
                    If ST_WIN.ShowModal2(sTpl, "展開") Then
                        Call SetSql(sTpl, ActiveCell.Offset(1, 0))
                    End If
                    If cnTplMode = ccTPL_CRE Then
                        cnTplMode = ccTPL_CLR
                    End If
                End If
                If cnTplMode = ccTPL_CLR Then
                    cnTplMode = ccTPL_SEL
                Else
                    cnTplMode = cnTplMode + 1
                End If
            End If
        End If
    Else
        Dim nRow As Long, nCol As Long, nMinRow As Long, nMaxRow As Long, i As Integer
        Dim bSQL As Boolean, bESC As Boolean, bVAL As Boolean
        Dim sVAL As String, sChar As String
        If ActiveCell <> "" Then
            nRow = ActiveCell.Row
            nCol = ActiveCell.Column
            '文取得
            Do Until nRow = 1 Or ActiveSheet.Cells(nRow, nCol) = ""
                nRow = nRow - 1
            Loop
            If ActiveSheet.Cells(nRow, nCol) = "" Then
                nRow = nRow + 1
            End If
            nMinRow = nRow
            Do Until ActiveSheet.Cells(nRow, nCol) = ""
                sSQL = sSQL & ActiveSheet.Cells(nRow, nCol) & vbCrLf
                nRow = nRow + 1
            Loop
            
            If GetKeyState(VK_SHIFT) < 0 Then
                'SQL文の整形
                sTpl = SqlFmt(sSQL)
            Else
                'SQL生成コード変換/逆変換
                
'                If InStr(sSQL, """") <> 0 And InStr(sSQL, "&") <> 0 Then
'                    '変換
'                    If InStr(sSQL, "vbCrLf") <> 0 Then
'                        sSQL = Replace(sSQL, vbCrLf, "")
'                        sSQL = Replace(sSQL, "vbCrLf", vbCrLf)
'                    End If
'                    For i = 1 To Len(sSQL)
'                        Select Case Mid(sSQL, i, 1)
'                        Case """"
'                            If Not bESC Then
'                                bSQL = Not bSQL
'                                If bSQL Then
'                                    sVAL = Replace(sVAL, " ", "")
'                                    If Left(sVAL, 1) = "&" And InStr(sVAL, vbCr) = 0 Then
'                                        If Right(sVAL, 1) = "&" Then
'                                            sVAL = Left(sVAL, Len(sVAL) - 1)
'                                        End If
'                                        sTpl = sTpl & sVAL
'                                    End If
'                                    sVAL = ""
'                                End If
'                                bESC = True
'                            Else
'                                sTpl = sTpl & """"
'                                bESC = False
'                            End If
'                        Case Else
'                            If bSQL Or Mid(sSQL, i, 1) = vbLf Then
'                                MsgBox sVAL
'                                sTpl = sTpl & Mid(sSQL, i, 1)
'                            Else
'                                sVAL = sVAL & Mid(sSQL, i, 1)
'                            End If
'                            bESC = False
'                        End Select
'                    Next
                
                If InStr(sSQL, """") <> 0 And InStr(sSQL, "&") <> 0 Then
                    '変換
                    If InStr(sSQL, "vbCrLf") <> 0 Then
                        sSQL = Replace(sSQL, vbCrLf, "")
                        sSQL = Replace(sSQL, "vbCrLf", vbCrLf)
                    End If
                    For i = 1 To Len(sSQL)
                        sChar = Mid(sSQL, i, 1)
                        Select Case sChar
                        Case """"
                            If Not bESC Then
                                bSQL = Not bSQL
                                If bSQL Then
                                    sVAL = Replace(sVAL, " ", "")
                                    If Left(sVAL, 1) = "&" And InStr(sVAL, vbCr) = 0 Then
                                        If Right(sVAL, 1) = "&" Then
                                            sVAL = Left(sVAL, Len(sVAL) - 1)
                                        End If
                                        sTpl = sTpl & sVAL
                                    End If
                                    sVAL = ""
'                                    sTpl = sTpl & "1 2"
                                End If
                                bESC = True
                            Else
                                sTpl = sTpl & """"
                                bESC = False
                            End If
                        Case Else
                            If bSQL Or sChar = vbLf Then
                                If sChar = vbLf And sVAL <> " & " & vbCr And sVAL <> vbCr Then
                                    sVAL = Replace(sVAL, " ", "")
                                    If Len(sVAL) > 1 Then
                                        sVAL = Left(sVAL, Len(sVAL) - 1)
                                    End If
                                    If Right(sVAL, 1) = "&" Then
                                        sVAL = Left(sVAL, Len(sVAL) - 1)
                                    End If
                                    sTpl = sTpl & sVAL
                                    sVAL = ""
                                End If
                                sTpl = sTpl & sChar
                            Else
                                sVAL = sVAL & sChar
                            End If
                            bESC = False
                        End Select
                    Next
                Else
                    '逆変換
                    
'                    sTpl = "sSQL = sSQL & """
'                    For i = 1 To Len(sSQL)
'                        Select Case Mid(sSQL, i, 1)
'                        Case vbCr
'                        Case vbLf
'                            sTpl = sTpl & """ & vbCrLf" & vbCrLf
'                            If i <> Len(sSQL) Then
'                                sTpl = sTpl & "sSQL = sSQL & """
'                            End If
'                        Case "&"
'                            If bVAL Then
'                                sTpl = sTpl & sVAL & " & "
'                                sVAL = ""
'                            Else
'                                sTpl = sTpl & """ & "
'                            End If
'                            bVAL = True
'                        Case Else
'                            If bVAL Then
'                                Select Case Mid(sSQL, i, 1)
'                                Case "'", " "
'                                    sTpl = sTpl & sVAL & " & """ & Mid(sSQL, i, 1)
'                                    bVAL = False
'                                    sVAL = ""
'                                Case Else
'                                    sVAL = sVAL & Mid(sSQL, i, 1)
'                                End Select
'                            Else
'                                sTpl = sTpl & Mid(sSQL, i, 1)
'                            End If
'                        End Select
'                    Next
                    
                    sTpl = "sSQL = sSQL & """
                    For i = 1 To Len(sSQL)
                        sChar = Mid(sSQL, i, 1)
                        Select Case sChar
                        Case vbLf
                            If Not bESC Then
                                sTpl = sTpl & """"
                            End If
                            sTpl = sTpl & " & vbCrLf" & vbCrLf
                            If i <> Len(sSQL) Then
                                sTpl = sTpl & "sSQL = sSQL & """
                            End If
                            bESC = False
                        Case "&"
                            If bVAL Then
                                sTpl = sTpl & sVAL & " & "
                                sVAL = ""
                            Else
                                sTpl = sTpl & """ & "
                            End If
                            bVAL = True
                        Case Else
                            If bVAL Then
                                Select Case sChar
                                Case "'", " ", vbCr
                                    If sChar <> vbCr Then
                                        sTpl = sTpl & sVAL & " & """ & sChar
                                    Else
                                        sTpl = sTpl & sVAL
                                        bESC = True
                                    End If
                                    bVAL = False
                                    sVAL = ""
                                Case Else
                                    sVAL = sVAL & sChar
                                End Select
                            Else
                                If sChar <> vbCr Then
                                    sTpl = sTpl & sChar
                                End If
                            End If
                        End Select
                    Next
                End If
            End If
            If sTpl <> "" Then
                If GetCfgVal("SQLの抽出, 整形, 生成") = 2 Then
                    Call SetSql(sTpl, ActiveSheet.Cells(nMinRow, nCol))
                Else
                    If ST_WIN.ShowModal2(sTpl, "展開") Then
                        Call SetSql(sTpl, ActiveSheet.Cells(nMinRow, nCol))
                    End If
                End If
            End If
        End If
    End If
End Sub

'SQL*Plusの起動
Sub ST_Plus()
    Dim sConStr As String, nPID As Long, hwnd As Long, sSQL As String
    Dim sCmdPara As String, sStaFile As String, nFn As Integer
    sConStr = GetConStr
    sSQL = GetSql
    
    'ALT同時押しの場合は AUTOTRACE設定
    If GetKeyState(VK_MENU) < 0 Then
        With GetCfgSheet
            sStaFile = .Parent.Path & "\" & Left(.Parent.Name, Len(.Parent.Name) - 4) & ".sql"
        End With
        nFn = FreeFile
        Open sStaFile For Output As nFn
        Print #nFn, "set trims on;"
        Print #nFn, "set lines 120;"
        Print #nFn, "set autotrace traceonly explain;"
        Print #nFn, "col PLAN_PLUS_EXP format a100;"
        Print #nFn, sSQL
        Print #nFn, "/"
        Close nFn
        sCmdPara = " ""@" & sStaFile & """"
    End If
    
    'SQL*Plusの起動
    nPID = Shell("sqlplusw " & sConStr & sCmdPara, vbNormalFocus)
    
    'ウインドウタイトルの変更
    hwnd = FindWindow(vbNullString, "Oracle SQL*Plus")
    If hwnd <> 0 Then
        Call SetWindowText(hwnd, GetConDisp(sConStr))
    End If

    'セル位置のＳＱＬを送信
    Application.Wait DateAdd("s", 1.5, Now) 'ウェイト秒数の調整要
    If sCmdPara = "" And ActiveSheet.Range("A1") <> "QPARAMS" Then
        SendKeys sSQL
    End If
End Sub

'ライブラリメニュー表示
Sub ST_Library(Optional argMode = "")
    
    If GetKeyState(VK_SHIFT) < 0 And GetKeyState(VK_MENU) < 0 Then
        With GetLibSheet
            If .Visible = False Then
                .Cells.AutoFilter Field:=2, Criteria1:="<>*SQLTOOL内部使用*", Operator:=xlAnd
                .Visible = True
                .Parent.Activate
                .Select
            Else
                .Visible = False
            End If
        End With
    Else
        If argMode = "URL" Then
            cnLibMode = 5 'お気に入りＵＲＬメニュー
        Else
            If GetKeyState(VK_MENU) < 0 Then
                cnLibMode = 1 'ライブラリSQLをクエリ実行
            ElseIf GetKeyState(VK_SHIFT) < 0 Then
                If ActiveSheet.Name = ccLibShtName Then
                    Exit Sub
                Else
                    cnLibMode = 2 'セル位置にライブラリSQLを挿入
                End If
            Else
                If ActiveSheet.Name = ccLibShtName Then
                    cnLibMode = 4 'お気に入りブック選択メニュー
                Else
                    cnLibMode = 3 'お気に入りブックメニュー
                End If
            End If
        End If
        Dim nRow As Long
        Dim oSht As Worksheet
        Dim oWbk As Workbook
        Dim oPop As CommandBar
        Dim oMnu
        On Error Resume Next
        Set oPop = Application.CommandBars(ccAppTitle)
        On Error GoTo 0
        If Not oPop Is Nothing Then
            oPop.Delete
        End If
        Set oPop = Application.CommandBars.Add(Name:=ccAppTitle, Position:=msoBarPopup, Temporary:=True)
        Set oSht = GetLibSheet
        If cnLibMode = 4 Then
            If oSht.Cells(ActiveCell.Row, 5) <> "" Then
                Exit Sub
            End If
            For Each oWbk In Workbooks
                If oWbk.Name <> oSht.Parent.Name And InStr(oWbk.FullName, ":") Then
                    With oPop.Controls.Add(msoControlButton)
                        .Caption = Left(oWbk.Name, Len(oWbk.Name) - 4)
                        .OnAction = "ST_Library_OnAction"
                        .Parameter = oWbk.Name
                    End With
                End If
            Next
        Else
            nRow = 2
            Do Until oSht.Cells(nRow, 3) = ""
                If (oSht.Cells(nRow, 4) = "M" And (cnLibMode = 1 Or cnLibMode = 2)) Or _
                    (oSht.Cells(nRow, 4) = "Q" And cnLibMode = 1) Or _
                    (oSht.Cells(nRow, 4) = "S" And cnLibMode = 2) Or _
                    (oSht.Cells(nRow, 4) = "B" And cnLibMode = 3) Or _
                    (oSht.Cells(nRow, 4) = "U" And cnLibMode = 5) _
                Then
                    If oSht.Cells(nRow, 2) <> "" Then
                        Set oMnu = oPop.FindControl(Type:=msoControlPopup, Tag:=oSht.Cells(nRow, 2))
                        If oMnu Is Nothing Then
                            Set oMnu = oPop.Controls.Add(msoControlPopup)
                            With oMnu
                                .Tag = oSht.Cells(nRow, 2)
                                .Caption = oSht.Cells(nRow, 2)
                            End With
                        End If
                    Else
                        Set oMnu = oPop
                    End If
                    With oMnu.Controls.Add(msoControlButton)
                        .Caption = oSht.Cells(nRow, 3)
                        .OnAction = "ST_Library_OnAction"
                        .Parameter = oSht.Cells(nRow, 3)
                    End With
                End If
                nRow = nRow + 1
            Loop
        End If
        If oPop.Controls.Count > 0 Then
            oPop.ShowPopup
        End If
    End If
    
End Sub

Private Sub ST_Library_OnAction()
    Dim sSQL As String, oPrms, sPrm, sInp
    With Application.CommandBars.ActionControl
        Select Case cnLibMode
        Case 1 'クエリ実行
            sSQL = GetLibSql(.Parameter)
            '&置換変数がある場合は入力プロンプト表示
            If GetSqlParams(sSQL, oPrms) Then
                For Each sPrm In oPrms.Keys
                    sInp = InputBox(sPrm, ccAppTitle)
                    sSQL = SetSqlParam(sSQL, sPrm, sInp)
                Next
            End If
            Call ST_Query(sSQL, .Parameter)
        Case 2 'SQL挿入
            sSQL = GetLibSql(.Parameter)
            ActiveCell.Value = sSQL
        Case 3 'ブックオープン
            On Error Resume Next
            Call Workbooks.Open(GetLibSql(.Parameter))
            On Error GoTo 0
        Case 4 'ブックパス設定
            Cells(ActiveCell.Row, 3) = Left(.Parameter, Len(.Parameter) - 4)
            Cells(ActiveCell.Row, 4) = "B"
            Cells(ActiveCell.Row, 5) = Workbooks(.Parameter).FullName
        Case 5 'URLブラウザ表示
            Dim oIE
            Set oIE = CreateObject("InternetExplorer.application")
            Call ShowWindow(oIE.hwnd, SW_MAXIMIZE)
            oIE.Visible = True
            Call oIE.Navigate(GetLibSql(.Parameter))
        End Select
    End With
End Sub

Sub ST_Knowledge()
    
    If GetKeyState(VK_MENU) < 0 Then
        'WfOU起動
        Dim oWfOU
        Set oWfOU = GetCfgSheet.Parent.Worksheets(ccWfOUShtName)
        Application.ScreenUpdating = False
        Set oWfOU = GetNewSheet(oWfOU.Name, True, False, oWfOU)
        oWfOU.Visible = True
        Application.ScreenUpdating = True
    ElseIf GetKeyState(VK_MENU) < 0 Then
    Else
        'URLメニュー表示
        Call ST_Library("URL")
    End If
    
End Sub

'接続文字列の取得
Private Function GetConStr() As String
    Dim i As Integer
    Dim arLines, oCfgSht
    Set oCfgSht = GetCfgSheet
    With ActiveSheet
        Select Case .Range("A1")
        Case "QPARAMS" '抽出条件入力
            GetConStr = GetCmtVal(ActiveSheet.Range("A1"))
        Case "MAINTEN" 'メンテナンス
            arLines = Split(GetCmtVal(ActiveSheet.Range("A1")), vbLf)
            GetConStr = arLines(0)
        Case Else
            Dim sCmt As String
            sCmt = GetCmtVal(.Range("A1"))
            If InStr(UCase(sCmt), "SELECT") <> 0 Then
                arLines = Split(sCmt, vbLf)
                GetConStr = Replace(arLines(0), "接続先: ", "")
                Exit Function
            End If
            Const cNoConIds = "No. TABLE_NAME INDEX_NAME NAME"
            If .Name <> ccCfgShtName And .Range("A1") <> "" And InStr(cNoConIds, .Range("A1")) = 0 Then
                GetConStr = .Cells(1, 1)
                Exit Function
            End If
            For i = 1 To cnCfgMaxRow
                If oCfgSht.Cells(i, 2) <> "" Then
                    GetConStr = oCfgSht.Cells(i, 1)
                    Exit Function
                End If
            Next
        End Select
    End With
End Function

'シェルコマンドの取得
Private Function GetShCmd() As String
    
    'ここも取りあえずこれで
    GetShCmd = GetCfgSheet.Range("A27")
    
End Function

'デフォルト実行モードの取得
Private Function GetDefExecMode() As Integer
    
    '取りあえずこれで
    Dim j As Integer
    Dim nValiType As Integer
    Dim aValiList, oCfgSht
    Set oCfgSht = GetCfgSheet
    With oCfgSht.Range("E1")
        On Error Resume Next
        nValiType = .Validation.Type
        On Error GoTo 0
        If nValiType = xlValidateList Then
            aValiList = Split(.Validation.Formula1, ",")
            For j = 0 To UBound(aValiList)
                If aValiList(j) = .Value Then
                    GetDefExecMode = j + 1
                    Exit Function
                End If
            Next
        End If
    End With
    GetDefExecMode = -1
    
End Function

'カレントディレクトリの取得
Private Function GetCurDir() As String
    
    '取りあえずこれで
    Dim i As Integer
    Dim oCfgSht
    Set oCfgSht = GetCfgSheet
    With oCfgSht
        For i = cnCurDirRowS To cnCurDirRowE
            If .Cells(i, 2) <> "" Then
                GetCurDir = .Cells(i, 1)
                Exit Function
            End If
        Next
    End With
    
End Function

Private Function GetConObj(argConStr, Optional argIsUpd = False)
    Dim sDSN As String, sUID As String, sPWD As String
    Dim aConStr, oCon, nConMode, nCmtMode, oConItem
    
    Set oCon = Nothing
    nConMode = GetCfgVal("接続モード")
    nCmtMode = GetCfgVal("コミットモード")
    
    '接続文字列の分解
    If InStr(argConStr, "@") <> 0 Then
        aConStr = Split(argConStr, "@")
        sUID = aConStr(0)
        sDSN = aConStr(1)
    Else
        sDSN = argConStr
    End If
    If InStr(sUID, "/") <> 0 Then
        aConStr = Split(sUID, "/")
        sUID = aConStr(0)
        sPWD = aConStr(1)
    End If
    
    Select Case nCmtMode
    Case "1" '即時コミット
        '毎回、接続オブジェクトを作成
    Case "2" 'メニューで制御
        '接続キーで検索し、存在すればセット
        With coConObjs
            If .Exists(nConMode & argConStr) Then
                Set oConItem = .Item(nConMode & argConStr)
                If argIsUpd And InStr(oConItem.Item("Disp"), "*") = 0 Then
                    oConItem.Item("Disp") = oConItem.Item("Disp") & "*"
                End If
                Set oCon = oConItem.Item("Obj")
            End If
        End With
    Case "3" '接続サーバ使用
        '接続サーバに委譲
    End Select
    
    '接続が無ければ
    If oCon Is Nothing Then
        '新たに生成
        Select Case nConMode
        Case "1" 'OO4O/SQL*Net
            Set oCon = CreateObject("OracleInProcServer.XoraSession"). _
                        OpenDatabase(sDSN, sUID & "/" & sPWD, 0&)
        Case "2" 'ADO/ODBC
            Set oCon = CreateObject("ADODB.Connection")
            If sUID <> "" Then
                Call oCon.Open("DSN=" & sDSN & ";DBQ=" & sDSN & ";UID=" & sUID & ";PWD=" & sPWD)
            Else
                Call oCon.Open("DSN=" & sDSN)
            End If
        End Select
        '即時コミット以外なら
        If nCmtMode <> 1 Then
            '保持リストに登録
            Set oConItem = New Prop
            If argIsUpd Then
                oConItem.Item("Disp") = GetConDisp(argConStr) & "*"
            Else
                oConItem.Item("Disp") = GetConDisp(argConStr)
            End If
            Set oConItem.Item("Obj") = oCon
            Set coConObjs.Item(nConMode & argConStr) = oConItem
            'トランザクション開始
            oCon.BeginTrans
        End If
    End If
    
    Set GetConObj = oCon
End Function

Private Function GetConDisp(argConStr) As String
    Dim arConStr
    If InStr(argConStr, "/") <> 0 Then
        arConStr = Split(argConStr, "/")
        GetConDisp = arConStr(0)
        If InStr(argConStr, "@") <> 0 Then
            arConStr = Split(arConStr(1), "@")
            GetConDisp = GetConDisp & "@" & arConStr(1)
        End If
    Else
        GetConDisp = argConStr
    End If
End Function

'設定シートの取得
Function GetCfgSheet() As Worksheet
    Dim oWbk, oSht
    If coSTBook Is Nothing Then
        For Each oWbk In Workbooks
            For Each oSht In oWbk.Worksheets
                If oSht.CodeName = "ST_CFG" Then
                    Set coSTBook = oWbk
                    Set GetCfgSheet = oSht
                    Exit Function
                End If
            Next
        Next
        Set GetCfgSheet = Sheets(ccCfgShtName)
    Else
        Set GetCfgSheet = coSTBook.Sheets(ccCfgShtName)
    End If
End Function

'ライブラリシートの取得
Function GetLibSheet() As Worksheet
    Dim oWbk, oSht
    If coSTBook Is Nothing Then
        For Each oWbk In Workbooks
            For Each oSht In oWbk.Worksheets
                If oSht.CodeName = "ST_LIB" Then
                    Set coSTBook = oWbk
                    Set GetLibSheet = oSht
                    Exit Function
                End If
            Next
        Next
        Set GetLibSheet = Sheets(ccLibShtName)
    Else
        Set GetLibSheet = coSTBook.Sheets(ccLibShtName)
    End If
End Function

'設定項目の取得
Private Function GetCfgVal(argPropName) As String
    Dim i As Integer
    Dim j As Integer
    Dim nValiType As Integer
    Dim aValiList, oCfgSht
    Set oCfgSht = GetCfgSheet
    For i = 1 To cnCfgMaxRow
        If oCfgSht.Cells(i, 4) = argPropName Then
            With oCfgSht.Cells(i, 5)
                On Error Resume Next
                nValiType = .Validation.Type
                On Error GoTo 0
                If nValiType = xlValidateList Then
                    aValiList = Split(.Validation.Formula1, ",")
                    For j = 0 To UBound(aValiList)
                        If aValiList(j) = .Value Then
                            GetCfgVal = j + 1
                            Exit For
                        End If
                    Next
                Else
                    GetCfgVal = .Value
                End If
            End With
            Exit Function
        End If
    Next
End Function

'ＳＱＬ文の取得
Private Function GetSql(Optional argSQLType = "", Optional argSqlTitle) As String
    Dim nRow As Long
    Dim nMaxRow As Long
    Dim nCol As Long
    Dim arStmt
    Dim sVAL As String
    Dim sSQL As String
    nRow = ActiveCell.Row
    nCol = ActiveCell.Column
    sVAL = Trim(ActiveCell.Value)
    
    If ActiveSheet.Name = ccLibShtName Then
        'ライブラリシート
        argSQLType = "SELECT"
        sSQL = ActiveSheet.Cells(nRow, nCol)
    ElseIf ActiveSheet.Range("A1") = "QPARAMS" Then
        '抽出条件シート
        sSQL = GetSqlQpsSheet(ActiveSheet, argSQLType, argSqlTitle)
    Else
        If Selection.Areas.Count = 1 And Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
            '☆現在位置からのSQL取得は改善の余地あり
            If sVAL <> "" And sVAL <> "No." Then
                '文開始位置取得
                Do Until nRow = 1 Or ActiveSheet.Cells(nRow, nCol) = ""
                    nRow = nRow - 1
                Loop
                If ActiveSheet.Cells(nRow, nCol) = "" Then
                    nRow = nRow + 1
                End If
                'SQL文開始キーワード検索
                Do Until ActiveSheet.Cells(nRow, nCol) = ""
                    arStmt = Split(UCase(Trim(ActiveSheet.Cells(nRow, nCol))), " ")
                    Select Case arStmt(0)
                    Case "SELECT"
                        argSQLType = "SELECT"
                        Exit Do
                    Case "INSERT", "UPDATE", "DELETE", "MERGE"
                        argSQLType = "DML"
                        Exit Do
                    Case "CREATE", "ALTER", "DROP", "RENAME", "TRUNCATE", "COMMENT"
                        argSQLType = "DDL"
                        Exit Do
                    Case "GRANT", "REVOKE"
                        argSQLType = "DCL"
                        Exit Do
'                    Case "DECLARE"
'                        argSQLType = "BLOCK"
'                        Exit Do
                    End Select
                    nRow = nRow + 1
                Loop
                'タイトル取得
                If Not IsMissing(argSqlTitle) And nRow - 1 > 0 Then
                    If argSQLType <> "" And ActiveSheet.Cells(nRow - 1, nCol) <> "" Then
                        argSqlTitle = ActiveSheet.Cells(nRow - 1, nCol)
                    End If
                End If
                'SQL文取得
                Do Until ActiveSheet.Cells(nRow, nCol) = ""
                    sSQL = sSQL & ActiveSheet.Cells(nRow, nCol) & vbCrLf
                    nRow = nRow + 1
                Loop
            Else
                'クエリ結果表示シートの再抽出
                Dim sCmt As String, arLines, nLineNo As Integer
                sCmt = GetCmtVal(ActiveSheet.Range("A1"))
                If InStr(UCase(sCmt), "SELECT") <> 0 Then
                    arLines = Split(sCmt, vbLf)
                    For nLineNo = 1 To UBound(arLines)
                        sSQL = sSQL & arLines(nLineNo) & vbCrLf
                    Next
                    argSqlTitle = ActiveSheet.Name
                    argSQLType = "RELOAD"
                End If
            End If
            'SQL文が取得できない場合、表引き可能なオブジェクト名なら抽出条件指定
            If sSQL = "" Then
'                sSQL = GetTemplateSql(sVAL, ccTPL_SEL, "nowhere")
'                argSQLType = IIf(sSQL <> "", "SELECT", "")
                Dim sObjType
                sObjType = GetDBObjectType(sVAL)
                Select Case sObjType
                Case "TABLE", "VIEW"
                    argSQLType = "TABLE"
                    '☆
                    argSqlTitle = sVAL
                    sSQL = GetTemplateSql(sVAL, ccTPL_SEL, "nowhere")
                Case "FUNCTION", "PROCEDURE"
                    argSQLType = sObjType
                    argSqlTitle = sVAL
                    sSQL = sVAL
                End Select
            End If
        Else
            '選択範囲のSQL文を取得
            Dim oArea, oRow
            sSQL = ""
'            For Each oRow In GetSafeSelection.Columns(1).Rows
'                sSQL = sSQL & oRow.Value & vbCrLf
'            Next
            For Each oArea In GetSafeSelection
                For Each oRow In oArea.Columns(1).Rows
                    sSQL = sSQL & oRow.Value & vbCrLf
                Next
            Next
            If InStr(UCase(sSQL), "SELECT") <> 0 Then
                argSQLType = "SELECT"
            
                'SELECTより前に ( がある場合は削除
                Dim nStart As Integer, nSCnt As Integer, nEnd As Integer, nECnt As Integer, i As Integer
                nStart = InStr(UCase(sSQL), "SELECT")
                If nStart > 0 Then
                    '開始/終了 括弧の数をカウント
                    i = 1
                    Do
                        i = InStr(i, sSQL, "(")
                        If i = 0 Then Exit Do
                        nSCnt = nSCnt + 1
                        i = i + 1
                    Loop
                    i = 1
                    Do
                        i = InStr(i, sSQL, ")")
                        If i = 0 Then Exit Do
                        nECnt = nECnt + 1
                        i = i + 1
                    Loop
                    '終了括弧が足りない分を開始数に加算
                    nSCnt = nECnt - nSCnt
                    nECnt = 0
                    'SELECT前の括弧数をカウント
                    For i = 1 To nStart - 1
                        Select Case Mid(sSQL, i, 1)
                        Case "("
                            nSCnt = nSCnt + 1
                        Case ")"
                            nSCnt = nSCnt - 1
                        End Select
                    Next
                    'SQL文末からの括弧数が一致するまで削除
                    For i = Len(sSQL) To 1 Step -1
                        Select Case Mid(sSQL, i, 1)
                        Case ")"
                            nECnt = nECnt + 1
                        Case "("
                            nECnt = nECnt - 1
                        End Select
                        If nSCnt = nECnt Then
                            nEnd = i - 1
                            Exit For
                        End If
                    Next
                    sSQL = Mid(sSQL, nStart, nEnd - nStart + 1)
                End If
            Else
                argSQLType = "DML"
            End If
        End If
    End If
    '***
    'SQL文末の改行、;を削除
    If Right(sSQL, 2) = vbCrLf Then
        sSQL = Left(sSQL, Len(sSQL) - 2)
    End If
    If Right(sSQL, 1) = vbCr Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If Right(sSQL, 1) = vbLf Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    If Right(sSQL, 1) = ";" Then
        sSQL = Left(sSQL, Len(sSQL) - 1)
    End If
    '&置換変数がある場合は入力プロンプト表示
    Dim oPrms, sPrm, sInp
    If GetSqlParams(sSQL, oPrms) Then
        For Each sPrm In oPrms.Keys
            sInp = InputBox(sPrm, ccAppTitle)
            sSQL = SetSqlParam(sSQL, sPrm, sInp)
        Next
    End If
    GetSql = sSQL
End Function

'抽出条件指定シートからＳＱＬ文を生成
Function GetSqlQpsSheet(argSht As Worksheet, argSQLType, argSqlTitle)
    Dim nRow As Long, nMaxRow As Long
    Dim sSQL As String, sTmpStr As String, sFldType As String, sFldName As String
    Dim sSelFlds As String, sWheFlds As String, sAddWhes As String, sOrdFlds As String
    Dim nOpr As Integer, sOpr As String, sExp As String, sDat As String, arExp
    Dim i As Integer, bCmtFldName As Boolean, sTA As String, sFA As String, sTblName As String
    With argSht
        '列位置 1:データ型 3:項目名 4:項目名称 5:表示 6:ソート 7:条件値 8:追加SQL文
        sTblName = Replace(ActiveSheet.Name, ccQpsShtPref, "")
        '☆
        argSqlTitle = sTblName
        nMaxRow = .UsedRange.Rows.Count
        If argSQLType = "MAINTEN" Then
            sSQL = sSQL & "SELECT ROWID MAINTEN," & vbCrLf
        Else
            sSQL = sSQL & "SELECT" & vbCrLf
            If .Range("J2") <> "" Then
                bCmtFldName = True
            End If
        End If
        If .Range("E1") <> "" Then
            sTA = " " & .Range("E1")
            sFA = .Range("E1") & "."
        End If
        'SELECT指定フィールド
        sTmpStr = ""
        For nRow = 3 To nMaxRow
            If .Cells(nRow, 5) <> "" And .Rows(nRow).RowHeight > 0 Then
                If sTmpStr <> "" Or sSelFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                End If
                If Len(sTmpStr) > ccMaxSqlWidth Then
                    sSelFlds = sSelFlds & "  " & sTmpStr & vbCrLf
                    sTmpStr = ""
                End If
                If Not bCmtFldName Then
                    sTmpStr = sTmpStr & sFA & .Cells(nRow, 3)
                Else
                    If .Cells(nRow, 4) <> "" Then
                        sTmpStr = sTmpStr & sFA & .Cells(nRow, 3) & " """ & .Cells(nRow, 4) & """"
                    Else
                        sTmpStr = sTmpStr & sFA & .Cells(nRow, 3)
                    End If
                End If
            End If
        Next
        If sTmpStr <> "" Then
            sSelFlds = sSelFlds & "  " & sTmpStr
        End If
        '隠し追加SELECT指定フィールド
        If .Cells(2, 8) <> "" Then
            sSQL = sSQL & "  " & .Cells(2, 8) & ", " & vbCrLf
        End If
        sSQL = sSQL & sSelFlds & vbCrLf
        sSQL = sSQL & "FROM " & sTblName & sTA & vbCrLf
        'WHERE指定フィールド
        Const cOP_NORM = 0
        Const cOP_NULL = 1
        Const cOP_IN = 2
        Const cOP_OP = 3
        sTmpStr = ""
        sFldType = .Cells(3, 1)
        sFldName = .Cells(3, 3)
        For nRow = 3 To nMaxRow
            '項目型/項目名を退避
            If .Cells(nRow, 3) <> "" Then
                sFldType = .Cells(nRow, 1)
                sFldName = .Cells(nRow, 3)
            End If
            If .Cells(nRow, 7) <> "" And .Rows(nRow).RowHeight > 0 Then
                '項目名が空の場合は直前の項目のOR条件
                If .Cells(nRow, 3) = "" Then
                    If sTmpStr <> "" Then
                        sTmpStr = sTmpStr & " OR "
                    End If
                Else
                    'ORで繋いだ一つ前の条件があればWHERE句に追加
                    If sTmpStr <> "" Then
                        If sWheFlds <> "" Then
                            sWheFlds = sWheFlds & vbCrLf & "  AND "
                        End If
                        If InStr(sTmpStr, " OR ") = 0 Then
                            sWheFlds = sWheFlds & sTmpStr
                        Else
                            sWheFlds = sWheFlds & "(" & sTmpStr & ")"
                        End If
                    End If
                    '条件をリセット
                    sTmpStr = ""
                End If
                '変数に抽出条件をセット
                sExp = .Cells(nRow, 7)
                '演算子タイプの判定、条件値の切り出し
                If InStr(UCase(.Cells(nRow, 7)), "NULL") Then
                    nOpr = cOP_NULL
                ElseIf InStr(.Cells(nRow, 7), ",") <> 0 Or InStr(.Cells(nRow, 7), vbLf) <> 0 Then
                    nOpr = cOP_IN
                    If InStr(sExp, "<>") <> 0 Or InStr(sExp, "!") <> 0 Then
                        sOpr = "NOT IN"
                    Else
                        sOpr = "IN"
                    End If
                    sExp = Replace(sExp, "<", "")
                    sExp = Replace(sExp, ">", "")
                    sExp = Replace(sExp, "!", "")
                    sExp = Replace(sExp, "=", "")
                    If InStr(sExp, ",") <> 0 Then
                        If InStr(sExp, vbLf) <> 0 Then
                            arExp = Split(Replace(sExp, ",", ""), vbLf)
                        Else
                            arExp = Split(sExp, ",")
                        End If
                    Else
                        arExp = Split(sExp, vbLf)
                    End If
                    sExp = "("
                    For i = 0 To UBound(arExp)
                        If i <> 0 Then
                            sExp = sExp & ","
                        End If
                        Select Case sFldType
                        Case "VARCHAR2", "CHAR", "DATE"
                            sExp = sExp & "'" & arExp(i) & "'"
                        Case Else
                            sExp = sExp & arExp(i)
                        End Select
                    Next
                    sExp = sExp & ")"
                ElseIf InStr(.Cells(nRow, 7), "<") <> 0 Or InStr(.Cells(nRow, 7), ">") <> 0 _
                    Or InStr(.Cells(nRow, 7), "!") <> 0 Or InStr(.Cells(nRow, 7), "=") <> 0 _
                Then
                    nOpr = cOP_OP
                    sExp = Replace(sExp, "<", "")
                    sExp = Replace(sExp, ">", "")
                    sExp = Replace(sExp, "!", "")
                    sExp = Replace(sExp, "=", "")
                    sOpr = Replace(Trim(.Cells(nRow, 7)), Trim(sExp), "")
                Else
                    nOpr = cOP_NORM
                End If
                sExp = LTrim(sExp)
                'データ型に応じて条件式を作成
                Select Case sFldType
                Case "VARCHAR2", "CHAR"
                    sTmpStr = sTmpStr & sFA & sFldName
                    
                    Select Case nOpr
                    Case cOP_NORM
                        sTmpStr = sTmpStr & " LIKE '" & sExp & "%'"
                    Case cOP_NULL
                        sTmpStr = sTmpStr & " IS " & sExp
                    Case cOP_IN
                        'ワイルドカード指定の場合は LIKE OR展開
                        If InStr(sExp, "%") <> 0 Or InStr(sExp, "%") Then
                            Dim sLikeOrFld, sLikeOrOpr, sLikeExp
                            sLikeOrFld = sTmpStr
                            If sOpr = "IN" Then
                                sLikeOrOpr = " LIKE "
                            Else
                                sLikeOrOpr = " NOT LIKE "
                            End If
                            sLikeExp = sExp
                            sLikeExp = Replace(sLikeExp, "(", "")
                            sLikeExp = Replace(sLikeExp, ")", "")
                            sLikeExp = Replace(sLikeExp, ",", " OR " & sLikeOrFld & sLikeOrOpr)
                            sTmpStr = sLikeOrFld & sLikeOrOpr & sLikeExp
                        Else
                            sTmpStr = sTmpStr & " " & sOpr & " " & sExp
                        End If
                    Case cOP_OP
                        If sOpr = "<>" Or sOpr = "!=" Then
                            sTmpStr = sTmpStr & " NOT LIKE '" & sExp & "%'"
                        Else
                            sTmpStr = sTmpStr & " " & sOpr & " '" & sExp & "'"
                        End If
                    End Select
                Case "DATE" '☆ ADO対応要
                    If nOpr <> cOP_IN Then
                        sDat = sExp
                    Else
                        sDat = arExp(0)
                    End If
                    
                    Select Case Len(sDat)
                    Case 19
                        sTmpStr = sTmpStr & "TO_CHAR(" & sFA & sFldName & ",'YYYY/MM/DD HH24:MI:SS')"
                    Case 103
                        sTmpStr = sTmpStr & "TO_CHAR(" & sFA & sFldName & ",'YYYY/MM/DD')"
                    Case 7
                        sTmpStr = sTmpStr & "TO_CHAR(" & sFA & sFldName & ",'YYYY/MM')"
                    Case 8
                        sTmpStr = sTmpStr & "TO_CHAR(" & sFA & sFldName & ",'HH24:MI:SS')"
                    Case Else
                        sTmpStr = sTmpStr & sFA & sFldName
                    End Select
                    
                    Select Case nOpr
                    Case cOP_NORM
                        sTmpStr = sTmpStr & " = '" & sExp & "'"
                    Case cOP_NULL
                        sTmpStr = sTmpStr & " IS " & sExp
                    Case cOP_IN
                        sTmpStr = sTmpStr & " " & sOpr & " " & sExp
                    Case cOP_OP
                        sTmpStr = sTmpStr & " " & sOpr & " '" & sExp & "'"
                    End Select
                Case Else
                    sTmpStr = sTmpStr & sFA & sFldName
                    
                    Select Case nOpr
                    Case cOP_NORM
                        sTmpStr = sTmpStr & " = " & sExp
                    Case cOP_NULL
                        sTmpStr = sTmpStr & " IS " & sExp
                    Case cOP_IN
                        sTmpStr = sTmpStr & " " & sOpr & " " & sExp
                    Case cOP_OP
                        sTmpStr = sTmpStr & " " & sOpr & " '" & sExp & "'"
                    End Select
                End Select
            End If
        Next
        If sTmpStr <> "" Then
            If sWheFlds <> "" Then
                sWheFlds = sWheFlds & vbCrLf & "  AND "
            End If
            If InStr(sTmpStr, " OR ") = 0 Then
                sWheFlds = sWheFlds & sTmpStr
            Else
                sWheFlds = sWheFlds & "(" & sTmpStr & ")"
            End If
        End If
        If sWheFlds <> "" Then
            sSQL = sSQL & "WHERE " & sWheFlds & vbCrLf
        End If
        'WHERE指定追加SQL文
        For nRow = 3 To nMaxRow
            If .Cells(nRow, 8) <> "" And Left(Trim(.Cells(nRow, 8)), 2) <> "--" And .Rows(nRow).RowHeight > 0 Then
                If sAddWhes <> "" Then
                    sAddWhes = sAddWhes & vbCrLf & "  "
                End If
                sAddWhes = sAddWhes & .Cells(nRow, 8)
            End If
        Next
        If sAddWhes <> "" Then
            If sWheFlds <> "" Then
                sSQL = sSQL & "  AND " & sAddWhes & vbCrLf
            Else
                sSQL = sSQL & "WHERE " & sAddWhes & vbCrLf
            End If
        End If
        'ORDER BY指定フィールド
        sTmpStr = ""
        For nRow = 3 To nMaxRow
            If (.Cells(nRow, 6) = "昇順" Or .Cells(nRow, 6) = "降順") And .Rows(nRow).RowHeight > 0 Then
                If sTmpStr <> "" Or sOrdFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                End If
                If Len(sTmpStr) > ccMaxSqlWidth Then
                    sOrdFlds = sOrdFlds & "  " & sTmpStr & vbCrLf
                    sTmpStr = ""
                End If
                sTmpStr = sTmpStr & sFA & .Cells(nRow, 3) & _
                    IIf(.Cells(nRow, 6) = "降順", " DESC", "")
            End If
        Next
        If sTmpStr <> "" Then
            sOrdFlds = sOrdFlds & "  " & sTmpStr
        End If
        If sOrdFlds <> "" Then
            sSQL = sSQL & "ORDER BY" & vbCrLf
            sSQL = sSQL & sOrdFlds & vbCrLf
        End If
        If .Range("K2") <> "" Then
            If Not ST_WIN.ShowModal2(sSQL, "抽出") Then
                Exit Function
            End If
        End If
    End With
    argSQLType = "SELECT"
    GetSqlQpsSheet = sSQL
End Function

Function GetTemplateSql(argObjName, argTplMode, Optional argTplOptStr = "", Optional argTplRs)
    Dim sObjType As String, oRs, oFlds, oOpts
    Dim sSQL As String, sTmpStr As String, sSelFlds As String, sWheFlds As String, sOrdFlds As String
    Dim sSetFlds As String, sKeyFlds As String, sTmpSt2 As String
    Dim nMaxSqlWidth As Integer, sMaxSqlWidth As String
    sMaxSqlWidth = GetCfgVal("SQLの桁折り幅")
    If sMaxSqlWidth <> "" Then
        nMaxSqlWidth = sMaxSqlWidth
    Else
        nMaxSqlWidth = ccMaxSqlWidth
    End If
    sObjType = GetDBObjectType(argObjName)
    If argTplMode <> ccTPL_CLR Then
        Select Case sObjType
        Case "TABLE", "VIEW"
            If Not IsMissing(argTplRs) Then
                Set oRs = argTplRs
            Else
                Set oRs = Nothing
            End If
            If oRs Is Nothing Then
                Call GetDBObjectInfo(argObjName, sObjType, oRs, oFlds, oOpts)
                If Not IsMissing(argTplRs) Then
                    Set argTplRs = oRs
                End If
            End If
        End Select
    End If
    Select Case argTplMode
    Case ccTPL_SEL
        Select Case sObjType
        Case "TABLE", "VIEW"
            'SELECT指定フィールド
            If Not InStr(argTplOptStr, "noselect") <> 0 Then
                sSQL = sSQL & "SELECT" & vbCrLf
                sTmpStr = ""
                oRs.MoveFirst
                Do Until oRs.EOF
                    If sTmpStr <> "" Or sSelFlds <> "" Then
                        sTmpStr = sTmpStr & ", "
                    End If
                    If Len(sTmpStr) > nMaxSqlWidth Then
                        sSelFlds = sSelFlds & "  " & RTrim(sTmpStr) & vbCrLf
                        sTmpStr = ""
                    End If
                    sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value
                    oRs.MoveNext
                Loop
                If sTmpStr <> "" Then
                    sSelFlds = sSelFlds & "  " & sTmpStr
                End If
                sSQL = sSQL & sSelFlds & vbCrLf
                sSQL = sSQL & "FROM " & argObjName & vbCrLf
            Else
                sSQL = "SELECT * FROM " & argObjName & vbCrLf
            End If
            If sObjType = "TABLE" Then
                'WHERE指定フィールド
                If Not InStr(argTplOptStr, "nowhere") <> 0 Then
                    oRs.MoveFirst
                    Do Until oRs.EOF
                        If oRs("KEY").Value <> "" Then
    
                            If sWheFlds <> "" Then
                                sWheFlds = sWheFlds & vbCrLf & "  AND "
                            End If
                            
                            Select Case oRs("DATA_TYPE").Value
                            Case "VARCHAR2", "CHAR"
                                sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = ''"
                            Case "DATE" '☆ ADO対応要
'                                sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = TO_DATE('','YYYY/MM/DD HH24:MI:SS')"
                                sWheFlds = sWheFlds & "TO_CHAR(" & oRs("COLUMN_NAME").Value & ",'YYYY/MM/DD HH24:MI:SS') = ''"
                            Case Else
                                sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = 0"
                            End Select
    
                        End If
                        oRs.MoveNext
                    Loop
                    If sWheFlds <> "" Then
                        sSQL = sSQL & "WHERE " & sWheFlds & vbCrLf
                    End If
                End If
                'ORDER BY指定フィールド
                sTmpStr = ""
                oRs.MoveFirst
                Do Until oRs.EOF
                    If oRs("KEY").Value <> "" Then
                        If sTmpStr <> "" Or sOrdFlds <> "" Then
                            sTmpStr = sTmpStr & ", "
                        End If
                        If Len(sTmpStr) > nMaxSqlWidth Then
                            sOrdFlds = sOrdFlds & "  " & RTrim(sTmpStr) & vbCrLf
                            sTmpStr = ""
                        End If
                        sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value
                    End If
                    oRs.MoveNext
                Loop
                If sTmpStr <> "" Then
                    sOrdFlds = sOrdFlds & "  " & sTmpStr
                End If
                If sOrdFlds <> "" Then
                    sSQL = sSQL & "ORDER BY" & vbCrLf
                    sSQL = sSQL & sOrdFlds & vbCrLf
                End If
            End If
        End Select
    Case ccTPL_INS
        Select Case sObjType
        Case "TABLE"
            sSQL = sSQL & "INSERT INTO " & argObjName & " (" & vbCrLf
            sTmpStr = ""
            oRs.MoveFirst
            Do Until oRs.EOF
                If sTmpStr <> "" Or sSelFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                    sTmpSt2 = sTmpSt2 & ", "
                End If
                If Len(sTmpStr) > nMaxSqlWidth Or Len(sTmpSt2) > nMaxSqlWidth Then
                    sSelFlds = sSelFlds & "  " & RTrim(sTmpStr) & vbCrLf
                    sSetFlds = sSetFlds & "  " & RTrim(sTmpSt2) & vbCrLf
                    sTmpStr = ""
                    sTmpSt2 = ""
                End If
                sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value
                Select Case oRs("DATA_TYPE").Value
                Case "VARCHAR2", "CHAR"
                    sTmpSt2 = sTmpSt2 & "''"
                Case "DATE" '☆ ADO対応要
                    sTmpSt2 = sTmpSt2 & "TO_DATE('','YYYY/MM/DD HH24:MI:SS')"
                Case Else
                    sTmpSt2 = sTmpSt2 & "0"
                End Select
                oRs.MoveNext
            Loop
            If sTmpStr <> "" Then
                sSelFlds = sSelFlds & "  " & sTmpStr
                sSetFlds = sSetFlds & "  " & sTmpSt2
            End If
            sSQL = sSQL & sSelFlds & vbCrLf
            sSQL = sSQL & ") VALUES (" & vbCrLf
            sSQL = sSQL & sSetFlds & vbCrLf
            sSQL = sSQL & ")"
        End Select
    Case ccTPL_UPD
        Select Case sObjType
        Case "TABLE"
            sSQL = sSQL & "UPDATE " & argObjName & " SET" & vbCrLf
            sTmpStr = ""
            oRs.MoveFirst
            Do Until oRs.EOF
                If sTmpStr <> "" Or sSetFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                End If
                If Len(sTmpStr) > nMaxSqlWidth Then
                    sSetFlds = sSetFlds & "  " & RTrim(sTmpStr) & vbCrLf
                    sTmpStr = ""
                End If
                Select Case oRs("DATA_TYPE").Value
                Case "VARCHAR2", "CHAR"
                    sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value & " = ''"
                Case "DATE" '☆ ADO対応要
                    sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value & " = TO_DATE('','YYYY/MM/DD HH24:MI:SS')"
                Case Else
                    sTmpStr = sTmpStr & oRs("COLUMN_NAME").Value & " = 0"
                End Select
                oRs.MoveNext
            Loop
            If sTmpStr <> "" Then
                sSetFlds = sSetFlds & "  " & sTmpStr
            End If
            sSQL = sSQL & sSetFlds & vbCrLf
            oRs.MoveFirst
            Do Until oRs.EOF
                If oRs("KEY").Value <> "" Then

                    If sWheFlds <> "" Then
                        sWheFlds = sWheFlds & vbCrLf & "  AND "
                    End If
                    
                    Select Case oRs("DATA_TYPE").Value
                    Case "VARCHAR2", "CHAR"
                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = ''"
                    Case "DATE" '☆ ADO対応要
'                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = TO_DATE('','YYYY/MM/DD HH24:MI:SS')"
                        sWheFlds = sWheFlds & "TO_CHAR(" & oRs("COLUMN_NAME").Value & ",'YYYY/MM/DD HH24:MI:SS') = ''"
                    Case Else
                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = 0"
                    End Select

                End If
                oRs.MoveNext
            Loop
            If sWheFlds <> "" Then
                sSQL = sSQL & "WHERE " & sWheFlds & vbCrLf
            End If
        End Select
    Case ccTPL_DEL
        Select Case sObjType
        Case "TABLE"
            sSQL = sSQL & "DELETE FROM " & argObjName & vbCrLf
            oRs.MoveFirst
            Do Until oRs.EOF
                If oRs("KEY").Value <> "" Then

                    If sWheFlds <> "" Then
                        sWheFlds = sWheFlds & vbCrLf & "  AND "
                    End If
                    
                    Select Case oRs("DATA_TYPE").Value
                    Case "VARCHAR2", "CHAR"
                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = ''"
                    Case "DATE" '☆ ADO対応要
'                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = TO_DATE('','YYYY/MM/DD HH24:MI:SS')"
                        sWheFlds = sWheFlds & "TO_CHAR(" & oRs("COLUMN_NAME").Value & ",'YYYY/MM/DD HH24:MI:SS') = ''"
                    Case Else
                        sWheFlds = sWheFlds & oRs("COLUMN_NAME").Value & " = 0"
                    End Select

                End If
                oRs.MoveNext
            Loop
            If sWheFlds <> "" Then
                sSQL = sSQL & "WHERE " & sWheFlds & vbCrLf
            End If
        End Select
    Case ccTPL_CRE
        Select Case sObjType
        Case "TABLE"
            nMaxSqlWidth = 0
            sSQL = sSQL & "CREATE TABLE " & argObjName & " (" & vbCrLf
            sTmpStr = ""
            Dim sNameVal As String, sTypeVal As String
            Dim nNameMax As Integer, nTypeMax As Integer
            oRs.MoveFirst
            Do Until oRs.EOF
                sNameVal = oRs("COLUMN_NAME").Value
                sTypeVal = oRs("DATA_TYPE").Value
                If oRs("DATA_TYPE").Value <> "DATE" Then
                    sTypeVal = sTypeVal & "(" & oRs("WIDTH").Value
                    If oRs("SCALE").Value <> "" Then
                        sTypeVal = sTypeVal & "," & oRs("SCALE").Value
                    End If
                    sTypeVal = sTypeVal & ") "
                End If
                If Len(sNameVal) > nNameMax Then
                    nNameMax = Len(sNameVal)
                End If
                If Len(sTypeVal) > nTypeMax Then
                    nTypeMax = Len(sTypeVal)
                End If
                oRs.MoveNext
            Loop
            oRs.MoveFirst
            Do Until oRs.EOF
                'テーブル項目の取得
                If sTmpStr <> "" Or sSelFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                End If
                If Len(sTmpStr) > nMaxSqlWidth Then
                    sSelFlds = sSelFlds & "  " & RTrim(sTmpStr) & vbCrLf
                    sTmpStr = ""
                End If
                sNameVal = oRs("COLUMN_NAME").Value
                sNameVal = Left(sNameVal & Space(nNameMax), nNameMax)
                sTmpStr = sTmpStr & sNameVal & " "
                sTypeVal = oRs("DATA_TYPE").Value
                If oRs("DATA_TYPE").Value <> "DATE" Then
                    sTypeVal = sTypeVal & "(" & oRs("WIDTH").Value
                    If oRs("SCALE").Value <> "" Then
                        sTypeVal = sTypeVal & "," & oRs("SCALE").Value
                    End If
                    sTypeVal = sTypeVal & ") "
                End If
                sTypeVal = Left(sTypeVal & Space(nTypeMax), nTypeMax)
                sTmpStr = sTmpStr & sTypeVal & " "
                If oRs("NOTNULL").Value <> "" Then
                    sTmpStr = sTmpStr & "NOT NULL"
                Else
                    sTmpStr = sTmpStr & "NULL"
                End If
'                If oRs("KEY").Value <> "" Then
'                    sTmpStr = sTmpStr & "/* KEY */ "
'                End If
                
                'キー項目の取得
                If oRs("KEY").Value <> "" Then
                    If sTmpSt2 <> "" Or sKeyFlds <> "" Then
                        sTmpSt2 = sTmpSt2 & ", "
                    End If
                    If Len(sTmpSt2) > nMaxSqlWidth Then
                        sKeyFlds = sKeyFlds & "    " & RTrim(sTmpSt2) & vbCrLf
                        sTmpSt2 = ""
                    End If
                    sTmpSt2 = sTmpSt2 & oRs("COLUMN_NAME").Value
                End If
                oRs.MoveNext
            Loop
            If sTmpStr <> "" Then
                sSelFlds = sSelFlds & "  " & sTmpStr
            End If
            sSQL = sSQL & sSelFlds
            If sTmpSt2 <> "" Then
                sKeyFlds = sKeyFlds & "    " & sTmpSt2
            End If
            If sKeyFlds <> "" Then
                sSQL = sSQL & "," & vbCrLf & _
                    "  CONSTRAINT " & ccPriKeyPref & argObjName & ccPriKeySuff & " PRIMARY KEY ("
                sSQL = sSQL & vbCrLf & sKeyFlds
                sSQL = sSQL & vbCrLf & "  )"
            End If
            sSQL = sSQL & vbCrLf & ")"
        End Select
    Case ccTPL_CLR
        sSQL = vbCrLf
    End Select
    GetTemplateSql = sSQL
End Function

'ＳＱＬの実行
Private Function ExecSql(argConStr, argSql, Optional argRs, Optional argFlds, Optional argOpts)
    Dim oCon, oFld, oStmt, oPrm
    Dim nST, nET
    Dim nRecCnt
    Dim i As Integer
    Dim bKeyBreak As Boolean
    
    On Error Resume Next
'    Application.EnableCancelKey = xlErrorHandler
    
'    Debug.Print Replace(Replace(argSql, vbCr, "[CR]"), vbLf, "[LF]" & vbCrLf) & vbCrLf & String(100, "-")
    Debug.Print argSql & vbCrLf & String(100, "-")
    
    If GetKeyState(VK_SCROLL) <> 0 Then
        If Not ST_WIN.ShowModal2(argSql, "実行") Then
            GoTo END_FUNC
        End If
    End If
    
    Select Case GetCfgVal("接続モード")
    Case "1" 'OO4O/SQL*Net
    
        Set oCon = GetConObj(argConStr, IsMissing(argRs))
        
        If oCon Is Nothing Then
          Call MsgBoxST("接続に失敗しました" & vbCrLf & vbCrLf & argConStr)
          GoTo END_FUNC
        End If
        
        'パラメータの設定
        If Not IsMissing(argOpts) Then
        If Not IsEmpty(argOpts) Then
        If TypeName(argOpts.Item("パラメータ")) = "Prop" Then
            Dim sPrmName As String, sIOType As String, nIoType As Integer, nSvType As Integer
            For Each oPrm In argOpts.Item("パラメータ").Items
                sPrmName = oPrm.Item("Name")
                sIOType = oPrm.Item("InOut")
                If InStr(sIOType, "IN") <> 0 And InStr(sIOType, "OUT") <> 0 Then
                    nIoType = 3 'ORAPARM_BOTH
                ElseIf InStr(sIOType, "OUT") <> 0 Then
                    nIoType = 2 'ORAPARM_OUTPUT
                Else
                    nIoType = 1 'ORAPARM_INPUT
                End If
                Select Case oPrm.Item("Type")
                Case "VARCHAR2"
                    nSvType = 1 'ORATYPE_VARCHAR2
                Case "NUMBER"
                    nSvType = 2 'ORATYPE_NUMBER
                Case "CHAR"
                    nSvType = 96 'ORATYPE_CHAR
                Case Else
                    Call MsgBoxST(sPrmName & " 未対応の入力パラメータです。")
                    Exit Function
                End Select
                oCon.Parameters.Add sPrmName, oPrm.Item("Value"), nIoType, nSvType
            Next
            If argOpts.Item("戻り値") <> "" Then
                Select Case argOpts.Item("戻り値")
                Case "VARCHAR2"
                    nSvType = 1 'ORATYPE_VARCHAR2
                Case "NUMBER"
                    nSvType = 2 'ORATYPE_NUMBER
                Case "CHAR"
                    nSvType = 96 'ORATYPE_CHAR
                Case Else
                    Call MsgBoxST(" 未対応の戻り値です。")
                    Exit Function
                End Select
                oCon.Parameters.Add "RESULT", "", 2, nSvType
            End If
        End If
        End If
        End If
        
        nST = Timer()

        '同期モードでのクエリ発行
        If Not IsMissing(argRs) Then
            'レコードセットの取得
            Set argRs = oCon.CreateDynaset(argSql, &H2 + &H4)
            'ORADYN_NO_BLANKSTRIP, ORADYN_READONLY
        Else
            'DML/DDLの実行
'            nRecCnt = oCon.ExecuteSQL(Replace(argSql, vbCr, "")) '☆
            nRecCnt = oCon.ExecuteSQL(argSql)
        End If
        
'    On Error GoTo KEYBREAK
'
'        '非同期モードでのクエリ発行
'        If Not IsMissing(argRs) Then
'            'レコードセットの取得
'            oCon.Parameters.Add "C", 0, 2, 102    'ORAPARM_OUTPUT, ORATYPE_CURSOR
'            oCon.Parameters.Add "S", argSql, 1, 1 'ORAPARM_INPUT , ORATYPE_VARCHAR2
'            Set oStmt = oCon.CreateSQL("DECLARE BEGIN OPEN :C FOR :S; END;", &H4) 'ORASQL_NONBLK
'        Else
'            'DML/DDLの実行
''            nRecCnt = oCon.ExecuteSQL(argSql)
'            Set oStmt = oCon.CreateSQL(argSql, &H4) 'ORASQL_NONBLK
'        End If
'        '非同期実行したクエリの終了待ちループ
'        bKeyBreak = False
'        Do While oStmt.NonBlockingState = -3123 'ORASQL_STILL_EXECUTING
'            '中断キー押下を検知
'            If bKeyBreak Then
'                If MsgBoxST("中断しますか？", vbYesNo) = vbYes Then
'                    oStmt.Cancel
'                    Exit Do
'                End If
'                bKeyBreak = False
'            End If
'        Loop
'        '実行完了後処理
'        If Not IsMissing(argRs) Then
'            'レコードセットの場合はパラメータのカーソル変数を取得
'            Set argRs = oCon.Parameters("C").Value
'            oCon.Parameters.Remove "C"
'            oCon.Parameters.Remove "S"
'        Else
'            'DML/DDLの場合は影響レコード件数を取得
'            nRecCnt = oStmt.RecordCount
'        End If
        
        nET = Timer()
        
        'パラメータの結果取得、削除
        If Not IsMissing(argOpts) Then
        If Not IsEmpty(argOpts) Then
        If TypeName(argOpts.Item("パラメータ")) = "Prop" Then
            For Each oPrm In argOpts.Item("パラメータ").Items
                If InStr(oPrm.Item("InOut"), "OUT") <> 0 Then
                    oPrm.Item("Value") = oCon.Parameters(oPrm.Item("Name")).Value
                End If
                oCon.Parameters.Remove oPrm.Item("Name")
            Next
            If argOpts.Item("戻り値") <> "" Then
                argOpts.Item("戻り値") = oCon.Parameters("RESULT").Value
                oCon.Parameters.Remove "RESULT"
            End If
        End If
        End If
        End If
        
        If (Err <> 0 Or oCon.LastServerErr <> 0) And oCon.LastServerErr <> 24344 Then '☆
          Call MsgBoxST(Err.Number & ":" & Err.Description & vbCrLf & oCon.LastServerErrText)
          GoTo END_FUNC
        End If
        
    Case "2" 'ADO/ODBC
        
        Set oCon = GetConObj(argConStr, IsMissing(argRs))
        
        If oCon.State <> 1 Then 'adStateOpen
          Call MsgBoxST("接続に失敗しました" & vbCrLf & vbCrLf & argConStr)
          GoTo END_FUNC
        End If
        
        nST = Timer()
        If Not IsMissing(argRs) Then
            Set argRs = oCon.Execute(argSql)
        Else
            'DML/DDLの実行
            Call oCon.Execute(argSql, nRecCnt) 'adCmdTextを指定した方がいいかも？
        End If
        nET = Timer()
        
        If Err <> 0 Then
          Call MsgBoxST(Err.Number & ":" & Err.Description)
          GoTo END_FUNC
        End If
    
    End Select
    
    Application.EnableCancelKey = xlDisabled
    If Not IsMissing(argFlds) Then
        'フィールド情報のセット
        Dim oFlds As New Prop
        With argRs.Fields
            For i = 0 To .Count - 1
                Set oFld = New Prop
                oFld.Item("Title") = .Item(i).Name
                oFld.Item("Name") = .Item(i).Name
                oFld.Item("Type") = GetFldType(argRs, i)
                
                Set oFlds.Item() = oFld
            Next
            Set argFlds = oFlds
        End With
    End If
    
    If Not IsMissing(argOpts) Then
        'オプション情報のセット
        If IsEmpty(argOpts) Then
            Set argOpts = New Prop
        End If
        argOpts.Item("実行") = TimeElapsed(nST, nET)
        argOpts.Item("更新件数") = nRecCnt
    End If
    
    ExecSql = True
    GoTo END_FUNC
KEYBREAK:
    bKeyBreak = True
    Resume
END_FUNC:
    Application.EnableCancelKey = xlInterrupt
'    Application.StatusBar = False
End Function

'ExSQell暫定ロジック
'シェルコマンド実行
Private Function ExecShell(argCmd, argRs, argFlds)
    Dim sCmd, sWrk, WRK, sLine, i, oFld, sTmp, sShTmp, TMP, sCurDir
    
    If IsEmpty(coWSH) Then
        Set coWSH = CreateObject("WScript.Shell")
    End If
    
    If IsEmpty(coFSO) Then
        Set coFSO = CreateObject("Scripting.FileSystemObject")
    End If
    
    sWrk = ActiveWorkbook.Path & "\ExSQell.wrk"
    
    '初回とシェルコマンド変更時にunameを実行して実行環境を判別
    If csShCmd <> GetShCmd Then
        csShCmd = GetShCmd
        
        sCmd = csShCmd & " uname"
        Call ExecCmd(sCmd, sWrk)
        
        csUname = "BoW"
        
        If coFSO.FileExists(sWrk) Then
            Set WRK = coFSO.OpenTextFile(sWrk, 1, False)
            If Not WRK Is Nothing Then
                If Not WRK.AtEndOfStream Then
                    csUname = WRK.ReadLine
                End If
            End If
            WRK.Close
            Call coFSO.DeleteFile(sWrk)
        End If
    End If
    
    '実行環境に応じて、リダイレクトするテンポラリファイルパスを算出
    sTmp = coFSO.BuildPath(coFSO.GetSpecialFolder(2), coFSO.GetTempName)
    sShTmp = LCase(Left(sTmp, 1)) & "/" & Replace(Mid(sTmp, 4), "\", "/")
    If InStr(csUname, "CYGWIN") > 0 Then
        sShTmp = "/cygdrive/" & sShTmp
    ElseIf InStr(csUname, "MINGW") > 0 Then
        sShTmp = "/" & sShTmp
    Else 'Bash on Ubuntu on Windows
        sShTmp = "/mnt/" & sShTmp
    End If
    
    'カレントディレクトリ移動コマンド設定
    sCurDir = GetCurDir
    If sCurDir <> "" Then
        sCurDir = "cd " & sCurDir & "; "
    End If
    
    'シェルコマンドの実行
    sCmd = csShCmd & " """ & sCurDir & Replace(argCmd, """", """""") & " > " & sShTmp & """"
    Call coWSH.Run(sCmd, vbHide, True)
    
    'ローカルレコードセット作成
    Set argRs = CreateObject("ADODB.Recordset")
    
    
    '各種設定情報の取得
    Dim nOutCnv, sCharset, nLineSeparator, nFirstRowTitle, nFieldSep, sFieldName, arFields, nFieldCnt, nArrayCnt
    nOutCnv = GetCfgVal("出力文字コード変換")
    nFirstRowTitle = GetCfgVal("1行目をタイトル行")
    nFieldSep = GetCfgVal("フィールド区切り")
    
    'テンポラリファイルの読み込み方式に応じて前準備
    Select Case nOutCnv
    Case 2 To 5 '「UTF-8 → SJIS」～「SJIS  → SJIS」
        Select Case nOutCnv
        Case 2 'UTF-8 → SJIS
            sCharset = "utf-8"
            nLineSeparator = 10 'adLF
        Case 3 'EUC   → SJIS
            sCharset = "euc-jp"
            nLineSeparator = 10 'adLF
        Case 4 'JIS   → SJIS
            sCharset = "iso-2022-jp"
            nLineSeparator = 10 'adLF
        Case 5 'SJIS  → SJIS
            sCharset = "shift_jis"
            nLineSeparator = -1 'adCRLF
        End Select
        
        'テンポラリファイルを読み込んでローカルレコードセットに追加
        Set TMP = CreateObject("ADODB.Stream")
        With TMP
            .Charset = sCharset
            .LineSeparator = nLineSeparator
            .Open
            Call .LoadFromFile(sTmp)
            
            '1行目の読み込み
            If Not .EOS Then
                sLine = .Readtext(-2) 'adReadLine
                'Debug.Print sLine
                
                'フィールド区切り
                If nFieldSep = 1 Then 'なし
                    '1行目をタイトル行
                    If nFirstRowTitle = 1 Then 'としない
                        sFieldName = "Line"
                    Else
                        sFieldName = sLine
                    End If
                    argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, 最大桁数
                    nFieldCnt = 1
                Else
                    '行をフィールドに分割
                    Select Case nFieldSep
                    Case 2 '連続した空白文字
                        arFields = SplitSh(sLine)
                    Case 3 'カンマ
                        arFields = SplitCsv(sLine)
                    Case 4 'タブ文字
                        arFields = Split(sLine, " ")
                    End Select
                    
                    'フィールド数に応じてローカルレコードセットにフィールド追加
                    For i = 0 To UBound(arFields)
                        '1行目をタイトル行
                        If nFirstRowTitle = 1 Then 'としない
                            sFieldName = "Filed" & (i + 1)
                        Else
                            sFieldName = arFields(i)
                        End If
                        argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, 最大桁数
                    Next
                    nFieldCnt = argRs.Fields.Count
                End If
                
                'ローカルレコードセットをオープン
                argRs.Open
                
                '1行目をタイトル行としない場合
                If nFirstRowTitle = 1 Then
                    '1行目のレコード追加
                    argRs.AddNew
                    
                    'フィールド区切り
                    If nFieldSep = 1 Then 'なし
                        argRs(0) = sLine
                    Else
                        '列数に応じてフィールド値セット
                        For i = 0 To nFieldCnt - 1
                            argRs(i) = arFields(i)
                        Next
                    End If
                    
                    '1行目のレコード登録
                    argRs.Update
                End If
               
            Else '結果データレコード無し
                'フィールド区切り
                If nFieldSep = 1 Then 'なし
                    argRs.Fields.Append "Line", 202 'adVarWChar
                Else
                    argRs.Fields.Append "Field1", 202 'adVarWChar
                End If
                argRs.Open
            End If
            
            '2行目以降の読み込み
            Do Until .EOS
                sLine = .Readtext(-2) 'adReadLine
                'Debug.Print sLine
                
                argRs.AddNew
                
                'フィールド区切り
                If nFieldSep = 1 Then 'なし
                    argRs(0) = sLine
                Else
                    '行をフィールドに分割
                    Select Case nFieldSep
                    Case 2 '連続した空白文字
                        arFields = SplitSh(sLine)
                    Case 3 'カンマ
                        arFields = SplitCsv(sLine)
                    Case 4 'タブ文字
                        arFields = Split(sLine, " ")
                    End Select
                    
                    nArrayCnt = UBound(arFields) + 1
                    
                    '列数に応じてフィールド値セット
                    For i = 0 To nFieldCnt - 1
                        If i <= nArrayCnt - 1 Then
                            argRs(i) = arFields(i)
                        End If
                    Next
                End If
                
                argRs.Update
            Loop
            .Close
        End With
    Case Else '1 「何も変換しない」
        'テンポラリファイルを読み込んでローカルレコードセットに追加
        Set TMP = coFSO.OpenTextFile(sTmp, 1, False)

        '1行目の読み込み
        If Not TMP.AtEndOfStream Then
            sLine = TMP.ReadLine
            'Debug.Print sLine
            
            'フィールド区切り
            If nFieldSep = 1 Then 'なし
                '1行目をタイトル行
                If nFirstRowTitle = 1 Then 'としない
                    sFieldName = "Line"
                Else
                    sFieldName = sLine
                End If
                argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, 最大桁数
                nFieldCnt = 1
            Else
                '行をフィールドに分割
                Select Case nFieldSep
                Case 2 '連続した空白文字
                    arFields = SplitSh(sLine)
                Case 3 'カンマ
                    arFields = SplitCsv(sLine)
                Case 4 'タブ文字
                    arFields = Split(sLine, " ")
                End Select
                
                'フィールド数に応じてローカルレコードセットにフィールド追加
                For i = 0 To UBound(arFields)
                    '1行目をタイトル行
                    If nFirstRowTitle = 1 Then 'としない
                        sFieldName = "Filed" & (i + 1)
                    Else
                        sFieldName = arFields(i)
                    End If
                    argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, 最大桁数
                Next
                nFieldCnt = argRs.Fields.Count
            End If
            
            'ローカルレコードセットをオープン
            argRs.Open
            
            '1行目をタイトル行としない場合
            If nFirstRowTitle = 1 Then
                '1行目のレコード追加
                argRs.AddNew
                
                'フィールド区切り
                If nFieldSep = 1 Then 'なし
                    argRs(0) = sLine
                Else
                    '列数に応じてフィールド値セット
                    For i = 0 To nFieldCnt - 1
                        argRs(i) = arFields(i)
                    Next
                End If
                
                '1行目のレコード登録
                argRs.Update
            End If
           
        Else '結果データレコード無し
            'フィールド区切り
            If nFieldSep = 1 Then 'なし
                argRs.Fields.Append "Line", 202 'adVarWChar
            Else
                argRs.Fields.Append "Field1", 202 'adVarWChar
            End If
            argRs.Open
        End If
        
        '2行目以降の読み込み
        Do Until TMP.AtEndOfStream
            sLine = TMP.ReadLine
            'Debug.Print sLine
            
            argRs.AddNew
            
            'フィールド区切り
            If nFieldSep = 1 Then 'なし
                argRs(0) = sLine
            Else
                '行をフィールドに分割
                Select Case nFieldSep
                Case 2 '連続した空白文字
                    arFields = SplitSh(sLine)
                Case 3 'カンマ
                    arFields = SplitCsv(sLine)
                Case 4 'タブ文字
                    arFields = Split(sLine, " ")
                End Select
                
                nArrayCnt = UBound(arFields) + 1
                
                '列数に応じてフィールド値セット
                For i = 0 To nFieldCnt - 1
                    If i <= nArrayCnt - 1 Then
                        argRs(i) = arFields(i)
                    End If
                Next
            End If
            
            argRs.Update
        Loop
        TMP.Close
    End Select
    
    argRs.MoveFirst
    
    'フィールド情報の設定
    Dim oFlds As New Prop
    With argRs.Fields
        For i = 0 To .Count - 1
            Set oFld = New Prop
            oFld.Item("Title") = .Item(i).Name
            oFld.Item("Name") = .Item(i).Name
            oFld.Item("Type") = GetFldType(argRs, i)
            Set oFlds.Item() = oFld
        Next
        Set argFlds = oFlds
    End With
    
    Call coFSO.DeleteFile(sTmp)
    
    ExecShell = True
End Function

'コマンド実行 ※vbs経由で非表示同期実行、標準出力をワークファイルへ書き出し
Private Function ExecCmd(argCmd, argWrk)
    Dim sVbs, oVbs, sRun

    With coFSO
        sVbs = .GetParentFolderName(argWrk) & "\" & .GetBaseName(argWrk) & ".vbs"
        Set oVbs = .OpenTextFile(sVbs, 2, True)
    End With
    
    With oVbs
      .WriteLine "Dim WSH, FSO, WRK"
      .WriteLine "Set WSH = WScript.CreateObject(""WScript.Shell"")"
      .WriteLine "Set FSO = WScript.CreateObject(""Scripting.FileSystemObject"")"
'      .WriteLine "WSH.CurrentDirectory = ""C:\cygwin64\bin"""
      .WriteLine "Set PRC = WSH.Exec(""" & Replace(argCmd, """", """""") & """)"
      .WriteLine "Set WRK = FSO.OpenTextFile(""" & argWrk & """,2,True)"
      .WriteLine "Do Until PRC.StdOut.AtEndOfStream"
      .WriteLine "  WRK.WriteLine PRC.StdOut.ReadLine"
      .WriteLine "Loop"
      .WriteLine "WRK.Close"
'      .WriteLine "WScript.Echo """ & Replace(argCmd, """", """""") & """"
      .Close
    End With
    sRun = "CSCRIPT //NOLOGO """ & sVbs & """"
    ExecCmd = coWSH.Run(sRun, vbHide, True) 'Run(sRunCmd, [intWindowStyle], [bWaitOnReturn])
    Call coFSO.DeleteFile(sVbs)
End Function

'クエリの実行
Private Function ExecQuery(argConStr, argSql, argSQLType, argRs, argFlds, argOpts)
    Dim sCntSql As String, oCntRs
    Dim sNoOrdSql As String, i As Integer, nLastOrd As Integer, nSCnt As Integer, nECnt As Integer
    If GetCfgVal("警告抽出件数") <> "" Then
        'ALT同時押し以外、テーブル抽出、メンテモードの場合は抽出件数を確認
        If Not GetKeyState(VK_MENU) < 0 Or argSQLType = "TABLE" Or argSQLType = "MAINTEN" Then
            '抽出件数を取得
            sCntSql = GetLibSql("件数取得")
            '実行SQL文にORDER BYが含まれている場合は以降を削除
            sNoOrdSql = UCase(argSql)
            i = 1
            Do
                i = InStr(i, sNoOrdSql, "ORDER ")
                If i = 0 Then Exit Do
                nLastOrd = i
                i = i + 6
            Loop
            If nLastOrd <> 0 Then
                i = nLastOrd
                Do
                    i = InStr(i, sNoOrdSql, "(")
                    If i = 0 Then Exit Do
                    nSCnt = nSCnt + 1
                    i = i + 1
                Loop
                i = nLastOrd
                Do
                    i = InStr(i, sNoOrdSql, ")")
                    If i = 0 Then Exit Do
                    nECnt = nECnt + 1
                    i = i + 1
                Loop
                '対応づいていない括弧が無い場合
                If nSCnt < nECnt Then
                    sNoOrdSql = argSql
                Else
                    '無い場合はORDER BY以降を削除
                    sNoOrdSql = Left(argSql, nLastOrd - 1)
                End If
            Else
                sNoOrdSql = argSql
            End If
            sCntSql = SetSqlParam(sCntSql, "取得対象", "(" & sNoOrdSql & ")")
            Application.StatusBar = "件数確認中..."
            If Not ExecSql(argConStr, sCntSql, oCntRs, argFlds, argOpts) Then
                Application.StatusBar = ""
                Exit Function
            End If
            Application.StatusBar = ""
            argOpts.Item("件数") = oCntRs("CNT").Value
            'SHIFT同時押しの場合は抽出件数ダイアログ表示
            If GetKeyState(VK_SHIFT) < 0 And oCntRs("CNT").Value > 0 Then
                Select Case MsgBoxST("抽出件数は " & oCntRs("CNT").Value & " 件です。" & vbCrLf & vbCrLf & _
                    "[全件表示↓]　　[サンプル数まで↓]", vbYesNoCancel + vbDefaultButton3)
                Case vbNo
                    argOpts.Item("サンプル抽出件数") = GetCfgVal("サンプル抽出件数")
                    argOpts.Item("件数") = argOpts.Item("サンプル抽出件数")
                Case vbCancel
                    Exit Function
                End Select
            End If
            '警告抽出件数以上ならワーニング
            If oCntRs("CNT").Value > CLng(GetCfgVal("警告抽出件数")) And argOpts.Item("サンプル抽出件数") = "" Then
                Select Case MsgBoxST("抽出件数が " & oCntRs("CNT").Value & " 件、警告件数以上です。" & vbCrLf & vbCrLf & _
                    "[全件表示↓]　　[サンプル数まで↓]", vbYesNoCancel + vbDefaultButton3)
                Case vbNo
                    argOpts.Item("サンプル抽出件数") = GetCfgVal("サンプル抽出件数")
                    argOpts.Item("件数") = argOpts.Item("サンプル抽出件数")
                Case vbCancel
                    Exit Function
                End Select
            End If
        End If
    End If
        
        
    'DB接続、SQL実行
    Application.StatusBar = "SQL実行中..."
    If Not ExecSql(argConStr, argSql, argRs, argFlds, argOpts) Then
        Application.StatusBar = ""
        Exit Function
    End If
    Application.StatusBar = ""

    If argSQLType <> "MAINTEN" Then
        If argRs.EOF Then
            Call MsgBoxST("抽出件数が０件です")
            Exit Function
        End If
    End If

    ExecQuery = True
End Function

'新規シート作成
Private Function GetNewSheet(argShtName, Optional argRenFlg = True, Optional argNewFlg = False, Optional argTplSht = Nothing) As Worksheet
    Dim n As Integer
    Dim sChkShtName As String
    Dim sNewShtName As String
    Dim sCurShtName As String
    Dim nCurZoom
    
    Set coPrevSheet = ActiveSheet
    nCurZoom = ActiveWindow.Zoom
    
    sChkShtName = argShtName
    
    If argRenFlg Then
        '同名のシートがあった場合は、指定シート名n にリネーム
        n = 2
        sNewShtName = ""
        Do
            On Error Resume Next
            sNewShtName = Sheets(sChkShtName).Name
            On Error GoTo 0
            If sNewShtName = "" Then
                sNewShtName = sChkShtName
            Else
                sNewShtName = ""
'                If InStr("0123456789]", Right(argShtName, 1)) = 0 Then
                    sChkShtName = argShtName & n
'                Else
'                    sChkShtName = argShtName & "[" & n & "]"
'                End If
                n = n + 1
            End If
        Loop Until sNewShtName <> ""
        
        sChkShtName = sNewShtName
    End If
    
    '指定シート名の存在チェック
    sNewShtName = ""
    On Error Resume Next
    sNewShtName = Sheets(sChkShtName).Name
    On Error GoTo 0
    If sNewShtName <> "" Then
        '同名で存在する場合
        If argNewFlg Then
            Dim bFlg
            bFlg = Application.DisplayAlerts
            Application.DisplayAlerts = False
            Sheets(sChkShtName).Delete
            Application.DisplayAlerts = bFlg
            
            '新規シートの作成
'            sCurShtName = ActiveSheet.Name
'            Select Case GetCfgVal("新規シート追加位置")
'            Case 1
'                Worksheets.Add before:=Worksheets(sCurShtName)
'            Case 2
'                Worksheets.Add after:=Worksheets(sCurShtName)
'            Case Else
'                Worksheets.Add after:=Worksheets(Worksheets.Count)
'            End Select
'
'            Set GetNewSheet = ActiveSheet
'            GetNewSheet.Name = sChkShtName
            Set GetNewSheet = MakeSheet(sChkShtName, ActiveSheet, argTplSht)
        Else
            Set GetNewSheet = Sheets(sChkShtName)
            If GetNewSheet.Visible = False Then
                GetNewSheet.Visible = True
            End If
            GetNewSheet.Select
        End If
    Else
        '同名で存在しない場合、単純に新規追加
'        sCurShtName = ActiveSheet.Name
'        Select Case GetCfgVal("新規シート追加位置")
'        Case 1
'            Worksheets.Add before:=Worksheets(sCurShtName)
'        Case 2
'            Worksheets.Add after:=Worksheets(sCurShtName)
'        Case Else
'            Worksheets.Add after:=Worksheets(Worksheets.Count)
'        End Select
'
'        Set GetNewSheet = ActiveSheet
'        GetNewSheet.Name = sChkShtName
        Set GetNewSheet = MakeSheet(sChkShtName, ActiveSheet, argTplSht)
    End If
        
    ActiveWindow.Zoom = nCurZoom

End Function

Private Function MakeSheet(argNewShtName, argCurSht, argTplSht) As Worksheet
    If argTplSht Is Nothing Then
        Select Case GetCfgVal("新規シート追加位置")
        Case 1
            Worksheets.Add before:=argCurSht
        Case 2
            Worksheets.Add after:=argCurSht
        Case Else
            Worksheets.Add after:=Worksheets(Worksheets.Count)
        End Select
    Else
        Select Case GetCfgVal("新規シート追加位置")
        Case 1
            argTplSht.Copy before:=argCurSht
        Case 2
            argTplSht.Copy after:=argCurSht
        Case Else
            argTplSht.Copy after:=Worksheets(Worksheets.Count)
        End Select
    End If
    ActiveSheet.Name = argNewShtName
    Set MakeSheet = ActiveSheet
End Function

'クエリ実行結果を表組みして貼り付け
Private Function MakeList(argStart, argRs, argFlds, argOpts)
    Dim oSht As Worksheet
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nCol As Long, nMinCol As Long, nMaxCol As Long
    Dim nFldCnt As Integer
    Dim vDtFlds() As String
    Dim sNmFlds() As String
    Dim sSpFlds() As String
    Dim nST, nET
    Dim nFldIdx As Integer
    Dim oFld
    Dim nCnt As Long
    Dim nMaxCnt As Long
    Dim bKeyBreak As Boolean, b1LScrUpd As Boolean, bSimaBgCol As Boolean
    Dim nRsIdx As Integer
    Dim bPrevScrUpd As Boolean
    
    Application.EnableCancelKey = xlDisabled
        
    If TypeName(argStart) = "Worksheet" Then
        Set oSht = argStart
        nMinRow = 1
        nMinCol = 1
    ElseIf TypeName(argStart) = "Range" Then
        Set oSht = argStart.Worksheet
        nMinRow = argStart.Row
        nMinCol = argStart.Column
    Else
        MsgBoxST "MakeList: 開始位置の指定は Worksheet か Range"
        Exit Function
    End If
    
    Application.DisplayAlerts = False
    bPrevScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    nST = Timer()
    
    nFldCnt = argFlds.Count
    nMaxCol = nMinCol + nFldCnt - 1
    ReDim vDtFlds(1 To nFldCnt)
    ReDim sNmFlds(1 To nFldCnt)
    ReDim sSpFlds(1 To nFldCnt)
    
    With oSht
        nRow = nMinRow
        
        'タイトル行のセット
        If argOpts.Item("タイトル") <> "" Then
            For nFldIdx = 1 To nFldCnt
                sNmFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("Title")
            Next
            With .Range(.Cells(nRow, nMinCol), .Cells(nRow, nMaxCol))
                .Interior.ColorIndex = ccGREY
                If InStr(argOpts.Item("スタイル"), "自動列幅データ") = 0 Then
                    .Value = sNmFlds
                End If
                If InStr(argOpts.Item("スタイル"), "自動列幅") = 0 Then
                    If InStr(argOpts.Item("タイトル"), "フィルタ") <> 0 Then
                        .AutoFilter
                    End If
                    .EntireColumn.AutoFit
                End If
            End With
            '列幅、横位置指定
            For nFldIdx = 1 To nFldCnt
                If argFlds.Item(nFldIdx).Item("列幅") <> "" Then
                    .Columns(nMinCol + nFldIdx - 1).ColumnWidth _
                        = argFlds.Item(nFldIdx).Item("列幅")
                End If
                If argFlds.Item(nFldIdx).Item("横位置") <> "" Then
                    With .Columns(nMinCol + nFldIdx - 1)
                        Select Case argFlds.Item(nFldIdx).Item("横位置")
                        Case "中央"
                            .HorizontalAlignment = xlCenter
                        Case "右寄"
                            .HorizontalAlignment = xlLeft
                        Case "左寄"
                            .HorizontalAlignment = xlRight
                        End Select
                    End With
                End If
            Next
            'データ型文字色
            If InStr(argOpts.Item("スタイル"), "項目型色付") <> 0 Then
                For nFldIdx = 1 To nFldCnt
                    Select Case argFlds.Item(nFldIdx).Item("Type")
                    Case ccFT_DATE
                        .Cells(1, nFldIdx).Font.ColorIndex = ccPURPLE
                    Case ccFT_ELSE
                        .Cells(1, nFldIdx).Font.ColorIndex = ccBLUE
                    Case ccFT_NONE
                        .Cells(1, nFldIdx).Font.ColorIndex = ccWHITE
                    Case ccFT_TEXT
                    Case Else
                        .Cells(1, nFldIdx).Font.ColorIndex = ccRED
                    End Select
                Next
            End If
            '印刷タイトル
#If VBA7 Then
            Application.PrintCommunication = False
#End If
            With .PageSetup
                .PrintTitleRows = argOpts.Item("行タイトル")
                .PrintTitleColumns = argOpts.Item("列タイトル")
            End With
#If VBA7 Then
            Application.PrintCommunication = True
#End If
            '枠固定
            If argOpts.Item("枠固定セル") <> "" Then
                .Range(argOpts.Item("枠固定セル")).Select
                ActiveWindow.FreezePanes = True
            End If
            nRow = nRow + 1
        End If
        
        '高速化の為、フィールド情報、スタイル等を変数に保存
        For nFldIdx = 1 To nFldCnt
            sSpFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("特殊列")
        Next
        '規定抽出件数
        If argOpts.Item("サンプル抽出件数") <> "" Then
            nMaxCnt = argOpts.Item("サンプル抽出件数")
        End If
        '縞模様セル
        If InStr(argOpts.Item("スタイル"), "縞模様セル") <> 0 Then
            bSimaBgCol = True
        End If
        '１行ごとに再描画
        If argOpts.Item("行再描画") <> "" Then
'            b1LScrUpd = True
        End If
        
        Dim sTotalCnt As String
        If argOpts.Item("件数") = "" Then
            sTotalCnt = " 件(ESCキーで中断)"
        Else
            sTotalCnt = " 件 / " & Format(argOpts.Item("件数"), "00000") & " 件(ESCキーで中断)"
        End If
        
        '列書式のセット
        For nFldIdx = 1 To nFldCnt
            Select Case argFlds.Item(nFldIdx).Item("Type")
            Case ccFT_TEXT, ccFT_DATE, ccFT_NONE
                '文字列書式をセット
                .Columns(nFldIdx).NumberFormatLocal = "@"
            Case ccFT_ELSE
                '標準書式のまま
            End Select
        Next
        
        
        'データ行のセット
        Application.EnableCancelKey = xlErrorHandler
        On Error GoTo KEYBREAK
        bKeyBreak = False
        Do Until argRs.EOF
            
            If nCnt Mod 10 = 0 Then
                Application.StatusBar = "データを抽出中... " & Format(nCnt, "00000") & sTotalCnt
            End If
            
            If nMaxCnt <> 0 And nMaxCnt <= nCnt Then
                Exit Do
            End If
            
            '中断キー押下を検知
            If bKeyBreak Then
                If MsgBoxST("中断しますか？", vbYesNo) = vbYes Then Exit Do
                bKeyBreak = False
            End If
            
            '配列に格納
            nRsIdx = 0
            For nFldIdx = 1 To nFldCnt
'                Set oFld = argFlds.Item(nFldIdx)
                Select Case sSpFlds(nFldIdx)
                Case "空白列"
                Case "行連番" '式でNo.表示
                Case "行番号" '値でNo.表示
                    vDtFlds(nFldIdx) = Format(nRow - 1, "00000")
                Case "データ"
                    vDtFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("データ").Item(nCnt + 1)
                Case "処理実行" '☆
                    vDtFlds(nFldIdx) = Run(argFlds.Item(nFldIdx).Item("処理名"), _
                                        vDtFlds(argFlds.Item(nFldIdx).Item("Ｐ位置")))
                Case Else 'レコードセットの項目
'                    With argRs.Fields(sNmFlds(nFldIdx))
                    With argRs.Fields(nRsIdx)
                        If GetFldType(argRs, nRsIdx) = ccFT_DATE Then
                            vDtFlds(nFldIdx) = Format(.Value, "yyyy/mm/dd hh:nn:ss")
                        Else
                            vDtFlds(nFldIdx) = Nvl(.Value)
                        End If
                    End With
                    nRsIdx = nRsIdx + 1
                End Select
            Next
            
            '行貼り付け
            With .Range(.Cells(nRow, nMinCol), .Cells(nRow, nMaxCol))
                .Value = vDtFlds
                If bSimaBgCol And nRow Mod 2 = 1 Then
                    .Interior.ColorIndex = ccYELLOW
                End If
            End With
            
            If b1LScrUpd Then
                Application.ScreenUpdating = True
                Application.ScreenUpdating = False
            End If
            
            argRs.MoveNext
            nRow = nRow + 1
            nCnt = nCnt + 1
            
        Loop
        On Error GoTo 0
        Application.EnableCancelKey = xlDisabled
        Application.StatusBar = ""
        nMaxRow = nRow - 1
        
        If InStr(argOpts.Item("スタイル"), "自動列幅") <> 0 Then
            With .Range(.Cells(nMinRow, nMinCol), .Cells(nMaxRow, nMaxCol))
                If InStr(argOpts.Item("タイトル"), "フィルタ") <> 0 Then
                    .AutoFilter
                End If
                .EntireColumn.AutoFit
            End With
            If argOpts.Item("タイトル") <> "" Then
                .Range(.Cells(nMinRow, nMinCol), .Cells(nMinRow, nMaxCol)).Value = sNmFlds
            End If
        End If
        
        'データ貼付後の整形
        For nFldIdx = 1 To nFldCnt
            Select Case sSpFlds(nFldIdx)
            Case "行連番" '式でNo.表示
                With .Range(.Cells(nMinRow + 1, nFldIdx), .Cells(nMaxRow, nFldIdx))
                    .HorizontalAlignment = xlHAlignLeft
                    .NumberFormatLocal = "00000"
                    .FormulaR1C1 = "=ROW()-1"
                    .Select
                End With
            Case Else
            End Select
        Next
        
        .Range("A1").Select
        With .Range(.Cells(nMinRow, nMinCol), .Cells(nMaxRow, nMaxCol))
            If InStr(argOpts.Item("スタイル"), "空白セル強調") <> 0 Then
'                With .FormatConditions.Add(Type:=xlExpression, _
'                    Formula1:="=IF(A$1<>"""",ISBLANK(A1),FALSE)")
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=IF(AND($A$1=""MAINTEN"",COLUMN()=2),FALSE,ISBLANK(A1))")
                    .Interior.ColorIndex = ccPINK
                End With
            End If
            If InStr(argOpts.Item("スタイル"), "縞模様書式") <> 0 Then
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=AND(MOD(SUBTOTAL(3,$A$1:$A1),2)=1,ROW()<>1)")
                    .Interior.ColorIndex = ccYELLOW
                End With
            End If
            If InStr(argOpts.Item("スタイル"), "罫線描画") <> 0 Then
                With .Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                With .Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                End With
                If nMinRow <> nMaxRow Then
                    With .Borders(xlInsideHorizontal)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End If
                If nMinCol <> nMaxCol Then
                    With .Borders(xlInsideVertical)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                    End With
                End If
            End If
        End With
        
        'デフォルトのページ設定
#If VBA7 Then
        Application.PrintCommunication = False
#End If
        With .PageSetup
            .LeftHeader = oSht.Name
            .LeftMargin = Application.InchesToPoints(0.393700787401575)
            .RightMargin = Application.InchesToPoints(0.393700787401575)
            .TopMargin = Application.InchesToPoints(0.78740157480315)
            .BottomMargin = Application.InchesToPoints(0.78740157480315)
        End With
#If VBA7 Then
        Application.PrintCommunication = True
#End If
    
        nET = Timer()
        
        argOpts.Item("抽出") = TimeElapsed(nST, nET)
        argOpts.Item("件数") = nCnt
        
        If InStr(argOpts.Item("コメント"), "件数、秒数") <> 0 Then
            Dim sCmtTxt As String
            If argOpts.Item("件数") <> "" Then
                sCmtTxt = "件数:" & argOpts.Item("件数") & "件"
            End If
            If argOpts.Item("実行") <> "" Then
                sCmtTxt = IIf(sCmtTxt = "", sCmtTxt, sCmtTxt & " ")
                sCmtTxt = sCmtTxt & "実行:" & argOpts.Item("実行") & "秒 抽出:" & argOpts.Item("抽出") & "秒"
            End If
            
            'コメント追加
            Call SetCmtVal(.Cells(1, 2), , sCmtTxt)
'            .Comment.Shape.Width = 180
'            .Comment.Shape.Height = 11
        
        End If
    End With
        
    MakeList = True
    GoTo END_FUNC
KEYBREAK:
    bKeyBreak = True
    Resume
END_FUNC:
    Application.EnableCancelKey = xlInterrupt
    Application.ScreenUpdating = bPrevScrUpd
    Application.DisplayAlerts = True
'    Application.StatusBar = False
End Function

Function MakeListRunFunc(argPara)
    MakeListRunFunc = "[" & argPara & "]"
End Function


'定型SQL文の取得
Function GetLibSql(argName) As String
    Dim nRow As Long
    With GetLibSheet
        nRow = 2
        Do Until .Cells(nRow, 3) = ""
            If .Cells(nRow, 3) = argName Then
                GetLibSql = .Cells(nRow, 5)
                Exit Function
            End If
            nRow = nRow + 1
        Loop
    End With
End Function

Function SetSqlParam(argSql, argName, argVal)
    SetSqlParam = Replace(argSql, "&" & argName, argVal)
End Function

Function GetSqlParams(argSql, argPrms)
    Dim arLines, i As Integer, j As Integer
    Dim nStart As Integer, nMatch As Integer, nMatch2 As Integer
    Dim sParam As String
    Set argPrms = New Prop
    arLines = Split(Replace(argSql, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(arLines)
        nStart = 1
        Do
            nMatch = InStr(nStart, arLines(i), "&")
            If nMatch = 0 Then Exit Do
            For j = nMatch + 1 To Len(arLines(i))
                If InStr("&', ", Mid(arLines(i), j, 1)) <> 0 Then Exit For
            Next
            sParam = Mid(arLines(i), nMatch + 1, j - nMatch - 1)
            If sParam <> "" Then argPrms.Item(sParam) = sParam
            nStart = j
        Loop
    Next
    If argPrms.Count > 0 Then
        GetSqlParams = True
    End If
End Function

'ＳＱＬ文をセルにセット
Private Function SetSql(argSql, Optional argCell, Optional argSqlTitle) As Integer
    Dim oCell, arLines, nLineNo As Integer
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long, nMinCol As Long
    
    If IsMissing(argCell) Then
        nMinCol = ActiveCell.Column
        nRow = ActiveCell.Row
        Do Until nRow = 1 Or ActiveSheet.Cells(nRow, nMinCol) = ""
            nRow = nRow - 1
        Loop
        If ActiveSheet.Cells(nRow, nMinCol) = "" Then
            nRow = nRow + 1
        End If
        nMinRow = nRow
'☆タイトルを考慮した貼り付け
'        Do Until Cells(nRow, nMinCol).Value = ""
'            If InStr(UCase(Cells(nRow, nMinCol)), "SELECT") <> 0 Then
'                nMinRow = nRow
'                Exit Do
'            End If
'            nRow = nRow + 1
'        Loop
        Set oCell = ActiveSheet.Cells(nMinRow, nMinCol)
    Else
        nMinCol = argCell.Column
        nMinRow = argCell.Row
        Set oCell = argCell
    End If
    nRow = nMinRow
    Do Until Cells(nRow, nMinCol).Value = ""
        nRow = nRow + 1
    Loop
    nMaxRow = nRow
    Range(Cells(nMinRow, nMinCol), Cells(nMaxRow, nMinCol)).Value = ""
    
    If Not IsMissing(argSqlTitle) Then
        argSql = argSqlTitle & vbLf & argSql
    End If
    If Mid(argSql, Len(argSql)) <> vbLf Then
        argSql = argSql & vbLf
    End If
    argSql = Replace(argSql, vbCrLf, vbLf)
    arLines = Split(argSql, vbLf)
    For nLineNo = 0 To UBound(arLines)
        oCell.Offset(nLineNo, 0) = arLines(nLineNo)
    Next
    SetSql = UBound(arLines) + 1
End Function

Private Function GetDBObjectType(argObjName) As String
    Dim sSQL As String
    Dim oRs
    Dim sObjName As String, sTypName As String, sCmtName As String
    
    
    If argObjName <> "" Then
        If InStr(argObjName, ",") <> 0 Then
            Dim arNames
            arNames = Split(argObjName, ",")
            sObjName = arNames(0)
            sTypName = arNames(1)
            If UBound(arNames) = 2 Then
                sCmtName = arNames(2)
            End If
        Else
            sObjName = argObjName
        End If
    
        If GetCfgVal("接続モード") = 1 Then 'OO4O/SQL*Net
            
            sSQL = GetLibSql("オブジェクト種別取得")
            
            sSQL = SetSqlParam(sSQL, "オブジェクト名", sObjName)
            sSQL = SetSqlParam(sSQL, "オブジェクト種別", sTypName)
            sSQL = SetSqlParam(sSQL, "コメント文字列", sCmtName)
            
            If Not ExecSql(GetConStr, sSQL, oRs) Then
                Exit Function
            End If
            
            Select Case oRs("CNT").Value
            Case Is > 1
                GetDBObjectType = "LIST"
            Case 1
                GetDBObjectType = oRs("OBJECT_TYPE").Value
            End Select
            
        Else 'ADO/ODBC
            GetDBObjectType = "ADO/ODBCモードのオブジェクト情報取得はまだ未対応です"
        End If
    End If
    
End Function

Private Function GetDBObjectList(argObjName, argRs, argFlds, argOpts)
    Dim sSQL As String, sExtSQL As String
    Dim oExtRs, oFld, oDat
    Dim sObjName As String, sTypName As String, sCmtName As String
    
    If argOpts.Item("拡張一覧") = "" Then
        sSQL = GetLibSql("オブジェクト一覧取得")
    Else
        sSQL = GetLibSql("オブジェクト拡張一覧取得")
    End If
    
    If InStr(argObjName, ",") <> 0 Then
        Dim arNames
        arNames = Split(argObjName, ",")
        sObjName = arNames(0)
        sTypName = arNames(1)
        sCmtName = arNames(2)
    Else
        sObjName = argObjName
    End If
    
    sSQL = SetSqlParam(sSQL, "オブジェクト名", sObjName)
    sSQL = SetSqlParam(sSQL, "オブジェクト種別", sTypName)
    sSQL = SetSqlParam(sSQL, "コメント文字列", sCmtName)
    
    If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
        Exit Function
    End If
    
    If argOpts.Item("拡張一覧") <> "" Then
        If GetCfgVal("拡張情報：状況") <> 1 Then
            Call argFlds.Remove("STATUS")
        End If
        If GetCfgVal("拡張情報：更新日時") <> 1 Then
            Call argFlds.Remove("LAST_DDL_TIME")
        End If
        If GetCfgVal("拡張情報：シノニム") <> 1 Then
            Call argFlds.Remove("FROM_SYNONYM")
        End If
        If GetCfgVal("拡張情報：件数") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If argRs("OBJECT_TYPE").Value = "TABLE" Then
                    sExtSQL = GetLibSql("件数取得")
                    sExtSQL = SetSqlParam(sExtSQL, "取得対象", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("CNT").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "件数"
            oFld.Item("特殊列") = "データ"
            Set oFld.Item("データ") = oDat
            Set argFlds.Item() = oFld
        End If
        If GetCfgVal("拡張情報：行数") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If argRs("OBJECT_TYPE").Value = "PROCEDURE" Or argRs("OBJECT_TYPE").Value = "FUNCTION" Then
                    sExtSQL = GetLibSql("ソース行数取得")
                    sExtSQL = SetSqlParam(sExtSQL, "オブジェクト名", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("CNT").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "行数"
            oFld.Item("特殊列") = "データ"
            Set oFld.Item("データ") = oDat
            Set argFlds.Item() = oFld
        End If
        If GetCfgVal("拡張情報：サイズ") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If InStr("TABLE,INDEX", argRs("OBJECT_TYPE").Value) <> 0 Then
                    sExtSQL = GetLibSql("サイズ取得")
                    sExtSQL = SetSqlParam(sExtSQL, "オブジェクト名", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("MB").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "サイズ(MB)"
            oFld.Item("特殊列") = "データ"
            Set oFld.Item("データ") = oDat
            Set argFlds.Item() = oFld
        End If
    End If
    
    GetDBObjectList = True
End Function

Private Function GetDBObjectInfo(argObjName, argObjType, argRs, argFlds, argOpts)
    Dim sSQL As String
    
    Select Case argObjType
    Case "TABLE", "VIEW"
        sSQL = GetLibSql("表定義取得")
        
        sSQL = SetSqlParam(sSQL, "オブジェクト名", argObjName)
        
        If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
            Exit Function
        End If
    Case "INDEX"
        sSQL = GetLibSql("索引定義取得")
        
        sSQL = SetSqlParam(sSQL, "オブジェクト名", argObjName)
        
        If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
            Exit Function
        End If
    Case "PROCEDURE", "FUNCTION", "PACKAGE", "PACKAGE BODY"
        sSQL = GetLibSql("ソース取得")
        
        sSQL = SetSqlParam(sSQL, "オブジェクト名", argObjName)
        
        If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
            Exit Function
        End If
    End Select
    
    GetDBObjectInfo = True
End Function

Private Function GetDBObjectExtInfo(argObjName, argObjType, argExtName)
    Dim oExtInfo As New Prop
    Dim oExtRs, sExtSQL As String
    With oExtInfo
        If InStr(argExtName, "状況") <> 0 Or InStr(argExtName, "更新日時") <> 0 _
        Or InStr(argExtName, "シノニム") <> 0 Then
            sExtSQL = GetLibSql("オブジェクト拡張一覧取得")
            sExtSQL = SetSqlParam(sExtSQL, "オブジェクト名", argObjName)
            If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                If InStr(argExtName, "状況") <> 0 Then
                    .Item("STATUS") = "状況: " & oExtRs("STATUS").Value
                End If
                If InStr(argExtName, "更新日時") <> 0 Then
                    .Item("LAST_DDL_TIME") = "更新: " & oExtRs("LAST_DDL_TIME").Value
                End If
                If InStr(argExtName, "シノニム") <> 0 And oExtRs("FROM_SYNONYM").Value <> "" Then
                    .Item("FROM_SYNONYM") = oExtRs("FROM_SYNONYM").Value
                End If
            End If
        End If
        If InStr(argExtName, "件数") <> 0 Then
            Select Case argObjType
            Case "TABLE", "VIEW"
                sExtSQL = GetLibSql("件数取得")
                sExtSQL = SetSqlParam(sExtSQL, "取得対象", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("COUNT") = "件数: " & oExtRs("CNT").Value
                End If
            End Select
        End If
        If InStr(argExtName, "行数") <> 0 Then
            Select Case argObjType
            Case "PROCEDURE", "FUNCTION"
                sExtSQL = GetLibSql("ソース行数取得")
                sExtSQL = SetSqlParam(sExtSQL, "オブジェクト名", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("LINES") = "行数: " & oExtRs("CNT").Value
                End If
            End Select
        End If
        If InStr(argExtName, "サイズ") <> 0 Then
            Select Case argObjType
            Case "TABLE", "INDEX"
                sExtSQL = GetLibSql("サイズ取得")
                sExtSQL = SetSqlParam(sExtSQL, "オブジェクト名", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("MB") = "サイズ: " & Trim(oExtRs("MB").Value) & "MB"
                End If
            End Select
        End If
    End With
    Set GetDBObjectExtInfo = oExtInfo
End Function

'コメント内容の取得
Private Function GetCmtVal(argCell, Optional argName) As String
    With argCell
        If .Comment Is Nothing Then
            GetCmtVal = ""
        Else
            If IsMissing(argName) Then
                GetCmtVal = .Comment.Text
            Else
                GetCmtVal = GetItemStr(.Comment.Text, argName)
            End If
        End If
    End With
End Function

'コメント内容の設定
Private Function SetCmtVal(argCell, Optional argName, Optional argVal = "") As Comment
    Dim sText As String
    Dim bPrevScrUpd As Boolean
    bPrevScrUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With argCell
        If .Comment Is Nothing Then
            .AddComment
        End If
        If IsMissing(argName) Then
            sText = argVal
        Else
            sText = SetItemStr(.Comment.Text, argName, argVal)
        End If
        
        Call .Comment.Text(sText)
        
        .Select
        .Comment.Visible = True
        Selection.Comment.Shape.Select True
        Selection.AutoSize = True
        .Comment.Visible = False
        
        Set SetCmtVal = .Comment
    End With
    Application.ScreenUpdating = bPrevScrUpd
End Function

Private Function SetItemStr(argStr, argName, argVal) As String
    Dim arStr
    Dim arNam0
    Dim arNam1
    Dim i
    Dim bHit
    If InStr(argStr, ":") <> 0 Then
        arStr = Split(argStr, ":")
        For i = 0 To UBound(arStr) - 1
            arNam0 = Split(Trim(Replace(arStr(i + 0), vbCrLf, " ")), " ")
            arNam1 = Split(Trim(Replace(arStr(i + 1), vbCrLf, " ")), " ")
            If arNam0(UBound(arNam0)) = argName Then
                arNam1(0) = argVal
                bHit = True
            End If
            SetItemStr = SetItemStr & IIf(i > 0, IIf(InStr(argStr, vbCrLf) = 0, " ", vbCrLf), "") & _
                arNam0(UBound(arNam0)) & ": " & arNam1(0)
        Next
        If Not bHit Then
            If SetItemStr <> "" And Right(SetItemStr, 2) <> vbCrLf Then
                SetItemStr = SetItemStr & IIf(InStr(argStr, vbCrLf) = 0, " ", vbCrLf)
            End If
            SetItemStr = SetItemStr & argName & ": " & argVal
        End If
    Else
        SetItemStr = argStr & IIf(argStr = "", "", IIf(InStr(argStr, vbCrLf) = 0, " ", vbCrLf)) & _
            argName & ": " & argVal
    End If
End Function

Private Function GetItemStr(argStr, argName) As String
    Dim arStr
    Dim arNam
    Dim i
    If InStr(argStr, ":") <> 0 Then
        arStr = Split(argStr, ":")
        For i = 0 To UBound(arStr) - 1
            arNam = Split(Trim(arStr(i)), " ")
            If arNam(UBound(arNam)) = argName Then
                arNam = Split(Trim(arStr(i + 1)), " ")
                GetItemStr = arNam(0)
                Exit For
            End If
        Next
    End If
End Function

Private Function MsgBoxST(argPrm, Optional argBtn)
    If IsMissing(argBtn) Then
        MsgBoxST = MsgBox(argPrm, Title:=ccAppTitle)
    Else
        MsgBoxST = MsgBox(argPrm, argBtn, ccAppTitle)
    End If
End Function

Private Function MsgBalloon(argTitle, argMsg)
    If Not IsEmpty(coBalloon) Then
        coBalloon.Close
    End If
    Set coBalloon = Application.Assistant.NewBalloon
    With coBalloon
'        .Button = 0 'msoButtonSetNone
'        .Button = 1
        .Mode = 1 'msoModeAutoDown
'        .Mode = 2 'msoModeModeless
'        .Callback = "MsgBalloonProc"
        .Heading = argTitle
        .Text = argMsg
        .Show
    End With
End Function

Public Sub MsgBalloonProc(oBln As Balloon, iBtn As Long, iId As Long)
    With oBln
        .Close
    End With
End Sub

Private Function Nvl(argVal)
    If IsNull(argVal) Then
        Nvl = ""
    Else
        Nvl = argVal
    End If
End Function


Private Function GetFldType(argRs, argFldIdxOrName)
    Dim nFldType
    If GetCfgVal("接続モード") = 1 Then
'式項目の項目型判定がおかしいので .type から .OraIDataType に変更
'        Select Case nFldType
'        Case vbEmpty
'            GetFldType = ccFT_NONE
'        Case 8
'            GetFldType = ccFT_DATE 'ORADB_DATE
'        Case 10, 11, 12
'            GetFldType = ccFT_TEXT 'ORADB_TEXT, ORADB_LONGBINARY, ORADB_MEMO
'        Case Else
'            GetFldType = ccFT_ELSE 'ORADB_INTEGER, ORADB_LONG, ORADB_DOUBLE, etc.
'        End Select
        On Error Resume Next
        nFldType = argRs(argFldIdxOrName).OraIDataType
        On Error GoTo 0
        Select Case nFldType
        Case vbEmpty
            GetFldType = ccFT_NONE
        Case 12
            GetFldType = ccFT_DATE 'ORATYPE_DATE
        Case 1, 96
            GetFldType = ccFT_TEXT 'ORATYPE_VARCHAR2, ORATYPE_CHAR
        Case Else
            GetFldType = ccFT_ELSE 'ORATYPE_NUMBER, etc.
        End Select
    Else
        On Error Resume Next
        nFldType = argRs(argFldIdxOrName).Type
        On Error GoTo 0
        Select Case nFldType
        Case vbEmpty
            GetFldType = ccFT_NONE
        Case 133, 134, 135
            GetFldType = ccFT_DATE 'adDBDate, adDBTime, adDBTimeStamp
        Case 8, 129, 200
            GetFldType = ccFT_TEXT 'adBSTR, adChar, adVarChar

        Case Else
            GetFldType = ccFT_ELSE 'etc.
        End Select
    End If
End Function

Public Function Split(sStr As Variant, sDelim As Variant) As Variant
    Dim vArray() As Variant
    Dim i As Long
    Dim bLp As Boolean
    Dim nPrevPos As Long, nNextPos As Long

    nPrevPos = 1: nNextPos = 0
    bLp = False

    Do
        nNextPos = InStr(nPrevPos, sStr, sDelim, vbBinaryCompare)
       
        If nNextPos = 0 Then: bLp = True: nNextPos = Len(sStr) + 1

        ReDim Preserve vArray(0 To i)
        vArray(i) = _
              CStr(Mid(sStr, nPrevPos, nNextPos - nPrevPos))
        nPrevPos = nNextPos + Len(sDelim)
        i = i + 1
    Loop Until bLp = True

    Split = vArray

End Function

'連続した空白文字での分割
Public Function SplitSh(sStr As Variant) As Variant
    Dim vArray() As Variant
    Dim i As Long
    Dim j As Long
    Dim sChr As String
    Dim sTmp As String
    
    i = 0
    
    For j = 1 To Len(sStr) + 1
        sChr = Mid(sStr, j, 1)
        If sChr = " " Or sChr = "　" Then
            If sTmp <> "" Then
                ReDim Preserve vArray(0 To i)
                vArray(i) = sTmp
                sTmp = ""
                i = i + 1
            End If
        Else
            sTmp = sTmp & sChr
        End If
    Next
    
    If sTmp <> "" Then
        ReDim Preserve vArray(0 To i)
        vArray(i) = sTmp
        sTmp = ""
        i = i + 1
    End If

    SplitSh = vArray

End Function

'CSV文字列の分割
Public Function SplitCsv(sStr As Variant) As Variant
    Dim vArray() As Variant
    Dim i As Long
    Dim j As Long
    Dim sChr As String
    Dim sNex As String
    Dim sTmp As String
    Dim nLen1 As Long
    Dim bQuo As Boolean
    
    i = 0
    nLen1 = Len(sStr) + 1
    
    For j = 1 To nLen1
        sChr = Mid(sStr, j, 1)
        If j + 1 <= nLen1 Then
            sNex = Mid(sStr, j + 1, 1)
        Else
            sNex = ""
        End If
        
        If Not bQuo Then
            If sChr = """" Then
                bQuo = True
            ElseIf sChr = "," Then
                ReDim Preserve vArray(0 To i)
                vArray(i) = Trim(sTmp)
                sTmp = ""
                i = i + 1
            Else
                sTmp = sTmp & sChr
            End If
        Else
            If sChr = """" And sNex = """" Then
                sTmp = sTmp & """"
                j = j + 1
            ElseIf sChr = """" Then
                bQuo = False
            Else
                sTmp = sTmp & sChr
            End If
        End If
    Next
    
    If sTmp <> "" Then
        ReDim Preserve vArray(0 To i)
        vArray(i) = Trim(sTmp)
        sTmp = ""
        i = i + 1
    End If

    SplitCsv = vArray

End Function

Public Function Replace(sStr As Variant, sOld As Variant, sNew As Variant) As String
    Dim nStart As Long, nMatch As Long, nOld As Long, nStr As Long
    Dim sTmp As String
    nStart = 1
    nOld = Len(sOld)
    nStr = Len(sStr)
    Do
        nMatch = InStr(nStart, sStr, sOld)
        If nMatch = 0 Then
            If nStart <= Len(sStr) Then
                sTmp = sTmp & Mid(sStr, nStart, nStr - nStart + 1)
            End If
            Exit Do
        Else
            If nStart < nMatch Then
                sTmp = sTmp & Mid(sStr, nStart, nMatch - nStart) & sNew
            Else
                sTmp = sTmp & sNew
            End If
            nStart = nMatch + nOld
        End If
    Loop
    Replace = sTmp
End Function

Function RPAD(argStr, argNum) As String
    RPAD = Left(argStr & Space(argNum), argNum)
End Function

Function TimeElapsed(argStart, argEnd)
    Dim nElapsed
    If argEnd >= argStart Then
        nElapsed = argEnd - argStart
    Else
        nElapsed = (86400 - argStart) + argEnd ' 真夜中の0時を跨いだときの対処
    End If
    TimeElapsed = Format(nElapsed, "0.00")
End Function

Private Sub ST_WIN_OnAction()
    Call ST_WIN.OnAction
End Sub

Sub test()
'    Dim coCon
'    Set coCon = CreateObject("ADODB.Connection")
'    coCon.Open "DSN=server;DBQ=server;UID=user;PWD=pass"
'    Set coCat = CreateObject("ADOX.Catalog")
'    coCat.ActiveConnection = coCon
'    MsgBox coCat.tables.Count
'
'    If GetKeyState(VK_SHIFT) < 0 Then
'        MsgBox "T"
'    Else
'        MsgBox "F"
'    End If
'
'Dim oProp As New Prop
'With oProp
'    .Item(1) = "A"
'    .Item(2) = "B"
'    .Item(3) = "C"
'    .Item(1) = "A2"
'    .Item(3) = "C2"
'    .Item(1, False) = "X"
'    .Item(2, False) = "Y"
'    .Item(6, False) = "Z"
'    .Item() = "1"
'    .Item() = "2"
'    .Item() = "3"
'    Set .Item("col") = New Collection
'    Call .Item("col").Add("zzz")
'    MsgBox .Item("col").Count
'End With
 
' Call GetNewSheet("test", True)
'
'Dim s
's = GetLibSql("オブジェクト一覧取得")
's = SetSqlParam(s, "オブジェクト名", "TBWE2")
'MsgBox s
'
'
'    Dim oPop As Office.CommandBar
'    With Application.CommandBars
'        .Item("test").Delete
'        Set oPop = .Add("test", msoBarPopup, False, True)
'        With oPop.Controls
'            With .Add(msoControlButton)
'                .Caption = "item1"
'            End With
'            With .Add(msoControlButton)
'                .Caption = "item2"
'            End With
'            With .Add(msoControlButton)
'                .Caption = "item3"
'            End With
'        End With
'        oPop.ShowPopup
'    End With
'
'Dim s, t, l
's = GetSql(t, l)
'MsgBox "[" & t & "]" & vbCrLf & "[" & l & "]" & vbCrLf & s
'
'MsgBox GetItemStr("A:a1 B:b1 C:c1", "B")z
'MsgBox SetItemStr("A:a1 B:b1 C:c1", "B", "z")
'
'Call SetCmtVal(ActiveCell, "A", "4")
'Call SetCmtVal(ActiveCell, "B", "5")
'Call SetCmtVal(ActiveCell, "C", "6")
'
'MsgBalloon "テスト"

'Dim RE
'Set RE = CreateObject("VBScript.RegExp")

'Call MakeSheet("ZZZ", ActiveSheet, coSTBook.Worksheets("WfOU_"))

'Call Shell("cmd /c C:\cygwin64\bin\bash --login -c uname > C:\ExSQell.tmp", vbHide)
'Call Shell("cmd /c ""C:\Program Files\Git\bin\bash"" --login -c uname > C:\ExSQell.tmp", vbHide)

Dim arFlds
'arFlds = SplitSh("a    b      c        d       e")
arFlds = SplitCsv("a , "" b,c "" , d")

End Sub
