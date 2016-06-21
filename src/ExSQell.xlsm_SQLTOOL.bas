'========================================================
' ExSQell.xlsm - Copyright (C) 2016 nmrmsys
' sqltools.xls - Copyright (C) 2005 M.Nomura
'========================================================

Option Explicit

'�萔��`
Const ccAppTitle = "ExSQell"
Const ccCfgShtName = "�ݒ�"
Const cnCfgMaxRow = 30
Const ccLibShtName = "���C�u����"
Const ccQryShtName = "���s����"
Const ccTblShtName = "�\��`"
Const ccIdxShtName = "������`"
Const ccSrcShtName = "�\�[�X"
Const ccWrkShtName = "��Ɨp"
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

'�ϐ���`
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

'���b�Z�[�W��`
Const ccUpdateDisclaim = "�X�V�@�\���g�p����ɂ́u�X�VSQL�̎��s�v�ݒ���m�F���ĉ�����" & vbCrLf _
                        & "�Ȃ��{�c�[���͖����J�����ׁ̈A�{�Ԋ��ւ̍X�V�͂��T��������"


Sub Auto_Open()
    '�}�N�����݃u�b�N��ޔ�
    Set coSTBook = ThisWorkbook
    '�}�N���N���V���[�g�L�[�o�^
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
    '�g�p�͈͂̍ő�s��擾
    With ActiveSheet.UsedRange
        nMaxRows = .Row - 1 + .Rows.Count
        nMaxCols = .Column - 1 + .Columns.Count
        sMaxCols = Cells(1, nMaxCols).Address(False, False)
        sMaxCols = Left(sMaxCols, Len(sMaxCols) - 1)
    End With
    '�s�S�I�����ǂ���
'    MsgBox ActiveSheet.UsedRange.Address
    If Selection.Rows.Count > nMaxRows Then
        '��S�I�����ǂ���
        If Selection.Columns.Count > nMaxCols Then
            '�V�[�g�I��
            Set GetSafeSelection = ActiveSheet.UsedRange
        Else
            '��I��
            For Each oArea In Selection.Areas
                If sRange <> "" Then sRange = sRange & ","
                sRange = sRange & Replace(oArea.Address, ":", "1:") & nMaxRows
            Next
            Set GetSafeSelection = Range(sRange)
        End If
    Else
        '��S�I�����ǂ���
        If Selection.Columns.Count > nMaxCols Then
            '�s�I��
            For Each oArea In Selection.Areas
                If sRange <> "" Then sRange = sRange & ","
                sRange = sRange & "$A" & Mid(Replace(oArea.Address, ":$", ":$" & sMaxCols), 2)
            Next
            Set GetSafeSelection = Range(sRange)
        Else
            '�͈͑I��
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
            sRowSep = InputBox("�s��؂蕶���̎w��", ccAppTitle)
            sColSep = ""
        Else
            sRowSep = ""
            sColSep = InputBox("���؂蕶���̎w��", ccAppTitle)
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
    If ST_WIN.ShowModal2(sSQL, "�W�J", "EDIT") Then
        Call SetSql(sSQL)
    End If
End Sub

Sub ST_ShtUniqChk()
    Dim nMaxRows As Long, nMaxCols As Long, nChkCol As Long, sFormula As String, oCol, oFC
    '�g�p�͈͂̍ő�s��擾
    With ActiveSheet.UsedRange
        nMaxRows = .Row - 1 + .Rows.Count
        nMaxCols = .Column - 1 + .Columns.Count
    End With
    If Selection.Rows.Count > nMaxRows Then
        nChkCol = nMaxCols + 1
        For Each oCol In Range(Cells(1, 1), Cells(1, nMaxCols))
            If Not (Range("A1") = "MAINTEN" And oCol.Column <= 2) Then
                If oCol = "�d��CHK" Or oCol = "" Then
                    nChkCol = oCol.Column
                    Exit For
                End If
            End If
        Next
        For Each oCol In GetSafeSelection.Columns
            If oCol.Rows(1) <> "�d��CHK" And oCol.Rows(1) <> "�d��KEY" Then
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
            .Value = "�d��CHK"
        End With
        With Range(Cells(2, nChkCol), Cells(nMaxRows, nChkCol))
            .NumberFormat = ""
            .FormulaR1C1 = "=IF(COUNTIF(R2C" & nChkCol + 1 & ":R" & nMaxRows & "C" & nChkCol + 1 & _
                            ",RC" & nChkCol + 1 & ")>1,""�d�����R�[�h"","""")"
        End With
        With Cells(1, nChkCol + 1)
            .ColumnWidth = 10.5
            .Value = "�d��KEY"
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
            If oCol = "�d��CHK" Then
                nChkCol = oCol.Column
                Exit For
            End If
        Next
        If nChkCol = 0 Then
            Call MsgBoxST("�d���`�F�b�N�������I�����Ă�������")
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
    Cells.Font.Name = "�l�r �S�V�b�N"
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

'�N�G���[�����s
Sub ST_Query(Optional argSql, Optional argSqlTitle)
    Dim sConStr As String
    Dim sSQL As String
    Dim sSQLType As String
    Dim sTitle As String, sObjName As String
    Dim sShtName As String, sShtMode As String
    Dim oRs, oFld, oFlds, oOpts, oSht
    Dim nFldType As Integer, nEffRecCnt As Long
    'ExSQell�b�胍�W�b�N
    Dim sCmd As String
    Dim nDefExecMode As Integer
    
    nDefExecMode = GetDefExecMode '2 Shell���s���[�h
    
    '�V�[�g��ʂ̔���
    sShtMode = ActiveSheet.Range("A1")
    Select Case sShtMode
    Case "MAINTEN"
        '�����e�i���X���f
        Call UpdateMaintenanceSheet
    Case "NAME" '�\�[�X��`
        Call ST_Maintenance '�R���p�C��
        Exit Sub
    Case "TABLE_NAME", "INDEX_NAME" '�\��`�A������`
        Exit Sub
    Case Else
        '�ʏ�N�G�����s�A���o��������
        
        'ExSQell�b�胍�W�b�N
        If nDefExecMode = 2 Or Left(ActiveCell.Value, 2) = "$ " Then
            
            'Shell���s���[�h�Ő擪 $ �������ꍇ�͕⊮
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
                '���݃Z���ʒu����r�p�k���A�^�C�v�A�V�[�g���̎擾
                sSQL = GetSql(sSQLType, sShtName)
                If sSQL = "" Then
                    Exit Sub
                End If
            Else
                '�r�p�k�����w�肵�Ď��s���ꂽ�ꍇ
                sSQL = argSql
                sSQLType = "SELECT"
                sShtName = argSqlTitle
            End If
        End If
        
        'ALT���������ȊO�̏ꍇ�͒��o�����w��
        If Not GetKeyState(VK_MENU) < 0 And Not GetKeyState(VK_SHIFT) < 0 And sSQLType = "TABLE" Then
            Call MakeQpsSheet(Trim(ActiveCell))
            Exit Sub
        End If
        
        Select Case sSQLType
        Case "SELECT", "TABLE", "RELOAD"
        
            'ExSQell�b�胍�W�b�N
            If sCmd <> "" Then
                '�V�F���R�}���h�̎��s
                If Not ExecShell(sCmd, oRs, oFlds) Then
                    Exit Sub
                End If
            Else
                '�N�G���̎��s
                If Not ExecQuery(sConStr, sSQL, sSQLType, oRs, oFlds, oOpts) Then
                    Exit Sub
                End If
            End If
        
            
            '���݃Z���ʒu����̃V�[�g�����󔒂̏ꍇ�̓f�t�H���g�̃V�[�g��
            sShtName = IIf(sShtName = "", ccQryShtName, sShtName)
            
            '�N�G�����ʃV�[�g�̍쐬
            If GetCfgVal("�s�ԍ��\��") <> 3 Then '���Ȃ��ȊO
                With oFlds
                    Set oFld = New Prop
                    oFld.Item("Title") = "No."
                    Select Case GetCfgVal("�s�ԍ��\��")
                    Case 1 '�l
                        oFld.Item("�����") = "�s�ԍ�"
                    Case 2 '����
                        oFld.Item("�����") = "�s�A��"
                    End Select
                    Set .Item(0, False) = oFld '�擪�ɒǉ�
                End With
            End If
            With oOpts
                .Item("�^�C�g��") = "�t�B���^"
                .Item("�s�^�C�g��") = "$1:$1"
                If GetCfgVal("�s�ԍ��\��") <> 3 Then '���Ȃ��ȊO
                    .Item("�g�Œ�Z��") = "B2"
                    .Item("��^�C�g��") = "$A:$A"
                End If
                .Item("�X�^�C��") = "�r���`��A���ڌ^�F�t"
                Select Case GetCfgVal("�Ȗ͗l�\��")
                Case 1 '�����t�������ŕ\��
                    .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�Ȗ͗l����"
                Case 2 '�Z���w�i�F�ŕ\��
                    .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�Ȗ͗l�Z��"
                End Select
                Select Case GetCfgVal("�󔒃Z������")
                Case 1 '�����t�������ŋ���
                    .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�󔒃Z������"
                End Select
                Select Case GetCfgVal("�񕝂̃T�C�Y����")
                Case 1 '��S�̂̕�����
                    .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�����񕝗�S��"
                Case 3 '�f�[�^�̕�����
                    .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�����񕝃f�[�^"
                End Select
                .Item("�R�����g") = "" & IIf(GetCfgVal("�����A�b���̃R�����g") = 1, "�A�����A�b��", "")
            End With
            
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            If sShtMode = "QPARAMS" Then
                If ActiveSheet.Range("I2") = "" Then
                    Select Case GetCfgVal("���o�����V�[�g")
                    Case 1 '���o���ɍ폜
                        ActiveSheet.Delete
                    Case 2 '���o���ɔ�\��
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
                Call SetCmtVal(.Range("A1"), , "�ڑ���: " & sConStr & vbLf & sSQL)
            End With
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
        
        Case "DML", "DDL", "DCL", "BLOCK"
            Dim arLines, nLineNo As Integer, nTotalCnt As Integer, bChk As Boolean, bRet As Boolean
            Dim sTotalCnt As String
            Select Case GetCfgVal("�X�VSQL�̎��s")
            Case 2, 3 '�m�F���Ď��s
                If MsgBoxST("�ڑ���F" & GetConDisp(sConStr) & vbCrLf & vbCrLf _
                    & "�X�VSQL�����s���܂����H", vbYesNo + vbDefaultButton2) = vbNo _
                Then
                    Exit Sub
                End If
                If GetCfgVal("�X�VSQL�̎��s") = 2 Then
                    bChk = True
                End If
            Case 4 '�ⓚ���p�Ŏ��s
                '�m�F����
            Case Else '���s���Ȃ�
                Call MsgBoxST(ccUpdateDisclaim)
                Exit Sub
            End Select
            'DB�ڑ��ASQL���s
            arLines = Split(sSQL, ";" & vbCrLf)
            nTotalCnt = UBound(arLines) + 1
            sTotalCnt = "/" & Format(UBound(arLines) + 1, "000")
            For nLineNo = 1 To UBound(arLines) + 1
                If nTotalCnt = 1 Then
                    Application.StatusBar = "�X�VSQL���s��..."
                    bRet = True
                Else
                    If bChk Then
                        bRet = ST_WIN.ShowModal2(arLines(nLineNo - 1), "���s")
                    Else
                        bRet = True
                    End If
                    Application.StatusBar = "�X�VSQL���s��... " & _
                        Format(nLineNo, "000") & sTotalCnt
                End If
                If bRet Then
                    If Not ExecSql(sConStr, arLines(nLineNo - 1), argOpts:=oOpts) Then
                        Application.StatusBar = ""
                        Exit Sub
                    End If
                    nEffRecCnt = oOpts.Item("�X�V����")
                    If nTotalCnt = 1 And nEffRecCnt <> 0 Then
                        MsgBoxST ("���s���������܂���" & vbCrLf _
                                & "�e�����R�[�h�����F" & nEffRecCnt & "��")
                    Else
                        Application.StatusBar = "�e�����R�[�h�����F" & nEffRecCnt & "��"
                    End If
                End If
            Next
            If nTotalCnt <> 1 Then
                MsgBoxST ("�X�VSQL�o�b�`���s���������܂���")
            End If
            Application.StatusBar = ""
        Case "FUNCTION", "PROCEDURE"
            '��ADO�Ή��v
            sObjName = ActiveCell
            If MsgBoxST(sSQLType & " " & sObjName & "�����s���܂����H", vbYesNo + vbDefaultButton2) = vbYes _
            Then
                Dim sSpSql As String, sSpMsg As String, sRetType As String, oPrms As New Prop, oPrm
                If sSQLType = "FUNCTION" Then
                    sSpSql = "BEGIN :RESULT := " & sObjName & "("
                Else
                    sSpSql = "BEGIN " & sObjName & "("
                End If
                sSQL = GetLibSql("�������擾")
                sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", sObjName)
                If Not ExecSql(sConStr, sSQL, oRs) Then
                    Exit Sub
                End If
                Do Until oRs.EOF
                    If oRs("POSITION") = 0 Then
                        sRetType = oRs("DATA_TYPE").Value
                    Else
                        '���sSQL�Ƀv���[�X�z���_�ǉ�
                        If Right(sSpSql, 1) <> "(" Then
                            sSpSql = sSpSql & ", "
                        End If
                        sSpSql = sSpSql & ":" & oRs("ARGUMENT_NAME")
                        '�p�����[�^�v���p�e�B�ǉ�
                        Set oPrm = New Prop
                        oPrm.Item("Name") = oRs("ARGUMENT_NAME").Value
                        oPrm.Item("InOut") = "[" & oRs("IN_OUT").Value & "]"
                        oPrm.Item("Type") = oRs("DATA_TYPE").Value
                        '���̓p�����[�^�̓_�C�A���O�\��
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
                Set oOpts.Item("�p�����[�^") = oPrms
                oOpts.Item("�߂�l") = sRetType
                If Not ExecSql(sConStr, sSpSql, argOpts:=oOpts) Then
'                    Exit Sub
                End If
                sSpMsg = sSpMsg & sSQLType & " " & sObjName & vbLf & vbLf
                For Each oPrm In oPrms.Items
                    sSpMsg = sSpMsg & RPAD(oPrm.Item("Name"), 18) & RPAD(oPrm.Item("Type"), 8) & _
                        RPAD(oPrm.Item("InOut"), 5) & " =  " & oPrm.Item("Value") & vbLf
                Next
                If sSQLType = "FUNCTION" Then
                    sSpMsg = sSpMsg & vbLf & "�߂�l " & sRetType & " = " & oOpts.Item("�߂�l") & vbLf
                End If
                Call ST_WIN.ShowModal2(sSpMsg)
            End If
        End Select
    End Select
End Sub

Sub ST_CommitMenu()
    Dim sConKey, nConMode, nCmtMode
    
    nConMode = GetCfgVal("�ڑ����[�h")
    nCmtMode = GetCfgVal("�R�~�b�g���[�h")
    
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
                    oItm.Caption = "�R�~�b�g"
                    oItm.OnAction = "ST_CommitMenu_OnAction"
                    oItm.Tag = sConKey
                    oItm.Parameter = "C"
                    Set oItm = oMnu.Controls.Add(msoControlButton)
                    oItm.Caption = "���[���o�b�N"
                    oItm.OnAction = "ST_CommitMenu_OnAction"
                    oItm.Tag = sConKey
                    oItm.Parameter = "R"
                    Set oItm = oMnu.Controls.Add(msoControlButton)
                    oItm.Caption = "�ؒf"
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
            sMsg = "�R�~�b�g"
            oCon.CommitTrans
            oCon.BeginTrans
            oConItem.Item("Disp") = Replace(oConItem.Item("Disp"), "*", "")
        Case "R"
            sMsg = "���[���o�b�N"
            If Left(.Tag, 1) = "1" Then
                oCon.RollBack
            Else
                oCon.RollBackTrans
            End If
            oCon.BeginTrans
            oConItem.Item("Disp") = Replace(oConItem.Item("Disp"), "*", "")
        Case "D"
            sMsg = "�ؒf"
            Set oCon = Nothing
            Call coConObjs.Remove(.Tag)
        End Select
        
        Call MsgBoxST(sMsg & "���������܂���")
        
    End With
End Sub

'�����e�i���X�V�[�g�̍X�V
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
        Select Case GetCfgVal("�X�VSQL�̎��s")
        Case 2, 3 '�m�F���Ď��s
            If MsgBoxST("�ڑ���F" & GetConDisp(sConStr) & vbCrLf & vbCrLf _
                & "�X�VSQL�����s���܂����H", vbYesNo + vbDefaultButton2) = vbNo _
            Then
                Exit Function
            End If
        Case 4 '�ⓚ���p�Ŏ��s
            '�m�F����
        Case Else '���s���Ȃ�
            Call MsgBoxST(ccUpdateDisclaim)
            Exit Function
        End Select
        '�t�B�[���h���擾�p
        If Not ExecSql(sConStr, "SELECT * FROM " & sTblName & " WHERE 0=1", oRs, oFlds) Then
            Exit Function
        End If
        For nRow = nMinRow To nMaxRow
            sSQL = ""
            Select Case UCase(.Cells(nRow, nMinCol - 1))
            Case "I"
                Application.StatusBar = "�ǉ��F" & Format(nRow, "00000") & "�s��"
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
                Application.StatusBar = "�X�V�F" & Format(nRow, "00000") & "�s��"
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
                Application.StatusBar = "�폜�F" & Format(nRow, "00000") & "�s��"
                sSQL = sSQL & "DELETE FROM " & sTblName & vbCrLf
                sSQL = sSQL & "WHERE ROWID = '" & .Cells(nRow, 1) & "'" & vbCrLf
            End Select
            If sSQL <> "" Then
                bFixFlg = True
                'DB�ڑ��ASQL���s
                If Not ExecSql(sConStr, sSQL, argOpts:=oOpts) Then
                    Exit Function
                End If
            End If
        Next
        Application.StatusBar = ""
    End With
    If bFixFlg And GetCfgVal("�����e�V�[�g�c�a���f��") = 1 Then '�����ōĒ��o
        Call ST_Maintenance
    End If
End Function

'�I�u�W�F�N�g�ꗗ/��`�̕\��
Sub ST_Which()
    Dim sWhich1 As String, sWhich2 As String, sWhich3 As String
    Dim sWhich As String, sWBpat As String, sWList As String
    Dim oRs, oFlds, oOpts As New Prop
    Dim sObjType As String, sObjType1 As String, sObjType2 As String
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nCol As Long, nMinCol As Long, nMaxCol As Long
    Dim nFldIdx As Integer
    
    'ExSQell�b�胍�W�b�N
    Dim sCmd As String
    Dim nDefExecMode As Integer
    
    nDefExecMode = GetDefExecMode '2 Shell���s���[�h
    
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
    
    'ExSQell�b�胍�W�b�N
    If nDefExecMode = 2 Or Left(ActiveCell.Value, 2) = "$ " Then
        
        'Shell���s���[�h�Ő擪 $ �������ꍇ�͕⊮
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
    
    '�܂��͊��S�}�b�`�Ō���
    sObjType1 = GetDBObjectType(sWList)
    
    'SHIFT���������̏ꍇ�͕����}�b�`�̌������L�����Z��
    If InStr(sObjType1, "ADO/ODBC") = 0 Then
        If Not GetKeyState(VK_SHIFT) < 0 Then
            sObjType2 = GetDBObjectType(sWBpat)
        End If
    End If
    
    'ALT���������̏ꍇ�͊g���ꗗ���ڂ�\��
    If GetKeyState(VK_MENU) < 0 Then
        oOpts.Item("�g���ꗗ") = "on"
    End If
    
    '���S ���� ����
    '�m   �m   �m
    '�m   �n   �k
    '�m   �k   �k
    '�n   �k   �k
    '�n   �n   �n
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
        If oOpts.Item("�g���ꗗ") <> "" Then
            For nFldIdx = 4 To oFlds.Count
                Cells(nMinRow, nMinCol + nFldIdx - 1) = oFlds.Item(nFldIdx).Item("Title")
            Next
            oOpts.Item("�s�ĕ`��") = "on"
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
                    .Item("�^�C�g��") = "�t�B���^"
                    .Item("�g�Œ�Z��") = "A2"
                    .Item("�s�^�C�g��") = "$1:$1"
                End If
                .Item("�X�^�C��") = "�r���`��A�Ȗ͗l����"
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
                    .Item("�^�C�g��") = "�t�B���^"
                    .Item("�g�Œ�Z��") = "A2"
                    .Item("�s�^�C�g��") = "$1:$1"
                End If
                .Item("�X�^�C��") = "�r���`��A�Ȗ͗l����"
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
                .Item(1).Item("��") = 10
                .Item(2).Item("��") = 12
                .Item(3).Item("��") = 7
                .Item(4).Item("��") = 2
                .Item(5).Item("��") = 85
            End With
            With oOpts
                If nRow = 1 Then
                    .Item("�^�C�g��") = "�t�B���^"
                    .Item("�g�Œ�Z��") = "A2"
                    .Item("�s�^�C�g��") = "$1:$1"
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

'�g�����擾
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
    
    If GetCfgVal("�g�����F��") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "��"
    End If
    If GetCfgVal("�g�����F�X�V����") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "�X�V����"
    End If
    If GetCfgVal("�g�����F�V�m�j��") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "�V�m�j��"
    End If
    If GetCfgVal("�g�����F����") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "����"
    End If
    If GetCfgVal("�g�����F�s��") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "�s��"
    End If
    If GetCfgVal("�g�����F�T�C�Y") = 1 Then
        sExtName = sExtName & IIf(sExtName = "", "", ", ") & "�T�C�Y"
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

'�f�[�^�����e�i���X
Sub ST_Maintenance()
    Dim sConStr As String, sTblName As String, sObjType As String, sObjName As String
    Dim sSQL As String, sErrSQL As String, sErrText As String, nErrLine As Integer
    Dim nRow As Long, nMinRow As Long, nMaxRow As Long
    Dim nMinCol As Long
    Dim oRs, oFld, oFlds, oOpts, oSht As Worksheet
    Dim arLines, nLineNo As Integer
    
    Select Case GetCfgVal("�X�VSQL�̎��s")
    Case 2, 3, 4 '�m�F���Ď��s�A�ⓚ���p�Ŏ��s
        '�m�F����
    Case Else '���s���Ȃ�
        Call MsgBoxST(ccUpdateDisclaim)
        Exit Sub
    End Select
    
    sConStr = GetConStr
    Select Case ActiveSheet.Range("A1")
    Case "QPARAMS" '���o��������
        sTblName = Replace(ActiveSheet.Name, ccQpsShtPref, "")
        '���͒��o��������r�p�k���𐶐�
        sSQL = GetSql("MAINTEN")
        If sSQL = "" Then
            Exit Sub
        End If
    Case "MAINTEN" '�����e�i���X
        '�V�[�g��̃R�����g����ڑ�������A�r�p�k�����擾
        arLines = Split(GetCmtVal(ActiveSheet.Range("A1")), vbLf)
        sTblName = Replace(ActiveSheet.Name, ccMntShtPref, "")
        For nLineNo = 1 To UBound(arLines)
            sSQL = sSQL & arLines(nLineNo)
        Next
    Case "NAME" '�\�[�X��`
        With ActiveSheet
            '�t�B���^�I������Ă��邩
            If .FilterMode Then
                If MsgBoxST("�X�g�A�h�t�@���N�V����/�v���V�[�W�����R���p�C�����܂����H", vbYesNo + vbDefaultButton2) = vbYes _
                Then '��ADO�Ή��v
                    '�R���p�C���\�[�X�̎擾
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
                    '�c�c�k���s
                    If Not ExecSql(sConStr, sSQL) Then
                        Exit Sub
                    End If
                    '�G���[�s��F�t��
                    sErrSQL = GetLibSql("�G���[���擾")
                    sErrSQL = SetSqlParam(sErrSQL, "�I�u�W�F�N�g��", sObjName)
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
                Call MsgBoxST("�R���p�C������\�[�X���t�B���^�I�����Ă�������")
            End If
        End With
        Exit Sub
    Case "TABLE_NAME", "INDEX_NAME" '�\��`�A������`
        Exit Sub
    Case Else '�ʏ�̃V�[�g��
        '�\�����\�I�u�W�F�N�g���ǂ���
        sTblName = Trim(ActiveCell.Value)
        sObjType = GetDBObjectType(sTblName)
        Select Case sObjType
        Case "TABLE"
        Case "FUNCTION", "PROCEDURE"
            If MsgBoxST("�X�g�A�h�t�@���N�V����/�v���V�[�W�����ăR���p�C�����܂����H", vbYesNo + vbDefaultButton2) = vbYes _
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
                                Call SetCmtVal(.Offset(0, 0), , "�f�[�^�x�[�X�G���[������")
                                .Font.ColorIndex = ccRED
                            Else
                                sErrSQL = GetLibSql("�G���[���擾")
                                sErrSQL = SetSqlParam(sErrSQL, "�I�u�W�F�N�g��", arLines(nLineNo))
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
                            Call SetCmtVal(.Offset(0, 0), , "�I�u�W�F�N�g�����݂��܂���")
                            .Font.ColorIndex = ccRED
                        End If
                    End With
                Next
            End If
            Exit Sub
        Case Else
'            Call MsgBoxST("�e�[�u�����w�肵�ĉ�����")
            Exit Sub
        End Select
    
        'ALT���������ȊO�̏ꍇ�͒��o�����w��
        If Not GetKeyState(VK_MENU) < 0 Then
            Call MakeQpsSheet(sTblName)
            Exit Sub
        End If
        
        nMinRow = ActiveCell.Row
        nMinCol = ActiveCell.Column
    
        sSQL = GetTemplateSql(sTblName, ccTPL_SEL, "noselect, nowhere")
        sSQL = Replace(sSQL, "SELECT ", "SELECT ROWID MAINTEN, " & sTblName & ".")
    End Select
    
    '�N�G���̎��s
    If Not ExecQuery(sConStr, sSQL, "MAINTEN", oRs, oFlds, oOpts) Then
        Exit Sub
    End If
        
    '�N�G�����ʃV�[�g�̍쐬
    With oFlds
        Set oFld = New Prop
        oFld.Item("Title") = ""
        oFld.Item("�����") = "�󔒗�"
        Set .Item(1, False) = oFld
        'ROWID��
        .Item(1).Item("��") = 0
        '�ǉ�/�X�V/�폜�}�[�N��
        .Item(2).Item("��") = 2.5
        .Item(2).Item("���ʒu") = "����"
    End With
    With oOpts
        .Item("�^�C�g��") = "�t�B���^"
        .Item("�g�Œ�Z��") = "C2"
        .Item("�s�^�C�g��") = "$1:$1"
        .Item("��^�C�g��") = "$A:$B"
        .Item("�X�^�C��") = "�r���`��A���ڌ^�F�t"
        Select Case GetCfgVal("�Ȗ͗l�\��")
        Case 1, 2 '�����t�������ŕ\���A�Z���w�i�F�ŕ\��
            .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�Ȗ͗l����"
        End Select
        Select Case GetCfgVal("�󔒃Z������")
        Case 1 '�����t�������ŋ���
            .Item("�X�^�C��") = .Item("�X�^�C��") & "�A�󔒃Z������"
        End Select
    End With
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If ActiveSheet.Range("A1") = "QPARAMS" Then
        Select Case GetCfgVal("���o�����V�[�g")
        Case 1 '���o���ɍ폜
            ActiveSheet.Delete
        Case 2 '���o���ɔ�\��
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
    Call SetCmtVal(ActiveSheet.Range("B1"), , "I:�ǉ� U:�X�V D:�폜")
    Application.StatusBar = GetConDisp(sConStr)
    '�L�[���ڂ̑����\���A�f�[�^�^�C�v�����F
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

'���o�����V�[�g����
Private Function MakeQpsSheet(argObjName)
    Dim sSQL As String, sShtName As String, sConStr As String
    Dim oRs, oFlds, oOpts As New Prop
    Dim nRow As Long
    Dim nMaxRow As Long
    sSQL = GetLibSql("���o�����V�[�g����")
    
    sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", argObjName)
    
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
            "���ڍs�̈ʒu���N�G���̗�ʒu�ɂȂ�܂��B��\���s�͕\��/�\�[�g/�������S�Ė���" & vbLf & _
            "��s��}�����ď����݂̂���͂���ƒ��O�̍��ڂ̏����Ƃ�OR�����ƂȂ�܂�")
        Call SetCmtVal(.Range("B2"), , _
            "���ږ��������̓L�[���ځA�����F������VARCHAR, CHAR�A����DATE�A�͂���ȊO")
        Call SetCmtVal(.Range("C2"), , _
            "���J�����R�����g��\���A���ږ��̃I�v�V������ON�ɂ���ƌ��ʂ̗�^�C�g�������̂ɂȂ�܂�")
        Call SetCmtVal(.Range("F1"), , _
            "�@������͑O����v�A���Z�q��<>!=���g�p�\" & vbLf & _
            "���J���}�A���s��؂�̕������IN�W�J����܂�")
        Call SetCmtVal(.Range("F2"), , _
            "��NULL�w��͐擪��IS�������Ďw�肵�܂�")
        Call SetCmtVal(.Range("G1"), , _
            "�@�ŏ���WHERE,AND��������SQL�����L�q" & vbLf & _
            "���擪 - - �ŃR�����g�A�E�g����܂�")
        With .Range("B1")
            .Value = "���ڏ����w��"
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
            .Value = "�ǉ�SQL���w��"
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
            .Value = "���b�N"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("J1")
            .Value = "���ږ���"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With .Range("K1")
            .Value = "SQL�\��"
            .Font.Bold = True
            .Interior.ColorIndex = ccGREY
        End With
        With oFlds
            .Item(1).Item("��") = 0
            .Item(2).Item("��") = 3
            .Item(3).Item("��") = 20
            .Item(4).Item("��") = 20
            .Item(5).Item("��") = 5
            .Item(5).Item("���ʒu") = "����"
            .Item(6).Item("��") = 5
            .Item(6).Item("���ʒu") = "����"
            .Item(7).Item("��") = 23
            .Item(8).Item("��") = 55
        End With
        With oOpts
            .Item("�^�C�g��") = "-"
            .Item("�X�^�C��") = "�r���`��"
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
            xlBetween, Formula1:="����,�~��"
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

'�e���v���[�g����
Sub ST_Template()
    Dim sSQL As String, sTpl As String
    If GetKeyState(VK_MENU) < 0 Then
        '�e���v���[�g����
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
                If GetCfgVal("SQL�̒��o, ���`, ����") = 2 Then
                    Call SetSql(sTpl, ActiveCell.Offset(1, 0))
                Else
                    If ST_WIN.ShowModal2(sTpl, "�W�J") Then
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
            '���擾
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
                'SQL���̐��`
                sTpl = SqlFmt(sSQL)
            Else
                'SQL�����R�[�h�ϊ�/�t�ϊ�
                
'                If InStr(sSQL, """") <> 0 And InStr(sSQL, "&") <> 0 Then
'                    '�ϊ�
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
                    '�ϊ�
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
                    '�t�ϊ�
                    
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
                If GetCfgVal("SQL�̒��o, ���`, ����") = 2 Then
                    Call SetSql(sTpl, ActiveSheet.Cells(nMinRow, nCol))
                Else
                    If ST_WIN.ShowModal2(sTpl, "�W�J") Then
                        Call SetSql(sTpl, ActiveSheet.Cells(nMinRow, nCol))
                    End If
                End If
            End If
        End If
    End If
End Sub

'SQL*Plus�̋N��
Sub ST_Plus()
    Dim sConStr As String, nPID As Long, hwnd As Long, sSQL As String
    Dim sCmdPara As String, sStaFile As String, nFn As Integer
    sConStr = GetConStr
    sSQL = GetSql
    
    'ALT���������̏ꍇ�� AUTOTRACE�ݒ�
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
    
    'SQL*Plus�̋N��
    nPID = Shell("sqlplusw " & sConStr & sCmdPara, vbNormalFocus)
    
    '�E�C���h�E�^�C�g���̕ύX
    hwnd = FindWindow(vbNullString, "Oracle SQL*Plus")
    If hwnd <> 0 Then
        Call SetWindowText(hwnd, GetConDisp(sConStr))
    End If

    '�Z���ʒu�̂r�p�k�𑗐M
    Application.Wait DateAdd("s", 1.5, Now) '�E�F�C�g�b���̒����v
    If sCmdPara = "" And ActiveSheet.Range("A1") <> "QPARAMS" Then
        SendKeys sSQL
    End If
End Sub

'���C�u�������j���[�\��
Sub ST_Library(Optional argMode = "")
    
    If GetKeyState(VK_SHIFT) < 0 And GetKeyState(VK_MENU) < 0 Then
        With GetLibSheet
            If .Visible = False Then
                .Cells.AutoFilter Field:=2, Criteria1:="<>*SQLTOOL�����g�p*", Operator:=xlAnd
                .Visible = True
                .Parent.Activate
                .Select
            Else
                .Visible = False
            End If
        End With
    Else
        If argMode = "URL" Then
            cnLibMode = 5 '���C�ɓ���t�q�k���j���[
        Else
            If GetKeyState(VK_MENU) < 0 Then
                cnLibMode = 1 '���C�u����SQL���N�G�����s
            ElseIf GetKeyState(VK_SHIFT) < 0 Then
                If ActiveSheet.Name = ccLibShtName Then
                    Exit Sub
                Else
                    cnLibMode = 2 '�Z���ʒu�Ƀ��C�u����SQL��}��
                End If
            Else
                If ActiveSheet.Name = ccLibShtName Then
                    cnLibMode = 4 '���C�ɓ���u�b�N�I�����j���[
                Else
                    cnLibMode = 3 '���C�ɓ���u�b�N���j���[
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
        Case 1 '�N�G�����s
            sSQL = GetLibSql(.Parameter)
            '&�u���ϐ�������ꍇ�͓��̓v�����v�g�\��
            If GetSqlParams(sSQL, oPrms) Then
                For Each sPrm In oPrms.Keys
                    sInp = InputBox(sPrm, ccAppTitle)
                    sSQL = SetSqlParam(sSQL, sPrm, sInp)
                Next
            End If
            Call ST_Query(sSQL, .Parameter)
        Case 2 'SQL�}��
            sSQL = GetLibSql(.Parameter)
            ActiveCell.Value = sSQL
        Case 3 '�u�b�N�I�[�v��
            On Error Resume Next
            Call Workbooks.Open(GetLibSql(.Parameter))
            On Error GoTo 0
        Case 4 '�u�b�N�p�X�ݒ�
            Cells(ActiveCell.Row, 3) = Left(.Parameter, Len(.Parameter) - 4)
            Cells(ActiveCell.Row, 4) = "B"
            Cells(ActiveCell.Row, 5) = Workbooks(.Parameter).FullName
        Case 5 'URL�u���E�U�\��
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
        'WfOU�N��
        Dim oWfOU
        Set oWfOU = GetCfgSheet.Parent.Worksheets(ccWfOUShtName)
        Application.ScreenUpdating = False
        Set oWfOU = GetNewSheet(oWfOU.Name, True, False, oWfOU)
        oWfOU.Visible = True
        Application.ScreenUpdating = True
    ElseIf GetKeyState(VK_MENU) < 0 Then
    Else
        'URL���j���[�\��
        Call ST_Library("URL")
    End If
    
End Sub

'�ڑ�������̎擾
Private Function GetConStr() As String
    Dim i As Integer
    Dim arLines, oCfgSht
    Set oCfgSht = GetCfgSheet
    With ActiveSheet
        Select Case .Range("A1")
        Case "QPARAMS" '���o��������
            GetConStr = GetCmtVal(ActiveSheet.Range("A1"))
        Case "MAINTEN" '�����e�i���X
            arLines = Split(GetCmtVal(ActiveSheet.Range("A1")), vbLf)
            GetConStr = arLines(0)
        Case Else
            Dim sCmt As String
            sCmt = GetCmtVal(.Range("A1"))
            If InStr(UCase(sCmt), "SELECT") <> 0 Then
                arLines = Split(sCmt, vbLf)
                GetConStr = Replace(arLines(0), "�ڑ���: ", "")
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

'�V�F���R�}���h�̎擾
Private Function GetShCmd() As String
    
    '��������肠���������
    GetShCmd = GetCfgSheet.Range("A27")
    
End Function

'�f�t�H���g���s���[�h�̎擾
Private Function GetDefExecMode() As Integer
    
    '��肠���������
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

'�J�����g�f�B���N�g���̎擾
Private Function GetCurDir() As String
    
    '��肠���������
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
    nConMode = GetCfgVal("�ڑ����[�h")
    nCmtMode = GetCfgVal("�R�~�b�g���[�h")
    
    '�ڑ�������̕���
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
    Case "1" '�����R�~�b�g
        '����A�ڑ��I�u�W�F�N�g���쐬
    Case "2" '���j���[�Ő���
        '�ڑ��L�[�Ō������A���݂���΃Z�b�g
        With coConObjs
            If .Exists(nConMode & argConStr) Then
                Set oConItem = .Item(nConMode & argConStr)
                If argIsUpd And InStr(oConItem.Item("Disp"), "*") = 0 Then
                    oConItem.Item("Disp") = oConItem.Item("Disp") & "*"
                End If
                Set oCon = oConItem.Item("Obj")
            End If
        End With
    Case "3" '�ڑ��T�[�o�g�p
        '�ڑ��T�[�o�ɈϏ�
    End Select
    
    '�ڑ����������
    If oCon Is Nothing Then
        '�V���ɐ���
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
        '�����R�~�b�g�ȊO�Ȃ�
        If nCmtMode <> 1 Then
            '�ێ����X�g�ɓo�^
            Set oConItem = New Prop
            If argIsUpd Then
                oConItem.Item("Disp") = GetConDisp(argConStr) & "*"
            Else
                oConItem.Item("Disp") = GetConDisp(argConStr)
            End If
            Set oConItem.Item("Obj") = oCon
            Set coConObjs.Item(nConMode & argConStr) = oConItem
            '�g�����U�N�V�����J�n
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

'�ݒ�V�[�g�̎擾
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

'���C�u�����V�[�g�̎擾
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

'�ݒ荀�ڂ̎擾
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

'�r�p�k���̎擾
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
        '���C�u�����V�[�g
        argSQLType = "SELECT"
        sSQL = ActiveSheet.Cells(nRow, nCol)
    ElseIf ActiveSheet.Range("A1") = "QPARAMS" Then
        '���o�����V�[�g
        sSQL = GetSqlQpsSheet(ActiveSheet, argSQLType, argSqlTitle)
    Else
        If Selection.Areas.Count = 1 And Selection.Rows.Count = 1 And Selection.Columns.Count = 1 Then
            '�����݈ʒu�����SQL�擾�͉��P�̗]�n����
            If sVAL <> "" And sVAL <> "No." Then
                '���J�n�ʒu�擾
                Do Until nRow = 1 Or ActiveSheet.Cells(nRow, nCol) = ""
                    nRow = nRow - 1
                Loop
                If ActiveSheet.Cells(nRow, nCol) = "" Then
                    nRow = nRow + 1
                End If
                'SQL���J�n�L�[���[�h����
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
                '�^�C�g���擾
                If Not IsMissing(argSqlTitle) And nRow - 1 > 0 Then
                    If argSQLType <> "" And ActiveSheet.Cells(nRow - 1, nCol) <> "" Then
                        argSqlTitle = ActiveSheet.Cells(nRow - 1, nCol)
                    End If
                End If
                'SQL���擾
                Do Until ActiveSheet.Cells(nRow, nCol) = ""
                    sSQL = sSQL & ActiveSheet.Cells(nRow, nCol) & vbCrLf
                    nRow = nRow + 1
                Loop
            Else
                '�N�G�����ʕ\���V�[�g�̍Ē��o
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
            'SQL�����擾�ł��Ȃ��ꍇ�A�\�����\�ȃI�u�W�F�N�g���Ȃ璊�o�����w��
            If sSQL = "" Then
'                sSQL = GetTemplateSql(sVAL, ccTPL_SEL, "nowhere")
'                argSQLType = IIf(sSQL <> "", "SELECT", "")
                Dim sObjType
                sObjType = GetDBObjectType(sVAL)
                Select Case sObjType
                Case "TABLE", "VIEW"
                    argSQLType = "TABLE"
                    '��
                    argSqlTitle = sVAL
                    sSQL = GetTemplateSql(sVAL, ccTPL_SEL, "nowhere")
                Case "FUNCTION", "PROCEDURE"
                    argSQLType = sObjType
                    argSqlTitle = sVAL
                    sSQL = sVAL
                End Select
            End If
        Else
            '�I��͈͂�SQL�����擾
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
            
                'SELECT���O�� ( ������ꍇ�͍폜
                Dim nStart As Integer, nSCnt As Integer, nEnd As Integer, nECnt As Integer, i As Integer
                nStart = InStr(UCase(sSQL), "SELECT")
                If nStart > 0 Then
                    '�J�n/�I�� ���ʂ̐����J�E���g
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
                    '�I�����ʂ�����Ȃ������J�n���ɉ��Z
                    nSCnt = nECnt - nSCnt
                    nECnt = 0
                    'SELECT�O�̊��ʐ����J�E���g
                    For i = 1 To nStart - 1
                        Select Case Mid(sSQL, i, 1)
                        Case "("
                            nSCnt = nSCnt + 1
                        Case ")"
                            nSCnt = nSCnt - 1
                        End Select
                    Next
                    'SQL��������̊��ʐ�����v����܂ō폜
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
    'SQL�����̉��s�A;���폜
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
    '&�u���ϐ�������ꍇ�͓��̓v�����v�g�\��
    Dim oPrms, sPrm, sInp
    If GetSqlParams(sSQL, oPrms) Then
        For Each sPrm In oPrms.Keys
            sInp = InputBox(sPrm, ccAppTitle)
            sSQL = SetSqlParam(sSQL, sPrm, sInp)
        Next
    End If
    GetSql = sSQL
End Function

'���o�����w��V�[�g����r�p�k���𐶐�
Function GetSqlQpsSheet(argSht As Worksheet, argSQLType, argSqlTitle)
    Dim nRow As Long, nMaxRow As Long
    Dim sSQL As String, sTmpStr As String, sFldType As String, sFldName As String
    Dim sSelFlds As String, sWheFlds As String, sAddWhes As String, sOrdFlds As String
    Dim nOpr As Integer, sOpr As String, sExp As String, sDat As String, arExp
    Dim i As Integer, bCmtFldName As Boolean, sTA As String, sFA As String, sTblName As String
    With argSht
        '��ʒu 1:�f�[�^�^ 3:���ږ� 4:���ږ��� 5:�\�� 6:�\�[�g 7:�����l 8:�ǉ�SQL��
        sTblName = Replace(ActiveSheet.Name, ccQpsShtPref, "")
        '��
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
        'SELECT�w��t�B�[���h
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
        '�B���ǉ�SELECT�w��t�B�[���h
        If .Cells(2, 8) <> "" Then
            sSQL = sSQL & "  " & .Cells(2, 8) & ", " & vbCrLf
        End If
        sSQL = sSQL & sSelFlds & vbCrLf
        sSQL = sSQL & "FROM " & sTblName & sTA & vbCrLf
        'WHERE�w��t�B�[���h
        Const cOP_NORM = 0
        Const cOP_NULL = 1
        Const cOP_IN = 2
        Const cOP_OP = 3
        sTmpStr = ""
        sFldType = .Cells(3, 1)
        sFldName = .Cells(3, 3)
        For nRow = 3 To nMaxRow
            '���ڌ^/���ږ���ޔ�
            If .Cells(nRow, 3) <> "" Then
                sFldType = .Cells(nRow, 1)
                sFldName = .Cells(nRow, 3)
            End If
            If .Cells(nRow, 7) <> "" And .Rows(nRow).RowHeight > 0 Then
                '���ږ�����̏ꍇ�͒��O�̍��ڂ�OR����
                If .Cells(nRow, 3) = "" Then
                    If sTmpStr <> "" Then
                        sTmpStr = sTmpStr & " OR "
                    End If
                Else
                    'OR�Ōq������O�̏����������WHERE��ɒǉ�
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
                    '���������Z�b�g
                    sTmpStr = ""
                End If
                '�ϐ��ɒ��o�������Z�b�g
                sExp = .Cells(nRow, 7)
                '���Z�q�^�C�v�̔���A�����l�̐؂�o��
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
                '�f�[�^�^�ɉ����ď��������쐬
                Select Case sFldType
                Case "VARCHAR2", "CHAR"
                    sTmpStr = sTmpStr & sFA & sFldName
                    
                    Select Case nOpr
                    Case cOP_NORM
                        sTmpStr = sTmpStr & " LIKE '" & sExp & "%'"
                    Case cOP_NULL
                        sTmpStr = sTmpStr & " IS " & sExp
                    Case cOP_IN
                        '���C���h�J�[�h�w��̏ꍇ�� LIKE OR�W�J
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
                Case "DATE" '�� ADO�Ή��v
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
        'WHERE�w��ǉ�SQL��
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
        'ORDER BY�w��t�B�[���h
        sTmpStr = ""
        For nRow = 3 To nMaxRow
            If (.Cells(nRow, 6) = "����" Or .Cells(nRow, 6) = "�~��") And .Rows(nRow).RowHeight > 0 Then
                If sTmpStr <> "" Or sOrdFlds <> "" Then
                    sTmpStr = sTmpStr & ", "
                End If
                If Len(sTmpStr) > ccMaxSqlWidth Then
                    sOrdFlds = sOrdFlds & "  " & sTmpStr & vbCrLf
                    sTmpStr = ""
                End If
                sTmpStr = sTmpStr & sFA & .Cells(nRow, 3) & _
                    IIf(.Cells(nRow, 6) = "�~��", " DESC", "")
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
            If Not ST_WIN.ShowModal2(sSQL, "���o") Then
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
    sMaxSqlWidth = GetCfgVal("SQL�̌��܂蕝")
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
            'SELECT�w��t�B�[���h
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
                'WHERE�w��t�B�[���h
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
                            Case "DATE" '�� ADO�Ή��v
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
                'ORDER BY�w��t�B�[���h
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
                Case "DATE" '�� ADO�Ή��v
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
                Case "DATE" '�� ADO�Ή��v
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
                    Case "DATE" '�� ADO�Ή��v
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
                    Case "DATE" '�� ADO�Ή��v
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
                '�e�[�u�����ڂ̎擾
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
                
                '�L�[���ڂ̎擾
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

'�r�p�k�̎��s
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
        If Not ST_WIN.ShowModal2(argSql, "���s") Then
            GoTo END_FUNC
        End If
    End If
    
    Select Case GetCfgVal("�ڑ����[�h")
    Case "1" 'OO4O/SQL*Net
    
        Set oCon = GetConObj(argConStr, IsMissing(argRs))
        
        If oCon Is Nothing Then
          Call MsgBoxST("�ڑ��Ɏ��s���܂���" & vbCrLf & vbCrLf & argConStr)
          GoTo END_FUNC
        End If
        
        '�p�����[�^�̐ݒ�
        If Not IsMissing(argOpts) Then
        If Not IsEmpty(argOpts) Then
        If TypeName(argOpts.Item("�p�����[�^")) = "Prop" Then
            Dim sPrmName As String, sIOType As String, nIoType As Integer, nSvType As Integer
            For Each oPrm In argOpts.Item("�p�����[�^").Items
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
                    Call MsgBoxST(sPrmName & " ���Ή��̓��̓p�����[�^�ł��B")
                    Exit Function
                End Select
                oCon.Parameters.Add sPrmName, oPrm.Item("Value"), nIoType, nSvType
            Next
            If argOpts.Item("�߂�l") <> "" Then
                Select Case argOpts.Item("�߂�l")
                Case "VARCHAR2"
                    nSvType = 1 'ORATYPE_VARCHAR2
                Case "NUMBER"
                    nSvType = 2 'ORATYPE_NUMBER
                Case "CHAR"
                    nSvType = 96 'ORATYPE_CHAR
                Case Else
                    Call MsgBoxST(" ���Ή��̖߂�l�ł��B")
                    Exit Function
                End Select
                oCon.Parameters.Add "RESULT", "", 2, nSvType
            End If
        End If
        End If
        End If
        
        nST = Timer()

        '�������[�h�ł̃N�G�����s
        If Not IsMissing(argRs) Then
            '���R�[�h�Z�b�g�̎擾
            Set argRs = oCon.CreateDynaset(argSql, &H2 + &H4)
            'ORADYN_NO_BLANKSTRIP, ORADYN_READONLY
        Else
            'DML/DDL�̎��s
'            nRecCnt = oCon.ExecuteSQL(Replace(argSql, vbCr, "")) '��
            nRecCnt = oCon.ExecuteSQL(argSql)
        End If
        
'    On Error GoTo KEYBREAK
'
'        '�񓯊����[�h�ł̃N�G�����s
'        If Not IsMissing(argRs) Then
'            '���R�[�h�Z�b�g�̎擾
'            oCon.Parameters.Add "C", 0, 2, 102    'ORAPARM_OUTPUT, ORATYPE_CURSOR
'            oCon.Parameters.Add "S", argSql, 1, 1 'ORAPARM_INPUT , ORATYPE_VARCHAR2
'            Set oStmt = oCon.CreateSQL("DECLARE BEGIN OPEN :C FOR :S; END;", &H4) 'ORASQL_NONBLK
'        Else
'            'DML/DDL�̎��s
''            nRecCnt = oCon.ExecuteSQL(argSql)
'            Set oStmt = oCon.CreateSQL(argSql, &H4) 'ORASQL_NONBLK
'        End If
'        '�񓯊����s�����N�G���̏I���҂����[�v
'        bKeyBreak = False
'        Do While oStmt.NonBlockingState = -3123 'ORASQL_STILL_EXECUTING
'            '���f�L�[���������m
'            If bKeyBreak Then
'                If MsgBoxST("���f���܂����H", vbYesNo) = vbYes Then
'                    oStmt.Cancel
'                    Exit Do
'                End If
'                bKeyBreak = False
'            End If
'        Loop
'        '���s�����㏈��
'        If Not IsMissing(argRs) Then
'            '���R�[�h�Z�b�g�̏ꍇ�̓p�����[�^�̃J�[�\���ϐ����擾
'            Set argRs = oCon.Parameters("C").Value
'            oCon.Parameters.Remove "C"
'            oCon.Parameters.Remove "S"
'        Else
'            'DML/DDL�̏ꍇ�͉e�����R�[�h�������擾
'            nRecCnt = oStmt.RecordCount
'        End If
        
        nET = Timer()
        
        '�p�����[�^�̌��ʎ擾�A�폜
        If Not IsMissing(argOpts) Then
        If Not IsEmpty(argOpts) Then
        If TypeName(argOpts.Item("�p�����[�^")) = "Prop" Then
            For Each oPrm In argOpts.Item("�p�����[�^").Items
                If InStr(oPrm.Item("InOut"), "OUT") <> 0 Then
                    oPrm.Item("Value") = oCon.Parameters(oPrm.Item("Name")).Value
                End If
                oCon.Parameters.Remove oPrm.Item("Name")
            Next
            If argOpts.Item("�߂�l") <> "" Then
                argOpts.Item("�߂�l") = oCon.Parameters("RESULT").Value
                oCon.Parameters.Remove "RESULT"
            End If
        End If
        End If
        End If
        
        If (Err <> 0 Or oCon.LastServerErr <> 0) And oCon.LastServerErr <> 24344 Then '��
          Call MsgBoxST(Err.Number & ":" & Err.Description & vbCrLf & oCon.LastServerErrText)
          GoTo END_FUNC
        End If
        
    Case "2" 'ADO/ODBC
        
        Set oCon = GetConObj(argConStr, IsMissing(argRs))
        
        If oCon.State <> 1 Then 'adStateOpen
          Call MsgBoxST("�ڑ��Ɏ��s���܂���" & vbCrLf & vbCrLf & argConStr)
          GoTo END_FUNC
        End If
        
        nST = Timer()
        If Not IsMissing(argRs) Then
            Set argRs = oCon.Execute(argSql)
        Else
            'DML/DDL�̎��s
            Call oCon.Execute(argSql, nRecCnt) 'adCmdText���w�肵���������������H
        End If
        nET = Timer()
        
        If Err <> 0 Then
          Call MsgBoxST(Err.Number & ":" & Err.Description)
          GoTo END_FUNC
        End If
    
    End Select
    
    Application.EnableCancelKey = xlDisabled
    If Not IsMissing(argFlds) Then
        '�t�B�[���h���̃Z�b�g
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
        '�I�v�V�������̃Z�b�g
        If IsEmpty(argOpts) Then
            Set argOpts = New Prop
        End If
        argOpts.Item("���s") = TimeElapsed(nST, nET)
        argOpts.Item("�X�V����") = nRecCnt
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

'ExSQell�b�胍�W�b�N
'�V�F���R�}���h���s
Private Function ExecShell(argCmd, argRs, argFlds)
    Dim sCmd, sWrk, WRK, sLine, i, oFld, sTmp, sShTmp, TMP, sCurDir
    
    If IsEmpty(coWSH) Then
        Set coWSH = CreateObject("WScript.Shell")
    End If
    
    If IsEmpty(coFSO) Then
        Set coFSO = CreateObject("Scripting.FileSystemObject")
    End If
    
    sWrk = ActiveWorkbook.Path & "\ExSQell.wrk"
    
    '����ƃV�F���R�}���h�ύX����uname�����s���Ď��s���𔻕�
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
    
    '���s���ɉ����āA���_�C���N�g����e���|�����t�@�C���p�X���Z�o
    sTmp = coFSO.BuildPath(coFSO.GetSpecialFolder(2), coFSO.GetTempName)
    sShTmp = LCase(Left(sTmp, 1)) & "/" & Replace(Mid(sTmp, 4), "\", "/")
    If InStr(csUname, "CYGWIN") > 0 Then
        sShTmp = "/cygdrive/" & sShTmp
    ElseIf InStr(csUname, "MINGW") > 0 Then
        sShTmp = "/" & sShTmp
    Else 'Bash on Ubuntu on Windows
        sShTmp = "/mnt/" & sShTmp
    End If
    
    '�J�����g�f�B���N�g���ړ��R�}���h�ݒ�
    sCurDir = GetCurDir
    If sCurDir <> "" Then
        sCurDir = "cd " & sCurDir & "; "
    End If
    
    '�V�F���R�}���h�̎��s
    sCmd = csShCmd & " """ & sCurDir & Replace(argCmd, """", """""") & " > " & sShTmp & """"
    Call coWSH.Run(sCmd, vbHide, True)
    
    '���[�J�����R�[�h�Z�b�g�쐬
    Set argRs = CreateObject("ADODB.Recordset")
    
    
    '�e��ݒ���̎擾
    Dim nOutCnv, sCharset, nLineSeparator, nFirstRowTitle, nFieldSep, sFieldName, arFields, nFieldCnt, nArrayCnt
    nOutCnv = GetCfgVal("�o�͕����R�[�h�ϊ�")
    nFirstRowTitle = GetCfgVal("1�s�ڂ��^�C�g���s")
    nFieldSep = GetCfgVal("�t�B�[���h��؂�")
    
    '�e���|�����t�@�C���̓ǂݍ��ݕ����ɉ����đO����
    Select Case nOutCnv
    Case 2 To 5 '�uUTF-8 �� SJIS�v�`�uSJIS  �� SJIS�v
        Select Case nOutCnv
        Case 2 'UTF-8 �� SJIS
            sCharset = "utf-8"
            nLineSeparator = 10 'adLF
        Case 3 'EUC   �� SJIS
            sCharset = "euc-jp"
            nLineSeparator = 10 'adLF
        Case 4 'JIS   �� SJIS
            sCharset = "iso-2022-jp"
            nLineSeparator = 10 'adLF
        Case 5 'SJIS  �� SJIS
            sCharset = "shift_jis"
            nLineSeparator = -1 'adCRLF
        End Select
        
        '�e���|�����t�@�C����ǂݍ���Ń��[�J�����R�[�h�Z�b�g�ɒǉ�
        Set TMP = CreateObject("ADODB.Stream")
        With TMP
            .Charset = sCharset
            .LineSeparator = nLineSeparator
            .Open
            Call .LoadFromFile(sTmp)
            
            '1�s�ڂ̓ǂݍ���
            If Not .EOS Then
                sLine = .Readtext(-2) 'adReadLine
                'Debug.Print sLine
                
                '�t�B�[���h��؂�
                If nFieldSep = 1 Then '�Ȃ�
                    '1�s�ڂ��^�C�g���s
                    If nFirstRowTitle = 1 Then '�Ƃ��Ȃ�
                        sFieldName = "Line"
                    Else
                        sFieldName = sLine
                    End If
                    argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, �ő包��
                    nFieldCnt = 1
                Else
                    '�s���t�B�[���h�ɕ���
                    Select Case nFieldSep
                    Case 2 '�A�������󔒕���
                        arFields = SplitSh(sLine)
                    Case 3 '�J���}
                        arFields = SplitCsv(sLine)
                    Case 4 '�^�u����
                        arFields = Split(sLine, " ")
                    End Select
                    
                    '�t�B�[���h���ɉ����ă��[�J�����R�[�h�Z�b�g�Ƀt�B�[���h�ǉ�
                    For i = 0 To UBound(arFields)
                        '1�s�ڂ��^�C�g���s
                        If nFirstRowTitle = 1 Then '�Ƃ��Ȃ�
                            sFieldName = "Filed" & (i + 1)
                        Else
                            sFieldName = arFields(i)
                        End If
                        argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, �ő包��
                    Next
                    nFieldCnt = argRs.Fields.Count
                End If
                
                '���[�J�����R�[�h�Z�b�g���I�[�v��
                argRs.Open
                
                '1�s�ڂ��^�C�g���s�Ƃ��Ȃ��ꍇ
                If nFirstRowTitle = 1 Then
                    '1�s�ڂ̃��R�[�h�ǉ�
                    argRs.AddNew
                    
                    '�t�B�[���h��؂�
                    If nFieldSep = 1 Then '�Ȃ�
                        argRs(0) = sLine
                    Else
                        '�񐔂ɉ����ăt�B�[���h�l�Z�b�g
                        For i = 0 To nFieldCnt - 1
                            argRs(i) = arFields(i)
                        Next
                    End If
                    
                    '1�s�ڂ̃��R�[�h�o�^
                    argRs.Update
                End If
               
            Else '���ʃf�[�^���R�[�h����
                '�t�B�[���h��؂�
                If nFieldSep = 1 Then '�Ȃ�
                    argRs.Fields.Append "Line", 202 'adVarWChar
                Else
                    argRs.Fields.Append "Field1", 202 'adVarWChar
                End If
                argRs.Open
            End If
            
            '2�s�ڈȍ~�̓ǂݍ���
            Do Until .EOS
                sLine = .Readtext(-2) 'adReadLine
                'Debug.Print sLine
                
                argRs.AddNew
                
                '�t�B�[���h��؂�
                If nFieldSep = 1 Then '�Ȃ�
                    argRs(0) = sLine
                Else
                    '�s���t�B�[���h�ɕ���
                    Select Case nFieldSep
                    Case 2 '�A�������󔒕���
                        arFields = SplitSh(sLine)
                    Case 3 '�J���}
                        arFields = SplitCsv(sLine)
                    Case 4 '�^�u����
                        arFields = Split(sLine, " ")
                    End Select
                    
                    nArrayCnt = UBound(arFields) + 1
                    
                    '�񐔂ɉ����ăt�B�[���h�l�Z�b�g
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
    Case Else '1 �u�����ϊ����Ȃ��v
        '�e���|�����t�@�C����ǂݍ���Ń��[�J�����R�[�h�Z�b�g�ɒǉ�
        Set TMP = coFSO.OpenTextFile(sTmp, 1, False)

        '1�s�ڂ̓ǂݍ���
        If Not TMP.AtEndOfStream Then
            sLine = TMP.ReadLine
            'Debug.Print sLine
            
            '�t�B�[���h��؂�
            If nFieldSep = 1 Then '�Ȃ�
                '1�s�ڂ��^�C�g���s
                If nFirstRowTitle = 1 Then '�Ƃ��Ȃ�
                    sFieldName = "Line"
                Else
                    sFieldName = sLine
                End If
                argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, �ő包��
                nFieldCnt = 1
            Else
                '�s���t�B�[���h�ɕ���
                Select Case nFieldSep
                Case 2 '�A�������󔒕���
                    arFields = SplitSh(sLine)
                Case 3 '�J���}
                    arFields = SplitCsv(sLine)
                Case 4 '�^�u����
                    arFields = Split(sLine, " ")
                End Select
                
                '�t�B�[���h���ɉ����ă��[�J�����R�[�h�Z�b�g�Ƀt�B�[���h�ǉ�
                For i = 0 To UBound(arFields)
                    '1�s�ڂ��^�C�g���s
                    If nFirstRowTitle = 1 Then '�Ƃ��Ȃ�
                        sFieldName = "Filed" & (i + 1)
                    Else
                        sFieldName = arFields(i)
                    End If
                    argRs.Fields.Append sFieldName, 202, ccMaxShFldLen 'adVarWChar, �ő包��
                Next
                nFieldCnt = argRs.Fields.Count
            End If
            
            '���[�J�����R�[�h�Z�b�g���I�[�v��
            argRs.Open
            
            '1�s�ڂ��^�C�g���s�Ƃ��Ȃ��ꍇ
            If nFirstRowTitle = 1 Then
                '1�s�ڂ̃��R�[�h�ǉ�
                argRs.AddNew
                
                '�t�B�[���h��؂�
                If nFieldSep = 1 Then '�Ȃ�
                    argRs(0) = sLine
                Else
                    '�񐔂ɉ����ăt�B�[���h�l�Z�b�g
                    For i = 0 To nFieldCnt - 1
                        argRs(i) = arFields(i)
                    Next
                End If
                
                '1�s�ڂ̃��R�[�h�o�^
                argRs.Update
            End If
           
        Else '���ʃf�[�^���R�[�h����
            '�t�B�[���h��؂�
            If nFieldSep = 1 Then '�Ȃ�
                argRs.Fields.Append "Line", 202 'adVarWChar
            Else
                argRs.Fields.Append "Field1", 202 'adVarWChar
            End If
            argRs.Open
        End If
        
        '2�s�ڈȍ~�̓ǂݍ���
        Do Until TMP.AtEndOfStream
            sLine = TMP.ReadLine
            'Debug.Print sLine
            
            argRs.AddNew
            
            '�t�B�[���h��؂�
            If nFieldSep = 1 Then '�Ȃ�
                argRs(0) = sLine
            Else
                '�s���t�B�[���h�ɕ���
                Select Case nFieldSep
                Case 2 '�A�������󔒕���
                    arFields = SplitSh(sLine)
                Case 3 '�J���}
                    arFields = SplitCsv(sLine)
                Case 4 '�^�u����
                    arFields = Split(sLine, " ")
                End Select
                
                nArrayCnt = UBound(arFields) + 1
                
                '�񐔂ɉ����ăt�B�[���h�l�Z�b�g
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
    
    '�t�B�[���h���̐ݒ�
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

'�R�}���h���s ��vbs�o�R�Ŕ�\���������s�A�W���o�͂����[�N�t�@�C���֏����o��
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

'�N�G���̎��s
Private Function ExecQuery(argConStr, argSql, argSQLType, argRs, argFlds, argOpts)
    Dim sCntSql As String, oCntRs
    Dim sNoOrdSql As String, i As Integer, nLastOrd As Integer, nSCnt As Integer, nECnt As Integer
    If GetCfgVal("�x�����o����") <> "" Then
        'ALT���������ȊO�A�e�[�u�����o�A�����e���[�h�̏ꍇ�͒��o�������m�F
        If Not GetKeyState(VK_MENU) < 0 Or argSQLType = "TABLE" Or argSQLType = "MAINTEN" Then
            '���o�������擾
            sCntSql = GetLibSql("�����擾")
            '���sSQL����ORDER BY���܂܂�Ă���ꍇ�͈ȍ~���폜
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
                '�Ή��Â��Ă��Ȃ����ʂ������ꍇ
                If nSCnt < nECnt Then
                    sNoOrdSql = argSql
                Else
                    '�����ꍇ��ORDER BY�ȍ~���폜
                    sNoOrdSql = Left(argSql, nLastOrd - 1)
                End If
            Else
                sNoOrdSql = argSql
            End If
            sCntSql = SetSqlParam(sCntSql, "�擾�Ώ�", "(" & sNoOrdSql & ")")
            Application.StatusBar = "�����m�F��..."
            If Not ExecSql(argConStr, sCntSql, oCntRs, argFlds, argOpts) Then
                Application.StatusBar = ""
                Exit Function
            End If
            Application.StatusBar = ""
            argOpts.Item("����") = oCntRs("CNT").Value
            'SHIFT���������̏ꍇ�͒��o�����_�C�A���O�\��
            If GetKeyState(VK_SHIFT) < 0 And oCntRs("CNT").Value > 0 Then
                Select Case MsgBoxST("���o������ " & oCntRs("CNT").Value & " ���ł��B" & vbCrLf & vbCrLf & _
                    "[�S���\����]�@�@[�T���v�����܂Ł�]", vbYesNoCancel + vbDefaultButton3)
                Case vbNo
                    argOpts.Item("�T���v�����o����") = GetCfgVal("�T���v�����o����")
                    argOpts.Item("����") = argOpts.Item("�T���v�����o����")
                Case vbCancel
                    Exit Function
                End Select
            End If
            '�x�����o�����ȏ�Ȃ烏�[�j���O
            If oCntRs("CNT").Value > CLng(GetCfgVal("�x�����o����")) And argOpts.Item("�T���v�����o����") = "" Then
                Select Case MsgBoxST("���o������ " & oCntRs("CNT").Value & " ���A�x�������ȏ�ł��B" & vbCrLf & vbCrLf & _
                    "[�S���\����]�@�@[�T���v�����܂Ł�]", vbYesNoCancel + vbDefaultButton3)
                Case vbNo
                    argOpts.Item("�T���v�����o����") = GetCfgVal("�T���v�����o����")
                    argOpts.Item("����") = argOpts.Item("�T���v�����o����")
                Case vbCancel
                    Exit Function
                End Select
            End If
        End If
    End If
        
        
    'DB�ڑ��ASQL���s
    Application.StatusBar = "SQL���s��..."
    If Not ExecSql(argConStr, argSql, argRs, argFlds, argOpts) Then
        Application.StatusBar = ""
        Exit Function
    End If
    Application.StatusBar = ""

    If argSQLType <> "MAINTEN" Then
        If argRs.EOF Then
            Call MsgBoxST("���o�������O���ł�")
            Exit Function
        End If
    End If

    ExecQuery = True
End Function

'�V�K�V�[�g�쐬
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
        '�����̃V�[�g���������ꍇ�́A�w��V�[�g��n �Ƀ��l�[��
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
    
    '�w��V�[�g���̑��݃`�F�b�N
    sNewShtName = ""
    On Error Resume Next
    sNewShtName = Sheets(sChkShtName).Name
    On Error GoTo 0
    If sNewShtName <> "" Then
        '�����ő��݂���ꍇ
        If argNewFlg Then
            Dim bFlg
            bFlg = Application.DisplayAlerts
            Application.DisplayAlerts = False
            Sheets(sChkShtName).Delete
            Application.DisplayAlerts = bFlg
            
            '�V�K�V�[�g�̍쐬
'            sCurShtName = ActiveSheet.Name
'            Select Case GetCfgVal("�V�K�V�[�g�ǉ��ʒu")
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
        '�����ő��݂��Ȃ��ꍇ�A�P���ɐV�K�ǉ�
'        sCurShtName = ActiveSheet.Name
'        Select Case GetCfgVal("�V�K�V�[�g�ǉ��ʒu")
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
        Select Case GetCfgVal("�V�K�V�[�g�ǉ��ʒu")
        Case 1
            Worksheets.Add before:=argCurSht
        Case 2
            Worksheets.Add after:=argCurSht
        Case Else
            Worksheets.Add after:=Worksheets(Worksheets.Count)
        End Select
    Else
        Select Case GetCfgVal("�V�K�V�[�g�ǉ��ʒu")
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

'�N�G�����s���ʂ�\�g�݂��ē\��t��
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
        MsgBoxST "MakeList: �J�n�ʒu�̎w��� Worksheet �� Range"
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
        
        '�^�C�g���s�̃Z�b�g
        If argOpts.Item("�^�C�g��") <> "" Then
            For nFldIdx = 1 To nFldCnt
                sNmFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("Title")
            Next
            With .Range(.Cells(nRow, nMinCol), .Cells(nRow, nMaxCol))
                .Interior.ColorIndex = ccGREY
                If InStr(argOpts.Item("�X�^�C��"), "�����񕝃f�[�^") = 0 Then
                    .Value = sNmFlds
                End If
                If InStr(argOpts.Item("�X�^�C��"), "������") = 0 Then
                    If InStr(argOpts.Item("�^�C�g��"), "�t�B���^") <> 0 Then
                        .AutoFilter
                    End If
                    .EntireColumn.AutoFit
                End If
            End With
            '�񕝁A���ʒu�w��
            For nFldIdx = 1 To nFldCnt
                If argFlds.Item(nFldIdx).Item("��") <> "" Then
                    .Columns(nMinCol + nFldIdx - 1).ColumnWidth _
                        = argFlds.Item(nFldIdx).Item("��")
                End If
                If argFlds.Item(nFldIdx).Item("���ʒu") <> "" Then
                    With .Columns(nMinCol + nFldIdx - 1)
                        Select Case argFlds.Item(nFldIdx).Item("���ʒu")
                        Case "����"
                            .HorizontalAlignment = xlCenter
                        Case "�E��"
                            .HorizontalAlignment = xlLeft
                        Case "����"
                            .HorizontalAlignment = xlRight
                        End Select
                    End With
                End If
            Next
            '�f�[�^�^�����F
            If InStr(argOpts.Item("�X�^�C��"), "���ڌ^�F�t") <> 0 Then
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
            '����^�C�g��
#If VBA7 Then
            Application.PrintCommunication = False
#End If
            With .PageSetup
                .PrintTitleRows = argOpts.Item("�s�^�C�g��")
                .PrintTitleColumns = argOpts.Item("��^�C�g��")
            End With
#If VBA7 Then
            Application.PrintCommunication = True
#End If
            '�g�Œ�
            If argOpts.Item("�g�Œ�Z��") <> "" Then
                .Range(argOpts.Item("�g�Œ�Z��")).Select
                ActiveWindow.FreezePanes = True
            End If
            nRow = nRow + 1
        End If
        
        '�������ׁ̈A�t�B�[���h���A�X�^�C������ϐ��ɕۑ�
        For nFldIdx = 1 To nFldCnt
            sSpFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("�����")
        Next
        '�K�蒊�o����
        If argOpts.Item("�T���v�����o����") <> "" Then
            nMaxCnt = argOpts.Item("�T���v�����o����")
        End If
        '�Ȗ͗l�Z��
        If InStr(argOpts.Item("�X�^�C��"), "�Ȗ͗l�Z��") <> 0 Then
            bSimaBgCol = True
        End If
        '�P�s���Ƃɍĕ`��
        If argOpts.Item("�s�ĕ`��") <> "" Then
'            b1LScrUpd = True
        End If
        
        Dim sTotalCnt As String
        If argOpts.Item("����") = "" Then
            sTotalCnt = " ��(ESC�L�[�Œ��f)"
        Else
            sTotalCnt = " �� / " & Format(argOpts.Item("����"), "00000") & " ��(ESC�L�[�Œ��f)"
        End If
        
        '�񏑎��̃Z�b�g
        For nFldIdx = 1 To nFldCnt
            Select Case argFlds.Item(nFldIdx).Item("Type")
            Case ccFT_TEXT, ccFT_DATE, ccFT_NONE
                '�����񏑎����Z�b�g
                .Columns(nFldIdx).NumberFormatLocal = "@"
            Case ccFT_ELSE
                '�W�������̂܂�
            End Select
        Next
        
        
        '�f�[�^�s�̃Z�b�g
        Application.EnableCancelKey = xlErrorHandler
        On Error GoTo KEYBREAK
        bKeyBreak = False
        Do Until argRs.EOF
            
            If nCnt Mod 10 = 0 Then
                Application.StatusBar = "�f�[�^�𒊏o��... " & Format(nCnt, "00000") & sTotalCnt
            End If
            
            If nMaxCnt <> 0 And nMaxCnt <= nCnt Then
                Exit Do
            End If
            
            '���f�L�[���������m
            If bKeyBreak Then
                If MsgBoxST("���f���܂����H", vbYesNo) = vbYes Then Exit Do
                bKeyBreak = False
            End If
            
            '�z��Ɋi�[
            nRsIdx = 0
            For nFldIdx = 1 To nFldCnt
'                Set oFld = argFlds.Item(nFldIdx)
                Select Case sSpFlds(nFldIdx)
                Case "�󔒗�"
                Case "�s�A��" '����No.�\��
                Case "�s�ԍ�" '�l��No.�\��
                    vDtFlds(nFldIdx) = Format(nRow - 1, "00000")
                Case "�f�[�^"
                    vDtFlds(nFldIdx) = argFlds.Item(nFldIdx).Item("�f�[�^").Item(nCnt + 1)
                Case "�������s" '��
                    vDtFlds(nFldIdx) = Run(argFlds.Item(nFldIdx).Item("������"), _
                                        vDtFlds(argFlds.Item(nFldIdx).Item("�o�ʒu")))
                Case Else '���R�[�h�Z�b�g�̍���
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
            
            '�s�\��t��
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
        
        If InStr(argOpts.Item("�X�^�C��"), "������") <> 0 Then
            With .Range(.Cells(nMinRow, nMinCol), .Cells(nMaxRow, nMaxCol))
                If InStr(argOpts.Item("�^�C�g��"), "�t�B���^") <> 0 Then
                    .AutoFilter
                End If
                .EntireColumn.AutoFit
            End With
            If argOpts.Item("�^�C�g��") <> "" Then
                .Range(.Cells(nMinRow, nMinCol), .Cells(nMinRow, nMaxCol)).Value = sNmFlds
            End If
        End If
        
        '�f�[�^�\�t��̐��`
        For nFldIdx = 1 To nFldCnt
            Select Case sSpFlds(nFldIdx)
            Case "�s�A��" '����No.�\��
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
            If InStr(argOpts.Item("�X�^�C��"), "�󔒃Z������") <> 0 Then
'                With .FormatConditions.Add(Type:=xlExpression, _
'                    Formula1:="=IF(A$1<>"""",ISBLANK(A1),FALSE)")
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=IF(AND($A$1=""MAINTEN"",COLUMN()=2),FALSE,ISBLANK(A1))")
                    .Interior.ColorIndex = ccPINK
                End With
            End If
            If InStr(argOpts.Item("�X�^�C��"), "�Ȗ͗l����") <> 0 Then
                With .FormatConditions.Add(Type:=xlExpression, _
                    Formula1:="=AND(MOD(SUBTOTAL(3,$A$1:$A1),2)=1,ROW()<>1)")
                    .Interior.ColorIndex = ccYELLOW
                End With
            End If
            If InStr(argOpts.Item("�X�^�C��"), "�r���`��") <> 0 Then
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
        
        '�f�t�H���g�̃y�[�W�ݒ�
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
        
        argOpts.Item("���o") = TimeElapsed(nST, nET)
        argOpts.Item("����") = nCnt
        
        If InStr(argOpts.Item("�R�����g"), "�����A�b��") <> 0 Then
            Dim sCmtTxt As String
            If argOpts.Item("����") <> "" Then
                sCmtTxt = "����:" & argOpts.Item("����") & "��"
            End If
            If argOpts.Item("���s") <> "" Then
                sCmtTxt = IIf(sCmtTxt = "", sCmtTxt, sCmtTxt & " ")
                sCmtTxt = sCmtTxt & "���s:" & argOpts.Item("���s") & "�b ���o:" & argOpts.Item("���o") & "�b"
            End If
            
            '�R�����g�ǉ�
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


'��^SQL���̎擾
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

'�r�p�k�����Z���ɃZ�b�g
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
'���^�C�g�����l�������\��t��
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
    
        If GetCfgVal("�ڑ����[�h") = 1 Then 'OO4O/SQL*Net
            
            sSQL = GetLibSql("�I�u�W�F�N�g��ʎ擾")
            
            sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", sObjName)
            sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g���", sTypName)
            sSQL = SetSqlParam(sSQL, "�R�����g������", sCmtName)
            
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
            GetDBObjectType = "ADO/ODBC���[�h�̃I�u�W�F�N�g���擾�͂܂����Ή��ł�"
        End If
    End If
    
End Function

Private Function GetDBObjectList(argObjName, argRs, argFlds, argOpts)
    Dim sSQL As String, sExtSQL As String
    Dim oExtRs, oFld, oDat
    Dim sObjName As String, sTypName As String, sCmtName As String
    
    If argOpts.Item("�g���ꗗ") = "" Then
        sSQL = GetLibSql("�I�u�W�F�N�g�ꗗ�擾")
    Else
        sSQL = GetLibSql("�I�u�W�F�N�g�g���ꗗ�擾")
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
    
    sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", sObjName)
    sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g���", sTypName)
    sSQL = SetSqlParam(sSQL, "�R�����g������", sCmtName)
    
    If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
        Exit Function
    End If
    
    If argOpts.Item("�g���ꗗ") <> "" Then
        If GetCfgVal("�g�����F��") <> 1 Then
            Call argFlds.Remove("STATUS")
        End If
        If GetCfgVal("�g�����F�X�V����") <> 1 Then
            Call argFlds.Remove("LAST_DDL_TIME")
        End If
        If GetCfgVal("�g�����F�V�m�j��") <> 1 Then
            Call argFlds.Remove("FROM_SYNONYM")
        End If
        If GetCfgVal("�g�����F����") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If argRs("OBJECT_TYPE").Value = "TABLE" Then
                    sExtSQL = GetLibSql("�����擾")
                    sExtSQL = SetSqlParam(sExtSQL, "�擾�Ώ�", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("CNT").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "����"
            oFld.Item("�����") = "�f�[�^"
            Set oFld.Item("�f�[�^") = oDat
            Set argFlds.Item() = oFld
        End If
        If GetCfgVal("�g�����F�s��") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If argRs("OBJECT_TYPE").Value = "PROCEDURE" Or argRs("OBJECT_TYPE").Value = "FUNCTION" Then
                    sExtSQL = GetLibSql("�\�[�X�s���擾")
                    sExtSQL = SetSqlParam(sExtSQL, "�I�u�W�F�N�g��", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("CNT").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "�s��"
            oFld.Item("�����") = "�f�[�^"
            Set oFld.Item("�f�[�^") = oDat
            Set argFlds.Item() = oFld
        End If
        If GetCfgVal("�g�����F�T�C�Y") = 1 Then
            Set oDat = New Prop
            Do Until argRs.EOF
                If InStr("TABLE,INDEX", argRs("OBJECT_TYPE").Value) <> 0 Then
                    sExtSQL = GetLibSql("�T�C�Y�擾")
                    sExtSQL = SetSqlParam(sExtSQL, "�I�u�W�F�N�g��", argRs("OBJECT_NAME").Value)
                    Call ExecSql(GetConStr, sExtSQL, oExtRs)
                    oDat.Item() = oExtRs("MB").Value
                Else
                    oDat.Item() = ""
                End If
                argRs.MoveNext
            Loop
            argRs.MoveFirst
            Set oFld = New Prop
            oFld.Item("Title") = "�T�C�Y(MB)"
            oFld.Item("�����") = "�f�[�^"
            Set oFld.Item("�f�[�^") = oDat
            Set argFlds.Item() = oFld
        End If
    End If
    
    GetDBObjectList = True
End Function

Private Function GetDBObjectInfo(argObjName, argObjType, argRs, argFlds, argOpts)
    Dim sSQL As String
    
    Select Case argObjType
    Case "TABLE", "VIEW"
        sSQL = GetLibSql("�\��`�擾")
        
        sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", argObjName)
        
        If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
            Exit Function
        End If
    Case "INDEX"
        sSQL = GetLibSql("������`�擾")
        
        sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", argObjName)
        
        If Not ExecSql(GetConStr, sSQL, argRs, argFlds, argOpts) Then
            Exit Function
        End If
    Case "PROCEDURE", "FUNCTION", "PACKAGE", "PACKAGE BODY"
        sSQL = GetLibSql("�\�[�X�擾")
        
        sSQL = SetSqlParam(sSQL, "�I�u�W�F�N�g��", argObjName)
        
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
        If InStr(argExtName, "��") <> 0 Or InStr(argExtName, "�X�V����") <> 0 _
        Or InStr(argExtName, "�V�m�j��") <> 0 Then
            sExtSQL = GetLibSql("�I�u�W�F�N�g�g���ꗗ�擾")
            sExtSQL = SetSqlParam(sExtSQL, "�I�u�W�F�N�g��", argObjName)
            If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                If InStr(argExtName, "��") <> 0 Then
                    .Item("STATUS") = "��: " & oExtRs("STATUS").Value
                End If
                If InStr(argExtName, "�X�V����") <> 0 Then
                    .Item("LAST_DDL_TIME") = "�X�V: " & oExtRs("LAST_DDL_TIME").Value
                End If
                If InStr(argExtName, "�V�m�j��") <> 0 And oExtRs("FROM_SYNONYM").Value <> "" Then
                    .Item("FROM_SYNONYM") = oExtRs("FROM_SYNONYM").Value
                End If
            End If
        End If
        If InStr(argExtName, "����") <> 0 Then
            Select Case argObjType
            Case "TABLE", "VIEW"
                sExtSQL = GetLibSql("�����擾")
                sExtSQL = SetSqlParam(sExtSQL, "�擾�Ώ�", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("COUNT") = "����: " & oExtRs("CNT").Value
                End If
            End Select
        End If
        If InStr(argExtName, "�s��") <> 0 Then
            Select Case argObjType
            Case "PROCEDURE", "FUNCTION"
                sExtSQL = GetLibSql("�\�[�X�s���擾")
                sExtSQL = SetSqlParam(sExtSQL, "�I�u�W�F�N�g��", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("LINES") = "�s��: " & oExtRs("CNT").Value
                End If
            End Select
        End If
        If InStr(argExtName, "�T�C�Y") <> 0 Then
            Select Case argObjType
            Case "TABLE", "INDEX"
                sExtSQL = GetLibSql("�T�C�Y�擾")
                sExtSQL = SetSqlParam(sExtSQL, "�I�u�W�F�N�g��", argObjName)
                If ExecSql(GetConStr, sExtSQL, oExtRs) Then
                    .Item("MB") = "�T�C�Y: " & Trim(oExtRs("MB").Value) & "MB"
                End If
            End Select
        End If
    End With
    Set GetDBObjectExtInfo = oExtInfo
End Function

'�R�����g���e�̎擾
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

'�R�����g���e�̐ݒ�
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
    If GetCfgVal("�ڑ����[�h") = 1 Then
'�����ڂ̍��ڌ^���肪���������̂� .type ���� .OraIDataType �ɕύX
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

'�A�������󔒕����ł̕���
Public Function SplitSh(sStr As Variant) As Variant
    Dim vArray() As Variant
    Dim i As Long
    Dim j As Long
    Dim sChr As String
    Dim sTmp As String
    
    i = 0
    
    For j = 1 To Len(sStr) + 1
        sChr = Mid(sStr, j, 1)
        If sChr = " " Or sChr = "�@" Then
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

'CSV������̕���
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
        nElapsed = (86400 - argStart) + argEnd ' �^�钆��0�����ׂ����Ƃ��̑Ώ�
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
's = GetLibSql("�I�u�W�F�N�g�ꗗ�擾")
's = SetSqlParam(s, "�I�u�W�F�N�g��", "TBWE2")
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
'MsgBalloon "�e�X�g"

'Dim RE
'Set RE = CreateObject("VBScript.RegExp")

'Call MakeSheet("ZZZ", ActiveSheet, coSTBook.Worksheets("WfOU_"))

'Call Shell("cmd /c C:\cygwin64\bin\bash --login -c uname > C:\ExSQell.tmp", vbHide)
'Call Shell("cmd /c ""C:\Program Files\Git\bin\bash"" --login -c uname > C:\ExSQell.tmp", vbHide)

Dim arFlds
'arFlds = SplitSh("a    b      c        d       e")
arFlds = SplitCsv("a , "" b,c "" , d")

End Sub
