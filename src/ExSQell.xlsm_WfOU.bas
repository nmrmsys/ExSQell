
Const ccWfOU = "Worksheet for Oracle Utility"

Const ccOU_POS = "B5"
Const ccPF_POS = "B7"

Const ccAUTO = xlColorIndexAutomatic
Const ccGREY = 15
Const ccEDIT = "notepad"

Public coUtilOpts

'Oracle Utility パス 参照ボタン
Private Sub CommandButton1_Click()
    Dim sFilename, sDrive, sPFold
    sDrive = GetDriveName(Range(ccOU_POS))
    sPFold = Range(ccOU_POS)
    On Error Resume Next
    If sDrive <> "" Then ChDrive (sDrive)
    If sPFold <> "" Then ChDir (sPFold)
    On Error GoTo 0
    sFilename = Application.GetOpenFilename("Oracle Utilityパス,*.exe")
    If sFilename <> False Then
        Range(ccOU_POS) = GetParentFolderName(sFilename)
    End If
End Sub

'パラメータファイル 参照ボタン
Private Sub CommandButton2_Click()
    Dim sFilename, sDrive, sPFold
    sDrive = GetDriveName(Range(ccPF_POS))
    sPFold = GetParentFolderName(Range(ccPF_POS))
    On Error Resume Next
    If sDrive <> "" Then ChDrive (sDrive)
    If sPFold <> "" Then ChDir (sPFold)
    On Error GoTo 0
    sFilename = Application.GetSaveAsFilename(Range(ccPF_POS), "パラメータファイル,*.par")
    If sFilename <> False Then
        Range(ccPF_POS) = sFilename
    End If
End Sub

'ダンプファイル 参照ボタン
Private Sub CommandButton3_Click()
    Call OpenFolder(Range(GetUtilOpts("FILE", "VPOS")))
End Sub

'ログファイル 確認ボタン
Private Sub CommandButton4_Click()
    Call EditFile(Range(GetUtilOpts("LOG", "VPOS")))
End Sub

'コントロールファイル 確認ボタン
Private Sub CommandButton5_Click()
    Call EditFile(Range(GetUtilOpts("CONTROL", "VPOS")))
End Sub

'データファイル 確認ボタン
Private Sub CommandButton12_Click()
    Call EditFile(Range(GetUtilOpts("DATA", "VPOS")))
End Sub

'不良ファイル 確認ボタン
Private Sub CommandButton6_Click()
    Call EditFile(Range(GetUtilOpts("BAD", "VPOS")))
End Sub

'廃棄ファイル 確認ボタン
Private Sub CommandButton7_Click()
    Call EditFile(Range(GetUtilOpts("DISCARD", "VPOS")))
End Sub

'読込ボタン
Private Sub CommandButton8_Click()
    If MsgBox("パラメータファイルからロードしますか？", vbYesNo + vbDefaultButton2, ccWfOU) = vbYes Then
        Call ParFileLoad(Range(ccPF_POS))
    End If
End Sub

'保存ボタン
Private Sub CommandButton9_Click()
    If MsgBox("パラメータファイルにセーブしますか？", vbYesNo + vbDefaultButton2, ccWfOU) = vbYes Then
        Call ParFileSave(Range(ccPF_POS))
    End If
End Sub

'表示ボタン
Private Sub CommandButton10_Click()
    Dim sTempFile
    sTempFile = GetBaseName(Range(ccPF_POS)) & ".tmp"
    Call ParFileSave(sTempFile)
    Call EditFile(sTempFile)
End Sub

'実行ボタン
Private Sub CommandButton11_Click()
    Dim sTempFile, sUtilCommand, sPs
    sTempFile = GetBaseName(Range(ccPF_POS)) & ".tmp"
    Call ParFileSave(sTempFile)
    Select Case GetFlag(Range(GetUtilOpts("*MODE", "VPOS")))
    Case "E"
        sUtilCommand = "EXP.EXE"
    Case "I"
        sUtilCommand = "IMP.EXE"
    Case "L"
        sUtilCommand = "SQLLDR.EXE"
    End Select
    If Range(ccOU_POS) <> "" Then
        If InStr(Range(ccOU_POS), "\") <> 0 Then
            sPs = "\"
        Else
            sPs = "/"
        End If
        sUtilCommand = Range(ccOU_POS) & sPs & sUtilCommand
    End If
    Dim sDrive, sPFold
    sDrive = GetDriveName(sTempFile)
    sPFold = GetParentFolderName(sTempFile)
    On Error Resume Next
    If sDrive <> "" Then ChDrive (sDrive)
    If sPFold <> "" Then ChDir (sPFold)
    On Error GoTo 0
'    Call ExecConsole(sUtilCommand & " PARFILE=" & sTempFile)
    Call ExecConsole(sUtilCommand & " PARFILE=" & GetFileName(sTempFile))
End Sub

'シート内容変更時
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
    Dim sAddress As String, sModeAdr As String, sDataAdr As String, sMode As String, sData As String
    Dim o
    
    sAddress = Target.Address(False, False)
    sModeAdr = GetUtilOpts("*MODE", "VPOS")
    sDataAdr = GetUtilOpts("*DATA", "VPOS")
    
    Select Case sAddress
    Case sModeAdr, sDataAdr
        '有効状態のセット
        sMode = GetFlag(Range(sModeAdr))
        sData = GetFlag(Range(sDataAdr))
        For Each o In GetUtilOpts
            If UtilOpt(o, "TYPE") <> "*" Then
                If IsEnabled(UtilOpt(o, "NPOS"), "NPOS") Then
                    Range(UtilOpt(o, "NPOS") & ":" & UtilOpt(o, "VPOS")).Font.ColorIndex = ccAUTO
                Else
                    Range(UtilOpt(o, "NPOS") & ":" & UtilOpt(o, "VPOS")).Font.ColorIndex = ccGREY
                End If
            End If
        Next
        '累積/増分のセット
        Range(GetUtilOpts("INCTYPE", "VPOS")) = ""
        Range(GetUtilOpts("RECORD", "VPOS")) = ""
        Select Case sMode
        Case "E"
            Call SetValidateList(GetUtilOpts("INCTYPE", "VPOS"), "完全,累積,増分")
        Case "I"
            Call SetValidateList(GetUtilOpts("INCTYPE", "VPOS"), "累積/増分情報,ユーザーデータ")
        End Select
    End Select
End Sub

'オプション情報取得
Private Function GetUtilOpts(Optional argKeyName, Optional argFldName)
    If IsEmpty(coUtilOpts) Or False Then
        Set coUtilOpts = New Collection
        With coUtilOpts
            '項目タイプ、有効状態、名称位置、値位置、パラメータ名、デフォルト値、キー
            Call .Add(Array("*", "EILFUT", "H6 ", "H7 ", "*MODE", "      "), "*MODE")
            
            Call .Add(Array("F", "EI_FUT", "B10", "B11", "FILE", "       "), "FILE")
            Call .Add(Array("F", "EILFUT", "B12", "B13", "LOG", "        "), "LOG")
            Call .Add(Array("F", "__LFUT", "B14", "B15", "CONTROL", "    "), "CONTROL")
            Call .Add(Array("F", "__LFUT", "B16", "B17", "DATA", "       "), "DATA")
            Call .Add(Array("F", "__LFUT", "B18", "B19", "BAD", "        "), "BAD")
            Call .Add(Array("F", "__LFUT", "B20", "B21", "DISCARD", "    "), "DISCARD")
            
            Call .Add(Array("S", "EILFUT", "H10", "H11", "USERID", "     "), "USERID")
            Call .Add(Array("*", "EI_FUT", "H12", "H13", "*DATA", "      "), "*DATA")
            Call .Add(Array("L", "_I_FUT", "H14", "H15", "SHOW", "      N"), "SHOW")
            
            Call .Add(Array("L", "EI_F__", "H18", "H19", "INCTYPE", "    "), "INCTYPE")
            Call .Add(Array("L", "E__F__", "H20", "H21", "RECORD", "    Y"), "RECORD")
            
            Call .Add(Array("L", "EI_FUT", "B24", "C24", "ROWS", "      Y"), "ROWS")
            Call .Add(Array("L", "EI_FUT", "B25", "C25", "INDEXES", "   Y"), "INDEXES")
            Call .Add(Array("L", "EI_FUT", "B26", "C26", "GRANTS", "    Y"), "GRANTS")
            Call .Add(Array("L", "E__FUT", "B27", "C27", "COMPRESS", "  Y"), "COMPRESS")
            Call .Add(Array("L", "E__FUT", "B28", "C28", "TRIGGERS", "  Y"), "TRIGGERS")
            Call .Add(Array("L", "E__FUT", "B29", "C29", "CONSISTENT", "N"), "CONSISTENT")
            Call .Add(Array("L", "_I_FUT", "B30", "C30", "IGNORE", "    N"), "IGNORE")
            Call .Add(Array("L", "_I_FUT", "B31", "C31", "COMMIT", "    N"), "COMMIT")
            Call .Add(Array("L", "E_LFUT", "B32", "C32", "DIRECT", "    N"), "DIRECT")
        
            Call .Add(Array("S", "EI_FUT", "B35", "C35", "FEEDBACK", "   "), "FEEDBACK")
            Call .Add(Array("S", "EI_FUT", "B36", "C36", "BUFFER", "     "), "BUFFER")
            Call .Add(Array("S", "EI_FUT", "B37", "C37", "FILESIZE", "   "), "FILESIZE")
            Call .Add(Array("S", "__LFUT", "B38", "C38", "ERRORS", "     "), "ERRORS")
            Call .Add(Array("S", "__LFUT", "B39", "C39", "DISCARDMAX", " "), "DISCARDMAX")
            Call .Add(Array("S", "__LFUT", "B40", "C40", "SKIP", "       "), "SKIP")
            Call .Add(Array("S", "__LFUT", "B41", "C41", "LOAD", "       "), "LOAD")
            
            Call .Add(Array("D", "EILFUT", "B46", "C54", "*DIRECT", "    "), "*DIRECT")
            
            Call .Add(Array("S", "E___U_", "E24", "F24", "OWNER", "      "), "OWNER")
            Call .Add(Array("S", "_I__UT", "E24", "F24", "FROMUSER", "   "), "FROMUSER")
            Call .Add(Array("S", "_I__UT", "E25", "F25", "TOUSER", "     "), "TOUSER")
        
            Call .Add(Array("T", "EI___T", "E28", "E43", "TABLES", "     "), "TABLES")
        
            Call .Add(Array("Q", "E____T", "E46", "E54", "QUERY", "      "), "QUERY")
        End With
    End If
    If IsMissing(argKeyName) Then
        Set GetUtilOpts = coUtilOpts
    ElseIf IsMissing(argFldName) Then
        On Error Resume Next
        GetUtilOpts = coUtilOpts(argKeyName)
        On Error GoTo 0
    Else
        On Error Resume Next
        GetUtilOpts = UtilOpt(coUtilOpts(argKeyName), argFldName)
        On Error GoTo 0
    End If
End Function

Private Function UtilOpt(argUtilOpt, argFldName)
    Dim nFldIndex
    Select Case argFldName
    Case "TYPE"
        nFldIndex = 0
    Case "EFLG"
        nFldIndex = 1
    Case "NPOS"
        nFldIndex = 2
    Case "VPOS"
        nFldIndex = 3
    Case "NAME"
        nFldIndex = 4
    Case "DVAL"
        nFldIndex = 5
    End Select
    UtilOpt = Trim(argUtilOpt(nFldIndex))
End Function

'フラグ/名称 変換
Private Function GetFlag(argName)
    Select Case argName
    Case "Export"
        GetFlag = "E"
    Case "Import"
        GetFlag = "I"
    Case "Loader"
        GetFlag = "L"
    Case "ＤＢ全体"
        GetFlag = "F"
    Case "ユーザー"
        GetFlag = "U"
    Case "テーブル"
        GetFlag = "T"
    Case "指定する"
        GetFlag = "Y"
    Case "指定しない"
        GetFlag = "N"
    Case "完全"
        GetFlag = "COMPLETE"
    Case "累積"
        GetFlag = "CUMULATIVE"
    Case "増分"
        GetFlag = "INCREMENTAL"
    Case "累積/増分情報"
        GetFlag = "SYSTEM"
    Case "ユーザーデータ"
        GetFlag = "RESTORE"
    Case ""
        GetFlag = ""
    End Select
End Function

Private Function GetName(argFlag)
    Select Case argFlag
    Case "E"
        GetName = "Export"
    Case "I"
        GetName = "Import"
    Case "L"
        GetName = "Loader"
    Case "F"
        GetName = "ＤＢ全体"
    Case "U"
        GetName = "ユーザー"
    Case "T"
        GetName = "テーブル"
    Case "Y"
        GetName = "指定する"
    Case "N"
        GetName = "指定しない"
    Case "COMPLETE"
        GetName = "完全"
    Case "CUMULATIVE"
        GetName = "累積"
    Case "INCREMENTAL"
        GetName = "増分"
    Case "SYSTEM"
        GetName = "累積/増分情報"
    Case "RESTORE"
        GetName = "ユーザーデータ"
    Case ""
        Stop
        GetName = ""
    Case Else
        Stop
    End Select
End Function

Private Function IsEnabled(argTest, Optional argFldName = "NAME")
    Dim sMode As String, sData As String, o
    
    sMode = GetFlag(Range(GetUtilOpts("*MODE", "VPOS")))
    sData = GetFlag(Range(GetUtilOpts("*DATA", "VPOS")))
    
    For Each o In GetUtilOpts
        If UtilOpt(o, argFldName) = argTest Then
            If InStr(UtilOpt(o, "EFLG"), sMode) <> 0 And InStr(UtilOpt(o, "EFLG"), sData) <> 0 Then
                IsEnabled = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub SetValidateList(argAddress As String, argList As String)
    With Range(argAddress).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=argList
    End With
End Sub

Private Function GetDriveName(argFilename) As String
    If InStr(argFilename, "\") <> 0 Then
        GetDriveName = Left(argFilename, 1)
    End If
End Function

Private Function GetParentFolderName(argFilename) As String
    Dim sPs As String, nLp As Integer, i As Integer
    If InStr(argFilename, "\") <> 0 Then
        sPs = "\"
    Else
        sPs = "/"
    End If
    For i = Len(argFilename) To 1 Step -1
        If Mid(argFilename, i, 1) = sPs Then
            nLp = i - 1
            Exit For
        End If
    Next
    If nLp <> 0 Then
        GetParentFolderName = Left(argFilename, nLp)
    End If
End Function

Private Function GetFileName(argFilename) As String
    Dim sPFold As String
    sPFold = GetParentFolderName(argFilename)
    If sPFold <> "" Then
        GetFileName = Mid(argFilename, Len(sPFold) + 2)
    Else
        GetFileName = argFilename
    End If
End Function

Private Function GetBaseName(argFilename) As String
    Dim nLp As Integer, i As Integer
    For i = Len(argFilename) To 1 Step -1
        If Mid(argFilename, i, 1) = "." Then
            nLp = i - 1
            Exit For
        End If
    Next
    If nLp <> 0 Then
        GetBaseName = Left(argFilename, nLp)
    Else
        GetBaseName = argFilename
    End If
End Function

Private Function GetExtentionName(argFilename) As String
    Dim nLp As Integer, i As Integer
    For i = Len(argFilename) To 1 Step -1
        If Mid(argFilename, i, 1) = "." Then
            nLp = i + 1
            Exit For
        End If
    Next
    If nLp <> 0 Then
        GetExtentionName = Mid(argFilename, nLp)
    End If
End Function

Private Function GetPathCompletionName(argFilename) As String
    Dim sTmpName As String, sTmpPath As String, sPs As String, sPFold As String
    If argFilename <> "" Then
        sTmpName = GetParentFolderName(argFilename)
        If sTmpName = "" Then
            sTmpPath = Range("B7")
            If InStr(sTmpPath, "\") <> 0 Then
                sPs = "\"
            Else
                sPs = "/"
            End If
            sPFold = GetParentFolderName(sTmpPath)
            If sPFold <> "" Then
                GetPathCompletionName = sPFold & sPs & argFilename
            Else
                GetPathCompletionName = argFilename
            End If
        Else
            GetPathCompletionName = argFilename
        End If
    End If
End Function

Private Function Split(sStr As Variant, sDelim As Variant) As Variant
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

Private Function Replace(sStr As Variant, sOld As Variant, sNew As Variant) As String
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

Private Sub EditFile(argFilename)
    Call Shell(ccEDIT & " " & GetPathCompletionName(argFilename), vbNormalFocus)
End Sub

Private Sub OpenFolder(argFilename)
'    Call Shell("EXPLORER.EXE /select," & GetPathCompletionName(argFilename), vbNormalFocus)
    Call Shell("EXPLORER.EXE /n," & _
        GetParentFolderName(GetPathCompletionName(argFilename)), vbNormalFocus)
End Sub

Private Sub ExecConsole(argCommand)
    Call Shell(Environ$("ComSpec") & " /k " & argCommand, vbNormalFocus)
End Sub

Private Sub ParFileSave(argFilename)
    Dim nFn As Integer, sFlag As String, bFull As Boolean, oRow, sLine As String
    nFn = FreeFile
    Open argFilename For Output As nFn
    For Each o In GetUtilOpts
        If IsEnabled(UtilOpt(o, "NAME")) Then
            Select Case UtilOpt(o, "TYPE")
            Case "F" 'ファイル名指定
                If Range(UtilOpt(o, "VPOS")) <> "" Then
                    Print #nFn, UtilOpt(o, "NAME") & "=" & _
                        Range(UtilOpt(o, "VPOS"))
'                        GetPathCompletionName(Range(UtilOpt(o, "VPOS")))
                End If
            Case "S" '文字列指定
                If Range(UtilOpt(o, "VPOS")) <> "" Then
                    Print #nFn, UtilOpt(o, "NAME") & "=" & _
                        Range(UtilOpt(o, "VPOS"))
                End If
            Case "L" 'リスト指定
                If Range(UtilOpt(o, "VPOS")) <> "" Then
                    sFlag = GetFlag(Range(UtilOpt(o, "VPOS")))
'                    If UtilOpt(o, "DVAL") <> sFlag Then
                        Print #nFn, UtilOpt(o, "NAME") & "=" & sFlag
'                    End If
                End If
            Case "D" '直接指定
                For Each oRow In Range(UtilOpt(o, "NPOS") & ":" & UtilOpt(o, "VPOS")).Rows
                    If oRow.Columns(1) <> "" And oRow.Columns(1).Font.ColorIndex <> ccGREY Then
                        Print #nFn, oRow.Columns(1) & "=" & oRow.Columns(2)
                    End If
                Next
            Case "T" 'テーブル指定
                sLine = ""
                For Each oRow In Range(UtilOpt(o, "NPOS") & ":" & UtilOpt(o, "VPOS")).Rows
                    If oRow.Columns(1) <> "" And oRow.Columns(1).Font.ColorIndex <> ccGREY Then
                        If sLine <> "" Then sLine = sLine & vbLf
                        sLine = sLine & Trim(oRow.Columns(1))
                    End If
                Next
                sLine = Replace(sLine, "," & vbLf, ",")
                sLine = Replace(sLine, vbLf, ",")
                sLine = Replace(sLine, " ", "")
                sLine = Replace(sLine, ",", "," & vbCrLf & " ")
                Print #nFn, "TABLES=(" & vbCrLf & " " & sLine & vbCrLf & ")"
            Case "Q" '抽出条件
                sLine = ""
                For Each oRow In Range(UtilOpt(o, "NPOS") & ":" & UtilOpt(o, "VPOS")).Rows
                    If oRow.Columns(1) <> "" And oRow.Columns(1).Font.ColorIndex <> ccGREY Then
                        If sLine <> "" Then sLine = sLine & vbLf
                        sLine = sLine & Trim(oRow.Columns(1))
                    End If
                Next
                sLine = Replace(sLine, vbLf, " AND ")
                If sLine <> "" Then
                    Print #nFn, "QUERY=""WHERE " & sLine & """"
                End If
            Case Else
                If UtilOpt(o, "NAME") = "*DATA" Then
                    If GetFlag(Range(UtilOpt(o, "VPOS"))) = "F" Then bFull = True
                End If
            End Select
        End If
    Next
    If bFull Then Print #nFn, "FULL=Y"
    Close nFn
End Sub

Private Sub ParFileLoad(argFilename)
    Dim nFn As Integer, sLine As String, arLine, sName As String
    Dim sMode As String, sData As String, sPFold As String
    Dim bTables As Boolean, sTables As String, arTables, nTables As Integer
    Dim arQuery, nDirect As Integer, i As Integer
    Range("I1").Select
    nFn = FreeFile
    Open argFilename For Input As nFn
    
    Range(GetUtilOpts("*DIRECT", "NPOS") & ":" & GetUtilOpts("*DIRECT", "VPOS")) = ""
    Range(GetUtilOpts("TABLES", "NPOS") & ":" & GetUtilOpts("TABLES", "VPOS")) = ""
    Range(GetUtilOpts("QUERY", "NPOS") & ":" & GetUtilOpts("QUERY", "VPOS")) = ""
    
    Do While Not EOF(nFn)
        Line Input #nFn, sLine
        arLine = Split(sLine, "=")
        sName = UCase(Trim(arLine(0)))
        
        'モード取得
        Select Case Left(GetUtilOpts(sName, "EFLG"), 3)
        Case "E__"
            sMode = "E"
        Case "_I_"
            sMode = "I"
        Case "__L"
            sMode = "L"
        End Select
    
        'データ範囲取得、値セット
        Select Case sName
        Case "FULL"
            If UCase(arLine(1)) = "Y" Then
                sData = "F"
            End If
        Case "OWNER", "FROMUSER", "TOUSER"
            If sData <> "T" Then sData = "U"
            Range(GetUtilOpts(sName, "VPOS")) = arLine(1)
        Case "TABLES"
            sData = "T"
            sTables = arLine(1)
            If InStr(arLine(1), ")") = 0 Then bTables = True
        Case "QUERY"
            arQuery = Split(Mid(sLine, InStr(UCase(sLine), "WHERE ") + 6), " AND ")
            For i = 0 To UBound(arQuery)
                Range(GetUtilOpts(sName, "NPOS")).Offset(i, 0) = Replace(arQuery(i), """", "")
            Next
        Case Else
            If InStr(sLine, ")") Then
                If bTables Then
                    sTables = sTables & sLine
                End If
                bTables = False
            End If
            If bTables Then
                sTables = sTables & sLine
            Else
                Select Case GetUtilOpts(sName, "TYPE")
                Case "F" 'ファイル指定
                    sPFold = GetParentFolderName(arLine(1))
                    If sPFold <> "" And sPFold = GetParentFolderName(Range(ccPF_POS)) Then
                        arLine(1) = Mid(arLine(1), Len(sPFold) + 2)
                    End If
                    Range(GetUtilOpts(sName, "VPOS")) = arLine(1)
                Case "S" '文字列指定
                    Range(GetUtilOpts(sName, "VPOS")) = arLine(1)
                Case "L" 'リスト指定
                    Range(GetUtilOpts(sName, "VPOS")) = GetName(arLine(1))
                Case Else
                    '直接指定
                    If UBound(arLine) = 1 Then
                        Range(GetUtilOpts("*DIRECT", "NPOS")).Offset(nDirect, 0) = arLine(0)
                        Range(GetUtilOpts("*DIRECT", "NPOS")).Offset(nDirect, 1) = arLine(1)
                        nDirect = nDirect + 1
                    End If
                End Select
            End If
        End Select
    Loop
    Close nFn
    'テーブル指定
    sTables = Replace(sTables, " ", "")
    sTables = Replace(sTables, "(", "")
    sTables = Replace(sTables, ")", "")
    arTables = Split(sTables, ",")
    sTables = ""
    For i = 0 To UBound(arTables)
        If Len(sTables & ", " & arTables(i)) >= 59 Then
            Range(GetUtilOpts("TABLES", "NPOS")).Offset(nTables, 0) = sTables & ","
            nTables = nTables + 1
            sTables = arTables(i)
        Else
            If sTables <> "" Then
                sTables = sTables & ", "
            End If
            sTables = sTables & arTables(i)
        End If
    Next
    If sTables <> "" Then
        Range(GetUtilOpts("TABLES", "NPOS")).Offset(nTables, 0) = sTables
    End If
    'モード、データ範囲設定
    If sMode <> "" Then
        Range(GetUtilOpts("*MODE", "VPOS")) = GetName(sMode)
    End If
    If sData <> "" Then
        Range(GetUtilOpts("*DATA", "VPOS")) = GetName(sData)
    End If
End Sub

Sub test()
    Call ExecConsole("dir c:\ ")
End Sub
