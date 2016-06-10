'===================================================================================================
'= FFT - File_Filter_Toolkit v0.4.2
'===================================================================================================
'
' オプション処理 GetOptArgs, SetOptDefs, GetOptChoice, GetOptPrompt, GetOptDialog 以降は未実装 GetOptIni, GetOptReg
' ファイル操作   AddFiles, SetExclude, SetExcludeList, SetModifyOnly, SetOrder, SetRename, SetBackup
' フィルタ定義   SetFilter, ClearFilter, DeleteFilter
' 組込フィルタ   Grep, Sed, Tr, Sort, Uniq, Empty, Contains 以降は未実装 Cut, Wc, Cat, Split, Tee, Head, Tail
' 処理実行、他   Execute, View, Show, Resources, LoadResourceFile
' ユーティリティ M, S, CreateObject, GetTempName, GPN, GFN, GBN, GEN, Split, Sort, Tokenize
'                GetWinFolder, GetSysFolder, GetTempFolder, GetTempFile, LenB, MidB, LeftB, RightB, LPAD, RPAD
'                OpenMDBFile, OpenXMLFile
'                TextFileクラス, IE_IPCクラス
'
' 正規表現の     BASP21 http://www.hi-ho.ne.jp/babaq/bregexp.html
' リファレンス   RegExp http://www.f3.dion.ne.jp/~element/msaccess/AcResSnippetRegExp.html
'
' SetOptDefsの定義例
'  sOptDefs = "OptionGroup1"              & vbCrLf _
'           & "    -a, --add"             & vbCrLf _
'           & "    -b NUM, --bar=NUM"     & vbCrLf _
'           & ""                          & vbCrLf _
'           & "OptionGroup2"              & vbCrLf _
'           & "    -c [STR], --cat[=STR]" & vbCrLf
'
'todo
' ファイル追加/フィルタ処理中の進捗通知イベントハンドラ
' 処理ファイルを除外するIf ProcessFileでブール値以外？ or FileSetsフィルタタイプ
' 前回より更新されているファイルだけを対象とする様なフィルタ データはローカルMDB？
' フィルタオプション、パラメータ廻りの整理 SetFilterOption  sort -r とか
' BASP21Mode Option化?
'
Class FileFilterToolkit
  Public FSO
  Public WSH
  Public Options
  Public Mode
  Public BSP
  Public VRE
  Public XLS
  Public RES
  Public Files
  Public FOrders
  Public Filters
  Public FMatchs
  Public PathMatchOption
  Public Name
  Public FullName
  Public Path
  Public ProcessFilePath
  Public CurrentFilter, CurrentFilterID
  Public RenPtn, RenRep
  Public BakPtn, BakRep
  Public ExcludeSet
  Public ModifyOnlySet, ModifyOnlyDBPath, ModifyOnlyDB
  Public ProcessProgress
  Public OptArgsPrefix
  Public OptDefs
  Public OptErrs
  
  '初期処理
  Private Sub Class_Initialize
    Set FSO     = CreateObject("Scripting.FileSystemObject")
    Set WSH     = CreateObject("WScript.Shell")
    Set Options = CreateObject("Scripting.Dictionary")
    Set Files   = CreateObject("Scripting.Dictionary")
    Set FOrders = CreateObject("Scripting.Dictionary")
    Set Filters = CreateObject("Scripting.Dictionary")
    Set FMatchs = CreateObject("Scripting.Dictionary")
    Set CurrentFilter = Nothing
    CurrentFilterID = ""
    
    If LoadOptionModule(BSP, "BASP21", False) Then
      Mode = "BASP21"
    Else
      Set VRE = CreateObject("VBScript.RegExp")
    End If
    
    With WScript
      Name     = .ScriptName
      FullName = .ScriptFullName
    End With
    Path = GPN(FullName)
    PathMatchOption = "ik"
    OptArgsPrefix = "--, -, /"
  End Sub
  
  '終了処理
  Private Sub Class_Terminate
    If ModifyOnlySet Then
      ModifyOnlyDB.Close
    End If
    If Not IsEmpty(XLS) Then
      XLS.Quit
    End If
    Set FSO = Nothing
  End Sub
  
' SetOptDefsの定義例
'  sOptDefs = "OptionGroup1"              & vbCrLf _
'           & "    -a, --add"             & vbCrLf _
'           & "    -b NUM, --bar=NUM"     & vbCrLf _
'           & ""                          & vbCrLf _
'           & "OptionGroup2"              & vbCrLf _
'           & "    -c [STR], --cat[=STR]" & vbCrLf

  '引数定義セット
  Public Sub SetOptDefs(argOptDefs)

  End Sub
  
  '引数解析
  Public Sub GetOptArgs(argArgs)
    '簡易オプション解析用
    sOptArgsPrefixRegExp = "^(" & Replace(Replace(OptArgsPrefix, ",", "|"), " ", "") & ")"

    'パラメータ引数を順に処理
    Dim i
    For i = 0 To argArgs.Count - 1
      If M(argArgs(i), sOptArgsPrefixRegExp, "") Then
        'オプションの場合
        sOptionName = S(argArgs(i), sOptArgsPrefixRegExp, "", "")
        Options(sOptionName) = True '取りあえず暫定でTrueをセットしとく
      ElseIf FileExists(argArgs(i)) Then
        'ファイルの場合
        If Not Options("GetOptArgsDisableAddFile") Then
          '前回より更新のみチェック
          If IsModifyOnly(argArgs(i)) Then
            '除外チェック
            If Not IsExclude(argArgs(i)) Then
              Call Files.Add(Files.Count,argArgs(i))
            End If
          End If
        Else
          MsgBox "ファイルは追加できません: " & argArgs(i)
        End If
      Else
        'フォルダの場合
        If Not Options("GetOptArgsDisableAddFolder") Then
          Call AddFiles(argArgs(i) & "\*.*", Not Options("GetOptArgsDisableRecursive"))
        Else
          MsgBox "フォルダは追加できません: " & argArgs(i)
        End If
      End If
    Next
    
  End Sub
  
  '入力選択
  Public Sub GetOptChoice(argOptName, argPrompt, argChoice, argDefault)
    Dim sValue, sPrompt, sChoice, sDefault, iRet, iButtons
    
    sChoice  = UCase(argChoice)
    sDefault = UCase(argDefault)
    
    '表示ボタン設定
    If InStr(sChoice,"YES") > 0 And InStr(sChoice,"NO") > 0 And InStr(sChoice,"CANCEL") > 0 Then
      iButtons = 3 'vbYesNoCancel
      Select Case sDefault
      Case "NO"     iButtons = iButtons + 256 'vbDefaultButton2
      Case "CANCEL" iButtons = iButtons + 512 'vbDefaultButton3
      End Select
    ElseIf InStr(sChoice,"YES") > 0 And InStr(sChoice,"NO") > 0 Then
      iButtons = 4 'vbYesNo
      Select Case sDefault
      Case "NO"     iButtons = iButtons + 256 'vbDefaultButton2
      End Select
    ElseIf InStr(sChoice,"RETRY") > 0 And InStr(sChoice,"CANCEL") > 0 Then
      iButtons = 5 'vbRetryCancel
      Select Case sDefault
      Case "CANCEL" iButtons = iButtons + 256 'vbDefaultButton2
      End Select
    ElseIf InStr(sChoice,"ABORT") > 0 Then
      iButtons = 2 'vbAbortRetryIgnore
      Select Case sDefault
      Case "RETRY"  iButtons = iButtons + 256 'vbDefaultButton2
      Case "IGNORE" iButtons = iButtons + 512 'vbDefaultButton3
      End Select
    ElseIf InStr(sChoice,"OK") > 0 And InStr(sChoice,"CANCEL") > 0 Then
      iButtons = 1 'vbOkCancel
      Select Case sDefault
      Case "CANCEL" iButtons = iButtons + 256 'vbDefaultButton2
      End Select
    Else
      iButtons = 0 'vbOk
    End If
    
    '表示アイコン設定
    If InStr(argPrompt,"[x]") > 0 Then
      iButtons = iButtons + 16 'vbCritical
      sPrompt = Replace(argPrompt,"[x]","")
    ElseIf InStr(argPrompt,"[?]") > 0 Then
      iButtons = iButtons + 32 'vbQuestion
      sPrompt = Replace(argPrompt,"[?]","")
    ElseIf InStr(argPrompt,"[!]") > 0 Then
      iButtons = iButtons + 48 'vbExclamation
      sPrompt = Replace(argPrompt,"[!]","")
    ElseIf InStr(argPrompt,"[i]") > 0 Then
      iButtons = iButtons + 64 'vbInformation
      sPrompt = Replace(argPrompt,"[i]","")
    End If
    
    'メッセージボックス表示
    iRet = MsgBox(sPrompt, iButtons, Name)
    
    '押下ボタン取得
    Select Case iRet
    Case 1 'vbOK
      sValue = "OK"
    Case 2 'vbCancel
      sValue = "Cancel"
    Case 3 'vbAbort
      sValue = "Abort"
    Case 4 'vbRetry
      sValue = "Retry"
    Case 5 'vbIgnore
      sValue = "Ignore"
    Case 6 'vbYes
      sValue = "Yes"
    Case 7 'vbNo
      sValue = "No"
    End Select
    
    Call Options.Add(argOptName, sValue)
  End Sub
  
  '入力プロンプト
  Public Sub GetOptPrompt(argOptName,argPrompt,argDefault)
    Dim sValue
    sValue = InputBox(argPrompt, Name, argDefault)
    Call Options.Add(argOptName, sValue)
  End Sub
  
  '入力ダイアログ
  Public Sub GetOptDialog(argDialogID)
    'ダイアログ定義の取得
    sDlgSrc = Resources(argDialogID)
    If sDlgSrc = "" Then
      MsgBox "ダイアログ定義がありません: " & argDialogID
      Exit Sub
    End If
    '</head>の前にFFTオブジェクト連携用スクリプト挿入
    sScrFFT = ""
    sScrFFT = sScrFFT & "" & vbCrLf
    sScrFFT = sScrFFT & "<script type=""text/vbscript"">" & vbCrLf
    sScrFFT = sScrFFT & "<!--" & vbCrLf
    sScrFFT = sScrFFT & "  Args = Split(HTA.commandLine, "" "")" & vbCrLf
    sScrFFT = sScrFFT & "  Port = Args(UBound(Args))" & vbCrLf
    sScrFFT = sScrFFT & "  Retry = 3" & vbCrLf
    sScrFFT = sScrFFT & "  Set oShell = CreateObject(""Shell.Application"")" & vbCrLf
    sScrFFT = sScrFFT & "  nTry = 0" & vbCrLf
    sScrFFT = sScrFFT & "  Do While nTry < Retry Or Retry = 0" & vbCrLf
    sScrFFT = sScrFFT & "    For Each oProc In oShell.Windows()" & vbCrLf
    sScrFFT = sScrFFT & "      If Not IsEmpty(oProc.GetProperty(""IE_IPC:"" & Port)) Then" & vbCrLf
    sScrFFT = sScrFFT & "        Set IE = oProc.GetProperty(""IE_IPC:"" & Port)" & vbCrLf
    sScrFFT = sScrFFT & "        Exit Do" & vbCrLf
    sScrFFT = sScrFFT & "      End If" & vbCrLf
    sScrFFT = sScrFFT & "    Next" & vbCrLf
    sScrFFT = sScrFFT & "    Call Sleep(1)" & vbCrLf
    sScrFFT = sScrFFT & "    nTry = nTry + 1" & vbCrLf
    sScrFFT = sScrFFT & "  Loop" & vbCrLf
    sScrFFT = sScrFFT & "  If IsEmpty(IE) Then" & vbCrLf
    sScrFFT = sScrFFT & "    MsgBox ""IE_IPC:"" & Port & ""に接続できませんでした""" & vbCrLf
    sScrFFT = sScrFFT & "  Else" & vbCrLf
    sScrFFT = sScrFFT & "    Set FFT = IE.GetProperty(""FFT"")" & vbCrLf
    sScrFFT = sScrFFT & "  End If" & vbCrLf
    sScrFFT = sScrFFT & "  " & vbCrLf
    sScrFFT = sScrFFT & "  Sub Sleep(argSec)" & vbCrLf
    sScrFFT = sScrFFT & "    Call CreateObject(""WScript.Shell"").Run(""ping.exe localhost -n "" & (argSec+1), 0, True)" & vbCrLf
    sScrFFT = sScrFFT & "  End Sub" & vbCrLf
    sScrFFT = sScrFFT & "-->" & vbCrLf
    sScrFFT = sScrFFT & "</script>" & vbCrLf
    sScrFFT = sScrFFT & "" & vbCrLf
    nEndHead = InStr(1, sDlgSrc, "</head>", 1)
    If nEndHead = 0 Then
      MsgBox "<head>タグを記述して下さい"
      Exit Sub
    End If
    sDlgSrc = Mid(sDlgSrc, 1, nEndHead-1) & sScrFFT & Mid(sDlgSrc, nEndHead)
    'HTAファイル作成
    sHtaName = GetTempName
    Set oHtaFile = OpenWriteTextFile(sHtaName)
    Call oHtaFile.WriteLine(sDlgSrc)
    oHtaFile.Close
    'プロセス間通信用
    Set oIE_IPC = New IE_IPC
    With oIE_IPC
      Call .Listen
      Set .Item("FFT") = Me
      'HTA起動
      Call WSH.Run("MSHTA.EXE """ & sHtaName & """ " & .Port, 1, True)
      Call .Close
    End With
    'HTAファイル削除
    Call DeleteFile(sHtaName)
  End Sub
  
  'フィルタを追加
  Public Sub SetFilter(argFltId, argFltObj)
    Dim n
    n = FOrders.Count
    Call FOrders.Add(n, argFltId)
    Call Filters.Add(argFltId, argFltObj)
    Call FMatchs.Add(argFltId, "")
  End Sub
  
  'マッチフィルタを追加(ファイルパス一致のみ処理)
  Public Sub SetMatchFilter(argFltId, argFltObj, argPattern)
    Dim n
    n = FOrders.Count
    Call FOrders.Add(n, argFltId)
    Call Filters.Add(argFltId, argFltObj)
    Call FMatchs.Add(argFltId, argPattern)
  End Sub
  
  'フィルタを削除
  Public Sub DeleteFilter(argFltId)
    Dim i
    For i = 0 To FOrders.Count - 1
      If FOrders(i) = argFltId Then
        Call FOrders.Remove(i)
        Call Filters.Remove(argFltId)
        Call FMatchs.Remove(argFltId)
        Exit Sub
      End If
    Next
  End Sub
  
  'フィルタをクリア
  Public Sub ClearFilter
    Call FOrders.RemoveAll
    Call Filters.RemoveAll
    Call FMatchs.RemoveAll
  End Sub
  
  'フィルタ処理実行
  Public Sub Execute
    Dim sImpFile, sWrkFile, sOutFile, i, j, k, WrkFiles, OutFiles, TmpFiles, bRetWrk, sFMatch, bFMatch
    
    'フィルタを順に処理
    For i = 0 To Filters.Count - 1
      CurrentFilterID = FOrders(i)
      Set CurrentFilter = Filters(CurrentFilterID)
      With Filters(FOrders(i))
        'フィルタ初期処理を実行
        Call .Initialize(Me)
        
        '作業ファイル、一時ファイルを初期化
        Set WrkFiles   = CreateObject("Scripting.Dictionary")
        Set TmpFiles   = CreateObject("Scripting.Dictionary")
        
        '入力ファイルを順に処理
        For j = 0 To Files.Count - 1
          '作業ファイル名を確保
          Call WrkFiles.Add(WrkFiles.Count, GetTempName)
          k = WrkFiles.Count - 1
          '出力ファイル名を退避
          If i = 0 Then
            Call TmpFiles.Add(TmpFiles.Count, Files(j))
          Else
            Call TmpFiles.Add(TmpFiles.Count, OutFiles(j))
          End If
          
          WrkFile         = WrkFiles(k) 'フィルタ内から作業ファイルパス変更用
          ProcessFilePath = TmpFiles(k) 'フィルタ内から処理ファイルパス取得用
          ProcessProgress = Right("000" & i  * Files.Count + j + 1,Len(Filters.Count * Files.Count)) & _
                          "/" & Right("000" & Filters.Count * Files.Count,Len(Filters.Count * Files.Count)) & _
                          " " & Right("  " & CInt((i  * Files.Count + j + 1) / (Filters.Count * Files.Count) * 100),3) & "%"
          
          'マッチフィルタのファイルパス一致判定
          sFMatch = FMatchs(FOrders(i))
          If Left(sFMatch,1) <> "!" Then
            bFMatch = M(ProcessFilePath, sFMatch, PathMatchOption)
          Else 'パターンの先頭１桁が!の場合は不一致の時のみ
            sFMatch = Mid(sFMatch,2)
            bFMatch = Not M(ProcessFilePath, sFMatch, PathMatchOption)
          End If
          
          'マッチフィルタ未設定 or マッチフィルタ一致の場合、フィルタ処理実行
          If sFMatch = "" Or bFMatch Then
'            msgbox FOrders(i) & " T " & ProcessFilePath
            
            'フィルタ種別ごとの処理
            bRetWrk = False
            Select Case UCase(.FilterType)
            Case "LINE" '行フィルタ処理
              Dim oImpFile, oWrkFile, sLine
              If .OpenFile(Me, Files(j), WrkFile) Then
                Set oImpFile = FSO.OpenTextFile(Files(j),1,False)
                Set oWrkFile = FSO.OpenTextFile(WrkFile,2,True)
                
                Do Until oImpFile.AtEndOfStream
                  sLine = oImpFile.ReadLine
                  If .ProcessLine(Me, sLine) Then
                    oWrkFile.WriteLine sLine
                  End If
                Loop
                
                oImpFile.Close
                oWrkFile.Close
                bRetWrk = .CloseFile(Me, Files(j), WrkFile)
              End If
            Case "FILE" 'ファイルフィルタ処理
              bRetWrk = .ProcessFile(Me, Files(j), WrkFile)
            Case "COMMAND" 'コマンドフィルタ処理
              bRetWrk = .ProcessCommand(Me, Files(j), WrkFile)
            Case "XLS" 'Excelブックフィルタ処理
              If IsEmpty(XLS) Then
                On Error Resume Next
                Set XLS = CreateObject("Excel.Application")
                On Error GoTo 0
                If IsEmpty(XLS) Then
                  MsgBox "Excelが初期化できません"
                  Exit Sub
                End If
              End If
              XLS.Visible = Options("XLS_Visible")
              Dim oWorkbook
              XLS.Workbooks.Open Files(j)
              Set oWorkbook = XLS.Workbooks(GFN(Files(j)))
              oWorkbook.Application.EnableEvents = Options("XLS_EnableEvents")
              bRetWrk = .ProcessBook(Me, oWorkbook, WrkFile)
              oWorkbook.Close False
            Case Else
              MsgBox "未定義フィルタータイプです"
            End Select
            
            '作業ファイルを処理結果として返さなかった場合
            If Not bRetWrk Then
              '作業ファイル、結果ファイルを削除
              If FileExists(WrkFile) And WrkFile <> ProcessFilePath Then
                Call DeleteFile(WrkFile)
              End If
              Call WrkFiles.Remove(k)
              Call TmpFiles.Remove(k)
            Else
              WrkFiles(k) = WrkFile '作業ファイルがフィルタ内で変更されていた場合用
            End If
            
            '中間入力作業ファイルがある場合、削除
            If Files(j) <> ProcessFilePath Then
              If FileExists(Files(j)) Then
                Call DeleteFile(Files(j))
              End If
            End If
          Else 'フィルタ対象ファイルパスで無かった場合
'            msgbox FOrders(i) & " F " & ProcessFilePath
            WrkFiles(k) = Files(j)
          End If
          
        Next
        
        ProcessFilePath = ""
        ProcessProgress = ""
        
        '作業ファイル->入力ファイル、一時ファイル->出力ファイル
        Set Files    = WrkFiles
        Set OutFiles = TmpFiles
        
        'フィルタ終了処理を実行
        Call .Terminate(Me)
        Set CurrentFilter = Nothing
        CurrentFilterID = ""
      End With
    Next
    
    'フィルタ処理しているなら
    If Filters.Count > 0 Then
      '入力ファイル(作業ファイル)を出力ファイルにする
      For j = 0 To Files.Count - 1
        'バックアップ指定がある場合は上書き前の出力ファイルをコピー
        If BakPtn <> "" and BakRep <> "" Then
          BakFile = S(OutFiles(j), BakPtn, BakRep, PathMatchOption)
          If OutFiles(j) <> BakFile Then
            If FileExists(BakFile) Then
              Call DeleteFile(BakFile)
            End If
            Call CopyFile(OutFiles(j),BakFile)
          End If
        End If
        '最終の入力ファイルを出力ファイルにリネーム
        If RenPtn <> "" and RenRep <> "" Then
          RenFile = S(OutFiles(j), RenPtn, RenRep, PathMatchOption)
'          msgbox "Ptn: " & RenPtn & vbLf & "Rep: " & RenRep & vbLf & "Ret: " & RenFile
        Else
          RenFile = OutFiles(j)
        End If
        '入力ファイルを１回は処理しているならリネーム
        If Files(j) <> OutFiles(j) Then
          If FileExists(Files(j)) Then
            If FileExists(RenFile) Then
              Call DeleteFile(RenFile)
            End If
            Call MoveFile(Files(j),RenFile)
          End If
        Else
          '１回も処理して無くてもリネーム指定ありならコピー
          If OutFiles(j) <> RenFile Then
            If FileExists(RenFile) Then
              Call DeleteFile(RenFile)
            End If
            Call CopyFile(OutFiles(j),RenFile)
          End If
        End If
        OutFiles(j) = RenFile
      Next
      Set Files = OutFiles
    End If
    
  End Sub
  
  '結果確認
  Public Sub View
    '結果ファイルを順に処理
    For i = 0 To Files.Count - 1
      Call WSH.Run("NOTEPAD " & Files(i))
    Next
  End Sub
  
  'シェルコマンドを実行し、標準出力をファイルに出力
  Public Sub RunCommand(argCommand, argWrkFile)
    Const wsHide = 0
    Dim sVbs, oVbs, sCmd
    
    'VBSに引数を渡す方式は"のエスケープが上手く処理できず断念
    sVbs = S(argWrkFile,"\.tmp","\.vbs","i")
    Set oVbs = FSO.OpenTextFile(sVbs,2,True)
    With oVbs
      .WriteLine "Dim WSH, FSO, WRK"
      .WriteLine "Set WSH = WScript.CreateObject(""WScript.Shell"")"
      .WriteLine "Set FSO = WScript.CreateObject(""Scripting.FileSystemObject"")"
      .WriteLine "Set PRC = WSH.Exec(""" & S(argCommand,"""","""""","") & """"")"
      .WriteLine "Set WRK = FSO.OpenTextFile(""" & argWrkFile & """,2,True)"
      .WriteLine "Do Until PRC.StdOut.AtEndOfStream"
      .WriteLine "  WRK.WriteLine PRC.StdOut.ReadLine"
      .WriteLine "Loop"
      .WriteLine "WRK.Close"
'      .WriteLine "WScript.Echo """ & S(argCommand,"""","""""","") & """"
      .Close
    End With
    
    sCmd = "CSCRIPT //NOLOGO """ & sVbs & """"
    Call WSH.Run(sCmd, wsHide, True) 'Run(sCmd, [intWindowStyle], [bWaitOnReturn])
    
    If Not FileExists(argWrkFile) Then
      msgbox "コマンド出力結果ファイルが未作成" & vbLf & argCommand
    End If
    
    Call DeleteFile(sVbs)
  End Sub
  
  '選択使用のCOMサーバをロード
  Public Function LoadOptionModule(argObject, argClassName, argWarnMsg)
      If IsEmpty(argObject) Then
        On Error Resume Next
        Set argObject = CreateObject(argClassName)
        On Error GoTo 0
        If IsEmpty(argObject) Then
          If argWarnMsg Then MsgBox argClassName & "が初期化できません"
          Exit Function
        End If
      End If
      LoadOptionModule = True
  End Function
  
  '指定パスのファイルを追加 ファイル数が大量の時はBASP21モード推奨
  Public Sub AddFiles(argFilePath, argRecursive)
    Dim tmpGPN, tmpGFN, tmpGEN, tmpFolders, tmpFiles, i, n, oFolder, oFile
    
'    If Not FileExists(argFilePath) Then
'      argFilePath = argFilePath & "\*.*"
'    End If
    
    tmpGPN   = GPN(argFilePath)
    tmpGFN   = GFN(argFilePath)
    tmpGEN   = GEN(argFilePath)
    
    If InStr(UCase(Mode),"BASP21") > 0 Then
      'BASP21モード
      Call LoadOptionModule(BSP, "BASP21", True)
      
      'フォルダの検索
      tmpFolders = BSP.ReadDir(tmpGPN & "\*")
      For i = 0 to Ubound(tmpFolders)
        tmpPath = tmpGPN & "\" & tmpFolders(i)
        If FolderExists(tmpPath) And argRecursive Then
          '除外チェック
          If Not IsExclude(tmpPath) Then
            Call AddFiles(tmpPath & "\" & tmpGFN, argRecursive)
          End If
        End If
      Next
      
      'ファイルの検索
      tmpFiles = BSP.ReadDir(argFilePath & " :D")
      For i = 0 to Ubound(tmpFiles)
        tmpPath = tmpGPN & "\" & tmpFiles(i)
        '前回より更新のみチェック
        If IsModifyOnly(tmpPath) Then
          '除外チェック
          If Not IsExclude(tmpPath) Then
            Call Files.Add(Files.Count,tmpPath)
          End If
        End If
      Next
    Else
      'BASP21不使用モード
      
      'フォルダの検索
      Set tmpFolders = FSO.GetFolder(tmpGPN).SubFolders
      For Each oFolder In tmpFolders
        tmpPath = tmpGPN & "\" & oFolder.Name
        If FolderExists(tmpPath) And argRecursive Then
          '除外チェック
          If Not IsExclude(tmpPath) Then
            Call AddFiles(tmpPath & "\" & tmpGFN, argRecursive)
          End If
        End If
      Next
      
      'ファイルの検索
      Set tmpFiles = FSO.GetFolder(tmpGPN).Files
      For Each oFile In tmpFiles
        tmpPath = tmpGPN & "\" & oFile.Name
        '前回より更新のみチェック
        If IsModifyOnly(tmpPath) Then
          '除外チェック
          If Not IsExclude(tmpPath) And (GEN(oFile.Name) = tmpGEN Or tmpGEN = "*") Then '☆拡張子マッチのみ取りあえず対応
            Call Files.Add(Files.Count,tmpPath)
          End If
        End If
      Next
    End If
    
  End Sub
  
  'ファイル追加時に除外するパターンを指定
  Public Sub SetExclude(argPattern)
    Dim RegExp, PatIdx
    If argPattern <> "" Then
      If IsEmpty(Options("FFT_Exclude")) Then
        Set Options("FFT_Exclude") = CreateObject("Scripting.Dictionary")
      End If
      With Options("FFT_Exclude")
        PatIdx = .Count
        Call .Add(PatIdx, argPattern)
      End With
      ExcludeSet = True
    Else
      Call Options.Remove("FFT_Exclude")
      ExcludeSet = False
    End If
  End Sub
  
  'ファイル追加時に除外するパターンのファイルを読込
  Public Sub SetExcludeList(argFilePath)
    If argFilePath <> "" Then
      If InStr(argFilePath, "\") > 0 Then
        tmpExcludeList = argFilePath
      Else
        tmpExcludeList = GPN(FullName) & "\" & argFilePath
      End If
      If FileExists(tmpExcludeList) Then
        Set ELF = FSO.OpenTextFile(tmpExcludeList)
        Do Until ELF.AtEndOfStream
          Call SetExclude(ELF.ReadLine)
        Loop
        ELF.Close
      End If
    Else
      Call SetExclude("")
    End If
  End Sub
  
  Private Function IsExclude(argFilePath)
    If ExcludeSet Then
      Dim i
      With Options("FFT_Exclude")
        For i = 0 To .Count - 1
'          If M(argFilePath,.Item(i),PathMatchOption) Then
          If InStr(argFilePath,.Item(i)) > 0 Then
            IsExclude = True
            Exit Function
          End If
        Next
      End With
    End If
  End Function
  
  Public Function OpenMDBFile(argMDBFilePath)
    Set OpenMDBFile = CreateObject("ADODB.Connection")
    OpenMDBFile.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & argMDBFilePath & ";"
  End Function
  
  '前回処理時点より更新されているファイルのみ処理対象
  Public Sub SetModifyOnly(argDBPath)
    If ModifyOnlySet Then
      ModifyOnlyDB.Close
    End If
    ModifyOnlyDBPath = argDBPath
    If argDBPath <> "" Then
      'ファイルパス補完
      If InStr(ModifyOnlyDBPath,"\") = 0 Then
        ModifyOnlyDBPath = Path & "\" & ModifyOnlyDBPath
      End If
      'データベース接続
      Set ModifyOnlyDB = OpenMDBFile(ModifyOnlyDBPath)
      '最終チェック日時を更新
      Set Rs = ModifyOnlyDB.Execute("SELECT MODDATE FROM MODDB WHERE FILEPATH = '[最終チェック日時]'")
      If Rs.EOF Then
        Call ModifyOnlyDB.Execute("INSERT INTO MODDB(FILEPATH, MODDATE) VALUES('[最終チェック日時]', '" & Now & "')")
      Else
        Call ModifyOnlyDB.Execute("UPDATE MODDB SET MODDATE = '" & Now & "' WHERE FILEPATH = '[最終チェック日時]'")
      End If
      ModifyOnlySet = True
    Else
      ModifyOnlySet = False
    End If
  End Sub
  
  Private Function IsModifyOnly(argFilePath)
    If ModifyOnlySet Then
      'ファイルの最終変更日時取得
      dFileDate = CDate(GetFile(argFilePath).DateLastModified)
      '前回処理時点の最終変更日時取得
      Set Rs = ModifyOnlyDB.Execute("SELECT MODDATE FROM MODDB WHERE FILEPATH = '" & argFilePath & "'")
      If Rs.EOF Then
        dSaveDate = dFileDate - 1 '登録が無い場合はファイル日付より前にする
      Else
        dSaveDate = Rs("MODDATE")
      End If
      '比較してファイル最終変更日の方が新しければ処理対象
      If dFileDate > dSaveDate Then
        '処理時点の最終変更日時を保存
        If Rs.EOF Then
          Call ModifyOnlyDB.Execute("INSERT INTO MODDB(FILEPATH, MODDATE) VALUES('" & argFilePath & "', '" & dFileDate & "')")
        Else
          Call ModifyOnlyDB.Execute("UPDATE MODDB SET MODDATE = '" & dFileDate & "' WHERE FILEPATH = '" & argFilePath & "'")
        End If
        IsModifyOnly = True
      End If
    Else
      IsModifyOnly = True
    End If
  End Function
  
  Public Function OpenXMLFile(argXMLFilePath)
    Set OpenXMLFile = CreateObject("MSXML.DOMDocument")
    With OpenXMLFile
      .Async = False
      Call .Load(argXMLFilePath)
    End With
  End Function
  
  Public Sub LoadResourceFile(argResFilePath)
      Set RES = OpenXMLFile(argResFilePath)
  End Sub
  
  Public Function Resources(argResID)
    If IsEmpty(RES) Then
      Call LoadResourceFile(FullName)
    End If
    Set retNode = RES.selectSingleNode("//*[@id='" & argResID & "']")
    If Not retNode is Nothing Then
      Resources = retNode.Text
    Else
      Resources = ""
    End If
  End Function
  
  'ファイル一覧をソート
  Public Sub SetOrder(argItem, argOrder)
    Dim i, j, bCmp, sSwapFiles, sItem, sOrder
    sItem  = UCase(argItem)
    sOrder = UCase(argOrder)
    
    For i = 0 To Files.Count - 1
      j = Files.Count - 1
      Do While j > i
        Select Case sItem
        Case "CREATED"
          If GetFile(Files(j-1)).DateCreated _
             > GetFile(Files(j)).DateCreated Then
            bCmp = True
          Else
            bCmp = False
          End If
        Case "ACCESSED"
          If GetFile(Files(j-1)).DateLastAccessed _
             > GetFile(Files(j)).DateLastAccessed Then
            bCmp = True
          Else
            bCmp = False
          End If
        Case "MODIFIED"
          If GetFile(Files(j-1)).DateLastModified _
             > GetFile(Files(j)).DateLastModified Then
            bCmp = True
          Else
            bCmp = False
          End If
        Case "SIZE"
          If GetFile(Files(j-1)).Size _
             > GetFile(Files(j)).Size Then
            bCmp = True
          Else
            bCmp = False
          End If
        Case "EXT"
          If GEN(Files(j-1)) > GEN(Files(j)) Then
            bCmp = True
          ElseIf GEN(Files(j-1)) > GEN(Files(j)) Then
            bCmp = False
          Else
            If GBN(Files(j-1)) > GBN(Files(j)) Then
              bCmp = True
            Else
              bCmp = False
            End If
          End If
        Case "NAME"
          If GFN(Files(j-1)) > GFN(Files(j)) Then
            bCmp = True
          Else
            bCmp = False
          End If
        Case Else ' "PATH"
          If Files(j-1) > Files(j) Then
            bCmp = True
          Else
            bCmp = False
          End If
        End Select
        If sOrder = "DESC" Then
          bCmp = Not bCmp
        End If
        If bCmp Then
          sSwapFiles = Files(j)
          Files(j)   = Files(j-1)
          Files(j-1) = sSwapFiles
        End If
        j = j - 1
      Loop
    Next
  End Sub
  
  '出力ファイルのリネーム指定
  Public Sub SetRename(argPattern, argReplace)
    RenPtn = argPattern
    RenRep = argReplace
  End Sub
  
  '入力ファイルのバックアップ指定
  Public Sub SetBackup(argPattern, argReplace)
    BakPtn = argPattern
    BakRep = argReplace
  End Sub
  
  '=== ユーティリティ関数群 ===
  
  Public Function CreateObject(argClassName)
    Set CreateObject = WScript.CreateObject(argClassName)
  End Function
  
  Public Function GetWinFolder
    GetWinFolder = FSO.GetSpecialFolder(0)
  End Function
  
  Public Function GetSysFolder
    GetSysFolder = FSO.GetSpecialFolder(1)
  End Function
  
  Public Function GetTempFolder
    GetTempFolder = FSO.GetSpecialFolder(2)
  End Function
  
  Public Function GetTempFile
    GetTempFile = FSO.GetTempName()
  End Function
  
  Public Function GetTempName
    GetTempName = FSO.BuildPath(GetTempFolder,GetTempFile)
  End Function
  
  Public Function RegExpMatch(argTarget, argPattern, argOption)
    If InStr(UCase(Mode),"BASP21") > 0 Then
      'BASP21モード
      Call LoadOptionModule(BSP, "BASP21", True)
      
      If argTarget <> "" Then
        'グループ化を使うとマッチ文字列が返ってくるので・・
        'If BSP.Match("/" & argPattern & "/" & argOption, argTarget) = 1 Then
        If IsArray(BSP.MatchEx("/" & argPattern & "/" & argOption, argTarget)) Then
          If InStr(argOption,"n") > 0 Then '独自オプション n でマッチしないもののみ
            RegExpMatch = Not True
          Else
            RegExpMatch = True
          End If
        Else
          If InStr(argOption,"n") > 0 Then '独自オプション n でマッチしないもののみ
            RegExpMatch = Not False
          Else
            RegExpMatch = False
          End If
        End If
      End If
    Else
      If argPattern <> "" Then
        If IsEmpty(VRE) Then Set VRE = CreateObject("VBScript.RegExp")
        With VRE
          .Pattern = argPattern
          If InStr(argOption,"i") > 0 Then
            .IgnoreCase = True
          Else
            .IgnoreCase = False
          End If
          If InStr(argOption,"g") > 0 Then
            .Global = True
          Else
            .Global = False
          End If
          If InStr(argOption,"n") > 0 Then '独自オプション n でマッチしないもののみ
            RegExpMatch = Not .Test(argTarget)
          Else
            RegExpMatch = .Test(argTarget)
          End If
        End With
      Else
        If InStr(argOption,"n") > 0 Then '独自オプション n でマッチしないもののみ
          RegExpMatch = Not False
        Else
          RegExpMatch = False
        End If
      End If
    End If
  End Function
  
  Public Function M(argTarget, argPattern, argOption)
    M = RegExpMatch(argTarget, argPattern, argOption)
  End Function
  
  Public Function RegExpReplace(argTarget, argPattern, argReplace, argOption)
    If InStr(UCase(Mode),"BASP21") > 0 Then
      'BASP21モード
      Call LoadOptionModule(BSP, "BASP21", True)
      
      If argTarget <> "" Then
        RegExpReplace = BSP.Replace("s/" & argPattern & "/" & argReplace & "/" & argOption, argTarget)
      Else
        RegExpReplace = argTarget
      End If
    Else
      If IsEmpty(VRE) Then Set VRE = CreateObject("VBScript.RegExp")
      With VRE
        Dim sReplace
        If InStr(argReplace,"\") > 0 Then
'          '何故 VBScript.RegExpの置換文字は$なのだ・・てかperlも$だった\はsed方言・・
'          .IgnoreCase = False
'          .Global = True
'          '取りあえず \t と\数 は対応しとく
'          .Pattern = "\\t"
'          sReplace = .Replace(sReplace, vbTab)
'          '.Pattern = "\\(\d+)" 'これだと \\1 にマッチしてしまう
'          .Pattern = "[^\\]\\(\d+)" 'これも \2 を使うとメモリリークバグ
'          sReplace = .Replace(sReplace, "$$$1")
'          'エスケープも処理してくれないのか・・
'          .Pattern = "\\\."
'          sReplace = .Replace(sReplace, ".")
'          .Pattern = "\\\\"
'          sReplace = .Replace(sReplace, "\")
          '結局スキャンする事にする
          sReplace = VbsRegExpRepFix(argReplace)
        Else
          sReplace = argReplace
        End If
        .Pattern = argPattern
        If InStr(argOption,"i") > 0 Then
          .IgnoreCase = True
        Else
          .IgnoreCase = False
        End If
        If InStr(argOption,"g") > 0 Then
          .Global = True
        Else
          .Global = False
        End If
        '必要なら mオプション .MultiLineも追加 但しWSH5.5より
        RegExpReplace = .Replace(argTarget, sReplace)
      End With
    End If
  End Function
  
  Private Function VbsRegExpRepFix(argReplace)
    sTmp = ""
    n = Len(argReplace)
    bEsc = False
    For i = 1 To n
      c = Mid(argReplace,i,1)
      If bEsc Then
        Select Case c
        Case "1","2","3","4","5","6","7","8","9"
          sTmp = sTmp & "$" & c
        Case "t"
          sTmp = sTmp & vbTab
        Case "n"
          sTmp = sTmp & vbCrLf
        Case "r"
          sTmp = sTmp & vbCr
        Case "f"
          sTmp = sTmp & vbLf
        Case Else
          sTmp = sTmp & c
        End Select
        bEsc = False
      Else
        If c = "\" Then
          bEsc = True
        Else
          sTmp = sTmp & c
        End If
      End If
    Next
    VbsRegExpRepFix = sTmp
  End Function
  
  Public Function S(argTarget, argPattern, argReplace, argOption)
    S = RegExpReplace(argTarget, argPattern, argReplace, argOption)
  End Function
  
  Public Function GPN(argFilePath)
    GPN = FSO.GetParentFolderName(argFilePath)
  End Function
  
  Public Function GFN(argFilePath)
    GFN = FSO.GetFileName(argFilePath)
  End Function
  
  Public Function GBN(argFilePath)
    GBN = FSO.GetBaseName(argFilePath)
  End Function
  
  Public Function GEN(argFilePath)
    GEN = FSO.GetExtensionName(argFilePath)
  End Function
  
  Public Function GetParentFolderName(argFilePath)
    GetParentFolderName = GPN(argFilePath)
  End Function
  
  Public Function GetFileName(argFilePath)
    GetFileName = GFN(argFilePath)
  End Function
  
  Public Function GetBaseName(argFilePath)
    GetBaseName = GBN(argFilePath)
  End Function
  
  Public Function GetExtensionName(argFilePath)
    GetExtensionName = GEN(argFilePath)
  End Function
  
  Public Function FileExists(argFilePath)
    FileExists = FSO.FileExists(argFilePath)
  End Function
  
  Public Sub CreateFolder(argFolderPath)
    Dim sParentPath
    sParentPath = FSO.GetParentFolderName(argFolderPath)
    If FSO.FolderExists(sParentPath) Then
      If Not FSO.FolderExists(argFolderPath) Then
        FSO.CreateFolder argFolderPath
      End If
    Else
      Call CreateFolder(sParentPath)
      Call FSO.CreateFolder(argFolderPath)
    End If
  End Sub
  
  Public Function FolderExists(argFolderPath)
    FolderExists = FSO.FolderExists(argFolderPath)
  End Function
  
  Public Function GetFile(argFilePath)
    Set GetFile = FSO.GetFile(argFilePath)
  End Function
  
  Public Sub CopyFile(argSrc, argDst)
    Dim sParentPath
    sParentPath = FSO.GetParentFolderName(argDst)
    If Not FSO.FolderExists(sParentPath) Then
      Call CreateFolder(sParentPath)
    End If
    Call FSO.CopyFile(argSrc, argDst)
  End Sub
  
  Public Sub MoveFile(argSrc, argDst)
    Dim sParentPath
    sParentPath = FSO.GetParentFolderName(argDst)
    If Not FSO.FolderExists(sParentPath) Then
      Call CreateFolder(sParentPath)
    End If
    Call FSO.MoveFile(argSrc, argDst)
  End Sub
  
  Public Sub DeleteFile(argDel)
    Call FSO.DeleteFile(argDel)
  End Sub
  
  Public Function Split(sStr, sDelim)
    Dim vArray()
    Dim i
    Dim bLp
    Dim nPrevPos, nNextPos
    
    nPrevPos = 1: nNextPos = 0
    bLp = False
    
    Do
      nNextPos = InStr(nPrevPos, sStr, sDelim, vbBinaryCompare)
      
      If nNextPos = 0 Then: bLp = True: nNextPos = Len(sStr) + 1
      
      ReDim Preserve vArray(i)
      vArray(i) = CStr(Mid(sStr, nPrevPos, nNextPos - nPrevPos))
      nPrevPos = nNextPos + Len(sDelim)
      i = i + 1
    Loop Until bLp = True
    
    Split = vArray
  End Function
  
  Public Function Sort(argArray)
    Dim n, i, j, SortedArray(), SwapValue
    n = Ubound(argArray)
    Redim SortedArray(n)
    For i = 0 To n
      SortedArray(i) = argArray(i)
    Next
    For i = 0 To n
      j = n
      Do While j > i
        If SortedArray(j-1) > SortedArray(j) Then
          SwapValue        = SortedArray(j)
          SortedArray(j)   = SortedArray(j-1)
          SortedArray(j-1) = SwapValue
        End If
        j = j - 1
      Loop
    Next
    Sort = SortedArray
  End Function
  
  Public Function Tokenize(sStr)
    Dim vArray(), sChar, sToken, bLit
    n = 0
    ReDim vArray(n)
    For i = 1 To Len(sStr)
      sChar = Mid(sStr,i,1)
      Select Case sChar
      Case " "
        If bLit Then
          sToken = sToken & sChar
        Else
          If sToken <> "" Then
            n = n + 1
            ReDim Preserve vArray(n)
            vArray(n) = sToken
            sToken = ""
          End If
        End If
      Case "'", """"
        sToken = sToken & sChar
        bLit = Not bLit
      Case Else
        sToken = sToken & sChar
      End Select
    Next
    If sToken <> "" Then
      n = n + 1
      ReDim Preserve vArray(n)
      vArray(n) = sToken
    End If
    Tokenize = vArray
  End Function
  
  Function LenB(argStr)
    Dim nLen, sChar, i
    nLen = 0
    For i = 1 To Len(argStr)
      sChar = Mid(argStr, i, 1)
      If (Asc(sChar) And &HFF00) = 0 Then
        nLen = nLen + 1
      Else
        nLen = nLen + 2
      End If
    Next
    LenB = nLen
  End Function
  
  Function MidB(ByVal strSjis, ByVal lngStartPos, ByVal lngGetByte)
  ' strSjis:      切り出す文字列
  ' lngStartPos:  開始位置
  ' lngGetByte:   取得バイト数（"" 時は lngStartPos 以降全て）
  ' blnZenFlag:   全角文字が区切り位置で分割するときの動作
  '               True= スペースに変換, False= そのまま出力
  
  blnZenFlag = True
  
      Dim lngByte             ' バイト数
      Dim lngLoop             ' ループカウンタ
      Dim strChkBuff          ' 確認用バッファ
      Dim strLastByte         ' 最終バイト

      On Error Resume Next

      MidB = ""
      If lngGetByte = "" Then
          ' 最大文字数をセットしておく
          lngGetByte = Len(strSjis) * 2
      End If
      lngGetByte = CLng(lngGetByte)

      ' 開始位置
      lngByte = 0
      For lngLoop = 1 To Len(strSjis)
          strChkBuff = Mid(strSjis, lngLoop, 1)
          If (Asc(strChkBuff) And &HFF00) = 0 Then
              lngByte = lngByte + 1
          Else
              lngByte = lngByte + 2
              ' 全角の２バイト目が開始位置のとき
              If lngByte = lngStartPos Then
                  If blnZenFlag = True Then
                      MidB = " "
                  Else
                      MidB = Asc(strChkBuff) And &H00FF
                      If MidB < 0 Then
                          MidB = 256 + MidB
                      End If
                      MidB = ChrB(MidB)
                  End If
                  lngLoop = lngLoop + 1
              End If
          End If
          If lngByte >= lngStartPos Then
              Exit For
          End If
      Next

      ' 取得
      lngByte = LenB(MidB)
      If lngByte < lngGetByte Then
          For lngLoop = lngLoop To Len(strSjis)
              strChkBuff = Mid(strSjis, lngLoop, 1)
              MidB = MidB & strChkBuff
              If (Asc(strChkBuff) And &HFF00) = 0 Then
                  lngByte = lngByte + 1
              Else
                  lngByte = lngByte + 2
              End If
              If lngByte >= lngGetByte Then
                  Exit For
              End If
          Next
      End If

      lngByte = LenAscByte(MidB)
      If lngByte > lngGetByte Then
          ' 終端が全角１バイト目のとき。意味ないかも（笑）
          If blnZenFlag = True Then
              MidB = Mid(MidB, 1, Len(MidB) - 1) & " "
          Else
              strLastByte = Fix((Asc(Right(MidB, 1)) And &HFF00) / 256)
              If strLastByte < 0 Then
                  strLastByte = 256 + strLastByte
              End If
              MidB = Mid(MidB, 1, Len(MidB) - 1) & ChrB(strLastByte)
          End If
      End If
  End Function

  Function LeftB(ByVal strSjis, ByVal lngGetByte)
  
  blnZenFlag = True
  
      LeftB = MidB(strSjis, 1, lngGetByte)
  End Function

  Function RightB(ByVal strSjis, ByVal lngGetByte)
  
  blnZenFlag = True
  
      RightB = StrReverse(strSjis)
      RightB = MidB(RightB, 1, lngGetByte)
      RightB = StrReverse(RightB)
  End Function
  
  Function LPAD(argStr, argLen)
    Dim nLen
    nLen = LenB(argStr)
    If nLen < argLen Then
      LPAD = Space(argLen-nLen) & argStr
    Else
      LPAD = argStr
    End If
  End Function
  
  Function RPAD(argStr, argLen)
    Dim nLen
    nLen = LenB(argStr)
    If nLen < argLen Then
      RPAD = argStr & Space(argLen-nLen)
    Else
      RPAD = argStr
    End If
  End Function
  
  '=== 組み込みフィルタ関数 ===
  
  Public Function FFT_Grep(argPtn, argOpt)
    Set FFT_Grep = New FFTGrep
    With FFT_Grep
      .Ptn = argPtn
      .Opt = argOpt
    End With
  End Function
  
  Public Function FFT_Sed(argPtn, argRep, argOpt)
    Set FFT_Sed = New FFTSed
    With FFT_Sed
      .Ptn = argPtn
      .Rep = argRep
      .Opt = argOpt
    End With
  End Function
  
  Public Function FFT_Tr(argPtn, argRep, argOpt)
    Set FFT_Tr = New FFTTr
    With FFT_Tr
      .Re = "tr/" & argPtn & "/" & argRep & "/" & argOpt
    End With
  End Function
  
  Public Function FFT_Sort(argOpt)
    Set FFT_Sort = New FFTSort
    With FFT_Sort
      .Opt = Replace(argOpt,"-","/")
    End With
  End Function
  
  Public Function FFT_Uniq(argOpt)
    Set FFT_Uniq = New FFTUniq
    With FFT_Uniq
      .Opt = argOpt
    End With
  End Function
  
  Public Function FFT_Cut(argOpt)
    Set FFT_Cut = New FFTCut
    With FFT_Cut
      .Opt = argOpt
    End With
  End Function
  
  Public Function FFT_Empty()
    Set FFT_Empty = New FFTEmpty
  End Function
  
  Public Function FFT_Contains(argPtn, argOpt)
    Set FFT_Contains = New FFTContains
    With FFT_Contains
      .Ptn = argPtn
      .Opt = argOpt
    End With
  End Function
End Class

'=== 組み込みフィルタ ===

Class FFTGrep
  Public FilterName
  Public FilterType
  Public Ptn
  Public Opt
  Private Optv
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Grep"
      FilterType = "Line"
      If InStr(Opt,"v") <> 0 Then
        Optv = True
      Else
        Optv = False
      End If
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイルオープン
  Public Function OpenFile(argFFT, argImpFile, argWrkFile)
    OpenFile = True
  End Function
  'ファイルクローズ
  Public Function CloseFile(argFFT, argImpFile, argWrkFile)
    CloseFile = True
  End Function
  '行処理
  Public Function ProcessLine(argFFT, argLine)
    ProcessLine = argFFT.M(argLine, Ptn, Opt)
    If Optv Then
        ProcessLine = Not ProcessLine
    End If
  End Function
End Class

Class FFTSed
  Public FilterName
  Public FilterType
  Public Ptn
  Public Rep
  Public Opt
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Sed"
      FilterType = "File"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイル処理
  Public Function ProcessFile(argFFT, argImpFile, argWrkFile)
    If InStr(Ptn,"\n") = 0 Then
      sPtn = Ptn
      bWL = True
    Else
      sPtn = Replace(Ptn,"\n","")
      If sPtn = "" Then
        sPtn = "$"
      End If
      bWL = False
    End If
    Set ImpFile = OpenReadTextFile(argImpFile)
    Set WrkFile = OpenWriteTextFile(argWrkFile)
    If bWL Then
      Do Until ImpFile.AtEndOfStream
        Call WrkFile.WriteLine(argFFT.S(ImpFile.ReadLine, sPtn, Rep, Opt))
      Loop
    Else
      Do Until ImpFile.AtEndOfStream
        Call WrkFile.Write(argFFT.S(ImpFile.ReadLine, sPtn, Rep, Opt))
      Loop
    End If
    ImpFile.Close
    WrkFile.Close
    ProcessFile = True
  End Function
End Class

Class FFTTr
  Public FilterName
  Public FilterType
  Public Re
  Private Tr
  
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Tr"
      FilterType = "Line"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイルオープン
  Public Function OpenFile(argFFT, argImpFile, argWrkFile)
    OpenFile = True
  End Function
  'ファイルクローズ
  Public Function CloseFile(argFFT, argImpFile, argWrkFile)
    CloseFile = True
  End Function
  '行処理
  Public Function ProcessLine(argFFT, argLine)
    If argLine <> "" Then
      Call argFFT.BSP.Translate (Re, argLine, Tr)
      argLine     = Tr
    End If
    ProcessLine = True
  End Function
End Class

Class FFTSort
  Public FilterName
  Public FilterType
  Public Opt
  
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Sort"
      FilterType = "Command"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'コマンド処理
  Public Function ProcessCommand(argFFT, argImpFile, argWrkFile)
    With argFFT
      sCommand = "SORT " & Opt & " """ & argImpFile & """"
      Call .RunCommand(sCommand, argWrkFile)
    End With
    ProcessCommand = True
  End Function
End Class

Class FFTUniq
  Public FilterName
  Public FilterType
  Public Opt
  Private PrevLine, Count, OutputLine, IsFirst
  
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Uniq"
      FilterType = "Line"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイルオープン
  Public Function OpenFile(argFFT, argImpFile, argWrkFile)
    Count = 0
    OpenFile = True
    IsFirst = True
  End Function
  'ファイルクローズ
  Public Function CloseFile(argFFT, argImpFile, argWrkFile)
    If Not IsEmpty(PrevLine) Then
      Set oWrkFile = OpenTextFile(argWrkFile, OTF_APPEND)
      If InStr(Opt,"c") <> 0 Then
        OutputLine = Right("   " & Count,4) & " " & PrevLine
      Else
        OutputLine = PrevLine
      End If
      oWrkFile.WriteLine OutputLine
      oWrkFile.Close
    End If
    CloseFile = True
  End Function
  '行処理
  Public Function ProcessLine(argFFT, argLine)
    If PrevLine <> argLine And Not IsFirst Then
      If InStr(Opt,"c") <> 0 Then
        OutputLine = Right("   " & Count,4) & " " & PrevLine
      Else
        OutputLine = PrevLine
      End If
      PrevLine = argLine
      argLine = OutputLine
      ProcessLine = True
      Count = 0
    Else
      IsFirst = False
      PrevLine = argLine
    End If
    Count = Count + 1
  End Function
End Class

Class FFTCut 'まだ未実装
  Public FilterName
  Public FilterType
  Public Opt
  Private Mode, arCol
  
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Cut"
      FilterType = "Line"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイルオープン
  Public Function OpenFile(argFFT, argImpFile, argWrkFile)
    OpenFile = True
  End Function
  'ファイルクローズ
  Public Function CloseFile(argFFT, argImpFile, argWrkFile)
    CloseFile = True
  End Function
  '行処理
  Public Function ProcessLine(argFFT, argLine)
    ProcessLine = True
  End Function
End Class

Class FFTEmpty
  Public FilterName
  Public FilterType
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    FilterName = "FFTEmpty"
    FilterType = "File"
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイル処理
  Public Function ProcessFile(argFFT, argImpFile, argWrkFile)
    '何も処理せず入力ファイルを出力ファイルとする
    argWrkFile = argImpFile
    ProcessFile = True
  End Function
End Class

Class FFTContains
  Public FilterName
  Public FilterType
  Public Ptn
  Public Opt
  Private Hit
  'フィルタ初期処理
  Public Sub Initialize(argFFT)
    With argFFT
      FilterName = "FFT_Contains"
      FilterType = "File"
    End With
  End Sub
  'フィルタ終了処理
  Public Sub Terminate(argFFT)
  End Sub
  'ファイル処理
  Public Function ProcessFile(argFFT, argImpFile, argWrkFile)
    '入力ファイルに指定した正規表現文字列が含まれている場合のみ通す
    Hit= False
    Set ImpFile = OpenReadTextFile(argImpFile)
    Do Until ImpFile.AtEndOfStream
      If argFFT.M(ImpFile.ReadLine, Ptn, Opt) Then
        Hit = True
        Exit Do
      End If
    Loop
    ImpFile.Close
    If Hit Then
      argWrkFile = argImpFile
      ProcessFile = True
    Else
      ProcessFile = False
    End If
  End Function
End Class

'===== IEを利用したプロセス間通信 ======================================
Class IE_IPC
  Private IE
  Public  Port, Retry
  
  Private Sub Class_Initialize
    Retry = 3
  End Sub
  
  Public Sub Bind(argServiceName)
    Port = argServiceName
  End Sub
  
  Public Sub Listen
    Set IE = CreateObject("InternetExplorer.Application")
    If Port = "" Then
      Port = IE.hWnd
    End If
    Call IE.PutProperty("IE_IPC:" & Port, IE)
  End Sub
  
  Public Function Connect(argPort)
    Port = argPort
    Set oShell = CreateObject("Shell.Application")
    nTry = 0
    Do While nTry < Retry Or Retry = 0
      For Each oProc In oShell.Windows()
        If Not IsEmpty(oProc.GetProperty("IE_IPC:" & Port)) Then
          Set IE = oProc.GetProperty("IE_IPC:" & Port)
          Connect = True
          Exit Function
        End If
      Next
      WScript.Sleep 100
      nTry = nTry + 1
    Loop
    Connect = False
  End Function
  
  Public Sub Open(argPort)
    If Not .Connect(argPort) Then
      Call .Bind(argPort)
      Call .Listen
    End If
  End Sub
  
  Public Sub Close
    Call IE.Quit
  End Sub
  
  Public Property Get Item(argKey)
    If typename(IE) = "IWebBrowser2" Then
      If IsObject(IE.GetProperty(argKey)) Then
        Set Item = IE.GetProperty(argKey)
      Else
        Item = IE.GetProperty(argKey)
      End If
    Else
      MsgBox "接続されていません"
    End If
  End Property
  
  Public Property Let Item(argKey, argValue)
    If typename(IE) = "IWebBrowser2" Then
      Call IE.PutProperty(argKey, argValue)
    Else
      MsgBox "接続されていません"
    End If
  End Property
  
  Public Property Set Item(argKey, argValue)
    If typename(IE) = "IWebBrowser2" Then
      Call IE.PutProperty(argKey, argValue)
    Else
      MsgBox "接続されていません"
    End If
  End Property
End Class

'===== テキストファイルクラス ======================================
Class TextFile
  Public FSO, File, Path, Name, IsOpen
  
  Public Sub CreateTextFile(argPath, argOverWrite)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If argPath <> "" Then
      Set File = FSO.CreateTextFile(argPath, argOverWrite)
      Path = argPath
      Name = FSO.GetFileName(argPath)
      IsOpen = True
    End If
  End Sub
  
  Public Sub OpenTextFile(argPath, argMode)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If argPath <> "" Then
      Set File = FSO.OpenTextFile(argPath, argMode, True)
      Path = argPath
      Name = FSO.GetFileName(argPath)
      IsOpen = True
    End If
  End Sub
  
  Public Function AtEndOfLine
    If IsOpen Then
      AtEndOfLine = File.AtEndOfLine
    Else
      AtEndOfLine = True
    End If
  End Function
  
  Public Function AtEndOfStream
    If IsOpen Then
      AtEndOfStream = File.AtEndOfStream
    Else
      AtEndOfStream = True
    End If
  End Function
  
  Public Function Column
    If IsOpen Then
      Column = File.Column
    Else
      Column = 0
    End If
  End Function
  
  Public Function Line
    If IsOpen Then
      Line = File.Line
    Else
      Line = 0
    End If
  End Function
  
  Public Sub Close
    If IsOpen Then
      File.Close
      IsOpen = False
    End If
  End Sub
  
  Public Function Read(argCharN)
    If IsOpen Then
      Read = File.Read(argCharN)
    Else
      Read = ""
    End If
  End Function
  
  Public Function ReadAll
    If IsOpen Then
      ReadAll = File.ReadAll
    Else
      ReadAll = ""
    End If
  End Function
  
  Public Function ReadLine
    If IsOpen Then
      ReadLine = File.ReadLine
    Else
      ReadLine = ""
    End If
  End Function
  
  Public Sub Skip(argCharN)
    If IsOpen Then
      Call File.Skip(argCharN)
    End If
  End Sub
  
  Public Sub SkipLine
    If IsOpen Then
      File.SkipLine
    End If
  End Sub
  
  Public Sub Write(argStr)
    If IsOpen Then
      Call File.Write(argStr)
    End If
  End Sub
  
  Public Sub WriteBlankLines
    If IsOpen Then
      File.WriteBlankLines
    End If
  End Sub
  
  Public Sub WriteLine(argLine)
    If IsOpen Then
      Call File.WriteLine(argLine)
    End If
  End Sub
  
  Private Sub Class_Terminate
    If IsOpen Then
      File.Close
    End If
  End Sub
End Class

Function CreateTextFile(argPath, argOverWrite)
  Set CreateTextFile = new TextFile
  With CreateTextFile
    Call .CreateTextFile(argPath, argOverWrite)
  End With
End Function

Const OTF_READ   = 1
Const OTF_WRITE  = 2
Const OTF_APPEND = 8

Function OpenTextFile(argPath, argMode)
  Set OpenTextFile = new TextFile
  With OpenTextFile
    Call .OpenTextFile(argPath, argMode)
  End With
End Function

Function OpenReadTextFile(argPath)
  Set OpenReadTextFile = OpenTextFile(argPath, OTF_READ)
End Function

Function OpenWriteTextFile(argPath)
  Set OpenWriteTextFile = OpenTextFile(argPath, OTF_WRITE)
End Function

Function OpenAppendTextFile(argPath)
  Set OpenAppendTextFile = OpenTextFile(argPath, OTF_APPEND)
End Function

'===================================================================================================
