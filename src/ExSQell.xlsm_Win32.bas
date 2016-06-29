
'キー押下状態取得用
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_CAPITAL = &H14
Public Const VK_INSERT = &H2D
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
'NT系OSのみ有効
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

'リサイズ可能ユーザーフォーム用
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000 '最大化ボタン
Public Const WS_MINIMIZEBOX = &H20000 '最小化ボタン
Public Const WS_THICKFRAME = &H40000  'サイズ変更

'ウインドウ表示用
Public Const SW_MAXIMIZE = 3


#If VBA7 Then

'ウインドウタイトル設定用
Public Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

'キー押下状態取得用
Public Declare PtrSafe Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

'リサイズ可能ユーザーフォーム用
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
      (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
      (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long

Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                       (ByVal hwnd As LongPtr, ByVal lpOperation As String, _
                        ByVal lpFile As String, ByVal lpParameters As String, _
                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr

'ウインドウ表示用
Public Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

'SQL整形
Public Declare PtrSafe Function iCopyToPreparedStr Lib "SqlFmt" _
    (ByVal szOrgStmt As String, ByVal szIdtStr As String, _
     ByVal iMinNumComma As Long, ByVal iMinPartSize As Long, _
     ByVal szFmtStmt As String, ByRef iStrLen As Long, _
     ByVal iFmtStmt As Long) As Long

#Else

'ウインドウタイトル設定用
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'キー押下状態取得用
Public Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

'リサイズ可能ユーザーフォーム用
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
      (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
      (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                       (ByVal hwnd As Long, ByVal lpOperation As String, _
                        ByVal lpFile As String, ByVal lpParameters As String, _
                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'ウインドウ表示用
Public Declare Function ShowWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'SQL整形
Public Declare Function iCopyToPreparedStr Lib "SqlFmt" _
    (ByVal szOrgStmt As String, ByVal szIdtStr As String, _
     ByVal iMinNumComma As Long, ByVal iMinPartSize As Long, _
     ByVal szFmtStmt As String, ByRef iStrLen As Long, _
     ByVal iFmtStmt As Long) As Long

#End If

Public Function SqlFmt(argSql As String) As String
    Dim szFmtStmt As String       ' 整形後の文
    Dim iFmtStmt As Long          ' 整形後の文を格納するバッファの最大長
    Dim iStrLen As Long           ' 整形後の文に格納されたバイト数
    Dim iRet As Long              'インポート関数の戻り値
                                  '0: 正常終了 非0: 異常終了
    ' 変数の初期化
    iFmtStmt = 5000
    szFmtStmt = String(iFmtStmt, vbNullChar)

    ' 整形の実行
    On Error GoTo dll_err
    iRet = iCopyToPreparedStr(argSql, "  ", 14, 50, szFmtStmt, iStrLen, iFmtStmt)

    ' インポート関数のステータスをレポート
    If iRet = 0 Then
        SqlFmt = szFmtStmt
    Else
        MsgBox "文の整形に失敗しました。エラー番号：" & Str(iRet), vbOKOnly, "SqlFmt"
    End If
    Exit Function
dll_err:
    MsgBox "SqlFmt.dllをインストールして下さい" & vbCrLf & "http://www.vector.co.jp/soft/winnt/business/se299208.html"
End Function
