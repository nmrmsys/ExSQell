
'�L�[������Ԏ擾�p
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_CAPITAL = &H14
Public Const VK_INSERT = &H2D
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
'NT�nOS�̂ݗL��
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

'���T�C�Y�\���[�U�[�t�H�[���p
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000 '�ő剻�{�^��
Public Const WS_MINIMIZEBOX = &H20000 '�ŏ����{�^��
Public Const WS_THICKFRAME = &H40000  '�T�C�Y�ύX

'�E�C���h�E�\���p
Public Const SW_MAXIMIZE = 3


#If VBA7 Then

'�E�C���h�E�^�C�g���ݒ�p
Public Declare PtrSafe Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As LongPtr, ByVal lpString As String) As Long
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

'�L�[������Ԏ擾�p
Public Declare PtrSafe Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

'���T�C�Y�\���[�U�[�t�H�[���p
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

'�E�C���h�E�\���p
Public Declare PtrSafe Function ShowWindow Lib "user32" _
        (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long

'SQL���`
Public Declare PtrSafe Function iCopyToPreparedStr Lib "SqlFmt" _
    (ByVal szOrgStmt As String, ByVal szIdtStr As String, _
     ByVal iMinNumComma As Long, ByVal iMinPartSize As Long, _
     ByVal szFmtStmt As String, ByRef iStrLen As Long, _
     ByVal iFmtStmt As Long) As Long

#Else

'�E�C���h�E�^�C�g���ݒ�p
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" _
    (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'�L�[������Ԏ擾�p
Public Declare Function GetKeyState Lib "user32" _
    (ByVal nVirtKey As Long) As Integer

'���T�C�Y�\���[�U�[�t�H�[���p
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

'�E�C���h�E�\���p
Public Declare Function ShowWindow Lib "user32" _
        (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'SQL���`
Public Declare Function iCopyToPreparedStr Lib "SqlFmt" _
    (ByVal szOrgStmt As String, ByVal szIdtStr As String, _
     ByVal iMinNumComma As Long, ByVal iMinPartSize As Long, _
     ByVal szFmtStmt As String, ByRef iStrLen As Long, _
     ByVal iFmtStmt As Long) As Long

#End If

Public Function SqlFmt(argSql As String) As String
    Dim szFmtStmt As String       ' ���`��̕�
    Dim iFmtStmt As Long          ' ���`��̕����i�[����o�b�t�@�̍ő咷
    Dim iStrLen As Long           ' ���`��̕��Ɋi�[���ꂽ�o�C�g��
    Dim iRet As Long              '�C���|�[�g�֐��̖߂�l
                                  '0: ����I�� ��0: �ُ�I��
    ' �ϐ��̏�����
    iFmtStmt = 5000
    szFmtStmt = String(iFmtStmt, vbNullChar)

    ' ���`�̎��s
    On Error GoTo dll_err
    iRet = iCopyToPreparedStr(argSql, "  ", 14, 50, szFmtStmt, iStrLen, iFmtStmt)

    ' �C���|�[�g�֐��̃X�e�[�^�X�����|�[�g
    If iRet = 0 Then
        SqlFmt = szFmtStmt
    Else
        MsgBox "���̐��`�Ɏ��s���܂����B�G���[�ԍ��F" & Str(iRet), vbOKOnly, "SqlFmt"
    End If
    Exit Function
dll_err:
    MsgBox "SqlFmt.dll���C���X�g�[�����ĉ�����" & vbCrLf & "http://www.vector.co.jp/soft/winnt/business/se299208.html"
End Function
