
Public State As Boolean
Private cbBtnMode As Boolean

Private coEditMenu As CommandBar
Private coEditUndo As CommandBarButton
Private coEditRedo As CommandBarButton
Private coEditCopy As CommandBarButton
Private coEditCut As CommandBarButton
Private coEditPaste As CommandBarButton
Private coEditDelete As CommandBarButton
Private coEditSelAll As CommandBarButton

Private Sub UserForm_Initialize()
    Dim nHwnd As Long
    Dim nWlng As Long
'    myHWnd = FindWindow("ThunderDFrame", Me.Caption)
    nHwnd = FindWindow(vbNullString, Me.Caption)
    nWlng = GetWindowLong(nHwnd, GWL_STYLE)
    SetWindowLong nHwnd, GWL_STYLE, nWlng Or WS_MAXIMIZEBOX Or WS_THICKFRAME
    Call DrawMenuBar(nHwnd)
    On Error Resume Next
    Set coEditMenu = Application.CommandBars(Me.Name)
    On Error GoTo 0
    If Not coEditMenu Is Nothing Then
        coEditMenu.Delete
    End If
    Set coEditMenu = Application.CommandBars.Add(Name:=Me.Name, Position:=msoBarPopup, Temporary:=True)
    With coEditMenu
        Set coEditUndo = .Controls.Add(msoControlButton)
        With coEditUndo
            .Tag = "Undo"
            .Caption = "元に戻す(&U)"
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditRedo = .Controls.Add(msoControlButton)
        With coEditRedo
            .Tag = "Redo"
            .Caption = "やり直し(&R)"
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditCut = .Controls.Add(msoControlButton)
        With coEditCut
            .Tag = "Cut"
            .BeginGroup = True
            .Caption = "切り取り(&T)"
            .FaceId = 21
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditCopy = .Controls.Add(msoControlButton)
        With coEditCopy
            .Tag = "Copy"
            .Caption = "コピー(&C)"
            .FaceId = 19
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditPaste = .Controls.Add(msoControlButton)
        With coEditPaste
            .Tag = "Paste"
            .Caption = "貼り付け(&P)"
            .FaceId = 22
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditDelete = .Controls.Add(msoControlButton)
        With coEditDelete
            .Tag = "Delete"
            .Caption = "削除(&D)"
            .OnAction = "ST_WIN_OnAction"
        End With
        Set coEditSelAll = .Controls.Add(msoControlButton)
        With coEditSelAll
            .Tag = "SelAll"
            .BeginGroup = True
            .Caption = "すべて選択(&A)"
            .OnAction = "ST_WIN_OnAction"
        End With
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub UserForm_Resize()
    txtMAIN.Width = Me.InsideWidth
    If cbBtnMode Then
        txtMAIN.Height = Me.InsideHeight - 30
        btnDefault.Top = Me.InsideHeight - 26.5
        btnDefault.Left = Me.InsideWidth - 156
        btnCancel.Top = Me.InsideHeight - 26.5
        btnCancel.Left = Me.InsideWidth - 78
    Else
        txtMAIN.Height = Me.InsideHeight
    End If
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnDefault_Click()
    State = True
    Me.Hide
End Sub

Private Sub txtMAIN_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
    Case vbKeyTab
        With txtMAIN
            If .SelLength > 0 Then
                Dim nSelStart, sSelText, arLines, i
                nSelStart = .SelStart
                arLines = Split(.SelText, vbCr)
                If Not (Shift And 1) > 0 Then
                    For i = 0 To UBound(arLines)
                        If arLines(i) <> "" Then
                            arLines(i) = "  " & arLines(i)
                        End If
                    Next
                Else
                    For i = 0 To UBound(arLines)
                        If Left(arLines(i), 1) = " " Then
                            arLines(i) = Mid(arLines(i), 2)
                        End If
                        If Left(arLines(i), 1) = " " Then
                            arLines(i) = Mid(arLines(i), 2)
                        End If
                    Next
                End If
                For i = 0 To UBound(arLines)
                    If i > 0 Then sSelText = sSelText & vbCr
                    sSelText = sSelText & arLines(i)
                Next
                .SelText = sSelText
                .SelStart = nSelStart
                .SelLength = Len(sSelText)
            Else
                btnDefault.SetFocus
            End If
            KeyCode = 0
        End With
    Case vbKeyEscape
        If Not cbBtnMode Then
            Hide
            KeyCode = 0
        End If
    Case vbKeyReturn
        If (Shift And 1) > 0 Or (Shift And 2) > 0 Or (Shift And 4) > 0 Then
            State = True
            Hide
            KeyCode = 0
        End If
    End Select
End Sub

Private Sub txtMAIN_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = xlSecondaryButton Then
        coEditUndo.Enabled = CanUndo
        coEditRedo.Visible = CanRedo
        If txtMAIN.SelLength > 0 Then
            coEditCopy.Enabled = True
            coEditCut.Enabled = True
            coEditDelete.Enabled = True
        Else
            coEditCopy.Enabled = False
            coEditCut.Enabled = False
            coEditDelete.Enabled = False
        End If
        coEditPaste.Enabled = txtMAIN.CanPaste
        If Len(txtMAIN.Text) > 0 Then
            coEditSelAll.Enabled = True
        Else
            coEditSelAll.Enabled = False
        End If
        coEditMenu.ShowPopup
    End If
End Sub

Public Sub OnAction()
    With Application.CommandBars.ActionControl
        Select Case .Tag
        Case "Undo"
            UndoAction
        Case "Redo"
            RedoAction
        Case "Cut"
            SendKeys "^{x}"
        Case "Copy"
            SendKeys "^{c}"
        Case "Paste"
            SendKeys "^{v}"
        Case "Delete"
            SendKeys "{DELETE}"
        Case "SelAll"
            txtMAIN.SelStart = 0
            txtMAIN.SelLength = Len(txtMAIN.Text)
        End Select
    End With
End Sub

Public Function ShowModal2(argText, Optional argBtnCap, Optional argOpts = "")
    If IsMissing(argBtnCap) Then
        cbBtnMode = False
        btnDefault.Visible = False
        btnCancel.Visible = False
        txtMAIN.Height = Me.InsideHeight
    Else
        cbBtnMode = True
        btnDefault.Caption = argBtnCap
        btnDefault.Visible = True
        btnCancel.Visible = True
        txtMAIN.Height = Me.InsideHeight - 30
    End If
    If Not cbBtnMode Or InStr(UCase(argOpts), "EDIT") <> 0 Then
        txtMAIN.SetFocus
    Else
        btnDefault.SetFocus
    End If
    txtMAIN = argText
    txtMAIN.SelStart = 0
    State = False
    cbSetMode = True
    Show
    If State Then
        argText = txtMAIN.Text
    End If
    ShowModal2 = State
    StartUpPosition = 0
End Function

