
Private IItms As New Collection
Private IKeys As New Collection

Public Property Get Count()
    Count = IItms.Count
End Property

Public Property Get Exists(Key) As Boolean
    If Not TypeName(Me.Item(Key)) = "Empty" Then
        Exists = True
    End If
End Property

Public Sub Remove(Key)
    If TypeName(Key) = "String" Then
        Dim i
        On Error Resume Next
        Call IItms.Remove(Key)
        Call IKeys.Remove(Key)
        '�R���N�V�������v���p�e�B��Name�������w��L�[�̏ꍇ�͍폜
        For i = 1 To IItms.Count
            If InStr("Collection,Prop", TypeName(IItms.Item(i))) <> 0 Then
                If IItms.Item(i).Item("Name") = Key Then
                    Call IItms.Remove(i)
                    Call IKeys.Remove(i)
                    Exit For
                End If
            End If
        Next
        On Error GoTo 0
    Else
        Call IItms.Remove(Key)
        Call IKeys.Remove(Key)
    End If
End Sub

Public Property Get Item(Optional Key, Optional RepFlg = True)
    
    '�L�[�w�肪�����ꍇ�͂ǂ����邩...���̂����l����
    'RepFlg�� Let/Set�Ɗ֐���`�����킹��ׂ̃_�~�[
    
    On Error Resume Next
    If IsObject(IItms.Item(Key)) Then
        Set Item = IItms.Item(Key)
    Else
        Item = IItms.Item(Key)
    End If
    On Error GoTo 0

End Property

Public Property Let Item(Optional Key, Optional RepFlg = True, Value)
    
    If IsMissing(Key) Then
        '�L�[�w�肪�����ꍇ�͍Ō�ɒǉ�
        Call setItem("", Value)
    Else
        If IsMissing(RepFlg) Then
            Call setItem(Key, Value)
        Else
            Call setItem(Key, Value, RepFlg)
        End If
    End If

End Property

Public Property Set Item(Optional Key, Optional RepFlg = True, Value)
    If IsMissing(Key) Then
        '�L�[�w�肪�����ꍇ�͍Ō�ɒǉ�
        Call setItem("", Value)
    Else
        If IsMissing(RepFlg) Then
            Call setItem(Key, Value)
        Else
            Call setItem(Key, Value, RepFlg)
        End If
    End If
End Property

Private Function setItem(Key, Value, Optional RepFlg = True)
    Dim i As Integer
    
    If TypeName(Key) = "String" Then
        '�L�[�w��
        If RepFlg Then
            '�u���w��
            If Key <> "" Then
                On Error Resume Next
                Call IItms.Remove(Key)
                Call IKeys.Remove(Key)
                On Error GoTo 0
                Call IItms.Add(Value, Key)
                Call IKeys.Add(Key, Key)
            Else
                Call IItms.Add(Value)
                Call IKeys.Add(IItms.Count)
            End If
        Else
            '�ǉ��w��
            '�w��L�[�̈ʒu����������
            '���̌��ɒǉ�����\��
            'Call IItms.Add(Value, Key)
            'Call IKeys.Add(Key, Key)
            MsgBox "������"
        End If
    Else
        '�Y���w��
        If IItms.Count < Key Then
            '�w���菭�Ȃ��ꍇ�A�󔒍��ڂ�ǉ�
            For i = IItms.Count To Key - 2
                Call IItms.Add("")
            Next
        End If
        If RepFlg Then
            '�u���w��
            On Error Resume Next
            Call IItms.Remove(Key)
            On Error GoTo 0
            If IItms.Count < Key Then
                Call IItms.Add(Value)
            Else
                Call IItms.Add(Value, before:=Key)
            End If
        Else
            '�ǉ��w��
            If Key = 0 Then
                If IItms.Count = 0 Then
                    Call IItms.Add(Value)
                Else
                    Call IItms.Add(Value, before:=1)
                End If
            Else
                If IItms.Count < Key Then
                    Call IItms.Add("")
                    Call IItms.Add(Value)
                Else
                    Call IItms.Add(Value, after:=Key)
                End If
            End If
        End If
    End If

End Function

Public Property Get Keys() As Collection
    '�L�[�̕��я��͍��̏��A�ǉ����ꂽ��
    '�Y���A�N�Z�X�̃v���p�e�B�͖��T�|�[�g
    Set Keys = IKeys
End Property

Public Property Get Items() As Collection
    Set Items = IItms
End Property

