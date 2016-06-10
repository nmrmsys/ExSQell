
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
        'コレクションかプロパティでName属性が指定キーの場合は削除
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
    
    'キー指定が無い場合はどうするか...そのうち考える
    'RepFlgは Let/Setと関数定義を合わせる為のダミー
    
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
        'キー指定が無い場合は最後に追加
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
        'キー指定が無い場合は最後に追加
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
        'キー指定
        If RepFlg Then
            '置換指定
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
            '追加指定
            '指定キーの位置を検索して
            'その後ろに追加する予定
            'Call IItms.Add(Value, Key)
            'Call IKeys.Add(Key, Key)
            MsgBox "未実装"
        End If
    Else
        '添字指定
        If IItms.Count < Key Then
            '指定より少ない場合、空白項目を追加
            For i = IItms.Count To Key - 2
                Call IItms.Add("")
            Next
        End If
        If RepFlg Then
            '置換指定
            On Error Resume Next
            Call IItms.Remove(Key)
            On Error GoTo 0
            If IItms.Count < Key Then
                Call IItms.Add(Value)
            Else
                Call IItms.Add(Value, before:=Key)
            End If
        Else
            '追加指定
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
    'キーの並び順は今の所、追加された順
    '添字アクセスのプロパティは未サポート
    Set Keys = IKeys
End Property

Public Property Get Items() As Collection
    Set Items = IItms
End Property

