Attribute VB_Name = "�����l���J��Ԃ�"
'�I�������͈͂̋󔒂łȂ��Z�����e���A
'�󔒃Z���ɃR�s�[����
'
Sub �����l���J��Ԃ�_����()
    
    Dim r As Range
    Dim val As Variant
    Dim retVal As Integer
    
    Dim numOfRows As Long
    Dim numOfCols As Long
    
    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��Ēl�̏������݂��s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    '������
    numOfRows = Selection.Rows.Count
    numOfCols = Selection.Columns.Count
    
    '���s���[�v
    For colFocus = 1 To numOfCols
    
        val = Selection(1).Offset(0, colFocus - 1).Value
    
        For rowFocus = 2 To numOfRows
            
            If Selection(1).Offset(rowFocus - 1, colFocus - 1).Value <> "" Then
            
                val = Selection(1).Offset(rowFocus - 1, colFocus - 1).Value '�������ݒl���X�V
                
            Else
                
                Selection(1).Offset(rowFocus - 1, colFocus - 1).Value = val '�������ݒl�ŏ�������
                
            End If
            
        Next
    
    Next
    
    MsgBox "Done!"
    
End Sub
