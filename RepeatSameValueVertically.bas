Attribute VB_Name = "RepeatSameValueVertically"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'�I�������͈͂̋󔒂łȂ��Z�����e���A
'�󔒃Z���ɃR�s�[����
'
Sub RepeatSameValueVertically()
    
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
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    '������
    numOfRows = Selection.Rows.count
    numOfCols = Selection.Columns.count
    
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
