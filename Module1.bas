Attribute VB_Name = "Module1"
Public Sub Rbn_customUI_onLoad(ribbon As IRibbonUI)
' Code for onLoad callback. Ribbon control customUI
    
End Sub

Public Sub Rbn_�����c�[��_�t�@�C������_�t�H���_�ꊇ�쐬_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    �t�H���_�ꊇ�쐬.Show
    
End Sub

Public Sub Rbn_�����c�[��_�t�@�C������_�ʃu�b�N�쐬_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    �ʃu�b�N�쐬.Show
    
End Sub

Public Sub Rbn_�����c�[��_�Z���̃X�^�C��_�h��Ԃ��t�H�[��_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    �h��Ԃ��t�H�[��.Show
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_���o_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    ���o�t�H�[��.Show
    
End Sub

Public Sub Rbn_�����c�[��_�t�@�C������_�u�b�N�ꊇ�ړ�_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    �V�[�g�ꊇ����.Show
    
End Sub

Public Sub Rbn_�����c�[��_�Z���̃X�^�C��_������������_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    Call ������������.������������
            
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�V�[�g���ꊇ�ύX_onAction(control As IRibbonControl)

    CreateRenameSheet
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�w�b�_�[�ꊇ�ύX_onAction(control As IRibbonControl)

    �w�b�_�[�ύX�t�H�[��.Show
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_����Vlookup_onAction(control As IRibbonControl)

    ����Vlookup.Show vbModeless
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�S�ʂ葫���Z_onAction(control As IRibbonControl)

    �S�ʂ葫���Z.Show
    
End Sub

Public Sub Rbn_�����c�[��_�Z���̃X�^�C��_��������_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    Call ��������.��������
            
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�ꊇVlookup_onAction(control As IRibbonControl)

    �ꊇVlookup.Show vbModeless
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�摜_onAction(control As IRibbonControl)

    Call ChangePictureProperties
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_�ړ��`�[�ԍ�_onAction(control As IRibbonControl)

    Call renban
    
End Sub

Public Sub Rbn_�����c�[��_�f�[�^����_TA�`�[_onAction(control As IRibbonControl)

    Call TA
    
End Sub
