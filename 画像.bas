Attribute VB_Name = "�摜"
Sub ChangePictureProperties()

    Dim pic As Picture
    ' ���݂̃V�[�g�̑S�Ẳ摜�ɑ΂��ď������s��
    
    For Each pic In ActiveSheet.Pictures
        ' �摜�̃v���p�e�B���u�Z���ɍ��킹�Ĉړ���T�C�Y��ύX����v�ɐݒ�
        pic.Placement = xlMoveAndSize
    Next pic
    
    MsgBox "�S�Ẳ摜�̃v���p�e�B��ύX���܂����B", vbInformation
    
End Sub
