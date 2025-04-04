Attribute VB_Name = "�`�[�ԍ�"
Option Explicit

Sub renban()

    Dim lastrow As Long
    Dim ws As worksheet
    Dim currentValue As String
    Dim currentAlpha As String
    Dim currentNum As Long
    Dim i As Long
    Dim newValue As String

    Set ws = ActiveSheet

    ' H��̍ŏI�s���擾
    lastrow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    
    ' A2�̒l�i��: "A1"�j���擾
    currentValue = ws.Range("A2").value
    currentAlpha = Left(currentValue, 1) ' �A���t�@�x�b�g����
    currentNum = CLng(Mid(currentValue, 2)) ' ��������
    
    ' A���3�s�ڂ���ŏI�s�܂ŁAH��̒l���ς�邽�т�A��ɒl���R�s�[
    For i = 3 To lastrow
        If ws.Cells(i, "H").value <> ws.Cells(i - 1, "H").value Then
            currentNum = currentNum + 1 ' H��̒l���ς�����ꍇ�A����������+1
        End If
        
        ' �V�����l���쐬�i�A���t�@�x�b�g�����͂��̂܂܂Ő����������C���N�������g�j
        newValue = currentAlpha & currentNum
        ws.Cells(i, "A").value = newValue
    Next i
    
End Sub
