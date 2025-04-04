Attribute VB_Name = "画像"
Sub ChangePictureProperties()

    Dim pic As Picture
    ' 現在のシートの全ての画像に対して処理を行う
    
    For Each pic In ActiveSheet.Pictures
        ' 画像のプロパティを「セルに合わせて移動やサイズを変更する」に設定
        pic.Placement = xlMoveAndSize
    Next pic
    
    MsgBox "全ての画像のプロパティを変更しました。", vbInformation
    
End Sub
