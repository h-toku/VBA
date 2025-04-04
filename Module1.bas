Attribute VB_Name = "Module1"
Public Sub Rbn_customUI_onLoad(ribbon As IRibbonUI)
' Code for onLoad callback. Ribbon control customUI
    
End Sub

Public Sub Rbn_徳原ツール_ファイル操作_フォルダ一括作成_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    フォルダ一括作成.Show
    
End Sub

Public Sub Rbn_徳原ツール_ファイル操作_個別ブック作成_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    個別ブック作成.Show
    
End Sub

Public Sub Rbn_徳原ツール_セルのスタイル_塗りつぶしフォーム_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    塗りつぶしフォーム.Show
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_抽出_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    抽出フォーム.Show
    
End Sub

Public Sub Rbn_徳原ツール_ファイル操作_ブック一括移動_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    シート一括統合.Show
    
End Sub

Public Sub Rbn_徳原ツール_セルのスタイル_自動結合解除_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    Call 自動凍結解除.自動凍結解除
            
End Sub

Public Sub Rbn_徳原ツール_データ操作_シート名一括変更_onAction(control As IRibbonControl)

    CreateRenameSheet
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_ヘッダー一括変更_onAction(control As IRibbonControl)

    ヘッダー変更フォーム.Show
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_自動Vlookup_onAction(control As IRibbonControl)

    自動Vlookup.Show vbModeless
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_全通り足し算_onAction(control As IRibbonControl)

    全通り足し算.Show
    
End Sub

Public Sub Rbn_徳原ツール_セルのスタイル_自動結合_onAction(control As IRibbonControl)
' Code for onAction callback. Ribbon control button

    Call 自動結合.自動結合
            
End Sub

Public Sub Rbn_徳原ツール_データ操作_一括Vlookup_onAction(control As IRibbonControl)

    一括Vlookup.Show vbModeless
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_画像_onAction(control As IRibbonControl)

    Call ChangePictureProperties
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_移動伝票番号_onAction(control As IRibbonControl)

    Call renban
    
End Sub

Public Sub Rbn_徳原ツール_データ操作_TA伝票_onAction(control As IRibbonControl)

    Call TA
    
End Sub
