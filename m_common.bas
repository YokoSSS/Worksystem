Attribute VB_Name = "m_common"
Option Explicit

'******************************************************************************
'FunctionName:マクロ開始
'Specifications：マクロ開始時、処理スピードを高める為に無駄な動きをする動作を
'停止させる
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************

Sub マクロ開始()
    Application.ScreenUpdating = False '画面描画を停止
    Application.Cursor = xlWait 'ウエイトカーソル
    Application.EnableEvents = False 'イベントを抑止
    Application.DisplayAlerts = False '確認メッセージを抑止
    Application.Calculation = xlCalculationManual '計算を手動に
End Sub

'******************************************************************************
'FunctionName:マクロ終了
'Specifications：マクロ開始時、処理スピードを高める為に無駄な動きをする動作を
'停止させていたものを再稼働させる
'Arguments：nothing
'ReturnValue:nothing
'Note：
'******************************************************************************
Sub マクロ終了()
    Application.StatusBar = False 'ステータスバーを消す
    Application.Calculation = xlCalculationAutomatic '計算を自動に
    Application.DisplayAlerts = True '確認メッセージを開始
    Application.EnableEvents = True 'イベントを開始
    Application.Cursor = xlDefault '標準カーソル
    Application.ScreenUpdating = True '画面描画を開始
End Sub


