Attribute VB_Name = "Module1"
Option Explicit

Sub sample()
    Dim objIE As New InternetExplorer
    objIE.Visible = True '見えるようにするなら

    objIE.Navigate "https://excel-ubara.com/"
    Call untilReady(objIE) 'ロード待ち

    Dim objHtml As HTMLDocument
    Set objHtml = objIE.Document

    Dim elm As Object
    '左メニューボタンの「エクセル入門」をクリックします。
    Set elm = objHtml.querySelector("#sub > nav > ul:nth-child(2) > li:nth-child(1) > a")
    Debug.Print elm.innerText
    elm.Click

    'InternetExplorerは開いたままにしておく場合は、そのままVBAを終了します。
End Sub

Sub untilReady(objIE As Object)
    Do While objIE.Busy = True Or objIE.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
End Sub
