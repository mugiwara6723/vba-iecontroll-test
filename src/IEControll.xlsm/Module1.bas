Attribute VB_Name = "Module1"
Option Explicit

Sub sample()
    Dim objIE As New InternetExplorer
    objIE.Visible = True '������悤�ɂ���Ȃ�

    objIE.Navigate "https://excel-ubara.com/"
    Call untilReady(objIE) '���[�h�҂�

    Dim objHtml As HTMLDocument
    Set objHtml = objIE.Document

    Dim elm As Object
    '�����j���[�{�^���́u�G�N�Z������v���N���b�N���܂��B
    Set elm = objHtml.querySelector("#sub > nav > ul:nth-child(2) > li:nth-child(1) > a")
    Debug.Print elm.innerText
    elm.Click

    'InternetExplorer�͊J�����܂܂ɂ��Ă����ꍇ�́A���̂܂�VBA���I�����܂��B
End Sub

Sub untilReady(objIE As Object)
    Do While objIE.Busy = True Or objIE.ReadyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
End Sub
