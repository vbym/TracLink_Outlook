'/* **************************************************
' * Copyright (c) 2017 Yamamoto Yuma
' * Released under the MIT license
' * http://opensource.org/licenses/mit-license.php
' * *************************************************/

''' <summary>
''' プロジェクト共通ユーティリティクラス(現状はTracLink依存)
''' </summary>
Public Class yyutil

    Public Const APP_NAME As String = "TracLink4Outlook"
    Public Const SECTION_NAME As String = "UserSettings"

    ''' <summary>
    ''' ユーザ情報保存
    ''' </summary>
    ''' <param name="name">キー文字列</param>
    ''' <param name="value">設定値</param>
    ''' <returns>保存結果(0:成功、それ以外:失敗)</returns>
    Public Shared Function SetUserSetting(name As String, value As String)
        SaveSetting(APP_NAME, SECTION_NAME, name, value)
        Return 0
    End Function

    ''' <summary>
    ''' ユーザ情報取得
    ''' </summary>
    ''' <param name="name">キー文字列</param>
    ''' <returns>設定値</returns>
    Public Shared Function GetUserSetting(name As String)
        Return GetSetting(APP_NAME, SECTION_NAME, name)
    End Function
End Class
