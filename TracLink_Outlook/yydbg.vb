'/* **************************************************
' * Copyright (c) 2017 Yamamoto Yuma
' * Released under the MIT license
' * http://opensource.org/licenses/mit-license.php
' * *************************************************/

Imports System.Diagnostics


''' <summary>
''' デバッグおよびログ出力用クラス
''' </summary>
Public Class yydbg

#If DEBUG Then

    Private Const FORMAT_CALLER_INFO As String = "{0}.{1}:{2}"
    Private Const FORMAT_ENTER As String = "{0}:<"
    Private Const FORMAT_EXIT As String = "{0}:>"
    Private Const FORMAT_CALL_ENTER As String = "{0}:>>{1}"
    Private Const FORMAT_CALL_EXIT As String = "{0}:<<{1}"
    Private Const FORMAT_LOG_DEBUG As String = "{0}:D:{1}"
    Private Const FORMAT_LOG_INFO As String = "{0}:I:{1}"
    Private Const FORMAT_LOG_ERROR As String = "{0}:E:{1}"

    ''' <summary>
    ''' 呼び出し元メソッド情報取得
    ''' </summary>
    ''' <returns>aaa</returns>
    Private Shared Function GetCallerInfo()
        Dim stackFrame As StackFrame
        stackFrame = New StackTrace(True).GetFrame(2)
        Return String.Format(FORMAT_CALLER_INFO,
                             stackFrame.GetMethod.ReflectedType.Name,
                             stackFrame.GetMethod.Name,
                             stackFrame.GetFileLineNumber)
    End Function

    ''' <summary>
    ''' ログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Private Shared Sub Print(msg As String)
        '        MsgBox(msg)
        Debug.Print(msg)
    End Sub

    ''' <summary>
    ''' トレースログ出力(メソッド開始)
    ''' </summary>
    Public Shared Sub TraceEnter()
        Call Print(String.Format(FORMAT_ENTER, GetCallerInfo()))
    End Sub

    ''' <summary>
    ''' トレースログ出力(メソッド終了)
    ''' </summary>
    Public Shared Sub TraceExit()
        Call Print(String.Format(FORMAT_EXIT, GetCallerInfo()))
    End Sub

    ''' <summary>
    ''' 外部メソッド呼び出し開始ログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Public Shared Sub CallEnter(msg As String)
        Call Print(String.Format(FORMAT_CALL_ENTER, GetCallerInfo(), msg))
    End Sub

    ''' <summary>
    ''' 外部メソッド呼び出し終了ログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Public Shared Sub CallExit(msg As String)
        Call Print(String.Format(FORMAT_CALL_EXIT, GetCallerInfo(), msg))
    End Sub

    ''' <summary>
    ''' デバッグログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Public Shared Sub LogDebug(msg As String)
        Call Print(String.Format(FORMAT_LOG_DEBUG, GetCallerInfo(), msg))
    End Sub

    ''' <summary>
    ''' 通常ログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Public Shared Sub LogInfo(msg As String)
        Call Print(String.Format(FORMAT_LOG_INFO, GetCallerInfo(), msg))
    End Sub

    ''' <summary>
    ''' エラーログ出力
    ''' </summary>
    ''' <param name="msg">出力メッセージ</param>
    Public Shared Sub LogError(msg As String)
        Call Print(String.Format(FORMAT_LOG_ERROR, GetCallerInfo(), msg))
    End Sub

    Public Shared Sub DumpArray(objs() As Object)
        Dim i As Integer
        LogInfo(LBound(objs))
        LogInfo(UBound(objs))

        For i = LBound(objs) To UBound(objs)
            If objs(i) IsNot Nothing Then
                LogInfo(objs(i).ToString)
            End If
        Next
    End Sub

    Public Shared Sub DumpCollection(obj As Collection)
        Dim i As Integer

        LogInfo(obj.Count)

        For i = 1 To obj.Count
            LogInfo(obj.Item(i).ToString)
        Next
    End Sub

    Public Shared Sub DumpDictionary(dict As Dictionary(Of String, String))
        For Each key As String In dict.Keys
            LogInfo(key + " = " + dict(key))
        Next
    End Sub




#Else
    Public Sub TraceEnter()
    End Sub

    Public Sub TraceExit()
    End Sub

    Public Sub CallEnter(msg As String)
    End Sub

    Public Sub CallExit(msg As String)
    End Sub

    Public Sub LogDebug(msg As String)
    End Sub

    Public Sub LogInfo(msg As String)
    End Sub
#End If

End Class
