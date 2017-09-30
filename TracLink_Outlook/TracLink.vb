'/* **************************************************
' * Copyright (c) 2017 Yamamoto Yuma
' * Released under the MIT license
' * http://opensource.org/licenses/mit-license.php
' * *************************************************/

Imports Microsoft.Office.Interop.Outlook

Public Class TracLink

    ''' <summary>
    ''' 全プロジェクト更新
    ''' </summary>
    ''' <param name="name">管理対象フォルダ名</param>
    Public Sub updateAllProject(name As String)
        Dim folders As Folders
        Dim folder As Folder
        Dim tracFolder As Folder
        Dim item As TaskItem

        folders = CreateObject("Outlook.Application").GetNamespace("MAPI").folders

        tracFolder = Nothing
        For Each folder In folders
            Try
                tracFolder = folder.Folders(name)
            Catch
                '何もしない
            End Try
        Next
        If tracFolder Is Nothing Then
            yydbg.LogError(name + " folder not found.")
            Exit Sub
        End If

        For Each item In tracFolder.Items
            Try
                tracFolder.Folders.Add(item.Subject, OlDefaultFolders.olFolderTasks)
            Catch
                '何もしない
            End Try
            Try
                Call updateProject(item, tracFolder.Folders(item.Subject))
            Catch
                '何もしない
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 指定プロジェクトのチケット更新
    ''' </summary>
    ''' <param name="mngTask">プロジェクト管理用のタスク</param>
    ''' <param name="pFolder">プロジェクトのフォルダ</param>
    Private Sub updateProject(mngTask As TaskItem, pFolder As Folder)

        Dim ticketIds() As String
        Dim query As String
        Dim url As String
        Dim project As Project = New Project()

        ' 管理用タスクから情報取得
        Call project.Initialize(mngTask.Subject, mngTask.Body)
        Call project.dump()

        ' チケットの情報取得
        query = project.query
        query = Replace(query, "&", "&amp;")

        'If project.clear = "0" Then
        '    query = query & "&amp;modified=" & mngTask.StartDate & ".."
        'End If

        url = project.url & "/login/xmlrpc"
        ticketIds = queryTicket(url, query)
        If UBound(ticketIds) = 0 Then
            Exit Sub
        End If

        Dim ticket As Ticket
        For i = LBound(ticketIds) To UBound(ticketIds)
            ticket = getTicket(url, ticketIds(i))
            If ticket Is Nothing Then
                Exit Sub
            End If
            ticket.log = getTicketLog(url, ticketIds(i))
            'If ticket.log Is Nothing Then
            '    Exit Sub
            'End If
            Call UpdateTracTask(pFolder, ticket)
        Next

        ' 最終更新日を更新
        mngTask.StartDate = New Date

    End Sub

    ''' <summary>
    ''' 指定タスク(チケット)更新
    ''' </summary>
    ''' <param name="folder">Outlookのフォルダ</param>
    ''' <param name="ticket">チケット情報</param>
    Private Sub UpdateTracTask(folder As Folder, ticket As Ticket)
        Dim TracTask As TaskItem

        TracTask = Nothing

        ' すでにタスクがあるか確認
        For i = 1 To folder.Items.Count
            If InStr(1, folder.Items(i).Subject, "#" & ticket.member("id") & ":") = 1 Then
                TracTask = folder.Items(i)
                Exit For
            End If
        Next

        ' タスクがない場合は作成
        If TracTask Is Nothing Then
            TracTask = CreateObject("Outlook.Application").CreateItem(OlItemType.olTaskItem)
            TracTask.Subject = "#" + ticket.member("id") + ":" + ticket.member("summary")
        End If

        ' タスクの中身を更新
        On Error Resume Next
        TracTask.Body = ""
        For i = 0 To ticket.log.Count
            TracTask.Body = "--------------------" & vbCrLf & ticket.log.Item(i) & vbCrLf & TracTask.Body
        Next
        TracTask.Body = ticket.url & vbCrLf & vbCrLf & ticket.member("description") & vbCrLf & TracTask.Body
        TracTask.Categories = ticket.member("owner")
        If ticket.member("status") = "new" Then
            TracTask.Status = OlTaskStatus.olTaskNotStarted
        Else
            If ticket.member("status") = "closed" Then
                TracTask.Status = OlTaskStatus.olTaskComplete
            Else
                TracTask.Status = OlTaskStatus.olTaskInProgress
            End If
        End If
        TracTask.DueDate = ticket.member("due_close")
        TracTask.Save()
        TracTask.Move(folder)
        On Error GoTo 0
    End Sub

    ''' <summary>
    ''' 一覧情報取得
    ''' </summary>
    ''' <param name="url">取得元URL</param>
    ''' <param name="funcname">取得対象メソッド名(Tracのメソッド名)</param>
    ''' <returns>一覧情報</returns>
    Private Function getList(url As String, funcname As String) As String()
        Dim message As String
        Dim i As Integer
        Dim data() As String
        Dim args(0) As String
        Dim receiveXml As Object

        message = createXMLData(funcname, args)
        receiveXml = sendMessage(url, message)
        If isError(receiveXml) = True Then
            Return Nothing
            Exit Function
        End If

        Dim objects = receiveXml.getElementsByTagName("string")
        ReDim data(objects.Length - 1)
        On Error Resume Next
        For i = 0 To objects.Length - 1
            data(i) = objects.Item(i).ChildNodes(0).NodeValue
        Next
        On Error GoTo 0
        getList = data
    End Function

    ''' <summary>
    ''' チケットフィールド取得(未実装)
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns>チケットフィールド</returns>
    Private Function getTicketFields(url As String) As String()
        Return getList(url, "ticket.getTicketFields")
    End Function

    ''' <summary>
    ''' コンポーネント一覧取得 
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns></returns>
    Private Function getComponentList(url As String) As String()
        Return getList(url, "ticket.component.getAll")
    End Function

    ''' <summary>
    ''' チケット分類取得
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns></returns>
    Private Function getTypeList(url As String) As String()
        Return getList(url, "ticket.type.getAll")
    End Function

    ''' <summary>
    ''' マイルストーン一覧取得
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns></returns>
    Private Function getMilestoneList(url As String) As String()
        getMilestoneList = getList(url, "ticket.milestone.getAll")
    End Function

    ''' <summary>
    ''' 優先度一覧取得 
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns></returns>
    Private Function getPriorityList(url As String) As String()
        getPriorityList = getList(url, "ticket.priority.getAll")
    End Function

    ''' <summary>
    ''' 解決方法一覧取得
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <returns></returns>
    Private Function getResolutionList(url As String) As String()
        getResolutionList = getList(url, "ticket.resolution.getAll")
    End Function

    ''' <summary>
    ''' チケット一覧取得
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <param name="qstr">チケット検索条件(TracQuery)</param>
    ''' <returns>チケットID一覧</returns>
    Private Function queryTicket(url As String, qstr As String) As String()
        Dim i As Integer
        Dim args(0) As String
        Dim oStrs() As String
        Dim message As String
        Dim receiveXml As Object

        args(0) = qstr
        message = createXMLData("ticket.query", args)
        receiveXml = sendMessage(url, message)
        If isError(receiveXml) = True Then
            queryTicket = New String() {}
            Exit Function
        End If
        Dim mobjects = receiveXml.getElementsByTagName("int")
        ReDim oStrs(mobjects.Length - 1)
        For i = 0 To mobjects.Length - 1
            oStrs(i) = mobjects.Item(i).Text
        Next
        queryTicket = oStrs
    End Function

    ''' <summary>
    ''' チケット詳細取得
    ''' </summary>
    ''' <param name="url">TracサーバURL</param>
    ''' <param name="id">チケットID</param>
    ''' <returns></returns>
    Private Function getTicket(url As String, id As String) As Ticket
        Dim message As String
        Dim args(0) As String
        Dim receiveXml As Object
        Dim dict As New Dictionary(Of String, String)()
        Dim ticket As New Ticket(url, id)

        args(0) = id
        message = createXMLData("ticket.get", args)
        receiveXml = sendMessage(url, message)
        If isError(receiveXml) = True Then
            Return Nothing
            Exit Function
        End If

        ' 本来はここでフォーマットチェック

        ticket.member("id") = receiveXml.getElementsByTagName("int").Item(0).Text
        For Each obj In receiveXml.getElementsByTagName("member")
            If obj.ChildNodes.Length = 2 Then
                '                col.Add(obj.ChildNodes(1).Text, obj.ChildNodes(0).Text)
                ticket.member(obj.ChildNodes(0).Text) = obj.ChildNodes(1).Text
            End If
        Next
        '    col.Add receiveXml.XML, "rowxml"

        Return ticket
    End Function

    ' 指定したチケットの情報を取得
    Private Function getTicketLog(url As String, id As String) As Collection
        Dim col As Collection
        Dim message As String
        Dim args(1) As String
        Dim receiveXml As Object
        Dim str As String
        Dim isAttatch As Boolean
        Dim body As String
        Dim reporter As String

        args(0) = id
        args(1) = "0"
        message = createXMLData("ticket.changeLog", args)
        receiveXml = sendMessage(url, message)
        If isError(receiveXml) = True Then
            Return Nothing
            Exit Function
        End If

        col = New Collection
        Dim key = ""
        Dim arrayObjects = receiveXml.getElementsByTagName("array")
        isAttatch = False
        reporter = ""
        body = ""
        For i = 1 To arrayObjects.Length - 1
            Dim objects = arrayObjects.Item(i).getElementsByTagName("string")
            '        If objects.Item(1).Text = "comment" Then
            If (objects.Item(1).Text = "comment") Or (objects.Item(1).Text = "attachment") Then

                If objects.Item(1).Text = "comment" Then
                    If isAttatch = False Then
                        If i <> 1 Then
                            yydbg.LogDebug(key & "*" & body)
                            col.Add(reporter & body, key)
                        End If
                        key = objects.Item(2).Text
                        body = objects.Item(3).Text & vbCrLf
                        yydbg.LogDebug("key1 = " & key & vbCrLf)
                    Else
                        isAttatch = False
                        body = body & objects.Item(3).Text & vbCrLf
                    End If


                Else
                    isAttatch = True
                    If i <> 1 Then
                        col.Add(reporter & body, key)
                    End If
                    key = objects.Item(3).Text
                    body = "<" & Replace(url, "/login/xmlrpc", "") & "/raw-attachment/ticket/" & id & "/" & objects.Item(3).Text & ">" & vbCrLf
                    yydbg.LogDebug("key2 = " & key & vbCrLf)
                End If

                str = arrayObjects.Item(i).getElementsByTagName("dateTime.iso8601").Item(0).Text
                reporter = "[" & objects.Item(0).Text & "] "
                reporter = reporter & Left(str, 4) & "/" & Mid(str, 5, 2) & "/" & Mid(str, 7, 2) & " " & Right(str, 8) & vbCrLf
            Else
                If objects.Item(1).Text = "description" Then
                    body = objects.Item(1).Text & " : 変更あり" & vbCrLf & body
                Else
                    body = objects.Item(1).Text & " : 「" & objects.Item(2).Text & "」->「" & objects.Item(3).Text & "」" & vbCrLf & body
                End If
            End If
        Next
        If key <> "" Then
            col.Add(reporter & body, key)
        End If

        Return col
    End Function

    'Private Function createTicket(params() As String) As String
    '    Dim receiveXml As Object
    '    Dim message As String

    '    message = "<?xml version='1.0'?><methodCall><methodName>ticket.create</methodName><params>"
    '    ' summary
    '    message = message & "<param><value><string>" & params(0) & "</string></value></param>"
    '    ' description
    '    message = message & "<param><value><string>" & params(1) & "</string></value></param>"
    '    ' attribute
    '    message = message & "<param><value><struct>"
    '    message = message & "<member><name>type</name><value><string>" & params(2) & "</string></value></member>"
    '    message = message & "<member><name>component</name><value><string>" & params(3) & "</string></value></member>"
    '    message = message & "<member><name>due_assign</name><value><string>" & params(4) & "</string></value></member>"
    '    message = message & "<member><name>due_close</name><value><string>" & params(5) & "</string></value></member>"
    '    message = message & "</struct></value></param>"
    '    message = message & "</params></methodCall>"

    '    receiveXml = sendMessage(url, message)
    '    If isError(receiveXml) = True Then
    '        createTicket = "-1"
    '        Exit Function
    '    End If

    '    Dim objects = receiveXml.getElementsByTagName("int")

    '    If objects.Length > 0 Then
    '        createTicket = objects.Item(i).ChildNodes(0).NodeValue
    '    Else
    '        createTicket = "-1"
    '    End If
    'End Function

    ''' <summary>
    ''' 通知用XMLデータ生成
    ''' </summary>
    ''' <param name="funcname">メソッド名</param>
    ''' <param name="args">パラメータ(int/stringのみ対応)</param>
    ''' <returns>通知用XMLデータ</returns>
    Private Function createXMLData(funcname As String, args() As String) As String
        Dim message As String
        Dim i As Integer

        message = "<?xml version='1.0' encoding='utf-8'?>" & vbNewLine &
            "<methodCall><methodName>" & funcname & "</methodName>"

        If args(0) <> "" Then
            message = message & "<params>"
            For i = LBound(args) To UBound(args)
                '                yydbg.LogDebug(VarType(args(i)))
                If IsNumeric(args(i)) = True Then
                    message = message & "<param><value><int>" & args(i) & "</int></value></param>"
                Else
                    message = message & "<param><value><string>" & args(i) & "</string></value></param>"
                End If
            Next
            message = message & "</params>"
        End If

        message = message & "</methodCall>"
        createXMLData = message
    End Function

    ''' <summary>
    ''' エラーチェック処理(エラー時はメッセージボックスで通知)
    ''' </summary>
    ''' <param name="receiveXml">エラーチェック対象XMLデータ</param>
    ''' <returns>チェック結果(True/False)</returns>
    Private Function isError(receiveXml As Object) As Boolean
        Dim col As Collection

        If receiveXml Is Nothing Then
            Return True
            Exit Function
        End If

        If receiveXml.getElementsByTagName("fault").Length = 0 Then
            Return False
            Exit Function
        End If

        col = New Collection
        For Each obj In receiveXml.getElementsByTagName("member")
            If obj.ChildNodes.Length = 2 Then
                col.Add(obj.ChildNodes(1).Text, obj.ChildNodes(0).Text)
            End If
        Next

        yydbg.LogError("Trac エラー" & vbCrLf & "エラーコード：" & col.Item("faultCode") & vbCrLf & "エラー内容:" & vbCrLf & col.Item("faultString"))
        Return True
    End Function

    ''' <summary>
    ''' Tracへの要求送信処理
    ''' </summary>
    ''' <param name="url">要求先URL</param>
    ''' <param name="message">要求内容</param>
    ''' <returns>結果データ</returns>
    Private Function sendMessage(url As String, message As String) As Object
        Dim http As Object
        http = CreateObject("MSXML2.XMLHTTP")

        http.Open("POST", url, False,
                 yyutil.GetUserSetting("UserName"),
                 yyutil.GetUserSetting("Password"))
        http.setRequestHeader("Content-Type", "text/xml")
        http.setRequestHeader("Method", "POST " & url & " HTTP/1.1")

        Dim tmpmessage = message
        '        yydbg.LogDebug(tmpmessage)

        Try
            Call http.Send(tmpmessage)
        Catch
            MsgBox("エラー番号:" & Err.Number & vbCrLf &
               "エラーの内容:" & Err.Description, vbExclamation)
            Return Nothing
            Exit Function
        End Try

        If http.responseText = "Environment not found" Then
            yydbg.LogError("Environment not found")
            Return Nothing
            Exit Function
        End If

        yydbg.LogDebug(http.responseText)
        Return http.responseXML
    End Function


    ''' <summary>
    ''' Tracプロジェクト情報クラス
    ''' </summary>
    Private Class Project
        Public query As String
        Public name As String
        Public clear As String
        Public url As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        Public Sub New()
            name = ""
            query = ""
            clear = ""
            url = ""
        End Sub

        ''' <summary>
        ''' 初期設定
        ''' </summary>
        ''' <param name="nameStr">プロジェクト名</param>
        ''' <param name="input">プロジェクト情報</param>
        Public Sub Initialize(nameStr As String, input As String)
            Dim reg = CreateObject("VBScript.RegExp")
            Dim list() As String
            Dim key As String
            Dim value As String

            name = nameStr

            reg.Global = True
            reg.IgnoreCase = True

            list = Split(input, vbCrLf)
            For i = 0 To UBound(list)
                reg.Pattern = "^(URL|ID|Query|Clear):[ \t]*([^\s#]+)"
                Dim matches = reg.Execute(list(i))
                key = matches(0).submatches(0)
                value = matches(0).submatches(1)
                Select Case key
                    Case "Query"
                        query = value
                    Case "Clear"
                        clear = value
                    Case "URL"
                        url = value
                End Select
            Next
        End Sub

        ''' <summary>
        ''' プロジェクト情報表示
        ''' </summary>
        Public Sub dump()
            yydbg.LogDebug("name = " + name)
            yydbg.LogDebug("query = " + query)
        End Sub

    End Class

    ''' <summary>
    ''' チケット情報クラス
    ''' </summary>
    Private Class Ticket
        Public member As Dictionary(Of String, String)
        Public log As Collection
        Public url As String

        ''' <summary>
        ''' コンストラクタ
        ''' </summary>
        Public Sub New(base As String, id As String)
            member = New Dictionary(Of String, String)
            log = New Collection
            url = Replace(base, "/login/xmlrpc", "/ticket/") & id
        End Sub

        ''' <summary>
        ''' チケット情報表示
        ''' </summary>
        Public Sub dump()
            yydbg.DumpDictionary(member)
        End Sub

    End Class


End Class
