'/* **************************************************
' * Copyright (c) 2017 Yamamoto Yuma
' * Released under the MIT license
' * http://opensource.org/licenses/mit-license.php
' * *************************************************/

Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports TracLink_Outlook
Imports Microsoft.Office.Interop.Outlook

<TestClass()> Public Class TracLinkTest

    Private YYUTIL As New yyutil
    Private TRACLINK As New TracLink
    Private target As TracLink
    Private pobj As PrivateObject
    Private TESTURL1 = "http://traclinktest.yamayuma.com/projects/TracLinkTest1/login/xmlrpc"
    Private TESTURL2 = "http://traclinktest.yamayuma.com/projects/TracLinkTest2/login/xmlrpc"
    Private TESTURL3 = "http://traclinktest.yamayuma.com/projects/TracLinkTest3/login/xmlrpc"
    Private TESTURL4 = "http://traclinktest.yamayuma.com/projects/TracLinkTest4/login/xmlrpc"

    Private Sub initialize()
        YYUTIL.SetUserSetting("UserName", "user1")
        YYUTIL.SetUserSetting("Password", "password")
        target = New TracLink()
        pobj = New PrivateObject(target)

    End Sub

    ' テスト対象未実装
    '<TestMethod()> Public Sub getTicketFieldsTest()
    '    Dim data() As String

    '    Call initialize()

    '    ' 想定通りの一覧が取得できるかどうかの確認
    '    data = pobj.Invoke("getTicketFields", {TESTURL1})
    '    yydbg.DumpArray(data)
    '    CollectionAssert.AreEqual({"defect", "enhancement", "task"}, data)

    '    ' 対象項目がない場合の確認
    '    data = pobj.Invoke("getTicketFields", {TESTURL2})
    '    CollectionAssert.AreEqual({}, data)

    '    ' 権限がない場合の確認
    '    YYUTIL.SetUserSetting("UserName", "hogehoge")
    '    data = pobj.Invoke("getTicketFields", {TESTURL3})
    '    CollectionAssert.AreEqual({}, data)


    '    ' 対象プロジェクトがない場合の確認
    '    data = pobj.Invoke("getTicketFields", {TESTURL4})
    '    CollectionAssert.AreEqual({}, data)

    'End Sub

    <TestMethod()> Public Sub getComponentListTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("getComponentList", {TESTURL1})
        CollectionAssert.AreEqual({"component1", "component2"}, data)

        ' 対象項目がない場合の確認
        data = pobj.Invoke("getComponentList", {TESTURL2})
        CollectionAssert.AreEqual({}, data)

        ' 権限がない場合の確認
        YYUTIL.SetUserSetting("UserName", "hogehoge")
        data = pobj.Invoke("getComponentList", {TESTURL3})
        CollectionAssert.AreEqual({}, data)

        ' 対象プロジェクトがない場合の確認
        data = pobj.Invoke("getComponentList", {TESTURL4})
        Assert.IsNull(data)

    End Sub

    <TestMethod()> Public Sub getTypeListTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("getTypeList", {TESTURL1})
        CollectionAssert.AreEqual({"defect", "enhancement", "task"}, data)

        ' 対象項目がない場合の確認
        data = pobj.Invoke("getTypeList", {TESTURL2})
        CollectionAssert.AreEqual({}, data)

        ' 権限がない場合の確認
        YYUTIL.SetUserSetting("UserName", "hogehoge")
        data = pobj.Invoke("getTypeList", {TESTURL3})
        CollectionAssert.AreEqual({}, data)

        ' 対象プロジェクトがない場合の確認
        data = pobj.Invoke("getTypeList", {TESTURL4})
        Assert.IsNull(data)

    End Sub

    <TestMethod()> Public Sub getMilestoneListTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("getMilestoneList", {TESTURL1})
        yydbg.DumpArray(data)
        CollectionAssert.AreEqual({"milestone1", "milestone2", "milestone3", "milestone4"}, data)

        ' 対象項目がない場合の確認
        data = pobj.Invoke("getMilestoneList", {TESTURL2})
        CollectionAssert.AreEqual({}, data)

        ' 権限がない場合の確認
        YYUTIL.SetUserSetting("UserName", "hogehoge")
        data = pobj.Invoke("getMilestoneList", {TESTURL3})
        CollectionAssert.AreEqual({}, data)


        ' 対象プロジェクトがない場合の確認
        data = pobj.Invoke("getMilestoneList", {TESTURL4})
        Assert.IsNull(data)

    End Sub

    <TestMethod()> Public Sub getResolutionListTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("getResolutionList", {TESTURL1})
        CollectionAssert.AreEqual({"fixed", "invalid", "wontfix", "duplicate", "worksforme"}, data)

        ' 対象項目がない場合の確認
        data = pobj.Invoke("getResolutionList", {TESTURL2})
        CollectionAssert.AreEqual({}, data)

        ' 権限がない場合の確認
        YYUTIL.SetUserSetting("UserName", "hogehoge")
        data = pobj.Invoke("getResolutionList", {TESTURL3})
        CollectionAssert.AreEqual({}, data)


        ' 対象プロジェクトがない場合の確認
        data = pobj.Invoke("getResolutionList", {TESTURL4})
        Assert.IsNull(data)

    End Sub

    <TestMethod()> Public Sub getPriorityListTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("getPriorityList", {TESTURL1})
        CollectionAssert.AreEqual({"blocker", "critical", "major", "minor", "trivial"}, data)

        ' 対象項目がない場合の確認
        data = pobj.Invoke("getPriorityList", {TESTURL2})
        CollectionAssert.AreEqual({}, data)

        ' 権限がない場合の確認
        YYUTIL.SetUserSetting("UserName", "hogehoge")
        data = pobj.Invoke("getPriorityList", {TESTURL3})
        CollectionAssert.AreEqual({}, data)


        ' 対象プロジェクトがない場合の確認
        data = pobj.Invoke("getResolutionList", {TESTURL4})
        Assert.IsNull(data)

    End Sub

    <TestMethod()> Public Sub getTicketTest()
        Dim ticket As Object
        Dim member As New Dictionary(Of String, String) _
            From {{"id", "1"}, {"status", "new"}, {"changetime", "20170927T13:08:40"}, {"_ts", "1506517720997794"}, {"description", "detail"}, {"reporter", "user1"}, {"cc", "aaaa"}, {"resolution", ""}, {"time", "20170927T13:08:40"}, {"component", "component1"}, {"summary", "testtitle"}, {"priority", "major"}, {"keywords", "hogehoge"}, {"version", "2.0"}, {"milestone", "milestone2"}, {"owner", "somebody"}, {"type", "defect"}}

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        ticket = pobj.Invoke("getTicket", {TESTURL1, "1"})
        ticket.dump()
        CollectionAssert.AreEqual(member, ticket.member)

        ' 対象項目がない場合の確認
        ticket = pobj.Invoke("getTicket", {TESTURL1, "3"})
        '        CollectionAssert.AreEqual({}, ticket.member)
        Assert.IsNull(ticket)

        ticket = pobj.Invoke("getTicket", {TESTURL2, "1"})
        Assert.IsNull(ticket)

        ' 権限がない場合の確認
        ticket = pobj.Invoke("getTicket", {TESTURL3, "1"})
        Assert.IsNull(ticket)

        ' 対象プロジェクトがない場合の確認
        ticket = pobj.Invoke("getTicket", {TESTURL4, "1"})
        Assert.IsNull(ticket)


    End Sub

    <TestMethod()> Public Sub getTicketLogTest()
        Dim ticketLog As Collection
        Dim testData As New Collection From {
"[user1] 2017/09/30 02:49:46
testlog
",
"[user1] 2017/09/30 02:58:10
<http://traclinktest.yamayuma.com/projects/TracLinkTest1/raw-attachment/ticket/2/testdata.txt>

"}

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL1, "2"})
        CollectionAssert.AreEqual(testData, ticketLog)

        ' 対象項目がない場合の確認
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL1, "1"})
        CollectionAssert.AreEqual({}, ticketLog)
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL1, "3"})
        Assert.IsNull(ticketLog)
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL2, "2"})
        Assert.IsNull(ticketLog)
        ' 権限がない場合の確認
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL3, "2"})
        Assert.IsNull(ticketLog)
        ' 対象プロジェクトがない場合の確認
        ticketLog = pobj.Invoke("getTicketLog", {TESTURL4, "2"})
        Assert.IsNull(ticketLog)


    End Sub


    <TestMethod()> Public Sub queryTicketTest()
        Dim data() As String

        Call initialize()

        ' 想定通りの一覧が取得できるかどうかの確認
        data = pobj.Invoke("queryTicket", {TESTURL1, "modified=2014/01/01.."})
        yydbg.DumpArray(data)
        CollectionAssert.AreEqual({"1", "2"}, data)

    End Sub


    <TestMethod()> Public Sub updateProjectTest()
        Dim data() As String
        Dim folder As Folder
        Dim mngFolder As Folder
        Dim item As TaskItem

        Call initialize()

        folder = CreateObject("Outlook.Application").GetNamespace("MAPI").folders("TracLinkTest")
        Call folder.Folders("tmpUnitTest").Delete()
        folder.Folders.Add("tmpUnitTest", OlDefaultFolders.olFolderTasks)
        mngFolder = folder.Folders("tmpUnitTest")

        item = CreateObject("Outlook.Application").CreateItem(OlItemType.olTaskItem)
        item.Subject = "TracLinkTest1"
        item.Body = "Query:modified=2014/1/1.." & vbCrLf &
                    "URL:http://traclinktest.yamayuma.com/projects/TracLinkTest1" & vbCrLf &
                    "Clear:0"
        item.Save()
        item.Move(mngFolder)

        mngFolder.Folders.Add(item.Subject, OlDefaultFolders.olFolderTasks)
        data = pobj.Invoke("updateProject", {item, mngFolder.Folders(item.Subject)})

    End Sub

    <TestMethod()> Public Sub updateAllProjectTest()
        Dim data() As String

        Call initialize()

        ' 管理対象フォルダ名が間違っている場合
        data = pobj.Invoke("updateAllProject", {"UnitTest_Dummy"})

        ' プロジェクトが空の場合
        data = pobj.Invoke("updateAllProject", {"UnitTest_Null"})

        ' 通常
        data = pobj.Invoke("updateAllProject", {"UnitTest"})

    End Sub


End Class