﻿Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security

' アセンブリに関する一般情報は、以下の属性セットによって 
' 制御されます。アセンブリに関連付けられている情報を変更するには、
' これらの属性値を変更します。

' アセンブリ属性の値を確認します

<Assembly: AssemblyTitle("TracLink_Outlook")> 
<Assembly: AssemblyDescription("")> 
<Assembly: AssemblyCompany("")> 
<Assembly: AssemblyProduct("TracLink_Outlook")> 
<Assembly: AssemblyCopyright("Copyright ©  2017")> 
<Assembly: AssemblyTrademark("")> 

' ComVisible を false に設定すると、その型はこのアセンブリ内で COM コンポーネントから 
' 参照できなくなります。COM からこのアセンブリ内の型にアクセスする必要がある場合は、
' その型の ComVisible 属性を true に設定してください。
<Assembly: ComVisible(False)>

'このプロジェクトが COM に公開される場合、次の GUID が typelib の ID になります
<Assembly: Guid("0b9e5533-e170-48a2-9681-92d050e7834a")> 

' アセンブリのバージョン情報は次の 4 つの値で構成されています:
'
'      メジャー バージョン
'      マイナー バージョン 
'      ビルド番号
'      リビジョン
'
' すべての値を指定することも、下に示すように '*' を使用してビルドおよびリビジョン番号を
' 既定値にすることもできます。
' <Assembly: AssemblyVersion("1.0.*")> 

<Assembly: AssemblyVersion("1.0.0.0")> 
<Assembly: AssemblyFileVersion("1.0.0.0")> 

Friend Module DesignTimeConstants
    Public Const RibbonTypeSerializer As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
    Public Const RibbonBaseTypeSerializer As String = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design"
    Public Const RibbonDesigner As String = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a"
End Module
