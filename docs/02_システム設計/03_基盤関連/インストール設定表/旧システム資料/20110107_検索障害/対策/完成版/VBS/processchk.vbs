'#################################################
'###	取引履歴検索システム
'###
'###	processchk.vbs
'###
'###	[機能]
'###	引数で指定されたプロセスの状態を標準出力
'###
'###	[戻り値]
'###	0:指定プロセスは稼動中
'###	0以外:指定プロセスは非稼動中
'###
'###	rev.	date		author		comments
'###	1.0	2011/01/24	N.Matsuda		新規作成
'###
'#################################################
Option Explicit

if ( WScript.Arguments.Count <> 1 ) then
	WScript.Echo "Arguments Error"
	WScript.Echo "USAGE: cscript processchk.vbs <Target Process>"
	WScript.Quit 1
End if

Dim strProcName 	' 対象プロセス名
Dim objProcList	' プロセス一覧
Dim objProcess	' プロセス情報
Dim objProcCnt	' プロセス数
Dim retStat	' 戻り値

strProcName = WScript.Arguments.Item(0)
objProcCnt = 0
retStat = 0

Set objProcList = GetObject("winmgmts:").InstancesOf("win32_process")

For Each objProcess In objProcList
	If objProcess.Name = strProcName Then
		objProcCnt = objProcCnt + 1		
	End if
Next

If objProcCnt = 0 Then
	WScript.Echo strProcName & "プロセスが見つかりませんでした。"
	retStat = 1
Else
	WScript.Echo strProcName & "プロセスが<" & objProcCnt & ">稼動中です。"
End If

Set objProcList = Nothing
WScript.Quit retStat
