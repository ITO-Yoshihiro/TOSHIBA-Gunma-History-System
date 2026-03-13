'#################################################
'###	取引履歴検索システム
'###
'###	servicechk.vbs
'###
'###	[機能]
'###	引数で指定されたサービスの状態を標準出力
'###
'###	[戻り値]
'###	0:指定サービスは稼動中
'###	0以外:指定サービスは非稼動中
'###
'###	rev.	date		author		comments
'###	1.0	2011/01/24	N.Matsuda		新規作成
'###
'#################################################
Option Explicit

if ( WScript.Arguments.Count <> 1 ) then
	WScript.Echo "Arguments Error"
	WScript.Echo "USAGE: cscript servicechk.vbs <Target Service>"
	WScript.Quit 1
End if


Dim strQuery	' サービス取得ＳＱＬ
Dim strServiceName 	' サービス名
Dim lngServiceCount	' サービス数
Dim objServiceList	' 対象のサービス一覧
Dim objServiceInfo	' サービスの情報
Dim retStat		' 戻り値


strServiceName = WScript.Arguments.Item(0)
lngServiceCount = 0
retStat = 0

strQuery = "SELECT * FROM Win32_Service" & _
            " WHERE Name = '" & strServiceName & "'"

Set objServiceList = GetObject("winmgmts:").ExecQuery(strQuery)

For Each objServiceInfo In objServiceList
	If objServiceInfo.State = "Running" Then
		WScript.Echo objServiceInfo.Name & "は開始しています。"

	ElseIf objServiceInfo.State = "Stopped" Then
		WScript.Echo objServiceInfo.Name & "は停止しています。"
		retStat = 1

	ElseIf objServiceInfo.State = "Paused" Then
		WScript.Echo objServiceInfo.Name & "は一時停止しています。"
		retStat = 1

   	Else
		WScript.Echo objServiceInfo.Name & "の状態が分かりません。"
		retStat = 1

	End If

	lngServiceCount = lngServiceCount + 1
Next

If lngServiceCount = 0 Then
	WScript.Echo strServiceName & "サービスが見つかりませんでした。"
	retStat = 1

End If

Set objServiceList = Nothing
WScript.Quit retStat
