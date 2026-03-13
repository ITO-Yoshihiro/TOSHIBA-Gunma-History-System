'#################################################
'###	取引履歴検索システム
'###
'###	httpget.vbs
'###
'###	[機能]
'###	引数で指定されたURLへアクセスし、
'###	Text形式でhtmlソースを標準出力
'###
'###	[戻り値]
'###	0:正常
'###	1:IEオブジェクトの作成に失敗	
'###
'###	rev.	date		author		comments
'###	1.0	2011/01/24	N.Matsuda		新規作成
'###
'#################################################
Option Explicit
On Error Resume Next

if ( WScript.Arguments.Count <> 2 ) then
	WScript.Echo "Arguments Error"
	WScript.Echo "USAGE: cscript httpget.vbs <Target URL> <Target Keyword>"
	WScript.Quit 1
End if

Dim strURL	' 表示するページ
Dim strKEY	' ページを特定するキーワード
Dim objIE	' IE オブジェクト
Dim retStat	' 戻り値

retStat = 0

strURL = WScript.Arguments.Item(0)
strKEY = WScript.Arguments.Item(1)

Set objIE = WScript.CreateObject("InternetExplorer.Application")

If Err.Number = 0 Then
    objIE.Navigate strUrl

    ' ページが取り込まれるまで待つ
    Do While objIE.busy or objIE.Document.readyState <> "complete"
        WScript.Sleep(1000)
    Loop

	' テキスト形式で出力
'	WScript.Echo objIE.Document.Body.InnerText

	If InStr(objIE.Document.Body.InnerText , strKEY) = 0 Then
		WScript.Echo "[" & strUrl & "]の取得に失敗しました"
		retStat = 1
	End if

Else
	WScript.Echo "InternetExplorerオブジェクトの作成に失敗しました"
	retStat = 1
End If

http.Quit
Set http = Nothing
WScript.Quit retStat
