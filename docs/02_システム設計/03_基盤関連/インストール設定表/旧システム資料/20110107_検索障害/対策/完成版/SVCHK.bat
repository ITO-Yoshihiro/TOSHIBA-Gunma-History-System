@echo off
REM ##############################################
REM	###	取引履歴検索システム
REM ###
REM ### システム監視
REM	###	SVCHK.bat
REM ###
REM	### rev.	date		author		comments
REM	### 1.0		2011/01/24	N.Matsuda	新規作成
REM ###
REM ##############################################

REM パス定義（変更要）
set WORK_DIR=C:\temp\完成版\

REM パス定義（変更不要）
set ELPUT=C:\history\bin\ELPut.exe
set SERV_CHK=%WORK_DIR%VBS\servicechk.vbs
set PROC_CHK=%WORK_DIR%VBS\processchk.vbs
set HTTP_GET=%WORK_DIR%VBS\httpget.vbs

set CHK_LIST=%WORK_DIR%CHKLIST.inf

REM ##### Main
setlocal enabledelayedexpansion

echo ################################
echo ### システム監視処理開始
echo ################################

set ER_CNT=0
set ER_WORD=

if not exist %ELPUT% (
	set ER_CNT=2
	set ER_WORD=!ER_WORD![%ELPUT%]
)

if not exist %SERV_CHK% (
	set ER_CNT=2
	set ER_WORD=!ER_WORD![%SERV_CHK%]
)

if not exist %PROC_CHK% (
	set ER_CNT=2
	set ER_WORD=!ER_WORD![%PROC_CHK%]
)

if not exist %HTTP_GET% (
	set ER_CNT=2
	set ER_WORD=!ER_WORD![%HTTP_GET%]
)

if not exist %CHK_LIST% (
	set ER_CNT=2
	set ER_WORD=!ER_WORD![%CHK_LIST%]
)

if not "%ER_CNT%" == "0" (
	goto :END
)

echo ################################
echo ### Service Check
echo ################################

	for /f "delims=, tokens=1,2" %%i in (%CHK_LIST%) do (
		if "%%i" == "SERVICE" (
			cscript %SERV_CHK% "%%j"
			
			if !errorlevel! NEQ 0 (
				echo %%j サービス停止
				%ELPUT% -tE "%%j サービス停止" "History"
				set ER_CNT=1
			)
		)
	)
	
echo ################################
echo ### Process Check
echo ################################
	for /f "delims=, tokens=1,2" %%i in (%CHK_LIST%) do (
		if "%%i" == "PROCESS" (
			cscript %PROC_CHK% "%%j"
			
			if !errorlevel! NEQ 0 (
				echo %%j プロセス停止
				%ELPUT% -tE "%%j プロセス停止" "History"
				set ER_CNT=1
			)
		)
	)

echo ################################
echo ### HTTP Check
echo ################################
	for /f "delims=, tokens=1,2,3" %%i in (%CHK_LIST%) do (
		if "%%i" == "HTTP" (
			cscript %HTTP_GET% %%j %%k

			if !errorlevel! NEQ 0 (
				echo HTTPサービス無応答[%%j][%%k]
				%ELPUT% -tE "HTTPサービス無応答[%%j][%%k]" "History"
				set ER_CNT=1
			)
		)
	)

:END

if %ER_CNT% EQU 0 (
	echo システム監視監視　Status[OK]
	%ELPUT% -tI "システム監視　Status[OK]" "History"
)
if %ER_CNT% EQU 1 (
	echo システム監視　Status[NG]
	%ELPUT% -tE "システム監視　Status[NG]" "History"
)
if %ER_CNT% EQU 2 (
	echo Nothing %ER_WORD% 
)

echo ################################
echo ### システム監視処理終了
echo ################################

endlocal

exit
