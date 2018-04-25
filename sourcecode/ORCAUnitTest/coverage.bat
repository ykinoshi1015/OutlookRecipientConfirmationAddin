REM ###### Settings ######


REM NUnit, OpenCover, ReportGeneratorのインストール先
SET NUNIT=..\packages\NUnit.ConsoleRunner.3.8.0\tools\nunit3-console.exe
SET OPEN_COVER=..\packages\OpenCover.4.6.519\tools\OpenCover.Console.exe
SET REPORT_GEN=..\packages\ReportGenerator.3.1.2\tools\ReportGenerator.exe

REM カバレッジ測定の対象
SET PROJECT_NAME=OutlookRecipientConfirmationAddin

REM カバレッジレポート(XML)のファイル名
SET REPORT_NAME=result.xml

REM 実行するテストのアセンブリと、その格納先
SET TEST=.\bin\Debug\ORCAUnitTest.dll
SET COVERAGE_DIR=.\bin\Debug

REM カバレッジレポート(HTML)の出力先
SET OUTPUT_DIR=..\html\

REM #######################

call :EXECUTE "%TEST%"
pause
REM 生成したレポートの表示
start "" %OUTPUT_DIR%\index.htm

exit

:EXECUTE
%OPEN_COVER% -register:user -target:"%NUNIT%" -targetargs:"ORCAUnitTest.dll" -targetdir:"%COVERAGE_DIR%" -filter:"PROJECT_NAME" -output:"%REPORT_NAME%" -mergebyhash

%REPORT_GEN% --reports:"%REPORT_NAME%" --targetdir:%OUTPUT_DIR%

exit /b