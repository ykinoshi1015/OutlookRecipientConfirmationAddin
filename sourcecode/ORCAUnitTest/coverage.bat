REM ###### Settings ######

SET PROJECT_NAME=OutlookRecipientConfirmationAddin
SET NUNIT=.\packages\NUnit.ConsoleRunner.3.8.0\tools\nunit3-console.exe

SET REPORT_NAME=result.xml
SET OUTPUT_DIR=.\html

SET OPEN_COVER=.\packages\OpenCover.4.6.519\tools\OpenCover.Console.exe
SET REPORT_GEN=.\packages\ReportGenerator.3.1.2\tools\ReportGenerator.exe

SET TEST=.\ORCAUnitTest\bin\Debug\ORCAUnitTest.dll
SET COVERAGE_DIR=.\TestApp\bin\Debug\
SET FILTERS=+[TestApp]*

REM #######################

call :EXECUTE "%TEST%"

REM 生成したレポートの表示
start "" %OUTPUT_DIR%\index.htm

exit

:EXECUTE
%OPEN_COVER% -register:user -target:"%NUNIT%" -targetargs:"\"%~f1\"" -targetdir:%COVERAGE_DIR% -filter:"%FILTERS%" -output:%REPORT_NAME% -mergebyhash
%REPORT_GEN% --reports:"%REPORT_NAME%" --targetdir:%OUTPUT_DIR%

exit /b