REM ###### Settings ######


REM NUnit, OpenCover, ReportGenerator�̃C���X�g�[����
SET NUNIT=..\packages\NUnit.ConsoleRunner.3.8.0\tools\nunit3-console.exe
SET OPEN_COVER=..\packages\OpenCover.4.6.519\tools\OpenCover.Console.exe
SET REPORT_GEN=..\packages\ReportGenerator.3.1.2\tools\ReportGenerator.exe

REM �J�o���b�W����̑Ώ�
SET PROJECT_NAME=OutlookRecipientConfirmationAddin

REM �J�o���b�W���|�[�g(XML)�̃t�@�C����
SET REPORT_NAME=result.xml

REM ���s����e�X�g�̃A�Z���u���ƁA���̊i�[��
SET TEST=.\bin\Debug\ORCAUnitTest.dll
SET COVERAGE_DIR=.\bin\Debug

REM �J�o���b�W���|�[�g(HTML)�̏o�͐�
SET OUTPUT_DIR=..\html\

REM #######################

call :EXECUTE "%TEST%"
pause
REM �����������|�[�g�̕\��
start "" %OUTPUT_DIR%\index.htm

exit

:EXECUTE
%OPEN_COVER% -register:user -target:"%NUNIT%" -targetargs:"ORCAUnitTest.dll" -targetdir:"%COVERAGE_DIR%" -filter:"PROJECT_NAME" -output:"%REPORT_NAME%" -mergebyhash

%REPORT_GEN% --reports:"%REPORT_NAME%" --targetdir:%OUTPUT_DIR%

exit /b