@ECHO OFF
REM #################################################################################
REM # �������@�bShapingXMLfileTool�i�N���p�o�b�`�j
REM # �@�\�@�@�bPowerShell�N���p�̃o�b�`
REM #--------------------------------------------------------------------------------
REM # �@�@�@�@�b-
REM #################################################################################
ECHO *---------------------------------------------------------
ECHO *
ECHO *  ShapingXMLfileTool
ECHO *
ECHO *---------------------------------------------------------
ECHO.
ECHO.
SET RETURNCODE=0
powershell -NoProfile -ExecutionPolicy Unrestricted -File .\source\powershell\Main.ps1
SET RETURNCODE=%ERRORLEVEL%

ECHO.
ECHO �������I�����܂����B
ECHO �����ꂩ�̃L�[�������ƃE�B���h�E�����܂��B
PAUSE > NUL
EXIT %RETURNCODE%
