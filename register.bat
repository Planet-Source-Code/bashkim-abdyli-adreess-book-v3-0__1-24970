@ECHO Off 
CLS
ECHO -------------------------------------------------------------------------
ECHO -------------------------------------------------------------------------
ECHO Hi, This is Used To Register or UnRegister The File "vbSendMail.dll". 
ECHO Make sure that file is in the same path as this file.
ECHO -------------------------------------------------------------------------
ECHO -------------------------------------------------------------------------
ECHO .
ECHO 1 - Register "vbSendMail.dll". 
ECHO 2 - UNRegister "vbSendMail.dll". 
ECHO 3 - Exit.

CHOICE /C:123 /N Please choose a number : 
IF ERRORLEVEL 3 Goto End
IF ERRORLEVEL 2 GOTO UNREGISTER 
IF ERRORLEVEL 1 GOTO REGISTER 
:REGISTER
ECHO .
ECHO .
ECHO Registering "vbSendMail.dll"
 C:\Windows\System\regsvr32.exe vbSendMail.dll
GOTO End 
:UNREGISTER
ECHO .
ECHO .
ECHO UnRegistering "vbSendMail.dll"
C:\Windows\System\regsvr32.exe /u vbSendMail.dll
GOTO End 
:End 
Exit 
