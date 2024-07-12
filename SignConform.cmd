@echo on
@echo Setting up variables
call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat" x86

@echo Parameter: %1

echo "%1" | findstr /C:"uninst" 1>nul

if errorlevel 1 (
	echo Executables already signed on first pass... 
) ELSE (
	echo Signing compiled executables
	echo signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Debug\Conform.exe"
	signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Debug\Conform.exe"

	echo signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Release\Conform.exe"
	signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Release\Conform.exe"

	echo signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Debug\Conform.exe"
	signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Debug\Conform.exe"

	echo signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Release\Conform.exe"
	signtool sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Release\Conform.exe"
)

echo Signing installer %1

REM The delays provide time for the executable to be released by the previous process ready for the next to proceed
echo Starting 1 second wait
timeout /T 1
echo Finished wait
start "Sign with SHA256" /wait /b "signtool" sign /a /fd SHA256 /n "Peter Simpson" /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1

rem pause

