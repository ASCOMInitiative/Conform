@echo on
@echo Setting up variables
call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat" x86

@echo Parameter: %1

echo "%1" | findstr /C:"uninst" 1>nul

if errorlevel 1 (
	echo got errorlevel one -  pattern missing
) ELSE (
	echo got errorlevel zero - found pattern
	echo signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Debug\Conform.exe"
	signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Debug\Conform.exe"
	signtool sign /a /as /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Debug\Conform.exe"

	echo signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Release\Conform.exe"
	signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Release\Conform.exe"
	signtool sign /a /as /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\Release\Conform.exe"

	echo signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Debug\Conform.exe"
	signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Debug\Conform.exe"
	signtool sign /a /as /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Debug\Conform.exe"

	echo signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Release\Conform.exe"
	signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Release\Conform.exe"
	signtool sign /a /as /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 "J:\Conform\Conform\bin\x86\Release\Conform.exe" 
)

echo signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1
signtool sign /a /fd SHA256 /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1
signtool sign /a /as /tr http://rfc3161timestamp.globalsign.com/advanced /td SHA256 %1

