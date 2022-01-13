@echo on
@echo Setting up variables
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x86
@echo Selecting directory: Conform

SET ConformDirectory=%cd%
@echo Current directory: %ConformDirectory%

@echo Building Conform in directory %cd%
msbuild Setup\buildconform.msbuild
