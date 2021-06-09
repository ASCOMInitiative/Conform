@echo on
@echo Setting up variables
call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat" x86
@echo Selecting directory: Conform

SET ConformDirectory=%cd%
@echo Current directory: %ConformDirectory%

@echo Building Conform in directory %cd%
msbuild Setup\buildconform.msbuild
