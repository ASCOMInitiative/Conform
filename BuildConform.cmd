@echo on
@echo Setting up variables
call "C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\VC\Auxiliary\Build\vcvarsall.bat" x86
@echo Selecting directory: Conform

SET ConformDirectory=%cd%
@echo Current directory: %ConformDirectory%

cd %ConformDirectory%
@echo Building Release x86 in directory %cd%
msbuild /t:rebuild /p:configuration=Release;Platform=x86

cd %ConformDirectory%
@echo Building Release x64 in directory %cd%
msbuild /t:rebuild /p:configuration=Release;Platform=x64

cd %ConformDirectory%
@echo Building Release AnyCPU in directory %cd%
msbuild /t:rebuild /p:configuration=Release;Platform="Any CPU"

cd %ConformDirectory%
@echo Building Debug x86 in directory %cd%
msbuild /t:rebuild /p:configuration=Debug;Platform=x86

cd %ConformDirectory%
@echo Building Debug x64 in directory %cd%
msbuild /t:rebuild /p:configuration=Debug;Platform=x64

cd %ConformDirectory%

@echo Building Debug AnyCPU in directory %cd%
msbuild /t:rebuild /p:configuration=Debug;Platform="Any CPU"
 