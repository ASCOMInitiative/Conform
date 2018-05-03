@echo on
@echo Setting up variables
call "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community\VC\Auxiliary\Build\vcvarsall.bat" x86
@echo Selecting directory: C:\Users\Peter\Documents\Visual Studio Projects\Conform
cd "C:\Users\Peter\Documents\Visual Studio Projects\Conform"

@echo Building Release x86
msbuild /t:rebuild /p:configuration=Release;Platform=x86

@echo Building Release x64
msbuild /t:rebuild /p:configuration=Release;Platform=x64

@echo Building Release AnyCPU
msbuild /t:rebuild /p:configuration=Release;Platform="Any CPU"

@echo Building Debug x86
msbuild /t:rebuild /p:configuration=Debug;Platform=x86

@echo Building Debug x64
msbuild /t:rebuild /p:configuration=Debug;Platform=x64

@echo Building Debug AnyCPU
msbuild /t:rebuild /p:configuration=Debug;Platform="Any CPU"
