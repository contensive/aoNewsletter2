
@echo off

rem 
rem Must be run from the projects git\project\scripts folder - everything is relative
rem run >build [deploymentNumber]
rem deploymentNumber is YYMMDD.build-number, like 190824.5
rem
rem Setup deployment folder
rem

rem all paths are relative to the git scripts folder
rem
rem GIT folder
rem     -- aoSample
rem			-- collection
rem				-- Sample
rem					unzipped collection files, must include one .xml file describing the collection
rem			-- server 
rem 			(all files related to server code)
rem				-- aoSample (visual studio project folder)
rem			-- ui 
rem				(all files related to the ui
rem			-- etc 
rem				(all misc files)

rem -- the application on the local server where this collection will be installed
set appName=app200509

rem -- major version 5, minor does not matter set 1
set majorVersion=5
set minorVersion=1

rem -- name of the collection on the site (should NOT include ao prefix). This is the name as it appears on the navigator
set collectionName=Newsletter

rem -- name of the collection folder, (should NOT include ao prefix)
set collectionPath=..\collections\Newsletter\

rem -- name of the solution. SHOULD include ao prefix
set solutionName=aoNewsletter2.sln

rem -- name of the solution. SHOULD include ao prefix
set binPath=..\server\aoNewsletter2\bin\debug\

rem -- name of the solution. SHOULD include ao prefix
set msbuildLocation=C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\

rem -- name of the solution. SHOULD include ao prefix
set deploymentFolderRoot=C:\Deployments\aoNewsletter2\Dev\

rem Setup deployment folder

set year=%date:~12,4%
set month=%date:~4,2%
if %month% GEQ 10 goto monthOk
set month=%date:~5,1%
:monthOk
set day=%date:~7,2%
if %day% GEQ 10 goto dayOk
set day=%date:~8,1%
:dayOk
set versionMajor=%year%
set versionMinor=%month%
set versionBuild=%day%
set versionRevision=1
rem
rem if deployment folder exists, delete it and make directory
rem
:tryagain
set versionNumber=%versionMajor%.%versionMinor%.%versionBuild%.%versionRevision%
if not exist "%deploymentFolderRoot%%versionNumber%" goto :makefolder
set /a versionRevision=%versionRevision%+1
goto tryagain
:makefolder
md "%deploymentFolderRoot%%versionNumber%"



rem ==============================================================
rem
echo build 
rem
cd ..\server
"%msbuildLocation%msbuild.exe" %solutionName%
if errorlevel 1 (
   echo failure building
   pause
   exit /b %errorlevel%
)
cd ..\scripts


rem ==============================================================
rem
echo Build addon collection
rem

rem remove old DLL files from the collection folder
del "%collectionPath%"\*.DLL
del "%collectionPath%"\*.config

rem copy bin folder assemblies to collection folder
copy "%binPath%*.dll" "%collectionPath%"

rem create new collection zip file
c:
cd %collectionPath%
del "%collectionName%.zip" /Q
"c:\program files\7-zip\7z.exe" a "%collectionName%.zip"
xcopy "%collectionName%.zip" "%deploymentFolderRoot%%versionNumber%" /Y
cd ..\..\scripts

