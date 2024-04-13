@echo off
setlocal enabledelayedexpansion

:: Navigate to the script's directory
cd %~dp0

:: Recursively find all .deploy files and rename them by removing the .deploy extension
for /R %%i in (*.deploy) do (
    set "FILENAME=%%i"
    set "NEWNAME=!FILENAME:~0,-7!"
    echo Renaming: "!FILENAME!" to "!NEWNAME!"
    rename "%%i" "%%~ni"
)

echo All '.deploy' extensions have been removed.


@echo off
setlocal enabledelayedexpansion

:: Step 1: Loop through each main directory
for /d %%D in (*) do (
    pushd %%D
    
    :: Step 2: Delete files in the main subdirectory
    del /q *.*
    
    :: Step 3: Navigate to the "Application Files" directory and delete its contents
    if exist "Application Files" (
        pushd "Application Files"
        del /q *.*
        
        :: Step 4: Navigate to the first subdirectory under "Application Files"
        for /d %%F in (*) do (
            pushd %%F
            
            :: Step 5: Copy all items back two levels up
            xcopy /s /e /h /y *.* "..\..\"
            
            popd
        )
        
        :: Move back to the main subdirectory
        popd

        :: Step 6: Delete the "Application Files" folder entirely
        rmdir /s /q "Application Files"
    )
    popd
)

echo Operation completed.
