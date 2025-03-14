@echo off
echo ===================================================
echo      Gmail 예약 스크래퍼 설치 프로그램
echo ===================================================
echo.

:: 관리자 권한으로 실행 확인
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo 관리자 권한이 필요합니다. 관리자 권한으로 다시 실행해주세요...
    pause
    exit /b
)

echo 설치를 시작합니다. 잠시만 기다려주세요...
echo.

:: Python 설치 확인
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Python이 설치되어 있지 않습니다. Python을 설치합니다...
    echo 브라우저가 열리면 Python 설치 프로그램을 다운로드하고 실행하세요.
    echo 설치 시 "Add Python to PATH" 옵션을 반드시 체크하세요.
    timeout /t 3 > nul
    start https://www.python.org/downloads/
    echo Python 설치가 완료되면 이 창으로 돌아와서 아무 키나 누르세요.
    pause
    
    :: Python 설치 다시 확인
    python --version > nul 2>&1
    if %errorlevel% neq 0 (
        echo Python 설치에 실패했습니다. 수동으로 설치해주세요.
        pause
        exit /b
    )
)

echo Python이 설치되어 있습니다. 계속 진행합니다...
echo.

:: 필요한 라이브러리 설치
echo 필요한 라이브러리를 설치합니다...
pip install selenium webdriver-manager pandas openpyxl schedule xlwings python-dotenv > nul 2>&1
if %errorlevel% neq 0 (
    echo 라이브러리 설치에 실패했습니다. 다시 시도합니다...
    pip install selenium webdriver-manager pandas openpyxl schedule xlwings python-dotenv
)
echo 라이브러리 설치가 완료되었습니다.
echo.

:: Chrome 설치 확인
reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version > nul 2>&1
if %errorlevel% neq 0 (
    echo Chrome 브라우저가 설치되어 있지 않습니다. Chrome을 설치합니다...
    echo 브라우저가 열리면 Chrome 설치 프로그램을 다운로드하고 실행하세요.
    timeout /t 3 > nul
    start https://www.google.com/chrome/
    echo Chrome 설치가 완료되면 이 창으로 돌아와서 아무 키나 누르세요.
    pause
)

echo Chrome 브라우저가 설치되어 있습니다. 계속 진행합니다...
echo.

:: 로그인 정보 설정
echo 로그인 정보를 설정합니다...
echo.

:: .env 파일 생성
set /p gmail_email=Gmail 이메일 주소를 입력하세요: 
set /p gmail_password=Gmail 비밀번호를 입력하세요: 
set /p cloudbeds_email=Cloudbeds 이메일 주소를 입력하세요: 
set /p cloudbeds_password=Cloudbeds 비밀번호를 입력하세요: 

echo GMAIL_EMAIL=%gmail_email%> .env
echo GMAIL_PASSWORD=%gmail_password%>> .env
echo CLOUDBEDS_EMAIL=%cloudbeds_email%>> .env
echo CLOUDBEDS_PASSWORD=%cloudbeds_password%>> .env

echo 로그인 정보가 .env 파일에 저장되었습니다.
echo.

:: .gitignore 파일 생성
echo # 민감 정보 파일> .gitignore
echo .env>> .gitignore
echo config.py>> .gitignore
echo credentials.json>> .gitignore
echo chrome_profile/>> .gitignore

:: 바탕화면에 바로가기 생성
echo 바탕화면에 바로가기를 생성합니다...
echo @echo off> Run_GmailScraper.bat
echo cd /d "%~dp0">> Run_GmailScraper.bat
echo python GmailScraper.py>> Run_GmailScraper.bat
echo pause>> Run_GmailScraper.bat

:: 바로가기 생성
set SCRIPT="%TEMP%\create_shortcut.vbs"
echo Set oWS = WScript.CreateObject("WScript.Shell")> %SCRIPT%
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\Gmail 예약 스크래퍼.lnk">> %SCRIPT%
echo Set oLink = oWS.CreateShortcut(sLinkFile)>> %SCRIPT%
echo oLink.TargetPath = "%~dp0Run_GmailScraper.bat">> %SCRIPT%
echo oLink.WorkingDirectory = "%~dp0">> %SCRIPT%
echo oLink.Description = "Gmail 예약 스크래퍼">> %SCRIPT%
echo oLink.Save>> %SCRIPT%
cscript /nologo %SCRIPT%
del %SCRIPT%

echo 바탕화면에 바로가기가 생성되었습니다.
echo.

echo ===================================================
echo      설치가 완료되었습니다!
echo ===================================================
echo.
echo 바탕화면에 생성된 'Gmail 예약 스크래퍼' 바로가기를 더블클릭하여 프로그램을 실행하세요.
echo.
echo 문제가 발생하면 IT 부서에 문의하세요.
echo.
pause 