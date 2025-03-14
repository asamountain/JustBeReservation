# Gmail 예약 스크래퍼 (Gmail Reservation Scraper)

이 프로그램은 Gmail에서 예약 확정 이메일을 자동으로 스크래핑하여 게스트 이름, 날짜 등의 정보를 추출하고 Excel 파일로 저장합니다. 또한 Cloudbeds에 자동 로그인하는 기능도 포함되어 있습니다.

## 주요 기능

- Gmail에서 예약 확정 이메일 자동 스크래핑
- 이메일 제목에서 게스트 이름 추출
- 데이터를 Excel 파일로 저장
- 데이터를 GUI 창에 표시
- Cloudbeds 자동 로그인
- 1시간마다 자동 업데이트

## 초보자를 위한 간편 설치 방법 (원클릭 설치)

컴퓨터에 익숙하지 않은 분들을 위한 간편 설치 방법입니다:

1. 이 프로그램 폴더에 있는 `setup.bat` 파일을 더블클릭하세요.
2. 설치가 자동으로 진행됩니다. 설치 중에는 검은색 창(PowerShell)이 열리고 진행 상황이 표시됩니다.
3. 설치가 완료되면 로그인 정보를 입력하는 창이 나타납니다.
4. Gmail과 Cloudbeds 계정 정보를 입력하세요.
5. 설정이 완료되면 바탕화면에 `Gmail 예약 스크래퍼` 바로가기가 생성됩니다.
6. 이 바로가기를 더블클릭하여 프로그램을 실행할 수 있습니다.

## 수동 설치 방법 (고급 사용자용)

### 1. Python 설치

먼저 Python 3.8 이상을 설치해야 합니다.

1. [Python 공식 웹사이트](https://www.python.org/downloads/)에서 최신 버전의 Python을 다운로드합니다.
2. 설치 과정에서 **"Add Python to PATH"** 옵션을 반드시 체크하세요.
3. 설치가 완료되면 PowerShell을 열고 다음 명령어로 Python이 제대로 설치되었는지 확인합니다:
   ```
   python --version
   ```

### 2. PowerShell 열기

Windows에서 PowerShell을 여는 방법:
1. 키보드에서 `Windows 키 + R`을 누릅니다.
2. 실행 창에 `powershell`을 입력하고 Enter 키를 누릅니다.
3. 파란색 창(PowerShell)이 열립니다.

### 3. 필요한 라이브러리 설치

PowerShell 창에서 다음 명령어를 실행하여 필요한 모든 라이브러리를 한 번에 설치합니다:

```
pip install selenium webdriver-manager pandas openpyxl schedule xlwings
```

또는 프로젝트 폴더로 이동한 후 다음 명령어를 실행할 수도 있습니다:

```
cd "프로그램이 저장된 폴더 경로"
pip install -r requirements.txt
```

### 4. Chrome 웹 브라우저 설치

이 프로그램은 Chrome 웹 브라우저를 사용합니다. 아직 설치되어 있지 않다면 [Chrome 공식 웹사이트](https://www.google.com/chrome/)에서 다운로드하여 설치하세요.

## 보안 및 민감 정보 관리

이 프로그램은 Gmail과 Cloudbeds 계정 정보를 필요로 합니다. 보안을 위해 다음과 같은 방법으로 민감 정보를 관리하세요:

### 방법 1: 환경 변수 사용 (권장)

1. `config.py` 파일을 생성하여 다음과 같이 작성합니다:
   ```python
   import os

   # 환경 변수에서 로그인 정보 가져오기
   GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL", "")
   GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD", "")
   CLOUDBEDS_EMAIL = os.environ.get("CLOUDBEDS_EMAIL", "")
   CLOUDBEDS_PASSWORD = os.environ.get("CLOUDBEDS_PASSWORD", "")
   ```

2. `GmailScraper.py` 파일의 상단에 다음 코드를 추가합니다:
   ```python
   try:
       from config import GMAIL_EMAIL, GMAIL_PASSWORD, CLOUDBEDS_EMAIL, CLOUDBEDS_PASSWORD
   except ImportError:
       # 기본값 설정
       GMAIL_EMAIL = "your_email@example.com"
       GMAIL_PASSWORD = "your_password"
       CLOUDBEDS_EMAIL = "your_email@example.com"
       CLOUDBEDS_PASSWORD = "your_password"
   ```

3. 환경 변수 설정 방법:
   - Windows: PowerShell에서 다음 명령어 실행
     ```
     [Environment]::SetEnvironmentVariable("GMAIL_EMAIL", "your_email@example.com", "User")
     [Environment]::SetEnvironmentVariable("GMAIL_PASSWORD", "your_password", "User")
     [Environment]::SetEnvironmentVariable("CLOUDBEDS_EMAIL", "your_email@example.com", "User")
     [Environment]::SetEnvironmentVariable("CLOUDBEDS_PASSWORD", "your_password", "User")
     ```

### 방법 2: 로컬 설정 파일 사용

1. `.env` 파일을 생성하여 다음과 같이 작성합니다:
   ```
   GMAIL_EMAIL=your_email@example.com
   GMAIL_PASSWORD=your_password
   CLOUDBEDS_EMAIL=your_email@example.com
   CLOUDBEDS_PASSWORD=your_password
   ```

2. `.gitignore` 파일에 다음 내용을 추가하여 민감 정보가 GitHub에 업로드되지 않도록 합니다:
   ```
   # 민감 정보 파일
   .env
   config.py
   credentials.json
   ```

3. `python-dotenv` 라이브러리를 설치합니다:
   ```
   pip install python-dotenv
   ```

4. `GmailScraper.py` 파일의 상단에 다음 코드를 추가합니다:
   ```python
   import os
   from dotenv import load_dotenv

   # .env 파일에서 환경 변수 로드
   load_dotenv()

   # 환경 변수에서 로그인 정보 가져오기
   GMAIL_EMAIL = os.getenv("GMAIL_EMAIL", "your_email@example.com")
   GMAIL_PASSWORD = os.getenv("GMAIL_PASSWORD", "your_password")
   CLOUDBEDS_EMAIL = os.getenv("CLOUDBEDS_EMAIL", "your_email@example.com")
   CLOUDBEDS_PASSWORD = os.getenv("CLOUDBEDS_PASSWORD", "your_password")
   ```

## 사용 방법

1. 프로그램 파일(`GmailScraper.py`)을 원하는 폴더에 저장합니다.

2. 위의 '보안 및 민감 정보 관리' 섹션에 따라 로그인 정보를 설정합니다.

3. PowerShell을 열고 프로그램 파일이 있는 폴더로 이동한 후 다음 명령어를 실행합니다:
   ```
   python GmailScraper.py
   ```

4. 프로그램이 실행되면 다음과 같은 작업이 자동으로 수행됩니다:
   - Chrome 브라우저가 열리고 Gmail에 로그인합니다.
   - 예약 확정 이메일을 검색하고 데이터를 추출합니다.
   - 추출된 데이터를 Excel 파일(`gmail_data.xlsx`)로 저장합니다.
   - 데이터를 보여주는 GUI 창이 열립니다.
   - Cloudbeds에 자동으로 로그인합니다.
   - 1시간마다 위 과정을 반복합니다.

5. 프로그램을 종료하려면 PowerShell 창에서 `Ctrl+C`를 누르세요.

## 바로가기 만들기 (선택 사항)

프로그램을 더 쉽게 실행하기 위해 바탕화면에 바로가기를 만들 수 있습니다:

1. 메모장을 열고 다음 내용을 입력합니다:
   ```
   @echo off
   cd /d "프로그램이 저장된 폴더 경로"
   python GmailScraper.py
   pause
   ```

2. 이 파일을 `Run_GmailScraper.bat`로 저장합니다.
3. 이 배치 파일을 바탕화면에 복사하거나 바로가기를 만듭니다.
4. 이제 이 바로가기를 더블클릭하여 프로그램을 실행할 수 있습니다.

## 문제 해결

### Chrome 드라이버 오류

프로그램이 Chrome 드라이버를 찾지 못하는 경우, `webdriver-manager` 라이브러리가 자동으로 적절한 드라이버를 다운로드합니다. 그러나 문제가 계속되면 다음 명령어로 라이브러리를 업데이트해 보세요:

```
pip install --upgrade webdriver-manager
```

### 로그인 문제

Gmail이나 Cloudbeds 로그인에 문제가 있는 경우:
1. 로그인 정보가 정확한지 확인하세요.
2. 2단계 인증이 활성화된 경우, 앱 비밀번호를 생성하여 사용하세요.
3. 처음 실행 시 수동으로 로그인해야 할 수도 있습니다.

### 데이터 추출 문제

이메일 제목에서 게스트 이름이 제대로 추출되지 않는 경우, `extract_guest_name` 함수의 정규식 패턴을 확인하고 필요에 따라 수정하세요.

## 추가 정보

- 프로그램은 `chrome_profile` 폴더를 생성하여 로그인 세션을 저장합니다.
- Excel 파일은 프로그램이 실행되는 폴더에 저장됩니다.
- 각 실행마다 새로운 시트가 생성되며, 시트 이름에는 타임스탬프가 포함됩니다.

## 연락처

문제가 발생하거나 도움이 필요한 경우 IT 부서에 문의하세요. 