from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import os
import re
import subprocess
import platform
from datetime import datetime
import schedule
import sys
import xlwings as xw
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

# 환경 변수에서 로그인 정보 로드 시도
try:
    # .env 파일이 있는지 확인
    if os.path.exists('.env'):
        # python-dotenv 라이브러리 사용 시도
        try:
            from dotenv import load_dotenv
            load_dotenv()
            print("환경 변수를 .env 파일에서 로드했습니다.")
        except ImportError:
            print(".env 파일이 있지만 python-dotenv 라이브러리가 설치되지 않았습니다.")
            print("pip install python-dotenv 명령어로 설치하거나 로그인 정보를 직접 입력하세요.")
    
    # 환경 변수에서 로그인 정보 가져오기
    GMAIL_EMAIL = os.environ.get("GMAIL_EMAIL", "")
    GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD", "")
    CLOUDBEDS_EMAIL = os.environ.get("CLOUDBEDS_EMAIL", "")
    CLOUDBEDS_PASSWORD = os.environ.get("CLOUDBEDS_PASSWORD", "")
    
    # 로그인 정보가 비어있는지 확인
    if not all([GMAIL_EMAIL, GMAIL_PASSWORD, CLOUDBEDS_EMAIL, CLOUDBEDS_PASSWORD]):
        raise ValueError("로그인 정보가 비어있습니다.")
        
except Exception as e:
    print(f"환경 변수에서 로그인 정보를 로드하는 중 오류 발생: {str(e)}")
    print("로그인 정보를 직접 입력하세요.")
    
    # GUI로 로그인 정보 입력 받기
    def get_login_info():
        root = tk.Tk()
        root.title("로그인 정보 입력")
        root.geometry("400x250")
        
        # 입력 필드 생성
        tk.Label(root, text="Gmail 이메일:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        gmail_email = tk.Entry(root, width=30)
        gmail_email.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(root, text="Gmail 비밀번호:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        gmail_pw = tk.Entry(root, width=30, show="*")
        gmail_pw.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(root, text="Cloudbeds 이메일:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        cloudbeds_email = tk.Entry(root, width=30)
        cloudbeds_email.grid(row=2, column=1, padx=10, pady=10)
        
        tk.Label(root, text="Cloudbeds 비밀번호:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        cloudbeds_pw = tk.Entry(root, width=30, show="*")
        cloudbeds_pw.grid(row=3, column=1, padx=10, pady=10)
        
        # 저장 여부 체크박스
        save_var = tk.BooleanVar()
        save_check = tk.Checkbutton(root, text="로그인 정보 저장하기", variable=save_var)
        save_check.grid(row=4, column=0, columnspan=2, pady=10)
        
        # 결과 저장 변수
        result = {"submitted": False}
        
        # 확인 버튼 클릭 시 실행할 함수
        def on_submit():
            result["gmail_email"] = gmail_email.get()
            result["gmail_password"] = gmail_pw.get()
            result["cloudbeds_email"] = cloudbeds_email.get()
            result["cloudbeds_password"] = cloudbeds_pw.get()
            result["save"] = save_var.get()
            result["submitted"] = True
            root.destroy()
        
        # 확인 버튼
        submit_btn = tk.Button(root, text="확인", command=on_submit)
        submit_btn.grid(row=5, column=0, columnspan=2, pady=10)
        
        # 창이 닫힐 때 실행할 함수
        def on_closing():
            if messagebox.askokcancel("종료", "로그인 정보를 입력하지 않으면 프로그램이 종료됩니다.\n정말 종료하시겠습니까?"):
                root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # GUI 실행
        root.mainloop()
        
        return result
    
    # 로그인 정보 입력 받기
    login_info = get_login_info()
    
    if not login_info["submitted"]:
        print("로그인 정보가 입력되지 않았습니다. 프로그램을 종료합니다.")
        sys.exit(1)
    
    # 로그인 정보 설정
    GMAIL_EMAIL = login_info["gmail_email"]
    GMAIL_PASSWORD = login_info["gmail_password"]
    CLOUDBEDS_EMAIL = login_info["cloudbeds_email"]
    CLOUDBEDS_PASSWORD = login_info["cloudbeds_password"]
    
    # 로그인 정보 저장
    if login_info["save"]:
        try:
            with open(".env", "w") as f:
                f.write(f"GMAIL_EMAIL={GMAIL_EMAIL}\n")
                f.write(f"GMAIL_PASSWORD={GMAIL_PASSWORD}\n")
                f.write(f"CLOUDBEDS_EMAIL={CLOUDBEDS_EMAIL}\n")
                f.write(f"CLOUDBEDS_PASSWORD={CLOUDBEDS_PASSWORD}\n")
            print("로그인 정보가 .env 파일에 저장되었습니다.")
            
            # .gitignore 파일 생성
            if not os.path.exists(".gitignore"):
                with open(".gitignore", "w") as f:
                    f.write("# 민감 정보 파일\n")
                    f.write(".env\n")
                    f.write("config.py\n")
                    f.write("credentials.json\n")
                    f.write("chrome_profile/\n")
        except Exception as e:
            print(f"로그인 정보 저장 중 오류 발생: {str(e)}")

# URLs
GMAIL_URL = "https://mail.google.com/mail/u/0/#inbox"
CLOUDBEDS_URL = "https://hotels.cloudbeds.com/auth#/calendar"

def setup_driver():
    """WebDriver 설정 및 초기화"""
    # Chrome 사용자 프로필 디렉토리 설정
    user_data_dir = os.path.join(os.getcwd(), "chrome_profile")
    if not os.path.exists(user_data_dir):
        os.makedirs(user_data_dir)

    # Chrome WebDriver 설정
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def check_and_login_gmail(driver):
    """Gmail 로그인 상태 확인 및 필요시 로그인 수행"""
    try:
        # 5초 동안 로그인 필드를 찾아봄
        email_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "identifier"))
        )
        # 로그인 필드가 발견되면 로그인 진행
        email_field.send_keys(GMAIL_EMAIL)
        email_field.send_keys(Keys.RETURN)
        
        time.sleep(3)  # 페이지 전환 대기
        password_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.NAME, "Passwd"))
        )
        password_field.send_keys(GMAIL_PASSWORD)
        password_field.send_keys(Keys.RETURN)
        
        time.sleep(5)  # Gmail 메인 화면 로딩 대기
    except:
        # 로그인 필드를 찾지 못하면 이미 로그인된 상태
        print("Gmail: 이미 로그인되어 있거나 다른 상태입니다.")
        pass

def login_cloudbeds(driver):
    """Cloudbeds 로그인 수행"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Cloudbeds 로그인 시도...")
    
    try:
        # Cloudbeds 로그인 페이지로 이동
        driver.get(CLOUDBEDS_URL)
        time.sleep(3)  # 페이지 로딩 대기
        
        # 이메일 입력
        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']"))
        )
        email_field.send_keys(CLOUDBEDS_EMAIL)
        
        # Next 버튼 클릭
        next_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        next_button.click()
        time.sleep(2)
        
        # 비밀번호 입력
        password_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']"))
        )
        password_field.send_keys(CLOUDBEDS_PASSWORD)
        
        # Login 버튼 클릭
        login_button = driver.find_element(By.CSS_SELECTOR, "button[type='submit']")
        login_button.click()
        
        # 로그인 성공 확인 (캘린더 페이지 로딩 대기)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[class*='calendar']"))
        )
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Cloudbeds 로그인 성공")
        
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Cloudbeds 로그인 실패: {str(e)}")
        raise

def search_emails(driver, keyword):
    """Gmail에서 주어진 키워드로 이메일 검색"""
    search_box = driver.find_element(By.CSS_SELECTOR, "form[role='search'] input")
    search_box.clear()
    search_box.send_keys(keyword)
    search_box.send_keys(Keys.RETURN)
    time.sleep(5)  # 검색 결과 로딩 대기

def clean_link(link):
    """Clean up the Gmail link to get the actual Airbnb URL"""
    # Try to extract airbnb.co.kr URL
    airbnb_match = re.search(r'(https?://[^\s<>"]+?airbnb\.co\.kr[^\s<>"]*)', link)
    if airbnb_match:
        return airbnb_match.group(1)
    return link

def extract_guest_name(title):
    """이메일 제목에서 게스트 이름 추출"""
    try:
        print(f"Processing title: {title}")  # 디버그 로깅 추가
        
        # 스크린샷에서 확인된 실제 이메일 제목 패턴
        patterns = [
            # 예약 확정 패턴 (스크린샷에서 확인됨)
            r'예약 확정 - ([^님\s]+(?:\s[^님\s]+)*)님이 \d+월 \d+일에 체크인할 예정입니다',
            r'예약 확정 - ([^님\s]+(?:\s[^님\s]+)*)님의 \d+월 \d+일에 체크인할 예정입니다',
            
            # Reservation CONFIRMED 패턴 (스크린샷에서 확인됨)
            r'Reservation CONFIRMED - ([^-]+) - Reservation #\d+',
            r'Reservation CONFIRMED - ([^-\s]+(?:\s[^-\s]+)*) - Reservation',
            
            # 이름 뒤에 님이 오는 패턴 (스크린샷에서 확인됨)
            r'예약 확정 - ([^님\s]+(?:\s[^님\s]+)*)님이',
            r'예약 확정 - ([^님\s]+(?:\s[^님\s]+)*)님의',
            
            # 기존 패턴들
            r'Airbnb 예약 확인 - ([^님]+)님의 예약',
            r'예약 확인: ([^님]+)님의 예약',
            r'Airbnb 예약 확인 - ([^님]+)(?:님)?$',
            r'Airbnb 예약 확인: ([^님]+)님의 예약',
            r'Airbnb 예약 확인 - ([^님]+)님의 예약이 확정되었습니다',
            r'Airbnb 예약이 확정되었습니다 - ([^님]+)님의 예약',
            r'Airbnb 예약이 확정되었습니다: ([^님]+)님의 예약',
            r'Airbnb 예약이 확정되었습니다 - ([^님]+)님',
            r'Airbnb 예약이 확정되었습니다: ([^님]+)님',
            r'예약이 확정되었습니다 - ([^님]+)님의 예약',
            r'예약이 확정되었습니다: ([^님]+)님의 예약',
            r'예약이 확정되었습니다 - ([^님]+)님',
            r'예약이 확정되었습니다: ([^님]+)님',
            r'Airbnb 예약 확인 - ([^님]+)(?:님)?$',
            r'예약 확인 - ([^님]+)(?:님)?$',
            r'예약이 확정되었습니다 - ([^님]+)(?:님)?$',
            r'Airbnb 예약 - ([^님]+)님',
            r'예약 - ([^님]+)님',
            r'Airbnb 예약: ([^님]+)님',
            r'예약: ([^님]+)님'
        ]
        
        for pattern in patterns:
            name_match = re.search(pattern, title)
            if name_match:
                name = name_match.group(1).strip()
                print(f"Found name: {name} using pattern: {pattern}")  # 디버그 로깅 추가
                return name
        
        # 특수 케이스: "정입니다" 처리
        if "정입니다" in title:
            # 이름 추출 시도
            name_match = re.search(r'예약 확정 - ([^님\s]+(?:\s[^님\s]+)*)', title)
            if name_match:
                name = name_match.group(1).strip()
                print(f"Found name from '정입니다' case: {name}")
                return name
        
        # 이름을 찾지 못한 경우 제목에서 직접 추출 시도
        # 1. 예약 확정 - 다음에 오는 이름 추출
        reservation_match = re.search(r'예약 확정 - ([^-\s]+(?:\s[^-\s]+)*)', title)
        if reservation_match:
            name = reservation_match.group(1).strip()
            if len(name) > 1 and len(name) < 30:  # 합리적인 이름 길이 확인
                print(f"Found name after '예약 확정 - ': {name}")
                return name
        
        # 2. 제목에서 마지막으로 등장하는 한글 이름 추출 (2-4글자)
        korean_name_match = re.search(r'[가-힣]{1,4}(?:님|$)', title)
        if korean_name_match:
            name = korean_name_match.group(0).replace('님', '')
            print(f"Found Korean name: {name}")
            return name
            
        # 3. 영문 이름 추출 시도
        english_name_match = re.search(r'([A-Z][a-z]+(?:\s[A-Z][a-z]+)*)', title)
        if english_name_match:
            name = english_name_match.group(1).strip()
            if len(name) > 1:  # 합리적인 이름 길이 확인
                print(f"Found English name: {name}")
                return name
        
        print(f"Failed to extract name from title: {title}")
        return title.split(' - ')[1] if ' - ' in title else "게스트"
        
    except Exception as e:
        print(f"이름 추출 중 오류 발생: {str(e)}")
        print(f"문제가 발생한 제목: {title}")
        return "게스트"

def get_email_data(driver):
    """이메일 제목, 날짜, 게스트 이름을 가져오기"""
    # Wait for emails to load
    wait = WebDriverWait(driver, 10)
    
    # Wait for the main table container to be present
    table_container = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[role='main'] table"))
    )
    
    # Find tbody within the table
    try:
        tbody = table_container.find_element(By.TAG_NAME, "tbody")
        # Get all email rows from tbody
        email_rows = tbody.find_elements(By.CSS_SELECTOR, "tr")
        data = []
        
        for row in email_rows:
            try:
                # Get email title - using a more reliable selector
                title_element = row.find_element(By.CSS_SELECTOR, "td[role='gridcell'] span.bog")
                title = title_element.text.strip()
                
                # Get date - using a more reliable selector
                date_element = row.find_element(By.CSS_SELECTOR, "td[role='gridcell'] span[title]")
                date = date_element.get_attribute("title") or date_element.text.strip()
                
                # Extract guest name from title
                guest_name = extract_guest_name(title)
                
                if title and date:  # Only add if we have both title and date
                    print(f"Found email: {title[:50]}... | {date} | Guest: {guest_name}")  # Debug output
                    data.append([title, date, guest_name])
            
            except Exception as e:
                print(f"Error processing row: {str(e)}")
                continue
        
        return data
    
    except Exception as e:
        print(f"Error finding tbody element: {str(e)}")
        return []

def display_data_in_window(data, filename):
    """Display data in a tkinter window with a table view"""
    try:
        # Create the main window
        root = tk.Tk()
        root.title(f"Gmail Data - {os.path.basename(filename)}")
        root.geometry("1200x600")  # Set window size
        
        # Create a frame for the table
        frame = ttk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create Treeview
        columns = ("제목", "날짜", "게스트 이름")
        tree = ttk.Treeview(frame, columns=columns, show='headings')
        
        # Set column headings
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)  # Default width
        
        # Add scrollbars
        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        
        # Grid layout
        tree.grid(row=0, column=0, sticky='nsew')
        yscroll.grid(row=0, column=1, sticky='ns')
        xscroll.grid(row=1, column=0, sticky='ew')
        
        # Configure grid weights
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        
        # Insert data
        for row in data:
            tree.insert('', tk.END, values=row)
        
        # Auto-adjust column widths
        for col in columns:
            max_width = max(len(str(row[columns.index(col)])) for row in data)
            tree.column(col, width=min(max_width * 10, 400))  # Cap at 400 pixels
        
        # Add a button to close the window
        close_button = ttk.Button(root, text="Close", command=root.destroy)
        close_button.pack(pady=10)
        
        # Start the GUI event loop
        root.mainloop()
        
    except Exception as e:
        print(f"Error displaying data window: {str(e)}")
        messagebox.showerror("Error", f"Failed to display data: {str(e)}")

def save_to_excel(data, filename="gmail_data.xlsx"):
    """데이터를 엑셀 파일로 저장하고 GUI 창에 표시"""
    if not data:
        print("No data to save!")
        return
        
    try:
        # Get absolute path for the file
        abs_path = os.path.abspath(filename)
        print(f"Saving to: {abs_path}")
        
        # Generate a unique sheet name
        sheet_name = f"Data_{time.strftime('%Y%m%d_%H%M%S')}"
        
        try:
            # Create DataFrame
            df = pd.DataFrame(data, columns=["제목", "날짜", "게스트 이름"])
            
            # Save to Excel
            df.to_excel(filename, sheet_name=sheet_name, index=False)
            print(f"\n엑셀 파일 저장 완료: {filename}")
            print(f"총 {len(data)}개의 이메일 데이터가 저장되었습니다.")
            
            # Display data in GUI window
            display_data_in_window(data, filename)
            
        except Exception as e:
            print(f"Error during Excel operations: {str(e)}")
            # Try saving as CSV as fallback
            csv_filename = f"gmail_data_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            print(f"Attempting to save as CSV instead: {csv_filename}")
            
            # Use pandas to save CSV
            df.to_csv(csv_filename, index=False, encoding='utf-8-sig')
            print(f"Data saved to CSV: {csv_filename}")
            
            # Display data in GUI window
            display_data_in_window(data, csv_filename)
            
    except Exception as e:
        print(f"파일 저장 중 오류 발생: {str(e)}")
        print("\n데이터 미리보기:")
        for idx, row in enumerate(data[:5]):
            print(f"Row {idx + 1}: {row}")

def scrape_gmail():
    """Gmail 스크래핑 메인 함수"""
    print(f"\n[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 스크래핑 시작...")
    
    try:
        driver = setup_driver()
        
        # Gmail 처리
        driver.get(GMAIL_URL)
        check_and_login_gmail(driver)
        search_emails(driver, "label:.예약확정.")
        data = get_email_data(driver)
        
        # 엑셀 파일 저장
        if data:
            save_to_excel(data)
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 데이터 업데이트 완료")
        else:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 새로운 데이터가 없습니다")
            
        # Cloudbeds 로그인
        login_cloudbeds(driver)
        
    except Exception as e:
        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 오류 발생: {str(e)}")
    
    finally:
        try:
            driver.quit()
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 브라우저 종료")
        except:
            pass

def main():
    """메인 실행 함수"""
    print("Gmail 스크래퍼 시작...")
    print("1시간 간격으로 실행됩니다.")
    print(f"첫 실행 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 초기 실행
    scrape_gmail()
    
    # 1시간마다 실행되도록 스케줄 설정
    schedule.every(1).hours.do(scrape_gmail)
    
    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # 1분마다 스케줄 체크
            
    except KeyboardInterrupt:
        print("\n프로그램 종료 요청됨...")
        sys.exit(0)
    except Exception as e:
        print(f"\n예상치 못한 오류 발생: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
