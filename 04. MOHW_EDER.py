import os
import time
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, WebDriverException

def get_start_date():
    """
    사용자로부터 유효한 시작일을 입력받습니다.
    """
    while True:
        start_date_str = input("데이터 수집을 시작할 날짜를 입력하세요 (YYYY-MM-DD 형식): ")
        try:
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            if start_date > datetime.now().date():
                print("오류: 시작일은 미래 날짜일 수 없습니다. 다시 입력해주세요.")
                continue
            return start_date
        except ValueError:
            print("오류: 날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 입력해주세요.")

def setup_driver():
    """
    Selenium WebDriver를 설정하고 반환합니다.
    """
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    
    try:
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except (WebDriverException, ValueError) as e:
        print(f"WebDriver 설정 중 오류가 발생했습니다: {e}")
        print("\n[문제 해결 제안]")
        print("1. Chrome 브라우저가 최신 버전인지 확인해주세요.")
        print("2. 터미널에서 'pip install --upgrade webdriver-manager' 명령어로 라이브러리를 업데이트해보세요.")
        print("3. C:\\Users\\{사용자이름}\\.wdm\\drivers 폴더를 삭제하고 다시 시도해보세요.")
        return None

def scrape_data(driver, start_date):
    """
    Selenium을 사용하여 웹사이트 데이터를 스크래핑합니다.
    """
    base_url = "https://www.mohw.go.kr/law.es?mid=a10409010000"
    scraped_data = []
    page_num = 1
    stop_scraping = False

    driver.get(base_url)
    
    try:
        total_count_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".board_info .page .total b"))
        )
        total_count = int(total_count_element.text.replace('건', '').strip())
    except (TimeoutException, ValueError):
        print("전체 게시물 수를 가져올 수 없어 번호는 역순으로 기록됩니다.")
        total_count = 0

    while not stop_scraping:
        print(f"{page_num}페이지에서 데이터를 수집합니다...")
        
        url = f"{base_url}&pageNum={page_num}"
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".board_list > table > tbody > tr"))
            )
            rows = driver.find_elements(By.CSS_SELECTOR, ".board_list > table > tbody > tr")
            if not rows:
                print("더 이상 게시물이 없습니다.")
                break
        except TimeoutException:
            print(f"{page_num}페이지에서 데이터를 찾을 수 없어 스크래핑을 중단합니다.")
            break
        except WebDriverException as e:
            print(f"네트워크 오류 또는 페이지 로딩 오류 발생 (시도 3회 후 중단): {e}")
            break # 재시도 로직 추가 가능

        for i, row in enumerate(rows):
            try:
                cells = row.find_elements(By.TAG_NAME, "td")
                
                promulgation_date_str = cells[3].text.strip()
                try:
                    promulgation_date = datetime.strptime(promulgation_date_str, "%Y-%m-%d").date()
                except ValueError:
                    print(f"경고: 날짜 형식 오류 ('{promulgation_date_str}'). 해당 행을 건너뜁니다.")
                    continue

                if promulgation_date < start_date:
                    stop_scraping = True
                    print(f"수집 기간({start_date})을 벗어난 데이터를 발견하여 스크래핑을 중단합니다.")
                    break
                
                num = total_count - ((page_num - 1) * 10) - i if total_count > 0 else int(cells[0].text.strip())
                promulgation_num = cells[1].text.strip()
                
                title_tag = cells[2].find_element(By.TAG_NAME, 'a')
                title_text = re.sub(r'\s*새글\s*', '', title_tag.text).strip()
                link_url = title_tag.get_attribute('href').replace("&", "&")
                hyperlink = f'=HYPERLINK("{link_url}", "{title_text}")'
                
                enforcement_date = cells[4].text.strip()

                scraped_data.append([num, promulgation_num, hyperlink, promulgation_date_str, enforcement_date])

            except Exception as e:
                print(f"경고: 특정 행 파싱 중 오류 발생. 건너뜁니다. 오류: {e}")
                continue
        
        if stop_scraping: break
        page_num += 1
        time.sleep(0.5)

    return pd.DataFrame(scraped_data, columns=["번호", "공포번호", "법령명", "공포일자", "시행일"])

def save_to_excel(df):
    """
    수집된 데이터를 Excel 파일로 저장합니다.
    """
    columns = ["번호", "공포번호", "법령명", "공포일자", "시행일"]
    download_path = Path.home() / "Downloads"
    download_path.mkdir(exist_ok=True)

    # 스크립트 파일명을 직접 지정 (os.path.basename 사용 시 환경에 따라 문제 발생 가능)
    script_name = "mohw_scraper"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{script_name}_{timestamp}.xlsx"
    file_path = download_path / filename

    if df.empty:
        print("수집된 데이터가 없습니다. 빈 Excel 파일을 생성합니다.")
        df = pd.DataFrame([["수집된 내용이 없습니다."] + [''] * (len(columns) - 1)], columns=columns)

    try:
        df.to_excel(file_path, index=False, engine='openpyxl')
        print(f"\n성공: 데이터가 '{file_path}' 경로에 저장되었습니다.")
        print("참고: 법령명 링크는 브라우저에서 직접 접속 시 '페이지 접속 실패'가 발생할 수 있으며, 이는 정상입니다.")
    except Exception as e:
        print(f"오류: Excel 파일 저장에 실패했습니다. 오류: {e}")

if __name__ == '__main__':
    print("웹 스크래핑을 시작합니다.")
    
    start_date = get_start_date()
    driver = setup_driver()
    
    if driver:
        try:
            data_df = scrape_data(driver, start_date)
            if not data_df.empty:
                data_df = data_df.sort_values(by="번호", ascending=False).reset_index(drop=True)
            save_to_excel(data_df)
        finally:
            driver.quit()
    
    print("\nEND")
