import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from datetime import datetime, date
import os
from urllib.parse import urljoin

# --- Configuration ---
# 1. 스크래핑 대상 URL 정보
BASE_URL = "https://www.mohw.go.kr/board.es"
URL = f"{BASE_URL}?mid=a10503010100&bid=0027"

# 2. 결과 파일을 저장할 경로 (기본값: 사용자 홈 디렉토리의 Downloads 폴더)
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Downloads")

# 3. 출력 컬럼명
COLUMN_NAMES = ['번호', '제목', '담당부서', '작성일', '조회수']

# --- Functions ---

def get_start_date():
    """
    사용자로부터 유효한 시작일을 입력받습니다.
    """
    while True:
        input_date_str = input("▶ 스크래핑 시작일을 입력하세요 (YYYY-MM-DD 형식): ")
        try:
            start_dt = datetime.strptime(input_date_str, "%Y-%m-%d").date()
            if start_dt > date.today():
                print("오류: 시작일은 오늘보다 미래일 수 없습니다. 다시 입력해주세요.")
                continue
            return start_dt
        except ValueError:
            print("오류: 날짜 형식이 올바르지 않습니다. 'YYYY-MM-DD' 형식으로 다시 입력해주세요.")

def scrape_website(start_date):
    """
    웹사이트를 스크래핑하여 데이터를 수집합니다.
    """
    print("\n[1/3] 스크래핑을 시작합니다...")
    all_posts = []
    page_num = 1
    scraping_finished = False
    
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    })

    while not scraping_finished:
        print(f"  - {page_num} 페이지를 스크래핑 중...")
        
        page_url = f"{URL}&nPage={page_num}"
        response = None
        for attempt in range(3):
            try:
                response = session.get(page_url, timeout=10)
                response.raise_for_status()
                break
            except requests.exceptions.RequestException as e:
                print(f"    - 요청 실패 (시도 {attempt + 1}/3): {e}")
                if attempt < 2:
                    time.sleep(2)
                else:
                    print("    - 3회 재시도 실패. 스크래핑을 중단합니다.")
                    return None

        soup = BeautifulSoup(response.text, 'html.parser')
        rows = soup.select("div.board_list table.tstyle_list tbody tr")

        if not rows:
            print("  - 더 이상 게시글이 없어 스크래핑을 종료합니다.")
            break

        for row in rows:
            cells = row.select("td")
            
            post_num_text = cells[0].get_text(strip=True)
            if not post_num_text.isdigit():
                continue

            post_date_str = cells[3].get_text(strip=True)
            post_date = datetime.strptime(post_date_str, "%Y-%m-%d").date()
            
            if post_date < start_date:
                scraping_finished = True
                break
            
            number = int(post_num_text)
            
            # --- [수정된 부분] 제목 클리닝 로직 강화 ---
            title_tag = cells[1].select_one("a")
            link = urljoin(BASE_URL, title_tag['href'])

            # '새글' 아이콘(i)과 숨겨진 텍스트(span)를 모두 제거
            if title_tag.find('i'):
                title_tag.find('i').decompose()
            if title_tag.find('span', class_='sr_only'):
                title_tag.find('span', class_='sr_only').decompose()
            
            # 정리된 텍스트 추출
            title = title_tag.get_text(strip=True)
            
            # 제목에 큰따옴표(")가 있을 경우를 대비해 이스케이프 처리
            sanitized_title = title.replace('"', '""')
            hyperlink_formula = f'=HYPERLINK("{link}", "{sanitized_title}")'
            # --- 수정 끝 ---
            
            department = cells[2].get_text(strip=True)
            views = int(cells[4].get_text(strip=True).replace(",", ""))

            all_posts.append([number, hyperlink_formula, department, post_date_str, views])

        if scraping_finished:
            print(f"  - 시작일({start_date}) 이전의 게시물에 도달하여 스크래핑을 종료합니다.")
            break
        
        page_num += 1
        time.sleep(1)

    print(f"[1/3] 스크래핑 완료! 총 {len(all_posts)}개의 게시글을 수집했습니다.")
    return all_posts

def save_to_excel(data, script_filename):
    """
    수집된 데이터를 엑셀 파일로 저장합니다.
    """
    print("\n[2/3] 엑셀 파일 저장을 시작합니다...")
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{script_filename}_{timestamp}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)

    if not data:
        df = pd.DataFrame(["내용없음"])
        df.to_excel(filepath, index=False, header=False)
    else:
        df = pd.DataFrame(data, columns=COLUMN_NAMES)
        df.to_excel(filepath, index=False)

    print(f"[2/3] 엑셀 파일 저장 완료!")
    print(f"  - 파일 경로: {os.path.abspath(filepath)}")

def main():
    """
    메인 실행 함수
    """
    try:
        script_name = os.path.splitext(os.path.basename(__file__))[0]
    except NameError:
        script_name = "mohw_scraping_script"

    start_dt = get_start_date()
    scraped_data = scrape_website(start_dt)

    if scraped_data is not None:
        save_to_excel(scraped_data, script_name)

    print("\n[3/3] 작업이 모두 완료되었습니다.")
    input("▶ 엔터 키를 누르면 창이 닫힙니다...")

if __name__ == "__main__":
    main()