import os
import time
import re
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

# --- 1. 설정 (Configuration) ---
BASE_URL = "https://www.nhis.or.kr/nhis/together/wbhaec05600m01.do"
COLUMNS = ["번호", "제목", "첨부유무", "등록일", "조회수"]

def get_start_date():
    """
    사용자로부터 YYYY-MM-DD 형식의 시작일을 입력받고 유효성을 검사합니다.
    """
    while True:
        start_date_str = input("데이터 수집을 시작할 날짜를 입력하세요 (YYYY-MM-DD 형식): ")
        try:
            # 날짜 형식 검사
            start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            # 미래 날짜 검사 (오늘 자정까지 허용)
            if start_date > datetime.now().replace(hour=23, minute=59, second=59):
                print("오류: 미래 날짜는 입력할 수 없습니다. 다시 입력해주세요.")
                continue
            return start_date
        except ValueError:
            print("오류: 날짜 형식이 잘못되었습니다. YYYY-MM-DD 형식으로 다시 입력해주세요.")

def make_request(url, retries=3, delay=1):
    """
    지정된 URL에 대해 안정적인 HTTP GET 요청을 보냅니다. 실패 시 재시도합니다.
    """
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    for i in range(retries):
        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()  # 200 OK가 아니면 예외 발생
            return response
        except requests.exceptions.RequestException as e:
            print(f"네트워크 요청 실패: {e} ({i+1}/{retries} 번째 시도)")
            time.sleep(delay)
    print(f"오류: {retries}번의 재시도 후에도 페이지를 가져올 수 없습니다.")
    return None

def scrape_board():
    """
    게시판 데이터를 스크래핑하는 메인 함수입니다.
    """
    start_date = get_start_date()
    scraped_data = []
    page_offset = 0
    stop_scraping = False

    print("\n데이터 수집을 시작합니다...")

    while not stop_scraping:
        # 페이지 URL 구성 (article.offset은 0, 10, 20, ...)
        page_url = f"{BASE_URL}?mode=list&articleLimit=10&article.offset={page_offset}"
        print(f"페이지 스크래핑 중: {page_url}")

        response = make_request(page_url)
        if not response:
            break # 요청 실패 시 중단

        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 게시판 테이블의 tbody 요소를 선택
        table_body = soup.select_one('div.col-table > table > tbody')
        if not table_body:
            print("게시물 테이블을 찾을 수 없습니다. 스크립트를 종료합니다.")
            break

        rows = table_body.select('tr')
        if not rows:
            print("더 이상 게시물이 없습니다.")
            break

        for row in rows:
            try:
                cells = row.select('td')

                # 번호 - '공지' 등 숫자 아닌 경우 건너뛰기
                post_num_str = cells[0].get_text(strip=True)
                if not post_num_str.isdigit():
                    continue

                # 등록일 파싱 및 비교
                post_date_str = cells[3].get_text(strip=True).replace('.', '-')
                post_date = datetime.strptime(post_date_str, "%Y-%m-%d")

                if post_date < start_date:
                    stop_scraping = True
                    print(f"'{start_date.strftime('%Y-%m-%d')}' 이전의 게시물에 도달하여 수집을 중단합니다.")
                    break # 현재 페이지의 나머지 게시물 처리 중단

                # 데이터 추출 및 정제
                post_num = int(post_num_str)
                
                title_tag = cells[1].select_one('a')
                title_text = title_tag.get_text(strip=True)
                # 상대 경로를 절대 URL로 변환
                detail_link = requests.compat.urljoin(BASE_URL, title_tag['href'])
                excel_title = f'=HYPERLINK("{detail_link}", "{title_text}")'
                
                # 첨부파일 유무 확인
                has_attachment = 'Y' if cells[2].select_one('div.attach-list') else 'N'

                # 조회수에서 쉼표 제거 및 숫자로 변환
                views = int(cells[4].get_text(strip=True).replace(',', ''))

                scraped_data.append([post_num, excel_title, has_attachment, post_date_str, views])

            except Exception as e:
                # 개별 항목 파싱 실패 시 건너뛰기
                print(f"항목 파싱 오류 발생: {e}. 해당 항목을 건너뛰고 계속합니다.")
                continue

        # 다음 페이지로 이동
        page_offset += 10
        time.sleep(0.5) # 서버 부하를 줄이기 위한 짧은 대기

    return scraped_data

def save_to_excel(data):
    """
    수집된 데이터를 Excel 파일로 저장합니다.
    """
    # 사용자 PC의 다운로드 폴더 경로 확인
    try:
        downloads_path = str(Path.home() / "Downloads")
    except Exception:
        downloads_path = "." # 경로를 찾지 못할 경우 현재 폴더에 저장
        print("다운로드 폴더를 찾지 못해 현재 폴더에 저장합니다.")

    # 동적 파일명 생성
    script_name = os.path.basename(__file__).replace('.py', '')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{script_name}_{timestamp}.xlsx"
    output_path = os.path.join(downloads_path, filename)

    # 데이터프레임 생성
    if not data:
        print("수집된 내용이 없어 메시지를 포함한 파일을 생성합니다.")
        df = pd.DataFrame([{"결과": "수집된 내용이 없습니다."}])
    else:
        df = pd.DataFrame(data, columns=COLUMNS)

    # Excel 파일로 저장
    try:
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"\n성공: 데이터가 '{output_path}' 경로에 저장되었습니다.")
    except Exception as e:
        print(f"\n오류: Excel 파일 저장에 실패했습니다. - {e}")

# --- 4. 스크립트 실행 ---
if __name__ == "__main__":
    collected_data = scrape_board()
    save_to_excel(collected_data)
    print("END")