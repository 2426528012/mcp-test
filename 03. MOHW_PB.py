import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
import time
import sys
import os
from pathlib import Path

# --- 설정 (Constants) ---
BASE_URL = "https://www.mohw.go.kr"
BOARD_URL = "https://www.mohw.go.kr/board.es?mid=a10504000000&bid=0030&cg_code="
MAX_RETRIES = 3
RETRY_DELAY = 1  # seconds

def get_start_date():
    """사용자로부터 유효한 시작일을 YYYY-MM-DD 형식으로 입력받습니다."""
    while True:
        date_str = input(f"▶ 데이터 수집을 시작할 날짜를 입력하세요 (예: {datetime.now().strftime('%Y-%m-%d')}): ")
        try:
            start_date = datetime.strptime(date_str, "%Y-%m-%d").date()
            return start_date
        except ValueError:
            print("오류: 날짜 형식이 올바르지 않습니다. YYYY-MM-DD 형식으로 다시 입력해주세요.")

def make_request(url):
    """지정된 URL에 대해 재시도 로직을 포함하여 GET 요청을 보냅니다."""
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.get(url, headers=headers, timeout=10)
            response.raise_for_status()
            return response
        except requests.exceptions.RequestException as e:
            print(f"네트워크 요청 실패 (시도 {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
            else:
                print("오류: 웹사이트에 연결할 수 없습니다. 스크래핑을 중단합니다.")
                return None

def main():
    """메인 스크래핑 로직을 실행합니다."""
    start_date = get_start_date()
    scraped_data = []
    current_page = 1
    stop_scraping = False

    print("\n데이터 수집을 시작합니다...")

    while not stop_scraping:
        page_url = f"{BOARD_URL}&cpage={current_page}"
        print(f"페이지 {current_page} 수집 중... ({page_url})")

        response = make_request(page_url)
        if not response:
            break

        soup = BeautifulSoup(response.text, 'html.parser')
        
        # [수정됨] 제공해주신 HTML 구조 기반의 가장 정확한 선택자로 변경
        rows = soup.select('.board_list > table > tbody > tr')

        # --- [수정된 최종 중단 로직] ---
        # 페이지에 게시물(tr)이 하나도 없으면, 더 이상 볼 것이 없으므로 즉시 종료합니다.
        if not rows:
            print("-> 이 페이지에서 게시물을 찾을 수 없습니다. 스크래핑을 최종 종료합니다.")
            break

        for row in rows:
            try:
                # 위치 기반으로 각 셀(td)을 정확하게 선택
                num_cell = row.select_one('td:nth-of-type(1)')
                
                # '번호' 셀의 텍스트가 숫자가 아니면 (공지사항 등) 건너뜀
                if not num_cell or not num_cell.text.strip().isdigit():
                    continue

                date_cell = row.select_one('td:nth-of-type(4)')
                date_str = date_cell.text.strip()
                post_date = datetime.strptime(date_str, "%Y-%m-%d").date()

                # 날짜가 시작일보다 오래되었으면 중단 신호를 켜고, 현재 페이지의 나머지 게시물 확인을 중단
                if post_date < start_date:
                    stop_scraping = True
                    break

                # 데이터 수집
                num = int(num_cell.text.strip())
                
                title_cell = row.select_one('td:nth-of-type(2) > a')
                title = title_cell.text.strip().replace('"', "'") # Excel HYPERLINK 함수 오류 방지
                relative_link = title_cell['href']
                absolute_link = f"{BASE_URL}{relative_link}"
                title_hyperlink = f'=HYPERLINK("{absolute_link}", "{title}")'

                department = row.select_one('td:nth-of-type(3)').text.strip()
                views = int(row.select_one('td:nth-of-type(5)').text.strip().replace(',', ''))

                scraped_data.append({
                    '번호': num,
                    '제목': title_hyperlink,
                    '담당부서': department,
                    '등록일': date_str,
                    '조회수': views
                })

            except Exception as e:
                print(f"주의: 특정 행을 처리하는 중 오류 발생. 해당 항목을 건너뜁니다. (오류: {e})")
                continue
        
        # for loop를 탈출한 이유가 stop_scraping 때문이라면, while 루프도 탈출
        if stop_scraping:
            print(f"-> 설정한 시작일({start_date}) 이전 게시물에 도달하여 수집을 중단합니다.")
            break
        
        current_page += 1
        time.sleep(0.5)

    # --- 파일 저장 로직 (이하 동일) ---
    script_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"{script_name}_{timestamp}.xlsx"

    try:
        downloads_path = Path.home() / "Downloads"
        downloads_path.mkdir(exist_ok=True)
        save_path = downloads_path / filename
    except Exception:
        save_path = Path(filename)
        print(f"주의: 'Downloads' 폴더를 찾을 수 없어 스크립트가 있는 현재 폴더에 저장합니다.")

    if not scraped_data:
        print("\n수집된 데이터가 없습니다.")
        df = pd.DataFrame([{'결과': '지정된 기간 내에 수집된 내용이 없습니다.'}])
    else:
        print(f"\n총 {len(scraped_data)}개의 데이터를 수집했습니다.")
        df = pd.DataFrame(scraped_data, columns=['번호', '제목', '담당부서', '등록일', '조회수'])

    try:
        df.to_excel(save_path, index=False, engine='openpyxl')
        print(f"\n✅ 성공: 스크래핑 결과를 다음 경로에 저장했습니다.\n   {save_path.resolve()}")
    except Exception as e:
        print(f"오류: 파일을 저장하는 데 실패했습니다. ({e})")

if __name__ == '__main__':
    main()
    print("\nEND")