import requests
import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime
import os
import sys
import traceback

def scrape_mohw_board(from_date: str, to_date: str, output_filepath: str):
    """
    보건복지부 공지사항 게시판을 스크래핑하여 지정된 경로에 엑셀 파일로 저장합니다.
    """
    base_url = "https://www.mohw.go.kr"
    board_url = f"{base_url}/board.es"
    params = {"mid": "a10501010100", "bid": "0003"}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}

    all_posts = []
    page = 1
    stop_scraping = False
    from_date_obj = datetime.strptime(from_date, "%Y-%m-%d").date()
    to_date_obj = datetime.strptime(to_date, "%Y-%m-%d").date()

    print(f"\n'{from_date}'부터 '{to_date}'(오늘)까지의 게시글 스크래핑을 시작합니다...")

    while not stop_scraping:
        params["nPage"] = page
        print(f"페이지 {page} 스크래핑 중...")
        try:
            response = requests.get(board_url, params=params, headers=headers, timeout=10)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find('table', class_='tstyle_list')
            tbody = table.find('tbody') if table else None
            if not tbody:
                print("게시물 목록 테이블을 찾을 수 없습니다. 중단합니다.")
                break
            rows = tbody.find_all('tr')
            if not rows:
                print("더 이상 게시글이 없습니다. 스크래핑을 종료합니다.")
                break

            for row in rows:
                tds = row.find_all('td')
                if len(tds) < 6: continue

                post_number_str = tds[0].get_text(strip=True)
                if not post_number_str.isdigit(): continue

                post_date_str = tds[3].get_text(strip=True)
                try:
                    post_date = datetime.strptime(post_date_str.replace('.', '-'), "%Y-%m-%d").date()
                except ValueError: continue

                if post_date > to_date_obj: continue
                if post_date < from_date_obj:
                    stop_scraping = True
                    print(f"'{from_date_obj}' 이전 날짜({post_date})의 게시글에 도달하여 중단합니다.")
                    break

                title_tag = tds[2].find('a')
                if not title_tag: continue

                # --- [수정됨] '새글' 텍스트 제거 로직 ---
                # 'sr_only' 클래스를 가진 span 태그(스크린 리더용 '새글' 텍스트)를 찾아서 제거
                sr_only_span = title_tag.find('span', class_='sr_only')
                if sr_only_span:
                    sr_only_span.decompose()  # 태그를 트리에서 완전히 제거

                # 이제 순수한 제목 텍스트만 가져옴
                title_text = title_tag.get_text(strip=True)
                # --- 수정 끝 ---
                
                relative_url = title_tag['href']
                full_url = requests.compat.urljoin(base_url, relative_url)
                excel_hyperlink = f'=HYPERLINK("{full_url}", "{title_text.replace("\"", "\"\"")}")'

                post_data = {
                    "번호": int(post_number_str),
                    "분류": tds[1].get_text(strip=True),
                    "제목": excel_hyperlink,
                    "작성일": post_date_str,
                    "첨부파일 유무": bool(tds[4].find('img')),
                    "조회수": int(tds[5].get_text(strip=True).replace(',', '') or 0)
                }
                all_posts.append(post_data)

        except Exception as e:
            print(f"페이지 처리 중 예기치 않은 오류 발생: {e}")
            traceback.print_exc()
            break
            
        if stop_scraping: break
        page += 1
        time.sleep(1)

    if not all_posts:
        print(f"\n결과: 해당 기간('{from_date}' ~ '{to_date}')에 수집된 게시글이 없습니다.")
        return
        
    print(f"\n총 {len(all_posts)}개의 게시글을 수집했습니다.")
    df = pd.DataFrame(all_posts, columns=["번호", "분류", "제목", "작성일", "첨부파일 유무", "조회수"])
    
    try:
        output_dir = os.path.dirname(output_filepath)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            print(f"출력 폴더를 생성했습니다: {output_dir}")
        df.to_excel(output_filepath, index=False, engine='openpyxl')
        print(f"데이터를 성공적으로 저장했습니다: {output_filepath}")
    except Exception as e:
        print(f"엑셀 파일 저장에 실패했습니다: {e}")

def get_validated_date(prompt: str) -> str:
    """사용자로부터 'YYYY-MM-DD' 형식의 날짜를 입력받고 유효성을 검사하는 함수"""
    while True:
        date_str = input(prompt)
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
            return date_str
        except ValueError:
            print("오류: 날짜 형식이 잘못되었습니다. 'YYYY-MM-DD' 형식으로 다시 입력해주세요.")

if __name__ == "__main__":
    print("--- 보건복지부 공지사항 스크래핑 ---")
    try:
        today_date = datetime.now().date()
        while True:
            start_date_str = get_validated_date("시작일을 입력하세요 (YYYY-MM-DD): ")
            start_date_obj = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            if start_date_obj > today_date:
                print(f"오류: 시작일({start_date_str})이 오늘({today_date.strftime('%Y-%m-%d')})보다 미래일 수 없습니다. 다시 입력해주세요.\n")
            else:
                break
        
        end_date_str = today_date.strftime("%Y-%m-%d")
        
        script_name = os.path.splitext(os.path.basename(sys.argv[0]))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_filename = f"{script_name}_{timestamp}.xlsx"
        save_path = r"C:\Users\ildong\Downloads"
        final_filepath = os.path.join(save_path, output_filename)
        
        scrape_mohw_board(from_date=start_date_str, to_date=end_date_str, output_filepath=final_filepath)

    except Exception as e:
        print("\n---!!! 스크립트 실행 중 치명적인 오류 발생 !!!---")
        traceback.print_exc()
    
    finally:
        print("\n--- 스크립트 실행 종료 ---")
        input("엔터 키를 누르면 창이 닫힙니다...")