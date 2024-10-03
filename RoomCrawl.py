import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup as bs
import re
from itertools import product

def table_to_2d(table_tag):
    rowspans = []  # track pending rowspans
    rows = table_tag.find_all('tr')

    # first scan, see how many columns we need
    colcount = 0
    for r, row in enumerate(rows):
        cells = row.find_all(['td', 'th'], recursive=False)
        colcount = max(
            colcount,
            sum(int(c.get('colspan', 1)) or 1 for c in cells[:-1]) + len(cells[-1:]) + len(rowspans))
        rowspans += [int(c.get('rowspan', 1)) or len(rows) - r for c in cells]
        rowspans = [s - 1 for s in rowspans if s > 1]

    # build an empty matrix for all possible cells
    table = [[None] * colcount for row in rows]

    # fill matrix from row data
    rowspans = {}  # track pending rowspans
    for row, row_elem in enumerate(rows):
        span_offset = 0  # how many columns are skipped due to row and colspans 
        for col, cell in enumerate(row_elem.find_all(['td', 'th'], recursive=False)):
            col += span_offset
            while rowspans.get(col, 0):
                span_offset += 1
                col += 1

            # fill table data
            rowspan = rowspans[col] = int(cell.get('rowspan', 1)) or len(rows) - row
            colspan = int(cell.get('colspan', 1)) or colcount - col
            span_offset += colspan - 1
            value = cell.get_text(strip=True)
            for drow, dcol in product(range(rowspan), range(colspan)):
                try:
                    table[row + drow][col + dcol] = value
                    rowspans[col + dcol] = rowspan
                except IndexError:
                    pass

        # update rowspan bookkeeping
        rowspans = {c: s - 1 for c, s in rowspans.items() if s > 1}

    return table

# 크롬 드라이버 설정 및 실행
browser = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
browser.implicitly_wait(2)

# 네이버 로그인 페이지로 이동
url = 'https://nid.naver.com/nidlogin.login'
browser.get(url)

# 네이버 아이디와 비밀번호 입력
id = 'mrpn'
pw = 'Sippal0561'

browser.execute_script(f"document.getElementsByName('id')[0].value='{id}'")
browser.execute_script(f"document.getElementsByName('pw')[0].value='{pw}'")

# 로그인 버튼 클릭
login_button = browser.find_element(By.XPATH, '//*[@id="log.login"]')
login_button.click()
time.sleep(2)

# 카페 페이지로 이동
baseurl = 'https://cafe.naver.com/yullee2k/'
browser.get(baseurl)

# 특정 카테고리(menu_id=12)로 이동
menu_id = '12'
clubid = '13027583'
userDisplay = '50'

browser.get(f'{baseurl}ArticleList.nhn?search.clubid={clubid}&search.menuid={menu_id}&search.boardtype=L&userDisplay={userDisplay}')

# iframe으로 전환
browser.switch_to.frame('cafe_main')

# 페이지 소스를 파싱
soup = bs(browser.page_source, 'html.parser')

# 모든 게시글 링크 중 articleid를 추출하여 가장 큰 id를 가진 게시글 선택
articles = soup.select('a.article')
article_dict = {}

for article in articles:
    href = article.get('href')
    articleid = re.search(r'articleid=(\d+)', href)
    if articleid:
        articleid = int(articleid.group(1))
        article_dict[articleid] = href

# 가장 큰 articleid 찾기
latest_articleid = max(article_dict.keys())
latest_article_link = article_dict[latest_articleid]

# 최신 게시글로 이동하여 내용 크롤링
browser.get(baseurl + latest_article_link)

# 게시글 내용을 다시 iframe으로 접근하여 파싱
browser.switch_to.frame('cafe_main')
soup = bs(browser.page_source, 'html.parser')

# 지정된 CSS 선택자를 사용하여 <table> 요소 추출
table_content = soup.select_one('#SE-ced69d60-f048-4a36-9083-7cec56cb29ec > div > div > div > table')

# HTML 형식으로 저장
html_filename = 'latest_article_table.html'

# HTML 파일로 저장
with open(html_filename, 'w', encoding='utf-8') as file:
    if table_content:  # <table>이 존재할 경우에만 저장
        file.write(str(table_content))
    else:
        file.write('<p>No table found in the latest article.</p>')

# 테이블 내용을 2D 배열로 변환 후 DataFrame으로 생성
if table_content:
    table_2d = table_to_2d(table_content)  # 2D 배열로 변환
    df = pd.DataFrame(table_2d)

    # Excel 파일로 저장
    excel_filename = 'latest_article_table.xlsx'
    df.to_excel(excel_filename, index=False, header=False)  # 인덱스 없이 저장

    # Excel 파일에서 1~6행 삭제
    df = pd.read_excel(excel_filename, header=None)  # 인덱스 없이 읽기
    df = df.drop(index=[0, 1, 2, 3, 4, 5])  # 1~6행 삭제

    # 다시 저장
    df.to_excel(excel_filename, index=False, header=False)  # 인덱스 없이 저장
    print(f"Excel 파일에서 1~6행 삭제 후 저장 완료: {excel_filename}")

    # CSV 파일로 저장
    csv_filename = 'latest_article_table.csv'
    df.to_csv(csv_filename, index=False, header=False, encoding='utf-8-sig')  # 인덱스 없이 저장
    print(f"CSV 파일로 저장 완료: {csv_filename}")
else:
    print("테이블 내용이 없어서 Excel 파일로 저장하지 않았습니다.")

# 브라우저 닫기
browser.close()
