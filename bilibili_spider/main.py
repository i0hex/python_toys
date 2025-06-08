import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.webdriver import WebDriver
from selenium.webdriver.remote.webelement import WebElement
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

base_url = "https://www.bilibili.com/v/popular/rank/"

sub_urls = [
        {"all": "全部", "douga": "动画", "game": "游戏", "kichiku": "鬼畜", "music": "音乐", 
         "dance": "舞蹈", "cinephile": "影视", "ent": "娱乐", "knowledge": "知识", 
         "tech": "科技数码", "food": "美食", "car": "汽车", "fashion": "时尚美妆", 
         "sports": "体育运动", "animal": "动物"},
        {"anime": "番剧", "guochuang": "国创", "documentary": "纪录片", "movie": "电影", 
         "tv": "电视剧", "variety": "综艺"}
        ]

fetched_data = {}

headers = [
        ["排名", "标题", "链接", "UP", "UP链接", "播放量", "弹幕量"],
        ["排名", "标题", "链接", "更新状态", "播放量", "追番人数"]
        ]

def do_save(wb: Workbook, index: int, sub_type: int, sub_url: str, sub_data: list[list[str]]) -> None:
    ws = wb.create_sheet(sub_urls[sub_type][sub_url], index)
    ws.append(headers[sub_type])
    max_width_per_col: list[float] = [0] * len(headers[sub_type])
    headers_font = Font(bold=True, size=12)
    for cell in ws[1]:
        cell.font = headers_font

    min_width = 8
    max_width = 50
    font_factor = 1.1

    for data in sub_data:
        ws.append(data)
        for i, v in enumerate(data):
            v_len = max(min_width, min(max_width, len(v) * font_factor + 2))
            if v_len > max_width_per_col[i]:
                max_width_per_col[i] = v_len

    for column_index in range(1, ws.max_column + 1):
        column_letter = get_column_letter(column_index)
        ws.column_dimensions[column_letter].width = max_width_per_col[column_index - 1]


def save_data() -> None:

    print("Saving data...")

    wb: Workbook | None = None
    
    for index, (sub_url, sub_data) in enumerate(fetched_data.items()):
        if wb is None:
            wb = Workbook()

        if sub_url in sub_urls[0].keys():
            if len(sub_data) > 0:
                do_save(wb, index, 0, sub_url, sub_data)
                print(f"Save data from ['{sub_urls[0][sub_url]}']({base_url + sub_url}).")
        elif sub_url in sub_urls[1].keys():
            if len(sub_data) > 0:
                do_save(wb, index, 1, sub_url, sub_data)
                print(f"Save data from ['{sub_urls[1][sub_url]}']({base_url + sub_url}).")

    timpstamp: int = int(time.time())
    if wb is not None:
        file_name: str = "bilibili_rank_data_" + str(timpstamp) + ".xlsx"
        wb.save(file_name)
        print(f"The Bilibili ranking data has been saved. File: {file_name}")
    else:
        print("Err: No data has been saved.")


def fetch_data(driver: WebDriver, sub_url: str) -> None:
    # Implicit wait for 5 seconds
    driver.implicitly_wait(5)
    
    fetch_url = base_url + sub_url
    driver.get(fetch_url)

    fetched_data[sub_url] = []

    sub_type: int = -1
    if sub_url in sub_urls[0].keys():
        sub_type = 0
    elif sub_url in sub_urls[1].keys():
        sub_type = 1

    print(f"Fetch: ['{sub_urls[sub_type][sub_url]}']({fetch_url})")

    rank_elems: list[WebElement] = driver.find_elements(By.CLASS_NAME, "rank-item")
    for elem in rank_elems:
        idx = elem.get_attribute("data-rank")

        info_elem: WebElement = elem.find_element(By.CLASS_NAME, "info")
        title_elem: WebElement = info_elem.find_element(By.CLASS_NAME, "title")
        title: str | None = title_elem.get_attribute("title")
        link: str | None = title_elem.get_attribute("href")

        detail_elem: WebElement = info_elem.find_element(By.CLASS_NAME, "detail")
        detail_state_elem: WebElement = detail_elem.find_element(By.CLASS_NAME, "detail-state")
        detail_state_data_elems: list[WebElement] = detail_state_elem.find_elements(By.CLASS_NAME, "data-box")

        if sub_type == 0:
            up_info_elem: WebElement = detail_elem.find_element(By.TAG_NAME, "a")
            fetched_data[sub_url].append([idx, title, link, up_info_elem.text, up_info_elem.get_attribute("href"), 
                                          detail_state_data_elems[0].text, detail_state_data_elems[1].text])
        elif sub_type == 1:
            update_info_elem: WebElement = detail_elem.find_element(By.TAG_NAME, "span")
            fetched_data[sub_url].append([idx, title, link, update_info_elem.text, detail_state_data_elems[0].text, detail_state_data_elems[1].text])
    print(f"Fetch {len(rank_elems)} records from ['{sub_urls[sub_type][sub_url]}']({fetch_url}).")
    
def spider() -> None:
    options = Options()
    options.set_preference("general.useragent.override", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)\
            AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    # options.set_preference("intl.accept_languages", "en-US")
    options.add_argument("--headless")

    service = Service(executable_path='./geckodriver')
    driver = webdriver.Firefox(service=service, options=options)

    try:
        for v in sub_urls:
            for sub_url in v.keys():
                fetch_data(driver, sub_url)

        save_data()
        print("Done!")
    except Exception as e:
        print(f"Err: {str(e)}")
    finally:
        driver.quit()
        print("Browser was closed!")

if __name__ == "__main__":
    spider()

