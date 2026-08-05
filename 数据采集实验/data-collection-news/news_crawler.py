import re
import sqlite3
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


target_url = "https://news.bjfu.edu.cn/lsyw/index.html"
BASE_URL = "https://news.bjfu.edu.cn/"


def get_html(url):
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/145.0.0.0 Safari/537.36 Edg/145.0.0.0"
        )
    }
    response = requests.get(url, headers=headers, timeout=10)
    response.raise_for_status()
    response.encoding = response.apparent_encoding
    return response.text


def split_date_and_title(text):
    text = " ".join(text.split())
    match = re.match(r"^(\d{4}-\d{2}-\d{2})\s*(.+)$", text)
    if match:
        return match.group(2).strip(), match.group(1)
    return text.strip(), ""


def get_detail_info(url):
    try:
        detail_html = get_html(url)
    except requests.RequestException:
        return "", ""

    soup = BeautifulSoup(detail_html, "html.parser")
    title = ""
    for selector in ["h1", "h2"]:
        title_tag = soup.find(selector)
        if title_tag:
            title = title_tag.get_text(" ", strip=True)
            break
    if not title and soup.title:
        title = soup.title.get_text(" ", strip=True)

    page_text = soup.get_text(" ", strip=True)
    date_match = re.search(r"(发表时间|发布时间)[:：]?\s*(\d{4}[/-]\d{2}[/-]\d{2})", page_text)
    publish_time = date_match.group(2).replace("/", "-") if date_match else ""
    return title, publish_time


def parse_html(html, page_url=target_url):
    soup = BeautifulSoup(html, "html.parser")
    news_ul_items = soup.find_all("ul", attrs={"class": "news_ul"})
    news_data = []
    seen = set()

    for ul_item in news_ul_items:
        for a_tag in ul_item.find_all("a"):
            href = a_tag.get("href")
            raw_text = a_tag.get_text(" ", strip=True)
            if not href or not raw_text:
                continue

            title, publish_time = split_date_and_title(raw_text)
            full_news_url = urljoin(page_url or BASE_URL, href)
            li_tag = a_tag.find_parent("li")
            if li_tag:
                span_tag = li_tag.find("span")
                if span_tag:
                    span_text = span_tag.get_text(" ", strip=True)
                    date_match = re.search(r"\d{4}-\d{2}-\d{2}", span_text)
                    if date_match:
                        publish_time = date_match.group(0)

            if "..." in title or not publish_time:
                detail_title, detail_time = get_detail_info(full_news_url)
                title = detail_title or title
                publish_time = detail_time or publish_time

            key = (title, full_news_url)
            if key in seen:
                continue
            seen.add(key)
            news_data.append(
                {
                    "title": title,
                    "url": full_news_url,
                    "publish_time": publish_time,
                }
            )

    return news_data


def save_to_sqlite(data):
    conn = sqlite3.connect("news_data.db")
    cursor = conn.cursor()
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS news (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            link TEXT NOT NULL,
            publish_time TEXT
        )
        """
    )
    cursor.execute("DELETE FROM news")
    for news in data:
        cursor.execute(
            "INSERT INTO news (title, link, publish_time) VALUES (?, ?, ?)",
            (news["title"], news["url"], news["publish_time"]),
        )
    conn.commit()
    conn.close()


if __name__ == "__main__":
    html_content = get_html(target_url)
    news_result = parse_html(html_content, target_url)
    save_to_sqlite(news_result)
    print("数据抓取并存储完成")
    print(f"目标页面：{target_url}")
    print(f"共抓取新闻数量：{len(news_result)}条")
    for item in news_result[:5]:
        print(f"{item['publish_time'] or '无日期'} | {item['title']} | {item['url']}")
