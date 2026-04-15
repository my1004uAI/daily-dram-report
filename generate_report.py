import os
import re
import json
import urllib.parse
import datetime
from collections import defaultdict

import requests
import feedparser
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request as GoogleRequest
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# =========================
# 환경변수
# =========================
GOOGLE_CLIENT_ID = os.environ["GOOGLE_CLIENT_ID"]
GOOGLE_CLIENT_SECRET = os.environ["GOOGLE_CLIENT_SECRET"]
GOOGLE_REFRESH_TOKEN = os.environ["GOOGLE_REFRESH_TOKEN"]
GOOGLE_DRIVE_FOLDER_ID = os.environ["GOOGLE_DRIVE_FOLDER_ID"]

import os

print("GOOGLE_CLIENT_ID exists:", bool(os.getenv("GOOGLE_CLIENT_ID")))
print("GOOGLE_CLIENT_SECRET exists:", bool(os.getenv("GOOGLE_CLIENT_SECRET")))
print("GOOGLE_REFRESH_TOKEN exists:", bool(os.getenv("GOOGLE_REFRESH_TOKEN")))
print("GOOGLE_DRIVE_FOLDER_ID exists:", bool(os.getenv("GOOGLE_DRIVE_FOLDER_ID")))
print("GOOGLE_REFRESH_TOKEN length:", len(os.getenv("GOOGLE_REFRESH_TOKEN", "")))

# =========================
# 기본 설정
# =========================
SEARCH_TERMS = ["DRAM", "HBM", "DDR5", "DDR6", "LPDDR5", "LPDDR6"]
COMPANY_KEYWORDS = {
    "Samsung": ["samsung", "samsung electronics"],
    "SK hynix": ["sk hynix", "hynix"],
    "Micron": ["micron", "micron technology"],
    "CXMT": ["cxmt", "changxin", "changxin memory", "changxin memory technologies"],
}
MAJOR_PUBLISHER_HINTS = {
    "IEEE": ["ieee", "ieeexplore.ieee.org"],
    "ACM": ["acm", "dl.acm.org"],
    "Springer": ["springer", "springer.com", "link.springer.com"],
    "Elsevier": ["elsevier", "sciencedirect.com"],
    "Wiley": ["wiley", "onlinelibrary.wiley.com"],
    "Nature": ["nature", "nature.com"],
    "MDPI": ["mdpi", "mdpi.com"],
    "IOP": ["iop", "iopscience.iop.org"],
}
ARXIV_MAX_RESULTS = 30
CROSSREF_ROWS = 8

utc_now = datetime.datetime.utcnow()
cutoff = utc_now - datetime.timedelta(days=1)
today_ymd = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d")
file_name = f"DRAM_논문_특허_ChatGPT_{today_ymd}.docx"


def parse_arxiv_date(date_str):
    return datetime.datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ")


def build_google_scholar_url(title: str) -> str:
    return "https://scholar.google.com/scholar?q=" + urllib.parse.quote(title)


def normalize_text(*parts):
    return " ".join([p or "" for p in parts]).lower()


def norm_text(s):
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def score_title_match(a, b):
    a_n = norm_text(a)
    b_n = norm_text(b)
    if not a_n or not b_n:
        return 0.0
    if a_n == b_n:
        return 1.0
    a_words = set(a_n.split())
    b_words = set(b_n.split())
    overlap = len(a_words & b_words) / max(len(a_words), len(b_words)) if a_words and b_words else 0.0
    prefix_bonus = 0.15 if a_n[:80] == b_n[:80] else 0.0
    return min(1.0, overlap + prefix_bonus)


def detect_publisher_name(url="", container_title="", member_name=""):
    hay = " ".join([url or "", container_title or "", member_name or ""]).lower()
    for pub, hints in MAJOR_PUBLISHER_HINTS.items():
        if any(h in hay for h in hints):
            return pub
    return ""


def crossref_lookup(title, authors=""):
    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": CROSSREF_ROWS,
        "select": "DOI,title,author,URL,container-title,publisher"
    }
    headers = {"User-Agent": "daily-dram-report/1.0 (mailto:example@example.com)"}

    try:
        r = requests.get(url, params=params, headers=headers, timeout=30)
        r.raise_for_status()
        items = r.json().get("message", {}).get("items", [])
    except Exception:
        return None

    best = None
    best_score = 0.0
    first_author = authors.split(",")[0].strip().lower() if authors else ""

    for item in items:
        cr_title = (item.get("title") or [""])[0]
        score = score_title_match(title, cr_title)

        if first_author and item.get("author"):
            cr_authors = " ".join(
                f"{a.get('given','')} {a.get('family','')}".strip()
                for a in item.get("author", [])
            ).lower()
            last_token = first_author.split(" ")[-1] if first_author else ""
            if last_token and last_token in cr_authors:
                score += 0.1

        if score > best_score:
            best_score = score
            best = item

    if not best or best_score < 0.55:
        return None

    doi = best.get("DOI", "")
    official_url = best.get("URL", "")
    container_title = ((best.get("container-title") or [""])[0]).strip()
    publisher_name = (best.get("publisher") or "").strip()
    detected_pub = detect_publisher_name(official_url, container_title, publisher_name)

    return {
        "doi": doi,
        "doi_url": f"https://doi.org/{doi}" if doi else "",
        "official_url": official_url or "",
        "container_title": container_title,
        "publisher_name": publisher_name,
        "publisher_detected": detected_pub,
        "score": round(best_score, 3),
    }


def search_arxiv():
    query = " OR ".join([f'all:\"{term}\"' for term in SEARCH_TERMS])
    api_url = (
        "http://export.arxiv.org/api/query?"
        + urllib.parse.urlencode({
            "search_query": query,
            "start": 0,
            "max_results": ARXIV_MAX_RESULTS,
            "sortBy": "submittedDate",
            "sortOrder": "descending"
        })
    )
    feed = feedparser.parse(requests.get(api_url, timeout=30).text)

    all_papers = []
    recent_1d_papers = []

    for entry in feed.entries:
        title = entry.title.strip().replace("\n", " ")
        summary = (entry.summary or "").strip().replace("\n", " ")
        published = entry.published
        published_dt = parse_arxiv_date(published)
        authors = ", ".join(a.name for a in entry.authors) if hasattr(entry, "authors") else ""
        arxiv_id = entry.id.split("/abs/")[-1] if "/abs/" in entry.id else entry.id

        pdf_link = ""
        for link in entry.links:
            if getattr(link, "type", "") == "application/pdf":
                pdf_link = link.href
                break

        cr = crossref_lookup(title, authors)

        primary_link_type = "DOI" if cr and cr.get("doi") else "arXiv ID"
        primary_link_value = cr["doi"] if cr and cr.get("doi") else arxiv_id
        primary_url = cr["doi_url"] if cr and cr.get("doi") else entry.id
        official_url = (cr["official_url"] or cr["doi_url"]) if cr else ""

        record = {
            "title": title,
            "summary": summary,
            "authors": authors,
            "published": published,
            "published_dt": published_dt,
            "arxiv_id": arxiv_id,
            "abs_url": entry.id,
            "pdf_url": pdf_link,
            "scholar_url": build_google_scholar_url(title),
            "doi": cr["doi"] if cr else "",
            "doi_url": cr["doi_url"] if cr else "",
            "official_url": official_url,
            "publisher_name": cr["publisher_name"] if cr else "",
            "publisher_detected": cr["publisher_detected"] if cr else "",
            "container_title": cr["container_title"] if cr else "",
            "crossref_score": cr["score"] if cr else None,
            "primary_link_type": primary_link_type,
            "primary_link_value": primary_link_value,
            "primary_url": primary_url,
        }
        all_papers.append(record)

        if published_dt >= cutoff:
            recent_1d_papers.append(record)

    return all_papers, recent_1d_papers


def build_company_items(all_papers):
    company_hits = defaultdict(list)
    for paper in all_papers:
        hay = normalize_text(paper["title"], paper["summary"], paper["authors"])
        for company, kws in COMPANY_KEYWORDS.items():
            if any(kw in hay for kw in kws):
                company_hits[company].append(paper)

    company_items = {}
    for company in COMPANY_KEYWORDS.keys():
        if company_hits[company]:
            best = sorted(company_hits[company], key=lambda x: x["published_dt"], reverse=True)[0]
            company_items[company] = {
                "status": "확인",
                "title": best["title"],
                "link_type": best["primary_link_type"],
                "link_value": best["primary_link_value"],
                "url": best["primary_url"],
                "official_url": best["official_url"],
                "publisher": best["publisher_detected"] or best["publisher_name"],
            }
        else:
            company_items[company] = {"status": "링크 미확인"}
    return company_items


def create_docx(all_papers, recent_1d_papers, company_items, path):
    selected_papers = recent_1d_papers[:5] if recent_1d_papers else all_papers[:5]
    recent_summary = (
        f"신규 공개 논문 {len(recent_1d_papers)}건 확인"
        if recent_1d_papers else
        "신규 공개 논문: 확인되지 않음 (arXiv 최근 1일 기준)"
    )

    doc = Document()
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Malgun Gothic"
    font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")

    doc.add_heading(f"Daily DRAM Report ({today_ymd})", 0)
    doc.add_paragraph("실행 시각: 07:07 (Asia/Seoul 목표)")
    doc.add_paragraph("검색 범위: 최근 1일 DRAM / HBM / DDR5 / DDR6 / LPDDR5 / LPDDR6 관련 논문")
    doc.add_paragraph("검색 소스: arXiv 자동 검색 + Crossref DOI/공식 페이지 보강")

    doc.add_heading("1. 최근 1일 논문", level=1)
    doc.add_paragraph(recent_summary)

    doc.add_heading("2. 자동 검색된 최신 논문", level=1)
    if selected_papers:
        for idx, paper in enumerate(selected_papers, 1):
            lines = [
                f"{idx}) {paper['title']}",
                f"   - 저자: {paper['authors'] or '정보 없음'}",
                f"   - 공개시각(UTC): {paper['published']}",
                f"   - 기본 링크: {paper['primary_link_type']} / {paper['primary_link_value']}",
                f"   - 기본 링크 URL: {paper['primary_url']}",
                f"   - 공식 랜딩페이지: {paper['official_url'] or '링크 미확인'}",
                f"   - DOI: {paper['doi'] or '링크 미확인'}",
                f"   - arXiv 링크: {paper['abs_url']}",
                f"   - PDF 링크: {paper['pdf_url'] or '링크 미확인'}",
                f"   - Google Scholar: {paper['scholar_url']}",
            ]
            if paper["publisher_detected"] or paper["publisher_name"] or paper["container_title"]:
                lines.append(
                    f"   - 출판사/저널 힌트: "
                    f"{paper['publisher_detected'] or paper['publisher_name'] or paper['container_title']}"
                )
            doc.add_paragraph("\n".join(lines))
    else:
        doc.add_paragraph("자동 검색 결과 없음")

    doc.add_heading("3. 필수 기업 포함 항목", level=1)
    for company, item in company_items.items():
        if item["status"] == "확인":
            lines = [
                company,
                f"   - 제목: {item['title']}",
                f"   - 기본 링크: {item['link_type']} / {item['link_value']}",
                f"   - 링크: {item['url']}",
                f"   - 공식 랜딩페이지: {item['official_url'] or '링크 미확인'}",
            ]
            if item.get("publisher"):
                lines.append(f"   - 출판사 힌트: {item['publisher']}")
            doc.add_paragraph("\n".join(lines))
        else:
            doc.add_paragraph(f"{company}: 링크 미확인")

    doc.add_heading("4. 메모", level=1)
    doc.add_paragraph("- DOI가 확인되면 DOI 링크를 기본 링크로 사용했다.")
    doc.add_paragraph("- Crossref의 공식 URL이 있으면 출판사 랜딩페이지로 함께 기록했다.")
    doc.add_paragraph("- 공식 페이지를 찾지 못하면 arXiv 또는 Google Scholar 링크를 유지했다.")

    doc.save(path)


def get_drive_service():
    creds = Credentials(
        None,
        refresh_token=GOOGLE_REFRESH_TOKEN,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=GOOGLE_CLIENT_ID,
        client_secret=GOOGLE_CLIENT_SECRET,
        scopes=["https://www.googleapis.com/auth/drive.file"],
    )
    creds.refresh(GoogleRequest())
    return build("drive", "v3", credentials=creds)


def upload_to_drive(path):
    service = get_drive_service()
    media = MediaFileUpload(
        path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    file_metadata = {"name": os.path.basename(path), "parents": [GOOGLE_DRIVE_FOLDER_ID]}
    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,name,webViewLink"
    ).execute()
    print("Uploaded:", created["name"])
    print("Link:", created.get("webViewLink"))


def main():
    all_papers, recent_1d_papers = search_arxiv()
    company_items = build_company_items(all_papers)
    create_docx(all_papers, recent_1d_papers, company_items, file_name)
    upload_to_drive(file_name)


if __name__ == "__main__":
    main()
