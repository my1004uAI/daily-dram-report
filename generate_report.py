import os
import re
import urllib.parse
import datetime
from collections import defaultdict

import requests
import feedparser
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request as GoogleRequest
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# =========================
# 환경변수
# =========================
GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET", "")
GOOGLE_REFRESH_TOKEN = os.getenv("GOOGLE_REFRESH_TOKEN", "")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "")

print("GOOGLE_CLIENT_ID exists:", bool(GOOGLE_CLIENT_ID))
print("GOOGLE_CLIENT_SECRET exists:", bool(GOOGLE_CLIENT_SECRET))
print("GOOGLE_REFRESH_TOKEN exists:", bool(GOOGLE_REFRESH_TOKEN))
print("GOOGLE_DRIVE_FOLDER_ID exists:", bool(GOOGLE_DRIVE_FOLDER_ID))

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

KST = datetime.timezone(datetime.timedelta(hours=9))
now_kst = datetime.datetime.now(KST)
now_utc = datetime.datetime.utcnow()
cutoff_utc = now_utc - datetime.timedelta(days=1)

today_ymd = now_kst.strftime("%Y%m%d")
file_name = f"DRAM_논문_특허_ChatGPT_{today_ymd}.docx"

# =========================
# 유틸
# =========================
def parse_arxiv_date(date_str):
    return datetime.datetime.strptime(date_str, "%Y-%m-%dT%H:%M:%SZ")


def build_google_scholar_url(query: str) -> str:
    return "https://scholar.google.com/scholar?q=" + urllib.parse.quote(query)


def build_company_fallback(company: str) -> str:
    return build_google_scholar_url(f"{company} DRAM HBM DDR memory")


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
    if not a_words or not b_words:
        return 0.0
    overlap = len(a_words & b_words) / max(len(a_words), len(b_words))
    prefix_bonus = 0.15 if a_n[:80] == b_n[:80] else 0.0
    return min(1.0, overlap + prefix_bonus)


def detect_publisher_name(url="", container_title="", publisher_name=""):
    hay = " ".join([url or "", container_title or "", publisher_name or ""]).lower()
    for pub, hints in MAJOR_PUBLISHER_HINTS.items():
        if any(h in hay for h in hints):
            return pub
    return ""


def summarize_paper(title: str, summary: str) -> str:
    text = (summary or "").strip()
    if not text:
        return "초록 정보 없음"
    short = text[:260].strip()
    short = re.sub(r"\s+", " ", short)
    return short


def get_best_link(paper):
    if paper.get("doi_url"):
        return "DOI", paper["doi_url"]
    if paper.get("official_url"):
        return "공식 페이지", paper["official_url"]
    if paper.get("abs_url"):
        return "arXiv", paper["abs_url"]
    return "Google Scholar", paper["scholar_url"]


def require_env(name, value):
    if not value:
        raise RuntimeError(f"{name} is missing or empty")


# =========================
# DOCX 하이퍼링크
# =========================
def add_hyperlink(paragraph, text, url, color="0563C1", underline=True):
    """
    python-docx에서 클릭 가능한 외부 하이퍼링크 추가
    """
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")

    if color:
        c = OxmlElement("w:color")
        c.set(qn("w:val"), color)
        rPr.append(c)

    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single" if underline else "none")
    rPr.append(u)

    new_run.append(rPr)

    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)

    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_link_line(doc, label, url):
    p = doc.add_paragraph()
    p.add_run(label)
    add_hyperlink(p, url, url)


def add_paper_block(doc, idx, paper):
    doc.add_paragraph(f"{idx}) {paper['title']}")
    doc.add_paragraph(f"   - 저자: {paper['authors'] or '정보 없음'}")
    doc.add_paragraph(f"   - 공개시각(UTC): {paper['published']}")

    p = doc.add_paragraph()
    p.add_run(f"   - 대표 링크 ({paper['best_link_type']}): ")
    add_hyperlink(p, paper["best_link_url"], paper["best_link_url"])

    if paper["official_url"]:
        p = doc.add_paragraph()
        p.add_run("   - 공식 페이지: ")
        add_hyperlink(p, paper["official_url"], paper["official_url"])
    else:
        doc.add_paragraph("   - 공식 페이지: 링크 미확인")

    if paper["doi"]:
        doc.add_paragraph(f"   - DOI: {paper['doi']}")
    else:
        doc.add_paragraph("   - DOI: 링크 미확인")

    p = doc.add_paragraph()
    p.add_run("   - Google Scholar: ")
    add_hyperlink(p, paper["scholar_url"], paper["scholar_url"])

    doc.add_paragraph(f"   - 요약: {paper['short_summary']}")


def add_company_block(doc, company, item):
    doc.add_paragraph(company)
    doc.add_paragraph(f"   - 상태: {item['status']}")
    doc.add_paragraph(f"   - 제목/설명: {item['title']}")

    p = doc.add_paragraph()
    p.add_run(f"   - 대표 링크 ({item['link_type']}): ")
    add_hyperlink(p, item["url"], item["url"])

    if item["official_url"]:
        p = doc.add_paragraph()
        p.add_run("   - 공식 페이지: ")
        add_hyperlink(p, item["official_url"], item["official_url"])
    else:
        doc.add_paragraph("   - 공식 페이지: 링크 미확인")

    doc.add_paragraph(f"   - 비고: {item['note']}")


def add_venue_block(doc, paper, venue_hint):
    doc.add_paragraph(f"- {paper['title']}")
    doc.add_paragraph(f"  · 출판사/저널 힌트: {venue_hint}")
    p = doc.add_paragraph()
    p.add_run("  · 링크: ")
    add_hyperlink(p, paper["best_link_url"], paper["best_link_url"])


# =========================
# Crossref
# =========================
def crossref_lookup(title, authors=""):
    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": CROSSREF_ROWS,
        "select": "DOI,title,author,URL,container-title,publisher"
    }
    headers = {
        "User-Agent": "daily-dram-report/1.0 (mailto:example@example.com)"
    }

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


# =========================
# arXiv 검색
# =========================
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
            "official_url": (cr["official_url"] or cr["doi_url"]) if cr else "",
            "publisher_name": cr["publisher_name"] if cr else "",
            "publisher_detected": cr["publisher_detected"] if cr else "",
            "container_title": cr["container_title"] if cr else "",
            "crossref_score": cr["score"] if cr else None,
        }

        record["best_link_type"], record["best_link_url"] = get_best_link(record)
        record["short_summary"] = summarize_paper(title, summary)

        all_papers.append(record)

        if published_dt >= cutoff_utc:
            recent_1d_papers.append(record)

    return all_papers, recent_1d_papers


# =========================
# 회사별 항목
# =========================
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
                "link_type": best["best_link_type"],
                "url": best["best_link_url"],
                "official_url": best["official_url"],
                "note": "자동 검색 결과에서 회사 관련 항목 탐지",
            }
        else:
            company_items[company] = {
                "status": "fallback",
                "title": f"{company} 관련 DRAM/HBM 연구 검색",
                "link_type": "Google Scholar",
                "url": build_company_fallback(company),
                "official_url": "",
                "note": "최근 1일/자동 검색에서 직접 논문 미확인, 대체 검색 링크 제공",
            }

    return company_items


# =========================
# DOCX 생성
# =========================
def create_docx(all_papers, recent_1d_papers, company_items, path):
    fallback_papers = all_papers[:5]

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

    # 요약
    doc.add_heading("0. 오늘의 요약", level=1)
    doc.add_paragraph(f"- 최근 1일 신규 논문 수: {len(recent_1d_papers)}건")
    doc.add_paragraph(f"- 자동 검색 전체 논문 수: {len(all_papers)}건")
    doc.add_paragraph("- 필수 기업 4개는 별도 섹션에서 확인 결과 또는 대체 검색 링크를 제공")

    # 최근 1일 엄격 기준
    doc.add_heading("1. 최근 1일 논문 (엄격 기준)", level=1)
    if recent_1d_papers:
        for idx, paper in enumerate(recent_1d_papers[:5], 1):
            add_paper_block(doc, idx, paper)
    else:
        doc.add_paragraph("신규 공개 논문 없음")

    # 참고 논문
    doc.add_heading("2. 참고: 최근 공개 논문 (1일 초과)", level=1)
    if fallback_papers:
        for idx, paper in enumerate(fallback_papers, 1):
            add_paper_block(doc, idx, paper)
    else:
        doc.add_paragraph("참고 논문 없음")

    # 기업별
    doc.add_heading("3. 필수 기업 포함 항목", level=1)
    for company, item in company_items.items():
        add_company_block(doc, company, item)

    # 학회/저널 힌트
    doc.add_heading("4. 주요 학회/저널 힌트", level=1)
    shown = 0
    for paper in all_papers:
        venue_hint = paper["publisher_detected"] or paper["publisher_name"] or paper["container_title"]
        if venue_hint:
            add_venue_block(doc, paper, venue_hint)
            shown += 1
        if shown >= 5:
            break
    if shown == 0:
        doc.add_paragraph("자동 탐지 결과 없음")

    # 메모
    doc.add_heading("5. 메모", level=1)
    doc.add_paragraph("- 최근 1일 결과와 참고 논문을 분리해 혼동을 줄였다.")
    doc.add_paragraph("- 각 항목에는 DOI, 공식 페이지, arXiv, Google Scholar 중 최소 1개 링크를 보장하도록 구성했다.")
    doc.add_paragraph("- 회사 4개는 직접 탐지 실패 시에도 대체 검색 링크를 제공한다.")
    doc.add_paragraph("- 문서 내 URL은 클릭 가능한 하이퍼링크로 삽입했다.")

    doc.save(path)


# =========================
# Drive 업로드
# =========================
def get_drive_service():
    require_env("GOOGLE_CLIENT_ID", GOOGLE_CLIENT_ID)
    require_env("GOOGLE_CLIENT_SECRET", GOOGLE_CLIENT_SECRET)
    require_env("GOOGLE_REFRESH_TOKEN", GOOGLE_REFRESH_TOKEN)
    require_env("GOOGLE_DRIVE_FOLDER_ID", GOOGLE_DRIVE_FOLDER_ID)

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
    file_metadata = {
        "name": os.path.basename(path),
        "parents": [GOOGLE_DRIVE_FOLDER_ID]
    }

    created = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,name,webViewLink"
    ).execute()

    print("Uploaded:", created["name"])
    print("Link:", created.get("webViewLink"))


# =========================
# main
# =========================
def main():
    all_papers, recent_1d_papers = search_arxiv()
    company_items = build_company_items(all_papers)
    create_docx(all_papers, recent_1d_papers, company_items, file_name)
    upload_to_drive(file_name)


if __name__ == "__main__":
    main()
