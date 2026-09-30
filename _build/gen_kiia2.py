# gen_kiia2.py — 부록 E 실습 데이터 패키지: KIIA-2 폴더관리 세트(혼합 14종)
# 버전 중복·파일명/내용 불일치를 의도적으로 심어 파일 분류·중복 탐지 실습에 쓴다. 전량 가상.
import os, csv, datetime as dt
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = r"C:\Users\user\Documents\developer\7_audit\practice-package\KIIA-2_폴더관리"
NOTICE = "본 자료는 실습용 가상 데이터이며 실재 인물·기관·거래와 무관하다."
os.makedirs(ROOT, exist_ok=True)

# 반복 수의계약 단서용 공통 업체·계약
VENDOR = "주식회사 그린테크솔루션"

def new_doc():
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = "맑은 고딕"; st.font.size = Pt(10)
    try:
        st.element.rPr.rFonts.set(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia", "맑은 고딕")
    except Exception:
        pass
    return doc

def notice_line(doc):
    p = doc.add_paragraph()
    r = p.add_run(NOTICE); r.italic = True; r.font.size = Pt(8)

def add_heading(doc, text, size=14):
    p = doc.add_paragraph(); r = p.add_run(text); r.bold = True; r.font.size = Pt(size)

def styled_header(ws, headers, row=1):
    fill = PatternFill("solid", fgColor="1A2E5A"); font = Font(color="FFFFFF", bold=True, name="맑은 고딕")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h); cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")

def autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(w*1.6+2, 10), 44)

def p(name): return os.path.join(ROOT, name)

# ── 계약서 3종 (버전 중복 2 + 명칭/내용 불일치 1) ──────────────────
def contract_body(doc, contract_no, amount, sign_date, note_prior=True):
    add_heading(doc, "용역 계약서", 15)
    doc.add_paragraph(f"계약번호: {contract_no}")
    doc.add_paragraph(f"발주기관: ABC 제조(주) 감사실")
    doc.add_paragraph(f"수급업체: {VENDOR}")
    doc.add_paragraph(f"계약방식: 수의계약")
    doc.add_paragraph(f"계약금액: 금 {amount:,}원 (부가세 포함)")
    doc.add_paragraph(f"계약체결일: {sign_date}")
    doc.add_paragraph("계약기간: 체결일부터 3개월")
    doc.add_paragraph("")
    add_heading(doc, "제1조(목적)", 11)
    doc.add_paragraph("본 계약은 정보시스템 유지보수 및 운영 지원 용역의 수행에 관한 사항을 정함을 목적으로 한다.")
    add_heading(doc, "제2조(계약금액 및 지급)", 11)
    doc.add_paragraph("발주기관은 용역 완료 후 수급업체의 청구에 따라 계약금액을 지급한다. 선급금은 지급하지 아니한다.")
    add_heading(doc, "제3조(계약방식)", 11)
    txt = "본 계약은 수의계약으로 체결한다."
    if note_prior:
        txt += f" 동일 수급업체({VENDOR})와의 직전 유사 용역 계약(IT-OUT-2026-01, IT-OUT-2026-02)이 있으며, 합산 검토 여부는 별도로 확인한다."
    doc.add_paragraph(txt)
    add_heading(doc, "제4조(하자보수)", 11)
    doc.add_paragraph("수급업체는 용역 완료일부터 3개월간 하자를 무상으로 보수한다.")
    doc.add_paragraph("")
    doc.add_paragraph("발주기관: ABC 제조(주)   (인)")
    doc.add_paragraph(f"수급업체: {VENDOR}   (인)")

def contract_final_v3():
    doc = new_doc(); notice_line(doc)
    doc.add_paragraph("[문서 버전: v3 / 상태: 검토중]")
    contract_body(doc, "IT-OUT-2026-03", 48_000_000, "2026-03-10")
    doc.save(p("contract_final_v3.docx"))

def contract_jinjja():
    # v3와 거의 동일한 중복본(금액만 소액 변경) — 버전 중복 단서
    doc = new_doc(); notice_line(doc)
    doc.add_paragraph("[문서 버전: 진짜최종 / 상태: 서명대기]")
    contract_body(doc, "IT-OUT-2026-03", 48_000_000, "2026-03-12")
    doc.save(p("계약서_최종(진짜최종).docx"))

def contract_choejong_mismatch():
    # 파일명은 계약서이나 내용은 정산 내역서 — 명칭/내용 불일치
    doc = new_doc(); notice_line(doc)
    doc.add_paragraph("[주의: 파일명은 '계약서'이나 실제 내용은 정산 내역서다 — 분류 시 내용 기준으로 판정한다]")
    add_heading(doc, "용역대금 정산 내역서", 15)
    doc.add_paragraph(f"수급업체: {VENDOR}")
    doc.add_paragraph("정산 대상: 2026년 1분기 정보시스템 유지보수 용역")
    doc.add_paragraph("")
    tbl = doc.add_table(rows=1, cols=4); tbl.style = "Table Grid"
    hd = tbl.rows[0].cells
    for i, t in enumerate(["항목", "계약금액", "기성 인정", "지급액"]): hd[i].text = t
    for row in [["1월분", "16,000,000", "16,000,000", "16,000,000"],
                ["2월분", "16,000,000", "15,200,000", "15,200,000"],
                ["3월분", "16,000,000", "14,800,000", "14,800,000"],
                ["합계", "48,000,000", "46,000,000", "46,000,000"]]:
        c = tbl.add_row().cells
        for i, v in enumerate(row): c[i].text = v
    doc.add_paragraph("")
    doc.add_paragraph("정산 담당: 재무팀   확인: 감사실")
    doc.save(p("계약서_최종본_최종.docx"))

# ── 회의록·메모 ────────────────────────────────────────────────
def minutes():
    doc = new_doc(); notice_line(doc)
    add_heading(doc, "감사 착수 회의록", 14)
    doc.add_paragraph("일시: 2026-03-05 10:00 ~ 11:00")
    doc.add_paragraph("장소: 본관 3층 회의실")
    doc.add_paragraph("참석: 감사실장, 감사담당 2명, IT운영팀장")
    doc.add_paragraph("")
    add_heading(doc, "논의 사항", 11)
    for t in ["2026년 1분기 IT 외주용역 계약 자료 일체를 감사실로 이관한다.",
              "동일 업체 반복 수의계약 여부를 우선 점검 항목으로 둔다.",
              "제출 자료의 버전 중복과 파일명 불일치가 많아 분류 대장을 먼저 작성한다.",
              "스캔본은 내용 인식이 어려워 원본 확인이 필요하다."]:
        doc.add_paragraph(t, style="List Bullet")
    doc.save(p("회의록_20260305.docx"))

def memo():
    doc = new_doc(); notice_line(doc)
    add_heading(doc, "메모", 13)
    for t in ["계약서 파일이 세 개인데 어느 것이 최종인지 불명확함. 금액은 같고 날짜만 다름.",
              "'계약서_최종본_최종'은 열어 보니 정산 내역이었음. 파일명과 내용이 다름.",
              "vendor_master 신구본 비교 필요. 그린테크솔루션이 신규본에만 사업자번호가 갱신됨.",
              "감사로그(audit_log.csv)에서 계약 폴더 접근 기록 확인할 것."]:
        doc.add_paragraph("- " + t)
    doc.save(p("메모.docx"))

# ── XLSX 4종 ───────────────────────────────────────────────────
def invoice():
    wb = Workbook(); ws = wb.active; ws.title = "인보이스"
    styled_header(ws, ["청구번호", "청구일자", "공급자", "품목", "수량", "단가", "공급가액", "세액", "합계"])
    rows = [["INV-2026-0301", "2026-03-31", VENDOR, "유지보수 3월분", 1, 16000000, 16000000, 1600000, 17600000],
            ["INV-2026-0302", "2026-03-31", VENDOR, "긴급 장애대응", 1, 3000000, 3000000, 300000, 3300000],
            ["INV-2026-0303", "2026-03-31", "미래ICT", "라이선스 갱신", 2, 900000, 1800000, 180000, 1980000]]
    for r in rows: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(p("invoice_2026_03.xlsx"))

def budget():
    wb = Workbook(); ws = wb.active; ws.title = "예산계획"
    styled_header(ws, ["예산과목", "부서", "연간예산", "1분기 집행", "집행률(%)", "비고"])
    rows = [["정보화운영비", "IT운영팀", 200000000, 51000000, 25.5, ""],
            ["외주용역비", "IT운영팀", 120000000, 48000000, 40.0, "수의계약 비중 높음"],
            ["소모품비", "총무팀", 30000000, 6800000, 22.7, ""],
            ["교육훈련비", "인사팀", 25000000, 4100000, 16.4, ""]]
    for r in rows: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(p("budget_plan.xlsx"))

def settlement():
    wb = Workbook(); ws = wb.active; ws.title = "정산_1분기"
    styled_header(ws, ["정산번호", "업체", "계약금액", "기성인정", "지급액", "차액", "비고"])
    rows = [["ST-2026-Q1-01", VENDOR, 48000000, 46000000, 46000000, 2000000, "기성 미달분 차감"],
            ["ST-2026-Q1-02", "미래ICT", 1980000, 1980000, 1980000, 0, ""],
            ["ST-2026-Q1-03", VENDOR, 3300000, 3300000, 3300000, 0, "긴급 장애대응"]]
    for r in rows: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(p("settlement_Q1.xlsx"))

def checklist():
    wb = Workbook(); ws = wb.active; ws.title = "점검표"
    styled_header(ws, ["점검항목", "점검기준", "결과", "확인자", "비고"])
    rows = [["계약방식 적정성", "수의계약 사유 명시 여부", "확인 필요", "감사담당", "동일업체 반복"],
            ["합산 한도 검토", "유사·반복 계약 합산 여부", "미검토", "감사담당", ""],
            ["증빙 구비", "인보이스·정산서 일치", "일부 불일치", "감사담당", "파일명 불일치 존재"],
            ["문서 버전관리", "최종본 식별 가능 여부", "미흡", "감사담당", "중복본 3종"]]
    for r in rows: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(p("점검표.xlsx"))

# ── 벤더 마스터 신구본 ─────────────────────────────────────────
def vendor_master(newver):
    title = "vendor_master_new.xlsx" if newver else "vendor_master_old.xlsx"
    wb = Workbook(); ws = wb.active; ws.title = "벤더마스터"
    styled_header(ws, ["벤더코드", "업체명", "사업자번호", "대표자", "등록일", "상태"])
    old = [["V-001", VENDOR, "215-81-00000", "김대표", "2024-05-01", "정상"],
           ["V-002", "미래ICT", "120-81-11111", "이대표", "2023-02-10", "정상"],
           ["V-003", "한빛물류", "134-86-22222", "박대표", "2022-09-15", "정상"]]
    new = [["V-001", VENDOR, "215-81-99999", "김대표", "2026-01-05", "정상"],   # 사업자번호 갱신
           ["V-002", "미래ICT", "120-81-11111", "이대표", "2023-02-10", "정상"],
           ["V-003", "한빛물류", "134-86-22222", "박대표", "2022-09-15", "거래중지"],  # 상태 변경
           ["V-004", "정직건설", "220-88-33333", "최대표", "2026-02-20", "정상"]]      # 신규 추가
    for r in (new if newver else old): ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(p(title))

# ── 감사 로그 CSV ──────────────────────────────────────────────
def audit_log():
    rows = []
    base = dt.datetime(2026, 3, 5, 9, 0)
    events = [("hong.a", "OPEN", "contract_final_v3.docx"),
              ("hong.a", "OPEN", "계약서_최종(진짜최종).docx"),
              ("kim.b", "DOWNLOAD", "settlement_Q1.xlsx"),
              ("hong.a", "EDIT", "점검표.xlsx"),
              ("lee.c", "OPEN", "계약서_최종본_최종.docx"),
              ("kim.b", "COPY", "vendor_master_new.xlsx"),
              ("hong.a", "OPEN", "invoice_2026_03.xlsx"),
              ("admin", "DELETE", "vendor_master_temp.xlsx"),
              ("lee.c", "PRINT", "회의록_20260305.docx"),
              ("hong.a", "OPEN", "스캔문서.pdf")]
    for i, (u, act, f) in enumerate(events):
        ts = (base + dt.timedelta(minutes=i*37)).strftime("%Y-%m-%d %H:%M:%S")
        rows.append([ts, u, act, f, f"192.168.0.{20+i}"])
    with open(p("audit_log.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["timestamp", "user_id", "action", "target_file", "src_ip"])
        w.writerows(rows)
        w.writerow([])
        w.writerow([NOTICE])

# ── 스캔문서 안내(txt, PDF 자리) ───────────────────────────────
def scan_note():
    text = (
        "[스캔 이미지 PDF 자리 — 실습용 안내 파일]\n\n"
        "이 위치에는 원래 스캔 이미지 PDF(스캔문서.pdf)가 놓인다.\n"
        "문자 인식(OCR)이 되지 않는 이미지 전용 파일을 가정한 자리 표시 파일이다.\n"
        "실습에서 이 파일은 '내용 확인 불가, 별도 검토 필요'로 분류한다.\n\n"
        "분류 대장에 파일명은 올리되 핵심 내용 요약란은 '내용 인식 불가'로 남기고,\n"
        "감사인이 원본을 직접 열어 확인하는 대상으로 표시한다.\n\n"
        f"{NOTICE}\n"
    )
    with open(p("스캔문서_안내.txt"), "w", encoding="utf-8") as f:
        f.write(text)

# ── README ─────────────────────────────────────────────────────
def readme():
    text = f"""KIIA-2 폴더관리 실습 세트 — README

{NOTICE}

[실습 목적]
버전 중복과 파일명·내용 불일치가 뒤섞인 14개 파일을 유형별로 분류하고,
중복본과 최종본을 구분하여 문서 관리 대장을 만든다. 감사 착수 때 자료 전체를
넘겨받아 그 구성부터 파악하는 실제 장면을 재현한 세트다(따라하기 6-4 대응).

[분류 라벨]
- Contract  : 계약서
- Invoice   : 인보이스(청구)
- Settlement: 정산 내역
- Budget    : 예산 계획
- Vendor    : 벤더 마스터
- Minutes   : 회의록
- Memo      : 메모
- Checklist : 점검표
- Log       : 감사 로그
- Scan      : 스캔본(내용 인식 불가)

[파일 목록 · 14종]
1.  contract_final_v3.docx        [Contract]  용역계약서 v3(그린테크솔루션·수의계약)
2.  계약서_최종(진짜최종).docx     [Contract]  1번과 금액 동일·날짜만 다른 중복본
3.  계약서_최종본_최종.docx        [Settlement] 파일명은 계약서이나 내용은 정산 내역(명칭 불일치)
4.  invoice_2026_03.xlsx          [Invoice]   3월 청구 내역
5.  budget_plan.xlsx              [Budget]    연간 예산·1분기 집행
6.  settlement_Q1.xlsx            [Settlement] 1분기 정산
7.  점검표.xlsx                    [Checklist] 감사 점검 항목표
8.  vendor_master_old.xlsx        [Vendor]    벤더 마스터 구본
9.  vendor_master_new.xlsx        [Vendor]    벤더 마스터 신본(사업자번호·상태·신규 변동)
10. 회의록_20260305.docx          [Minutes]   감사 착수 회의록
11. 메모.docx                     [Memo]      담당자 메모
12. audit_log.csv                 [Log]       파일 접근 기록
13. 스캔문서_안내.txt              [Scan]      스캔 PDF 자리(내용 인식 불가)
14. README.txt                    [-]         본 안내 파일

[의도된 설계]
- 버전 중복: 1·2번 계약서는 금액이 같고 날짜만 다른 사실상 중복본이다.
- 명칭 불일치: 3번은 파일명이 계약서이나 실제 내용은 정산 내역이다.
- 반복 수의계약: 계약서 본문과 벤더 마스터에서 동일 업체(그린테크솔루션)의
  반복 수의계약 단서가 드러난다.
- 신구본 차이: 8·9번 벤더 마스터는 사업자번호 갱신, 상태 변경, 신규 등록에서 차이가 난다.
- 인식 불가: 13번은 OCR이 되지 않는 스캔본을 가정한 자리 파일이다.
위 불일치와 중복은 오류가 아니라 분류·중복 탐지 실습을 위한 의도된 구성이다.

[기대 산출물]
문서 관리 대장(파일명·유형·핵심 내용·버전·특이사항)과 중복·불일치 목록.
대장의 파일 수가 실제 폴더의 파일 수와 일치하는지 대조하는 것이 감사인의 몫이다.
"""
    with open(p("README.txt"), "w", encoding="utf-8") as f:
        f.write(text)

if __name__ == "__main__":
    contract_final_v3(); contract_jinjja(); contract_choejong_mismatch()
    minutes(); memo()
    invoice(); budget(); settlement(); checklist()
    vendor_master(False); vendor_master(True)
    audit_log(); scan_note(); readme()
    files = sorted(os.listdir(ROOT))
    print(f"KIIA-2 생성 완료: {len(files)}개 파일")
    for f in files:
        print("  -", f)
