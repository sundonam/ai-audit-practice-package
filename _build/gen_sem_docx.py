# gen_sem_docx.py — 부록 E 실습 데이터 패키지: SEM-1·SEM-2 조서 실습 DOCX 생성
# A4 / 맑은 고딕 / 표는 표로. 전량 가상 데이터. 책 ch07 서술과 일치.
# 프롬프트 전문(4·3. 프롬프트 일체)은 practice-materials.md 원문을 라인 슬라이스로 verbatim 추출한다.
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\user\Documents\developer\7_audit\practice-package"
MATERIALS = r"C:\Users\user\Documents\developer\7_audit\source\extracted\practice-materials.md"
NOTICE = "본 자료는 실습용 가상 데이터이며 실재 인물·기관·거래와 무관하다."
NAVY = "1A2E5A"
CODEBG = "F2F3F5"
GRAY = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ── 저수준 헬퍼 ────────────────────────────────────────────────
def set_font(run, name="맑은 고딕", size=10.5, bold=False, color=None, ascii_name=None):
    run.font.name = ascii_name or name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), name)
    if ascii_name:
        rfonts.set(qn('w:ascii'), ascii_name); rfonts.set(qn('w:hAnsi'), ascii_name)

def shade(el_tcPr_or_pPr, hexcolor):
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexcolor)
    el_tcPr_or_pPr.append(shd)

def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = "맑은 고딕"; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), "맑은 고딕")
    sec = doc.sections[0]
    sec.page_width = Cm(21.0); sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.4); sec.right_margin = Cm(2.4)
    return doc

def add_notice(doc):
    p = doc.add_paragraph()
    r = p.add_run(NOTICE)
    set_font(r, size=8.5, color=GRAY)
    r.italic = True
    p.paragraph_format.space_after = Pt(10)

def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_font(r, size=16, bold=True)
    p.paragraph_format.space_after = Pt(14)

def add_heading(doc, text, size=12, space_before=10):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_font(r, size=size, bold=True)
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(4)
    return p

def add_body(doc, text, size=10.5, bold=False, indent=None, space_after=4):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_font(r, size=size, bold=bold)
    if indent is not None:
        p.paragraph_format.left_indent = Cm(indent)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.3
    return p

def add_data_table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        shade(hdr[i]._tc.get_or_add_tcPr(), NAVY)
        p = hdr[i].paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h); set_font(r, size=9.5, bold=True, color=WHITE)
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val)); set_font(r, size=9.5)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t

def add_code_block(doc, text):
    """프롬프트 전문을 단일 셀 음영 박스에 고정폭 글꼴로 배치한다."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = 'Table Grid'
    cell = t.rows[0].cells[0]
    shade(cell._tc.get_or_add_tcPr(), CODEBG)
    cell.paragraphs[0].text = ""
    lines = text.split("\n")
    for idx, ln in enumerate(lines):
        p = cell.paragraphs[0] if idx == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(ln if ln else "")
        set_font(r, name="맑은 고딕", ascii_name="Consolas", size=9)
    return t

def save(doc, folder, filename):
    d = os.path.join(ROOT, folder)
    os.makedirs(d, exist_ok=True)
    path = os.path.join(d, filename)
    doc.save(path)
    print("  saved:", filename)
    return path

# ── 프롬프트 전문 verbatim 추출 ────────────────────────────────
def load_prompts():
    with open(MATERIALS, encoding="utf-8") as f:
        lines = f.read().split("\n")
    # 파일 라인 212..321 = SEM-1 6단계, 327..465 = SEM-2 7단계 (0-index 슬라이스)
    sem1 = "\n".join(lines[211:321]).rstrip("\n")
    sem2 = "\n".join(lines[326:465]).rstrip("\n")
    # 앵커 검증 (opening 210/325, closing 321/465)
    assert lines[210] == "```" and lines[321] == "```", "SEM-1 fence mismatch"
    assert lines[325] == "```" and lines[465] == "```", "SEM-2 fence mismatch"
    assert sem1.startswith("1. 실습 프롬프트(사전 단계)"), "SEM-1 start mismatch"
    assert sem2.startswith("1단계(사전단계):"), "SEM-2 start mismatch"
    return sem1, sem2

# ══════════════════════════════════════════════════════════════
# SEM-1 외주용역 계약분할 조서
# ══════════════════════════════════════════════════════════════
SEM1_DIR = "SEM-1_외주용역_계약분할_조서"

def sem1_file1():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "외주 용역 계약 내역 요약 자료")
    add_body(doc, "IT운영부가 2025년 상반기에 체결한 외주 용역 계약 현황이다. 계약 3건은 모두 동일 업체와 수의계약으로 체결되었으며, 약 1개월 간격으로 분산되어 있다.")
    add_data_table(
        doc,
        ["계약번호", "계약명", "업체", "계약방식", "계약금액", "계약일", "계약기간"],
        [
            ["IT-OUT-2025-01", "서버 운영 유지보수 용역", "주식회사 넥스트솔루션", "수의계약", "46,000,000원", "2025-01-15", "2025-01-16 ~ 2025-04-15"],
            ["IT-OUT-2025-02", "네트워크 보안 점검 용역", "주식회사 넥스트솔루션", "수의계약", "49,000,000원", "2025-02-20", "2025-02-21 ~ 2025-05-20"],
            ["IT-OUT-2025-03", "시스템 통합 관제 용역", "주식회사 넥스트솔루션", "수의계약", "47,000,000원", "2025-03-18", "2025-03-19 ~ 2025-06-18"],
        ],
        widths=[2.6, 3.6, 3.0, 1.8, 2.4, 2.0, 3.6],
    )
    add_body(doc, "3건 합계: 142,000,000원", bold=True, space_after=2)
    save(doc, SEM1_DIR, "1. 외주 용역 계약 내역 요약 자료.docx")

def sem1_file2():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "내부 계약 규정 발췌본")
    add_body(doc, "다음은 외주 용역 계약과 관련된 내부 계약 규정의 발췌본이다.", space_after=8)

    add_heading(doc, "제12조 (경쟁입찰 원칙)")
    add_body(doc, "① 용역계약은 원칙적으로 일반경쟁입찰로 체결한다.", indent=0.4)
    add_body(doc, "② 제15조의 요건을 충족하는 경우에 한하여 예외적으로 수의계약으로 체결할 수 있다.", indent=0.4)

    add_heading(doc, "제15조 (수의계약 기준 및 합산 검토)")
    add_body(doc, "① 추정가격이 5천만원 이하인 용역은 수의계약으로 체결할 수 있다.", indent=0.4)
    add_body(doc, "② 동일하거나 유사한 용역을 일정 기간 내에 분할하여 계약하는 경우에는 각 계약금액을 합산하며, 합산액이 3천만원을 초과하는 때에는 제12조에 따라 경쟁입찰에 부친다.", indent=0.4)
    add_body(doc, "③ 계약담당자는 수의계약 체결 시 합산 대상 여부를 검토한 기록을 남긴다.", indent=0.4)

    add_heading(doc, "제16조 (유사 용역의 판단)")
    add_body(doc, "용역의 목적, 대상 시스템, 수행 업체, 계약 시기의 근접성을 종합적으로 고려하여 동일하거나 유사한 용역에 해당하는지를 판단한다.", indent=0.4)
    save(doc, SEM1_DIR, "2. 내부 계약 규정 발췌본.docx")

def sem1_file3():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "감사 사실 메모")
    add_body(doc, "현장 확인 및 계약 서류 검토를 통해 확인한 사실이다.", space_after=8)
    facts = [
        "1. IT운영부는 2025년 상반기(1월~3월)에 주식회사 넥스트솔루션과 외주 용역 계약 3건(IT-OUT-2025-01·02·03)을 체결하였다.",
        "2. 3건 모두 수의계약으로 체결되었으며, 각 계약금액은 46,000,000원에서 49,000,000원으로 수의계약 한도(5천만원) 이하이다.",
        "3. 계약일은 1월 15일, 2월 20일, 3월 18일로 약 1개월 간격으로 분산되어 있다.",
        "4. 3건에 대한 경쟁입찰 검토 또는 계약금액 합산 검토 기록은 확인되지 않았다.",
        "5. 3건의 과업 내용은 서버 운영, 보안 점검, 통합 관제로 모두 동일 정보시스템을 대상으로 한 유사 용역으로 확인되었다.",
    ]
    for f in facts:
        add_body(doc, f, space_after=6)
    save(doc, SEM1_DIR, "3. 감사 사실 메모.docx")

def sem1_file4(sem1_prompt):
    doc = new_doc(); add_notice(doc)
    add_title(doc, "프롬프트 일체")
    add_body(doc, "외주용역 계약분할 조서 작성 실습의 6단계 프롬프트 전문이다. 원문 그대로 옮긴다.", space_after=8)
    add_code_block(doc, sem1_prompt)
    save(doc, SEM1_DIR, "4. 프롬프트 일체.docx")

def sem1_file5():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "감사조서 표준 양식")
    add_body(doc, "계약분할 사안을 채운 완성 예시 조서다. 기대 산출물의 구조와 문체를 확인하는 기준으로 활용한다.", space_after=8)

    add_heading(doc, "내부감사 결과보고서", size=13, space_before=4)
    add_body(doc, "제목: IT운영부 외주 용역 수의계약 분할 체결 및 합산 검토 미이행", bold=True, space_after=6)
    add_body(doc, "감사 대상 부서: IT운영부", space_after=2)
    add_body(doc, "감사 대상 업무: 2025년 상반기 외주 용역 계약 체결 및 수의계약 관리", space_after=2)
    add_body(doc, "감사 대상 기간: 2025년 1월 ~ 2025년 6월", space_after=6)

    add_heading(doc, "1. 감사목적")
    add_body(doc, "외주 용역 계약의 수의계약 적정성과 유사 용역 합산 검토의 이행 여부를 점검한다.", indent=0.4)

    add_heading(doc, "2. 관련 내부 규정")
    add_body(doc, "- 계약 규정 제12조(경쟁입찰 원칙)", indent=0.4)
    add_body(doc, "- 계약 규정 제15조(수의계약 기준 및 합산 검토)", indent=0.4)
    add_body(doc, "- 계약 규정 제16조(유사 용역의 판단)", indent=0.4)

    add_heading(doc, "3. 감사결과 요지")
    add_body(doc, "그럼에도 불구하고 IT운영부는 동일 업체와 목적이 유사한 외주 용역을 2025년 상반기 중 3건으로 나누어 각각 수의계약으로 체결하였음에도 불구하고, 각 계약금액을 합산하여 경쟁입찰 여부를 검토한 기록을 남기지 않은 사실이 확인되었다.", indent=0.4)

    add_heading(doc, "4. 영향 분석")
    add_body(doc, "- 3건의 계약금액 합산액은 142,000,000원으로 유사 용역 합산 검토 기준(3천만원)을 초과한다.", indent=0.4)
    add_body(doc, "- 합산 검토 없이 수의계약으로 체결되어 경쟁입찰을 통한 가격 비교와 업체 선정의 공정성 확보 기회가 제한되었다.", indent=0.4)
    add_body(doc, "- 동일 업체와의 계약이 반복되어 특정 업체에 대한 의존도가 높아질 수 있다.", indent=0.4)

    add_heading(doc, "5. 개선요구 사항")
    add_body(doc, "유사하거나 반복되는 용역에 대한 합산 검토 절차를 계약 체결 단계의 필수 통제로 규정에 반영하고, 수의계약 체결 시 합산 대상 여부를 검토한 기록을 의무적으로 보존하도록 관리 체계를 개선하기 바란다.", indent=0.4)

    add_heading(doc, "6. 유의사항")
    add_body(doc, "향후 동일 업체와의 반복 계약 시 계약금액 합산과 경쟁입찰 요건을 사전에 검토하여 수의계약 관리에 철저를 기하기 바란다.", indent=0.4)
    save(doc, SEM1_DIR, "5. 감사조서 표준 양식.docx")

# ══════════════════════════════════════════════════════════════
# SEM-2 IT보안 인터뷰 조서
# ══════════════════════════════════════════════════════════════
SEM2_DIR = "SEM-2_IT보안_인터뷰_조서"

def sem2_file1():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "면담 데이터")
    add_body(doc, "IT 보안 운영 프로세스 점검을 위해 수행한 면담 일지 4건이다. 진술은 구어 그대로 옮겼으며 사실, 의견, 추정이 혼재되어 있다.", space_after=8)

    interviews = [
        {
            "no": "면담 1",
            "meta": [("일시", "2026-05-12 10:00"), ("대상", "IT운영팀 담당자 A"), ("면담자", "감사부 김OO")],
            "qa": [
                ("지난 3월 장애를 언제 처음 인지하셨나요?",
                 "\"아 그날 아침에 모니터링 화면이 좀 이상하다 싶었어요. 정확히 몇 시라고 하긴 그런데, 한 9시 반쯤이었나. 처음엔 일시적인 건 줄 알고 좀 지켜봤죠.\""),
                ("인지하신 뒤에 어떻게 조치하셨나요?",
                 "\"일단 팀장님한테 구두로 바로 말씀드렸어요. 시스템 등록은... 솔직히 그날은 정신이 없어서 못 하고, 다음 날 아침에 등록했던 것 같아요.\""),
                ("장애 등급은 어떻게 정하셨나요?",
                 "\"그건 저희가 보고 판단하는 건데, 그때는 그렇게 심각한 건 아니라고 봐서 중간 등급 정도로 잡았어요. 근데 사실 등급 기준이 좀 애매하긴 해요.\""),
            ],
        },
        {
            "no": "면담 2",
            "meta": [("일시", "2026-05-12 11:00"), ("대상", "IT운영팀 팀장 B"), ("면담자", "감사부 김OO")],
            "qa": [
                ("담당자에게 보고를 언제 받으셨나요?",
                 "\"당일 오전에 구두로 들었죠. 근데 그때는 큰일 아니라고 해서 저도 그냥 알겠다 하고 넘어갔어요.\""),
                ("상부 보고나 IT보안팀 공유는 하셨나요?",
                 "\"그날 바로는 안 했고, 상황 좀 보고 하려다가 타이밍을 놓쳤네요. 보안팀에는 며칠 뒤에 공유됐을 거예요, 아마.\""),
                ("장애 등록 절차는 지켜졌다고 보시나요?",
                 "\"규정상 바로 등록해야 하는 건 아는데, 현장에서는 급하면 일단 처리부터 하고 등록은 나중에 하는 경우가 종종 있어요.\""),
            ],
        },
        {
            "no": "면담 3",
            "meta": [("일시", "2026-05-13 14:00"), ("대상", "IT보안팀 C"), ("면담자", "감사부 김OO")],
            "qa": [
                ("이번 장애를 언제 공유받으셨나요?",
                 "\"저희는 한참 뒤에 알았어요. 시스템에 등록이 돼 있어야 저희가 보는 건데, 등록이 늦으니까 저희도 늦게 본 거죠.\""),
                ("장애 등급 판단에 보안팀도 관여하나요?",
                 "\"원래는 같이 봐야 하는데, 실제로는 운영팀에서 먼저 정하고 저희한테는 사후에 알려주는 식이에요. 그래서 등급이 좀 낮게 잡히는 경우도 있는 것 같고요.\""),
            ],
        },
        {
            "no": "면담 4",
            "meta": [("일시", "2026-05-13 16:00"), ("대상", "경영지원부 D"), ("면담자", "감사부 김OO")],
            "qa": [
                ("장애 보고 규정은 어떻게 되어 있나요?",
                 "\"규정상으로는 인지하면 30분 안에 시스템에 등록하고, 등급별로 보고 라인이 정해져 있어요. 상위 등급이면 경영진까지 올라가야 하고요.\""),
                ("실제로도 그렇게 운영되나요?",
                 "\"솔직히 말씀드리면 규정이랑 현실이랑 좀 차이가 있어요. 등록이 늦거나 빠지는 경우도 있고, 등급도 담당자가 정하다 보니 들쭉날쭉한 편이에요.\""),
            ],
        },
    ]
    for iv in interviews:
        add_heading(doc, iv["no"], size=12)
        meta = " / ".join(f"{k}: {v}" for k, v in iv["meta"])
        add_body(doc, meta, size=9.5, space_after=4)
        for q, a in iv["qa"]:
            add_body(doc, "질문. " + q, bold=True, indent=0.4, space_after=1)
            add_body(doc, "답변. " + a, indent=0.4, space_after=6)
    save(doc, SEM2_DIR, "1. 면담 데이터.docx")

def sem2_file2():
    doc = new_doc(); add_notice(doc)
    add_title(doc, "내부 규정 발췌본")
    add_body(doc, "가상 「IT 장애 대응 및 보고 규정」의 발췌본이다. 규정상 요구되는 절차(To-Be)의 판단 기준으로 활용한다.", space_after=8)
    arts = [
        ("제1조 (목적)", "이 규정은 정보시스템 장애의 신속한 대응과 보고 체계를 정하여 업무 연속성을 확보함을 목적으로 한다."),
        ("제2조 (적용범위)", "정보시스템 운영 중 발생하는 모든 장애에 적용한다."),
        ("제3조 (장애 등급 분류)", "장애는 1급(전사 서비스 중단), 2급(주요 기능 장애), 3급(부분·경미 장애)으로 분류하며, 정해진 기준에 따라 등급을 판단한다."),
        ("제4조 (장애 인지 및 등록)", "장애를 인지한 담당자는 인지 후 30분 이내에 장애관리시스템에 장애 내용을 등록한다."),
        ("제5조 (등급 판단 주체)", "장애 등급은 IT운영팀과 IT보안팀이 공동으로 판단한다."),
        ("제6조 (보고 절차)", "장애 등록 즉시 등급별 보고 라인에 따라 보고한다. 1급은 경영진, 2급은 부서장, 3급은 팀장까지 보고한다."),
        ("제7조 (부서 간 공유)", "장애 등록 내용은 IT보안팀과 관련 부서에 지체 없이 공유한다."),
        ("제8조 (대응 및 조치 기록)", "대응 과정과 조치 내용을 장애관리시스템에 기록한다."),
        ("제9조 (사후 보고)", "장애 종료 후 3일 이내에 원인과 재발 방지 대책을 담은 사후 보고서를 작성한다."),
        ("제10조 (점검)", "감사부서는 장애 대응 및 보고의 적정성을 정기적으로 점검한다."),
    ]
    for title, body in arts:
        add_heading(doc, title, size=11, space_before=8)
        add_body(doc, body, indent=0.4)
    save(doc, SEM2_DIR, "2. 내부 규정 발췌본.docx")

def sem2_file3(sem2_prompt):
    doc = new_doc(); add_notice(doc)
    add_title(doc, "프롬프트 일체")
    add_body(doc, "IT보안 운영 인터뷰 조서 작성 실습의 7단계 프롬프트 전문이다. 원문 그대로 옮긴다.", space_after=8)
    add_code_block(doc, sem2_prompt)
    save(doc, SEM2_DIR, "3. 프롬프트 일체.docx")

# ══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    sem1_prompt, sem2_prompt = load_prompts()
    print("SEM-1 폴더:")
    sem1_file1(); sem1_file2(); sem1_file3(); sem1_file4(sem1_prompt); sem1_file5()
    print("SEM-2 폴더:")
    sem2_file1(); sem2_file2(); sem2_file3(sem2_prompt)
    print("\nSEM 조서 실습 DOCX 생성 완료.")
