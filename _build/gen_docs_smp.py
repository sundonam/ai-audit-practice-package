# gen_docs_smp.py — 부록 E 실습 데이터 패키지: SMP-2·4·5·6 문서·보고서(DOCX)
# 전량 가상 데이터. 책 6·9장의 실습 서술과 일치하도록 이상징후를 의도적으로 심는다.
import os
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\user\Documents\developer\7_audit\practice-package"
NOTICE = "본 자료는 실습용 가상 데이터이며 실재 인물·기관·거래와 무관하다."
FONT = "맑은 고딕"


def _set_run_font(run, size=None, bold=None, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def new_doc():
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(11)
    st.paragraph_format.line_spacing = 1.3
    for sec in doc.sections:
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)
        sec.top_margin = Mm(25)
        sec.bottom_margin = Mm(25)
        sec.left_margin = Mm(25)
        sec.right_margin = Mm(25)
    return doc


def notice(doc):
    p = doc.add_paragraph()
    r = p.add_run(NOTICE)
    _set_run_font(r, size=9, color=(0x80, 0x80, 0x80))
    r.italic = True
    p.paragraph_format.space_after = Pt(10)


def title(doc, text, sub=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    _set_run_font(r, size=18, bold=True)
    p.paragraph_format.space_after = Pt(4)
    if sub:
        ps = doc.add_paragraph()
        ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rs = ps.add_run(sub)
        _set_run_font(rs, size=11, color=(0x40, 0x40, 0x40))
        ps.paragraph_format.space_after = Pt(12)


def heading(doc, text, size=13):
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=size, bold=True, color=(0x1A, 0x2E, 0x5A))
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    return p


def body(doc, text, bold=False, space_after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    _set_run_font(r, size=11, bold=bold)
    p.paragraph_format.space_after = Pt(space_after)
    return p


def labeled(doc, label, text):
    p = doc.add_paragraph()
    r1 = p.add_run(f"{label}  ")
    _set_run_font(r1, size=11, bold=True, color=(0x1A, 0x2E, 0x5A))
    r2 = p.add_run(text)
    _set_run_font(r2, size=11)
    p.paragraph_format.space_after = Pt(4)
    return p


def _shade(cell, hexcolor):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tcpr.append(shd)


def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        r = hdr[i].paragraphs[0].add_run(h)
        _set_run_font(r, size=10, bold=True, color=(0xFF, 0xFF, 0xFF))
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(hdr[i], "1A2E5A")
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(val))
            _set_run_font(r, size=10)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Mm(w)
    return t


def save(doc, folder, name):
    path = os.path.join(ROOT, folder, name)
    doc.save(path)
    print(f"  저장: {folder}/{name}")


# ══════════════════════════════════════════════════════════════════
# SMP-2  HR 감사보고서 실습 (3종)
# ══════════════════════════════════════════════════════════════════
def smp2_past_report():
    doc = new_doc()
    notice(doc)
    title(doc, "인사(HR) 부문 내부감사 결과보고서", "감사대상기간 2025년 상반기 · 감사부서 감사실")
    body(doc, "본 보고서는 인사 부문 정기감사에서 확인한 발견사항을 정리한 과거 감사보고서다. "
              "이번 실습에서는 이 보고서를 현황과 원인의 확인 자료로 활용하되, 지적사항이 이미 조치되었는지 여부는 "
              "면담 기록과 규정으로 다시 확인한다.")

    heading(doc, "발견사항 1. 퇴직자 계정 관리 미흡")
    labeled(doc, "현황",
            "2025년 상반기 퇴직자 18명의 계정 회수 내역을 점검한 결과, 6명의 시스템 계정이 퇴직일 이후에도 "
            "활성 상태로 유지되고 있었다. 이 가운데 2명의 계정은 퇴직 후 2주 이상 로그인 가능 상태로 남아 있었다.")
    labeled(doc, "문제점",
            "인사관리규정 제12조는 퇴직자 계정을 퇴직 후 24시간 이내에 회수하도록 규정하고 있으나, 인사팀의 "
            "퇴직 처리와 정보시스템팀의 계정 회수가 연계되지 않아 회수 시점이 담당자 개인의 처리 순서에 따라 달라졌다.")
    labeled(doc, "권고",
            "퇴직 처리와 계정 회수를 연동하는 통제를 신설하고, 인사 시스템의 퇴직 확정과 동시에 계정 비활성화 요청이 "
            "자동으로 정보시스템팀에 전달되도록 절차를 정비한다.")
    labeled(doc, "결과",
            "퇴직자 계정이 회수되지 않은 채 유지되면 퇴직자에 의한 내부 시스템 무단 접근과 정보 유출 위험이 발생한다. "
            "회수 지연 계정에서 실제 접근 이력이 있었는지는 시스템 로그로 추가 확인이 필요하다.")

    heading(doc, "발견사항 2. 신규 입사자 보안교육 미이행")
    labeled(doc, "현황",
            "2025년 상반기 신규 입사자 22명 가운데 9명이 입사 1개월이 지나도록 정보보안 교육을 이수하지 않았다. "
            "이 중 4명은 개인정보를 취급하는 부서에 배치되어 있었다.")
    labeled(doc, "문제점",
            "인사관리규정 제18조는 신규 입사자가 입사 후 1개월 이내에 정보보안 교육을 이수하도록 규정하고 있으나, "
            "교육 대상자 명단이 부서별로 관리되어 이수 여부를 전사 차원에서 점검하는 절차가 없었다.")
    labeled(doc, "권고",
            "신규 입사자 보안교육을 입사 절차의 필수 항목으로 포함하고, 이수 현황을 인사부서가 월 단위로 취합해 "
            "미이수자에게 재안내하는 점검 주기를 마련한다.")
    labeled(doc, "결과",
            "보안교육을 이수하지 않은 인원이 개인정보 취급 업무에 투입되면 정보 오남용과 규정 위반의 소지가 커진다. "
            "미이수 기간 중 실제 개인정보 처리 이력이 있었는지는 별도 확인이 필요하다.")
    save(doc, "SMP-2_HR_감사보고서", "과거_감사_보고서_HR_감사.docx")


def smp2_interview():
    doc = new_doc()
    notice(doc)
    title(doc, "인사팀장 면담 녹취록", "면담일 2026년 3월 12일 · 면담자 감사실 · 피면담자 인사팀장")
    body(doc, "다음은 인사 부문 감사 과정에서 진행한 인사팀장 면담 녹취록이다. 진술에는 확인된 사실과 면담자의 "
              "주관적 의견이 함께 섞여 있으므로, 사실만을 근거로 삼고 의견은 별도로 구분해 읽어야 한다.", space_after=10)

    heading(doc, "문 1. 퇴직자 계정 관리에서 취약한 부분이 무엇이라고 보십니까")
    body(doc, "감사실:  퇴직자 계정 회수가 늦어지는 경우가 있다고 들었습니다. 실제로 어떤 상황인지 말씀해 주시겠습니까.")
    body(doc,
         "인사팀장:  네, 그건 좀 문제였죠. 솔직히 말씀드리면 저희가 퇴직 처리를 하면 그다음에 IT팀에 계정 닫아 "
         "달라고 따로 메일을 보내야 하는데, 이게 사람이 하는 일이다 보니까 바쁘면 하루 이틀 밀리기도 하고요. "
         "규정상으로는 24시간 안에 닫아야 하는 걸로 아는데, 솔직히 그 시간을 다 지키기가 쉽지 않았어요. "
         "제 생각에는 담당자가 특별히 게을러서 그런 건 아니고, 두 팀 사이에 자동으로 넘어가는 시스템이 없어서 "
         "그런 것 같습니다. 지난번에도 퇴직한 직원 계정이 2주 정도 열려 있었다고 감사에서 지적받은 적이 있었죠.",
         space_after=10)

    heading(doc, "문 2. 개선 방안으로 무엇이 필요하다고 생각하십니까")
    body(doc, "감사실:  이 부분을 개선하려면 어떤 조치가 필요하다고 보십니까.")
    body(doc,
         "인사팀장:  제일 좋은 건 인사 시스템에서 퇴직 처리를 하면 IT팀 쪽으로 계정 정지 요청이 자동으로 넘어가는 "
         "겁니다. 지금은 그걸 사람이 손으로 챙기니까 놓치는 거예요. 그리고 신규 입사자 보안교육도 비슷한데, 이것도 "
         "부서마다 알아서 챙기라고 하니까 안 듣는 사람이 꼭 나오더라고요. 제 개인적인 생각으로는 교육을 안 들으면 "
         "아예 시스템 접근 권한을 안 주는 식으로 강제하는 게 맞다고 봅니다. 물론 이건 IT팀이랑 협의를 해봐야 하는 "
         "부분이라 제가 확답을 드리긴 어렵습니다만.",
         space_after=8)
    save(doc, "SMP-2_HR_감사보고서", "녹취록_인사팀장_인터뷰.docx")


def smp2_hr_rule():
    doc = new_doc()
    notice(doc)
    title(doc, "인사관리규정 (발췌본)", "제정 2019. 3. 1. · 최종개정 2024. 7. 1.")
    body(doc, "본 규정은 실습용으로 발췌한 인사관리규정이다. 발견사항 판단의 기준으로만 활용하며, 실제 조직의 "
              "규정과는 무관하다.", space_after=10)

    heading(doc, "제3장 계정 및 정보보안 관리")

    body(doc, "제12조 (퇴직자 계정의 회수)", bold=True, space_after=2)
    body(doc, "① 인사부서는 임직원의 퇴직이 확정된 경우 지체 없이 그 사실을 정보시스템 담당 부서에 통보하여야 한다.")
    body(doc, "② 정보시스템 담당 부서는 퇴직자의 정보시스템 계정 및 접근 권한을 퇴직일로부터 24시간 이내에 회수하여야 한다.")
    body(doc, "③ 계정 회수 결과는 회수 대장에 기록하고 인사부서와 공유하여야 한다.", space_after=8)

    body(doc, "제18조 (신규 입사자의 정보보안 교육)", bold=True, space_after=2)
    body(doc, "① 신규 입사자는 입사일로부터 1개월 이내에 정보보안 교육을 이수하여야 한다.")
    body(doc, "② 인사부서는 신규 입사자의 교육 이수 여부를 관리하고, 미이수자에 대하여 재교육을 안내하여야 한다.")
    body(doc, "③ 개인정보 취급 업무에 종사하는 직원은 제1항의 교육을 이수하기 전까지 개인정보 처리 시스템에 대한 "
              "접근 권한을 부여받을 수 없다.", space_after=8)

    body(doc, "제19조 (교육 기록의 보존)", bold=True, space_after=2)
    body(doc, "정보보안 교육의 이수 기록은 교육 종료일로부터 3년간 보존한다.")
    save(doc, "SMP-2_HR_감사보고서", "인사_관리_규정.docx")


# ══════════════════════════════════════════════════════════════════
# SMP-4  계약서 검토 (전자부품 공급계약서)
# ══════════════════════════════════════════════════════════════════
def smp4_contract():
    doc = new_doc()
    notice(doc)
    title(doc, "전자부품 공급계약서")
    body(doc,
         "주식회사 미래제조(이하 “갑”이라 한다)와 주식회사 가온전자(이하 “을”이라 한다)는 전자부품 공급에 관하여 "
         "다음과 같이 계약을 체결한다.", space_after=10)

    body(doc, "제1조 (목적)", bold=True, space_after=2)
    body(doc, "본 계약은 갑이 을에게 발주하는 전자부품(모델 A123)의 공급 조건과 대금 지급, 검수 및 품질보증에 관한 "
              "사항을 정함을 목적으로 한다.", space_after=8)

    body(doc, "제2조 (계약 물품 및 금액)", bold=True, space_after=2)
    body(doc, "① 공급 물품은 전자부품(모델 A123) 1,000개로 한다.")
    body(doc, "② 계약 총액은 금 오천만원(₩50,000,000, 부가가치세 포함)으로 한다.")
    body(doc, "③ 단가는 개당 금 오만원(₩50,000)으로 한다.", space_after=8)

    body(doc, "제3조 (대금 지급)", bold=True, space_after=2)
    body(doc, "① 갑은 계약 체결과 동시에 계약 총액의 90퍼센트에 해당하는 금액을 을에게 선지급한다.")
    body(doc, "② 갑은 잔여 대금 10퍼센트를 납품 완료 후 을이 청구하는 즉시 지급한다.")
    body(doc, "③ 대금은 을이 지정하는 계좌로 현금 이체하여 지급한다.", space_after=8)

    body(doc, "제4조 (납품)", bold=True, space_after=2)
    body(doc, "① 을은 계약 체결일로부터 30일 이내에 계약 물품 전량을 갑이 지정하는 장소에 납품한다.")
    body(doc, "② 납품에 소요되는 운송비는 을이 부담한다.", space_after=8)

    body(doc, "제5조 (검수)", bold=True, space_after=2)
    body(doc, "① 갑은 납품일로부터 7일 이내에 물품의 수량과 상태를 검수한다.")
    body(doc, "② 갑이 제1항의 기간 내에 검수 결과를 통지하지 아니한 경우 검수에 합격한 것으로 본다.", space_after=8)

    body(doc, "제6조 (품질보증)", bold=True, space_after=2)
    body(doc, "① 을은 납품한 물품에 대하여 검수 완료일로부터 3개월간 품질을 보증한다.")
    body(doc, "② 보증 기간 내에 을의 책임 있는 사유로 하자가 발생한 경우 을은 해당 물품을 무상으로 교체한다.", space_after=8)

    body(doc, "제7조 (하자 처리)", bold=True, space_after=2)
    body(doc, "① 갑은 물품의 하자를 발견한 경우 납품일로부터 1개월 이내에 을에게 하자 보수를 청구하여야 한다.")
    body(doc, "② 제1항의 기간이 지난 후에 발견된 하자에 대하여 을은 책임을 지지 아니한다.", space_after=8)

    body(doc, "제8조 (분쟁의 해결 및 기타)", bold=True, space_after=2)
    body(doc, "① 본 계약과 관련하여 분쟁이 발생한 경우 갑과 을은 상호 협의하여 해결하며, 협의가 이루어지지 아니한 "
              "경우의 관할에 관하여는 별도로 정하지 아니한다.")
    body(doc, "② 본 계약에 정하지 아니한 사항은 관계 법령과 일반 상관례에 따른다.")
    body(doc, "③ 본 계약의 성립을 증명하기 위하여 계약서 2부를 작성하고 갑과 을이 각각 1부씩 보관한다.", space_after=10)

    body(doc, "2026년 3월 2일", space_after=4)
    body(doc, "갑  주식회사 미래제조     대표이사 (인)")
    body(doc, "을  주식회사 가온전자     대표이사 (인)")
    save(doc, "SMP-4_계약서_검토", "계약서 샘플.docx")


# ══════════════════════════════════════════════════════════════════
# SMP-5  구매규정 준수 (구매규정 6개 조항)
# ══════════════════════════════════════════════════════════════════
def smp5_rule():
    doc = new_doc()
    notice(doc)
    title(doc, "구매관리규정 (샘플)", "제정 2021. 1. 1. · 최종개정 2024. 3. 1.")
    body(doc, "본 규정은 실습용으로 작성한 구매관리규정이다. 조항별 점검을 통해 통제 공백과 개선 요소를 도출하는 "
              "실습 자료로 활용한다.", space_after=10)

    body(doc, "제1조 (목적)", bold=True, space_after=2)
    body(doc, "본 규정은 회사의 물품 및 용역 구매 업무에 관한 기준과 절차를 정하여 구매의 공정성과 효율성을 "
              "확보함을 목적으로 한다.", space_after=8)

    body(doc, "제2조 (적용 범위)", bold=True, space_after=2)
    body(doc, "본 규정은 회사가 수행하는 모든 물품 구매 및 용역 계약에 적용한다. 다만 긴급을 요하는 경우에는 "
              "별도의 기준에 따를 수 있다.", space_after=8)

    body(doc, "제3조 (구매 절차)", bold=True, space_after=2)
    body(doc, "① 구매를 요청하는 부서는 구매요청서를 작성하여 구매부서에 제출한다.")
    body(doc, "② 구매 예정 금액이 1,000만원 이상인 경우에는 구매심의위원회의 심의를 거쳐야 한다.")
    body(doc, "③ 구매부서는 견적을 받아 거래처를 선정하고 구매를 진행한다.", space_after=8)

    body(doc, "제4조 (계약 체결)", bold=True, space_after=2)
    body(doc, "① 구매부서는 선정된 거래처와 계약을 체결한다.")
    body(doc, "② 계약 조건이 통상의 범위를 벗어나는 경우 부서장의 승인을 받아 예외적으로 계약을 체결할 수 있다.", space_after=8)

    body(doc, "제5조 (윤리 준수)", bold=True, space_after=2)
    body(doc, "구매 업무를 담당하는 임직원은 거래처로부터 금품이나 향응을 제공받아서는 아니 되며, 공정하게 업무를 "
              "수행하여야 한다.", space_after=8)

    body(doc, "제6조 (기록 보관)", bold=True, space_after=2)
    body(doc, "구매와 관련한 문서는 계약 종료일로부터 5년간 보관한다.", space_after=10)

    body(doc, "부칙", bold=True, space_after=2)
    body(doc, "본 규정은 개정한 날부터 시행한다.")
    save(doc, "SMP-5_구매규정_준수", "샘플_구매규정.docx")


# ══════════════════════════════════════════════════════════════════
# SMP-6  감사이슈 보고서 분석 (4종 · 이슈 8건)
# ══════════════════════════════════════════════════════════════════
SMP6_REPORTS = [
    ("샘플 감사 이슈 보고서 1.docx", "감사 이슈 보고서 (1)", "총무부문·개인정보 보호", [
        {
            "name": "고정자산 관리대장과 실물 불일치",
            "dept": "총무팀",
            "cond": "고정자산 관리대장에 등재된 자산 320점 중 표본 40점을 실사한 결과 6점의 실물 소재가 확인되지 "
                    "않았으며, 관리대장에 없는 미등재 자산 3점이 발견되었다. 자산 실사는 최근 2년간 시행되지 않았다.",
            "rec": "연 1회 정기 자산 실사를 의무화하고, 자산 취득·이동·폐기 시 관리대장을 즉시 갱신하도록 "
                   "책임 부서와 갱신 절차를 명확히 한다.",
        },
        {
            "name": "개인정보 처리시스템 접근권한 과다 부여",
            "dept": "인사팀",
            "cond": "인사 개인정보 처리시스템의 접근 권한 보유자 28명 중 11명이 현재 업무상 해당 정보를 처리하지 "
                    "않는 인원이었다. 부서 이동 후에도 종전 권한이 회수되지 않은 사례가 다수 확인되었다.",
            "rec": "접근 권한을 업무 필요 범위로 한정하고, 반기마다 권한 보유 현황을 재검토하여 불필요한 권한을 "
                   "회수하는 절차를 신설한다.",
        },
    ]),
    ("샘플 감사 이슈 보고서 2.docx", "감사 이슈 보고서 (2)", "구매·외주 부문", [
        {
            "name": "외주계약 성과 점검 절차 부재",
            "dept": "IT운영팀",
            "cond": "2025년 체결된 외주 용역계약 12건 중 계약 종료 후 성과 점검 결과가 문서로 남아 있는 계약은 "
                    "3건에 불과했다. 나머지 계약은 산출물 검수 없이 대금이 지급되었다.",
            "rec": "외주계약 종료 시 산출물 검수와 성과 점검을 필수 절차로 규정하고, 점검 결과를 대금 지급의 "
                   "선행 요건으로 삼는다.",
        },
        {
            "name": "구매 승인 단계 생략",
            "dept": "구매팀",
            "cond": "1,000만원 이상 구매 42건 중 5건이 구매심의위원회 심의를 거치지 않고 집행되었다. 이 중 2건은 "
                    "동일 거래처에 대한 분할 발주로, 개별 금액은 한도 미만이나 합산 시 심의 대상에 해당하였다.",
            "rec": "심의 대상 여부를 동일 거래처·동일 기간 합산 기준으로 판단하도록 규정을 보완하고, 심의 누락 "
                   "건에 대한 사후 승인 절차와 책임을 명확히 한다.",
        },
    ]),
    ("샘플 감사 이슈 보고서 3.docx", "감사 이슈 보고서 (3)", "협력업체·정보보안 부문", [
        {
            "name": "협력업체 평가 미실시",
            "dept": "구매팀",
            "cond": "등록 협력업체 65개사 중 최근 1년간 정기 평가를 받은 업체는 22개사에 그쳤다. 평가 기준은 마련되어 "
                    "있으나 평가 시행 주기와 책임 부서가 규정에 명시되어 있지 않았다.",
            "rec": "협력업체 정기 평가의 주기와 책임 부서를 규정에 명시하고, 평가 결과를 재계약 및 신규 발주의 "
                   "판단 자료로 연계한다.",
        },
        {
            "name": "정보보안 교육 이수율 저조",
            "dept": "정보보안팀",
            "cond": "전 직원 대상 연간 정보보안 교육의 이수율이 61퍼센트에 머물렀다. 미이수자에 대한 재안내나 "
                    "이수 독려 절차가 운영되지 않았다.",
            "rec": "정보보안 교육 이수 현황을 부서별로 관리하고, 미이수자에 대한 재안내와 이수 기한을 설정하여 "
                   "이수율을 관리 지표로 점검한다.",
        },
    ]),
    ("샘플 감사 이슈 보고서 4.docx", "감사 이슈 보고서 (4)", "계약·검수 부문", [
        {
            "name": "계약서 서명 누락",
            "dept": "재무팀",
            "cond": "2025년 체결 계약 문서 58건을 점검한 결과 4건에서 계약 당사자 일방의 서명 또는 날인이 누락되어 "
                    "있었다. 서명 누락 계약도 대금 지급이 정상적으로 집행되었다.",
            "rec": "계약 체결 시 서명·날인 완결 여부를 확인하는 점검 항목을 마련하고, 서명이 완결되지 않은 계약은 "
                   "대금 집행 대상에서 제외한다.",
        },
        {
            "name": "납품 검수 기록 미비",
            "dept": "생산관리팀",
            "cond": "물품 납품 30건 중 9건에서 검수 조서가 작성되지 않았거나 검수자 서명이 누락되어 있었다. 검수 "
                    "없이 입고 처리된 물품에서 수량 부족이 뒤늦게 확인된 사례가 있었다.",
            "rec": "납품 검수를 입고 처리의 선행 절차로 규정하고, 검수 조서에 검수자·수량·상태를 기록하여 "
                   "검수 완료 후에만 입고와 대금 지급이 이루어지도록 통제한다.",
        },
    ]),
]


def smp6_reports():
    for fname, doctitle, scope, issues in SMP6_REPORTS:
        doc = new_doc()
        notice(doc)
        title(doc, doctitle, f"감사부서 감사실 · 대상범위 {scope}")
        body(doc, "본 보고서는 부서별 감사 과정에서 확인한 개별 이슈를 정리한 자료다. 이슈별로 이슈명, 부서, 현황, "
                  "개선권고를 기재하였으며, 위험 등급과 리스크 유형 분류는 통합 분석 단계에서 부여한다.", space_after=10)
        for idx, iss in enumerate(issues, 1):
            heading(doc, f"이슈 {idx}. {iss['name']}")
            labeled(doc, "부서", iss["dept"])
            labeled(doc, "현황", iss["cond"])
            labeled(doc, "개선권고", iss["rec"])
        save(doc, "SMP-6_감사이슈_보고서", fname)


if __name__ == "__main__":
    print("SMP-2 HR 감사보고서 실습")
    smp2_past_report()
    smp2_interview()
    smp2_hr_rule()
    print("SMP-4 계약서 검토")
    smp4_contract()
    print("SMP-5 구매규정 준수")
    smp5_rule()
    print("SMP-6 감사이슈 보고서 분석")
    smp6_reports()
    print("\nSMP 문서·보고서 생성 완료.")
