# gen_data.py — 부록 E 실습 데이터 패키지: 핵심 데이터 파일(CSV·XLSX) 생성
# 책의 실습 설명과 일치하도록 이상징후를 의도적으로 심는다. 전량 가상 데이터.
import csv, os, random, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

random.seed(42)
ROOT = r"C:\Users\user\Documents\developer\7_audit\practice-package"
NOTICE = "본 자료는 실습용 가상 데이터이며 실재 인물·기관·거래와 무관하다."

def styled_header(ws, headers, row=1):
    fill = PatternFill("solid", fgColor="1A2E5A"); font = Font(color="FFFFFF", bold=True, name="맑은 고딕")
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=c, value=h); cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")

def autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(w*1.6+2, 10), 42)

DEPTS = ["영업팀","인사팀","IT운영팀","재무팀","총무팀","생산관리팀","구매팀","경영지원팀"]

# ── SMP-3 급여 300행 CSV ─────────────────────────────────────────
def smp3():
    path = os.path.join(ROOT,"SMP-3_급여_데이터","sample_kor_hr_payroll_300.csv")
    rows=[];
    # 유령직원: 서로 다른 employee_id가 동일 계좌 공유 (2쌍)
    ghost = {107:"가상은행 3021-08-118677", 233:"가상은행 3021-08-118677",
             56:"가상은행 7742-11-390455", 288:"가상은행 7742-11-390455"}
    ot_outlier = {14, 61, 129, 205, 271}     # 초과근무 대비 수당 과다
    sal_outlier = {42, 190}                   # 기본급 이상치
    for i in range(1,301):
        emp=f"E{i:04d}"; dep=random.choice(DEPTS)
        hire=dt.date(random.randint(2012,2025), random.randint(1,12), random.randint(1,28))
        base=random.choice([2900000,3200000,3500000,3800000,4200000,4600000,5100000])
        base += random.randint(-150000,150000)
        oth=random.choice([0,0,4,8,10,12,16,18,22,26,30])
        otp = int(oth * random.uniform(22000,26000))
        if i in ot_outlier:            # 시간은 적은데 수당이 큼
            oth=random.choice([6,8,10]); otp=int(random.uniform(1600000,2400000))
        if i in sal_outlier:
            base=random.choice([8800000,9400000])
        if i in ghost: acct=ghost[i]
        else: acct=f"가상은행 {random.randint(1000,8999)}-{random.randint(10,99)}-{random.randint(100000,999999)}"
        rows.append([emp,dep,hire.isoformat(),base,oth,otp,acct,"2026-06-25"])
    with open(path,"w",encoding="utf-8-sig",newline="") as f:
        w=csv.writer(f)
        w.writerow(["employee_id","department","hire_date","base_salary","overtime_hours","overtime_pay","bank_account","payroll_date"])
        w.writerows(rows)
    print("SMP-3:", os.path.basename(path), len(rows),"rows (유령직원 2쌍·초과수당 이상 5·기본급 이상 2)")

# ── KIIA-5 경비 300행 XLSX (3시트) ──────────────────────────────
def kiia5():
    path=os.path.join(ROOT,"KIIA-5_5개부서_감사","expense_anomaly_data.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="경비지출_원본데이터"
    hdr=["지출번호","부서","거래처","지출일자","지출유형","금액","결재자","승인일자","비고"]
    styled_header(ws,hdr)
    vendors=["대한사무기기","한빛물류","미래ICT","우리케이터링","정직건설","세종문구","한결여행","다온전기"]
    etypes=["소모품","출장비","회의비","용역비","비품","유지보수"]
    approvers=["김결재","이승인","박전결","최부장","정팀장"]
    data=[]; num=1
    def wd(y,m,d): return dt.date(y,m,d)
    # 정상 240건 (총 300행 = 240 정상 + 60 이상)
    for _ in range(240):
        d=wd(2026,random.randint(1,6),random.randint(1,28))
        amt=random.choice([35000,68000,120000,250000,430000,620000,180000,95000,540000])+random.randint(-5000,5000)
        data.append([f"EXP-2026-{num:04d}",random.choice(DEPTS[:5]),random.choice(vendors),d.isoformat(),
                     random.choice(etypes),amt,random.choice(approvers),d.isoformat(),""]); num+=1
    # 금액이상(990,000 이상, 한도 1,000,000 회피) 18건 → 벤포드 9 과다
    for _ in range(18):
        d=wd(2026,random.randint(1,6),random.randint(1,28)); amt=random.choice([990000,995000,998000,992000,996000])
        data.append([f"EXP-2026-{num:04d}",random.choice(DEPTS[:5]),random.choice(vendors),d.isoformat(),
                     "용역비",amt,random.choice(approvers),d.isoformat(),""]); num+=1
    # 주말거래 12건
    for _ in range(12):
        base=wd(2026,random.randint(1,6),random.randint(1,22));
        while base.weekday()<5: base+=dt.timedelta(days=1)
        amt=random.choice([120000,260000,380000,540000])
        data.append([f"EXP-2026-{num:04d}",random.choice(DEPTS[:5]),random.choice(vendors),base.isoformat(),
                     "회의비",amt,random.choice(approvers),base.isoformat(),""]); num+=1
    # 결재자 없음 8건
    for _ in range(8):
        d=wd(2026,random.randint(1,6),random.randint(1,28)); amt=random.choice([210000,450000,330000])
        data.append([f"EXP-2026-{num:04d}",random.choice(DEPTS[:5]),random.choice(vendors),d.isoformat(),
                     "소모품",amt,"",d.isoformat(),""]); num+=1
    # 분할발주 5세트×3건=15(동일 거래처·같은 날, 합산 시 한도 초과)
    for _ in range(5):
        d=wd(2026,random.randint(1,6),random.randint(1,28)); v=random.choice(vendors)
        for _ in range(3):
            data.append([f"EXP-2026-{num:04d}",random.choice(DEPTS[:5]),v,d.isoformat(),"비품",
                         random.choice([480000,520000,610000]),random.choice(approvers),d.isoformat(),""]); num+=1
    random.shuffle(data)
    # 중복거래 7건: 기존 지출번호 재사용
    dup=random.sample(data,7)
    for r in dup:
        data.append([r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],"(중복 청구 의심)"])
    for r in data: ws.append(r)
    autofit(ws)
    # 이상거래_분석 시트
    ws2=wb.create_sheet("이상거래_분석")
    styled_header(ws2,["유형","판정 기준","건수","비고"])
    ws2.append(["분할발주","동일 거래처·같은 날 2건 이상","6세트","합산 시 한도 검토 필요"])
    ws2.append(["결재자없음","결재자 칸 공란",8,"승인 통제 미비"])
    ws2.append(["주말거래","토·일요일 지출",12,"업무 관련성 확인"])
    ws2.append(["금액이상","990,000원 이상(한도 1,000,000원 근접)",18,"한도 회피 의심"])
    ws2.append(["중복거래","동일 지출번호 재등장",5,"이중 청구 의심"])
    ws2.append([]); ws2.append([NOTICE])
    autofit(ws2)
    # 벤포드_분석 시트
    ws3=wb.create_sheet("벤포드_분석")
    styled_header(ws3,["첫째자리","벤포드 기대(%)","실제(%)","편차(%p)"])
    import math
    firsts=[int(str(r[5])[0]) for r in data if isinstance(r[5],int) and r[5]>0]
    n=len(firsts)
    for d in range(1,10):
        exp=math.log10(1+1/d)*100; act=firsts.count(d)/n*100
        ws3.append([d,round(exp,1),round(act,1),round(act-exp,1)])
    ws3.append([]); ws3.append(["※ 990,000원대 집중으로 첫째자리 9의 실제 비율이 기대치를 초과함"])
    autofit(ws3)
    wb.save(path)
    print("KIIA-5:", os.path.basename(path), len(data),"rows + 이상거래·벤포드 시트")

# ── KIIA-8 카드 80행 XLSX (2시트) ────────────────────────────────
def kiia8_card():
    path=os.path.join(ROOT,"KIIA-8_종합실습","corp_card_data.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="법인카드사용내역"
    hdr=["승인일시","사용자","부서","카드번호","가맹점명","업종","금액","승인번호","결재상태","사용목적","비고"]
    styled_header(ws,hdr)
    users=[f"김{c}" for c in "○"]+["이○○","박○○","최○○","정○○","강○○","조○○","윤○○"]
    users=["김○○","이○○","박○○","최○○","정○○","강○○","조○○","윤○○"]
    cards=["****-****-****-3045","****-****-****-7712","****-****-****-2290","****-****-****-8834"]
    normal_m=[("○○식당","음식점"),("△△마트","유통"),("□□주유소","주유"),("◇◇문구","사무"),
              ("☆☆카페","음식점"),("한빛서점","도서"),("미래택시","교통"),("우리사무기기","비품")]
    rows=[]; ap=1
    def stamp(day,hour,mn): return f"2026-04-{day:02d} {hour:02d}:{mn:02d}"
    for _ in range(66):  # 정상
        m=random.choice(normal_m); day=random.randint(1,28); hr=random.randint(9,20)
        rows.append([stamp(day,hr,random.randint(0,59)),random.choice(users),random.choice(DEPTS[:5]),
                     random.choice(cards),m[0],m[1],random.choice([12000,28000,45000,88000,150000,220000]),
                     f"A{ap:05d}",random.choice(["정상","정상","정상"]),"업무","" ]); ap+=1
    # 유흥주점 8건(정답 표시) + 심야
    bars=["블루문라운지","벨벳노래바","시티룸살롱","골드바","프리미엄단란주점"]
    for _ in range(8):
        day=random.randint(1,28); hr=random.choice([22,23,0,1,2])
        rows.append([stamp(day,hr,random.randint(0,59)),random.choice(users),random.choice(DEPTS[:5]),
                     random.choice(cards),random.choice(bars),"유흥",random.choice([180000,240000,320000,410000]),
                     f"A{ap:05d}","정상","접대","★유흥업소 사용"]); ap+=1
    # 고액·주말 6건
    for _ in range(6):
        base=dt.date(2026,4,random.randint(1,22))
        while base.weekday()<5: base+=dt.timedelta(days=1)
        rows.append([f"{base.isoformat()} {random.randint(11,16):02d}:{random.randint(0,59):02d}",random.choice(users),
                     random.choice(DEPTS[:5]),random.choice(cards),"백화점본점","백화점",random.choice([560000,720000,890000]),
                     f"A{ap:05d}","정상","비품","고액·주말 확인 필요"]); ap+=1
    random.shuffle(rows)
    for r in rows: ws.append(r)
    autofit(ws)
    ws2=wb.create_sheet("범례_및_분석지침")
    styled_header(ws2,["항목","설명"])
    for k,v in [("사용자","김○○ 형태로 비식별 처리"),("카드번호","****-****-****-뒤4자리 마스킹"),
                ("업종","음식점·유통·주유·유흥·백화점 등"),("비고 ★유흥업소 사용","탐지 정답 표시(실습 채점용)"),
                ("탐지 관점","부적절 업종(유흥), 심야·주말 사용, 고액, 특정 사용자 집중"),
                ("주의","비고의 정답 표시는 탐지 단계에서 가린 뒤 대조에 사용")]:
        ws2.append([k,v])
    ws2.append([]); ws2.append([NOTICE])
    autofit(ws2)
    wb.save(path)
    print("KIIA-8:", os.path.basename(path), len(rows),"rows (유흥 8건 정답표시)")

# ── KIIA-0 카드 444행 XLSX (대시보드용) ─────────────────────────
def kiia0_card():
    path=os.path.join(ROOT,"KIIA-0_카드_대시보드","corporate_card_transactions.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="카드사용내역"
    hdr=["거래ID","사용일시","사용자","부서","카드번호","가맹점명","업종","금액","승인상태","사용목적","비고"]
    styled_header(ws,hdr)
    users=["김○○","이○○","박○○","최○○","정○○","강○○","조○○","윤○○","임○○","한○○"]
    cards=["****-****-****-3045","****-****-****-7712","****-****-****-2290","****-****-****-8834","****-****-****-5561"]
    normal_m=[("○○식당","음식점"),("△△마트","유통"),("□□주유소","주유"),("◇◇문구","사무"),("☆☆카페","음식점"),
              ("미래택시","교통"),("우리사무기기","비품"),("한빛서점","도서"),("정성병원","의료")]
    rows=[]; tid=1
    def stamp(day,hr,mn): return f"2026-03-{day:02d} {hr:02d}:{mn:02d}"
    def row(day,hr,mn,mer,cat,amt,memo=""):
        nonlocal tid
        r=[f"T{tid:05d}",stamp(day,hr,mn),random.choice(users),random.choice(DEPTS[:6]),random.choice(cards),
           mer,cat,amt,"정상","업무",memo]; tid+=1; return r
    for _ in range(357):   # 정상 (총 444행)
        m=random.choice(normal_m); rows.append(row(random.randint(1,28),random.randint(9,20),random.randint(0,59),
            m[0],m[1],random.choice([9000,15000,32000,58000,90000,140000,230000])))
    # R1 심야 15
    for _ in range(15): rows.append(row(random.randint(1,28),random.choice([22,23,0,1,2,3]),random.randint(0,59),"심야식당","음식점",random.choice([45000,88000,120000]),"심야"))
    # R2 주말 20
    for _ in range(20):
        b=dt.date(2026,3,random.randint(1,22))
        while b.weekday()<5: b+=dt.timedelta(days=1)
        rows.append(row(b.day,random.randint(11,18),random.randint(0,59),"주말마트","유통",random.choice([60000,130000,210000]),"주말"))
    # R3 부적절 업종 14 (유흥·골프·면세·백화점)
    for m,c in [("로얄라운지","유흥"),("그린cc","골프"),("공항면세점","면세"),("제일백화점","백화점")]*3+[("시티룸살롱","유흥"),("스카이cc","골프")]:
        rows.append(row(random.randint(1,28),random.choice([13,19,21,22]),random.randint(0,59),m,c,random.choice([180000,350000,520000,760000]),"업종 확인"))
    # R4 분할결제 4세트(동일 사용자·가맹점 30분내 3건)
    for _ in range(4):
        u=random.choice(users); mer="분할상점"; day=random.randint(1,28); hr=random.randint(10,17); mn=random.randint(0,20)
        for k in range(3):
            r=[f"T{tid:05d}",stamp(day,hr,mn+k*8),u,random.choice(DEPTS[:6]),random.choice(cards),mer,"유통",290000,"정상","업무","분할결제 의심"]; tid+=1; rows.append(r)
    # R5 중복청구 4세트(동일 사용자·가맹점·금액 10분내)
    for _ in range(4):
        u=random.choice(users); mer="중복상점"; day=random.randint(1,28); hr=random.randint(10,17); mn=random.randint(0,40); amt=random.choice([120000,240000])
        for k in range(2):
            r=[f"T{tid:05d}",stamp(day,hr,mn+k*5),u,random.choice(DEPTS[:6]),random.choice(cards),mer,"음식점",amt,"정상","업무","중복청구 의심"]; tid+=1; rows.append(r)
    # R6 고액 12 (50만원+)
    for _ in range(12): rows.append(row(random.randint(1,28),random.randint(11,18),random.randint(0,59),"프리미엄전자","비품",random.choice([560000,680000,820000,1100000]),"고액"))
    # R7 한도임박 6 (한도 대비 95%+; 금액 950,000~990,000)
    for _ in range(6): rows.append(row(random.randint(1,28),random.randint(11,18),random.randint(0,59),"대형가전","비품",random.choice([950000,970000,985000]),"한도 임박"))
    random.shuffle(rows)
    for i,r in enumerate(rows,1): r[0]=f"T{i:05d}"; ws.append(r)
    autofit(ws)
    ws2=wb.create_sheet("안내")
    styled_header(ws2,["항목","설명"])
    for k,v in [("기간","2026년 3월"),("카드번호","****-****-****-뒤4자리 마스킹"),("사용자","김○○ 형태 비식별"),
                ("탐지 규칙","R1 심야·R2 주말·R3 부적절업종·R4 분할결제·R5 중복청구·R6 고액·R7 한도임박"),
                ("용도","PRD_card_audit_dashboard.md 규칙을 적용할 대시보드 실데이터")]:
        ws2.append([k,v])
    ws2.append([]); ws2.append([NOTICE]); autofit(ws2)
    wb.save(path)
    print("KIIA-0:", os.path.basename(path), len(rows),"rows (R1~R7 트리거 포함)")

# ── KIIA-6 부서정보 XLSX ────────────────────────────────────────
def kiia6_dept():
    path=os.path.join(ROOT,"KIIA-6_HWPX_양식","fake_audit_department_info.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="부서정보"
    styled_header(ws,["기관명","감사부서","담당자","직급","이메일","연락처"])
    data=[["한국디지털진흥원","감사실","홍○○","과장","audit1@example.org","02-000-3412"],
          ["대한에너지공사","감사부","김○○","차장","audit2@example.org","02-000-5521"],
          ["국립미래과학관","감사담당관실","이○○","사무관","audit3@example.org","02-000-1188"],
          ["한국복지자원공단","자체감사팀","박○○","팀장","audit4@example.org","02-000-9043"],
          ["중앙교통안전원","감사실","정○○","대리","audit5@example.org","02-000-2276"]]
    for r in data: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(path)
    print("KIIA-6:", os.path.basename(path), len(data),"기관")

# ── SMP-1 법인카드 사용내역 XLSX (표본 확대 ~32행) ──────────────
def smp1_card():
    path=os.path.join(ROOT,"SMP-1_법인카드_규정_사용내역","2025년_상반기_법인카드_사용내역_샘플.xlsx")
    wb=Workbook(); ws=wb.active; ws.title="사용내역"
    styled_header(ws,["사용일자","사용자","부서","사용처","금액","사용목적"])
    users=["김철수","이영희","최민지","박준호","정수아"]
    normal=[("정성식당","350000","팀 회식"),("한빛문구","82000","사무용품"),("미래주유소","110000","차량 주유"),
            ("우리호텔","220000","출장 숙박"),("대한서점","46000","도서 구입"),("세종택시","28000","시내 이동")]
    rows=[]
    for _ in range(26):
        d=dt.date(2025,random.randint(1,6),random.randint(1,28)); m=random.choice(normal)
        rows.append([d.isoformat(),random.choice(users),random.choice(DEPTS[:4]),m[0],int(m[1])+random.randint(-3000,3000),m[2]])
    # 점검 대상(용도 확인 필요) 6건
    rows.append([ "2025-03-15","김철수","영업팀","프리미엄골프장","480000","거래처 접대"])
    rows.append([ "2025-04-05","이영희","총무팀","백화점상품권","500000","명절 선물"])
    rows.append([ "2025-05-11","최민지","구매팀","주말리조트","390000","워크숍"])
    rows.append([ "2025-02-22","박준호","영업팀","고급주점","260000","거래처 접대"])
    rows.append([ "2025-06-08","정수아","인사팀","가전판매점","620000","비품 구입"])
    rows.append([ "2025-01-19","김철수","영업팀","출장호텔","540000","출장 숙박(2박)"])
    random.shuffle(rows)
    for r in rows: ws.append(r)
    ws.append([]); ws.append([NOTICE]); autofit(ws)
    wb.save(path)
    print("SMP-1:", os.path.basename(path), len(rows),"rows")

if __name__=="__main__":
    smp3(); kiia5(); kiia8_card(); kiia0_card(); kiia6_dept(); smp1_card()
    print("\n핵심 데이터 파일 생성 완료.")
