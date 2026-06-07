# CLAUDE.md — 작업 인수인계 가이드 (다른 PC의 Claude Code용)

> 이 파일은 Claude Code가 시작 시 자동으로 읽습니다. **다른 PC에서 작업을 이어받을 때 이 문서를 먼저 보세요.**
> 기준 시점: **V12가 최신 보고서**. 이 프로젝트는 공공데이터 ‘데이터 지도’ 거버넌스·인텔리전스 분석 + 그 결과 PPT를 반복 개선해 왔습니다.

---

## 0. 프로젝트 한 줄 요약
공공데이터포털 999개 데이터셋의 프로파일링(‘데이터 지도’)을 **비지도 데이터마이닝**으로 분석해 거버넌스 이슈·BI 규칙을 도출하고, 그 결과를 PPT 보고서로 만든 프로젝트. 전체 내용은 [`프로젝트_핵심정리.md`](프로젝트_핵심정리.md)와 [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md) 참조. PPT 구성·버전 히스토리는 [`report/PPT_슬라이드_구성안.md`](report/PPT_슬라이드_구성안.md)에 상세.

---

## 1. 현재 상태 (V12 기준)
- **최신 보고서**: [`report/데이터지도_분석보고서_V12.pptx`](report/데이터지도_분석보고서_V12.pptx) · **20장** · 16:9
- V1(원본)→V12까지 한 번에 한두 슬라이드씩 개선/삭제해 옴. **각 버전은 직전 버전 pptx를 base로** 변환 스크립트(`scripts/10~20`)가 생성. 다음 작업은 **V12를 base로 V13**을 만드는 식.
- ⚠️ 사용자가 V12를 **PowerPoint로 열어 직접 편집**하기도 함(슬라이드 8·10 일부는 사용자 수정본). 즉 **V12의 일부 슬라이드는 PowerPoint가 재저장**해 런(run)이 분할되어 있을 수 있음. 새 작업 전 **항상 대상 슬라이드의 현재 구조를 먼저 점검**할 것.

---

## 2. 환경 (다른 PC에서 매우 중요)
이 작업이 진행된 PC는 과학계산 라이브러리가 깨져 있어 **PPT 편집을 전부 순수 stdlib로** 했습니다. 다른 PC도 같은 제약일 수 있으니 먼저 확인하세요.

- **Python 실행기**: `python`/`py`는 Microsoft Store 스텁(실행 시 종료, exit 49)일 수 있음. 실제로는 **`C:\ProgramData\anaconda3\python.exe`**(base, Python 3.8) 사용. 다른 PC면 동등한 Python 경로를 먼저 찾을 것.
- **깨져 있던 것(중요)**: `numpy`·`matplotlib`·`sklearn`(DLL load failed: `_multiarray_umath`), `lxml.etree`(DLL), `python-pptx`(미설치, 오프라인 pip는 SSL 실패). conda env(`hoon`,`mypy39`)도 numpy 불가.
  - 결과 ① **scripts/06·09 재실행 불가**(KMeans/FP-Growth/matplotlib 필요) → 분석 수치는 **`analysis/*.md`·`*.json`에 이미 산출된 값**을 쓰거나, 단순 집계는 **stdlib(csv)로 직접 재계산**.
  - 결과 ② **python-pptx 사용 불가** → PPT 편집은 **zipfile + re + xml.etree.ElementTree(stdlib)**로 ZIP/XML 직접 편집.
  - 결과 ③ **raster 차트 생성 불가**(matplotlib) → 차트는 **PowerPoint 네이티브 도형(사각형 막대 등)으로 작도**.
- **셸**: Bash 도구가 가장 안정적. 한글 깨짐 방지로 출력은 `... 2>&1 | tr -d '\0'`, 인코딩은 `PYTHONIOENCODING=utf-8` 지정. PowerShell은 `$env:PYTHONIOENCODING="utf-8"` 필요하고 가끔 `EPERM uv_spawn` 발생(재시도 or Bash 사용).
- **렌더 확인 불가(이 PC)**: PowerPoint/LibreOffice headless 없음 → 레이아웃은 **좌표 계산으로 검증**하고, 시각 확인은 **사용자에게 열어보게 요청**. (사용자 PC엔 PowerPoint 있음.)

```bash
# 예: stdlib로 데이터 집계
PYTHONIOENCODING=utf-8 "/c/ProgramData/anaconda3/python.exe" - <<'PY' 2>&1 | tr -d '\0'
import csv, io
rows=list(csv.DictReader(io.open(r'd:\DataInt_2\data\column_map.csv',encoding='utf-8-sig')))
print(len(rows))
PY
```

---

## 3. ★ PPT 편집 방법론 (반드시 지킬 규칙)
PowerPoint가 비표준 XML을 ‘복구(repair)’하며 **도형을 삭제**하는 사고가 있었음(V3 콜아웃·캡션 소실). 아래 규칙을 지켜야 안전.

1. **직전 버전 base + 문자열 편집(기존 바이트 보존)**: `SRC=V_{n-1}.pptx`를 zip으로 읽고, **대상 슬라이드 XML만** 문자열/정규식으로 수정 후 다시 zip(`DST=V_{n}.pptx`). 나머지 파트는 그대로 재기록 → 대상 외 슬라이드는 byte-동일. **슬라이드 전체를 ElementTree로 재직렬화하지 말 것**(V3가 이걸로 repair 유발). ET는 **읽기/검사용**으로만.
2. **새 도형은 python-pptx 산출 구조를 그대로 복제**:
   - 채운 도형(rect/roundRect)은 **`<p:style>` 블록 포함 + 빈 txBody**(`<a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p>`). **텍스트는 절대 채운 도형 안에 넣지 말고 별도 텍스트박스**(txBox="1")로. (V3 콜아웃이 이걸 어겨 삭제됨.)
   - 텍스트박스 bodyPr: `wrap="square"` + `<a:normAutofit/>`.
   - 검증된 템플릿/헬퍼는 `scripts/15~20`의 `roundrect()/tbox()/frame()/_para()/STYLE/EMPTY_TX` 그대로 복붙해 쓰면 됨.
3. **도형은 텍스트가 아니라 좌표로 식별**: 한글은 PowerPoint 맞춤법검사로 여러 run으로 쪼개져 **부분문자열 매칭이 실패**함. `<a:off ... y="...">`로 식별:
   - 섹션태그 `y=109728` · 제목 `y=365760` · accent rule `y=1060704` · 코어/배너영역 `y≈1170432~1280432`(1.28″) · **본문 `y≥1828800`(2.0″)** · 푸터 페이지번호 `y=6437376`.
   - 2×2 쿼드(본문) 제거: `<p:sp>` 중 첫 `<a:off y>`가 **[1800000, 6000000)** 인 것 삭제.
4. **텍스트 XML 이스케이프**: run 텍스트의 `& < >`를 `&amp; &lt; &gt;`로(리터럴 `&`가 V10에서 XML을 깨뜨림 → `esc()` 추가).
5. **한글 잘림 방지(no-clip)**: 모든 `<p:txBody>`의 `<a:bodyPr>`에 `wrap="square"` + autofit(spAutoFit 유지/없으면 normAutofit, **noAutofit 금지**). 편집 마지막에 `fix_clipping()` 적용(scripts/14~20에 동일 함수 있음).
6. **그리드·색·폰트 규약**: EMU=inch×914400. 좌우 여백 0.62″, 본문 2.00~6.95″. 색 `NAVY=081F46`·`GREEN=00965F`·본문 `282D37`·연한배경 `EEF2F8`/`E8F5F0`·라인 `D0D8E4`·강조빨강 `C0392B`, 폰트 `맑은 고딕`. 제목 26pt·**메트릭 수치 36pt(가장 큼)**·카드 헤더 10pt(흰)·본문 8.5~9pt. 거버넌스=Navy/인텔리전스·검증=Green.
7. **이슈 슬라이드 카드 패턴**: 좌상/좌하/우상/우하 4카드(=roundRect 배경 + 헤더띠 roundRect + 흰 제목 텍스트박스 + 본문 텍스트박스). 상단 **메트릭 배너**(1.22~1.96″): ‘발견 패턴 ▸ …’ + 큰 수치(36pt) + 설명.

### 슬라이드 1장 삭제 절차(V11에서 사용)
`scripts/19_v11_delete_g2.py` 참고. ① `[Content_Types].xml`의 해당 Override ② `ppt/presentation.xml`의 `<p:sldId>` ③ `ppt/_rels/presentation.xml.rels`의 Relationship 제거 + slide 파트/그 `.rels` 미기록, ④ **이후 슬라이드 푸터 페이지번호 재번호**(y=6437376 텍스트 −1), ⑤ 교차참조(종합·로드맵) 정리. `docProps/app.xml`(슬라이드 수)은 PowerPoint가 저장 시 자동보정하므로 **건드리지 않음**(잘못 건드리면 repair 위험).

---

## 4. ★ 슬라이드 ↔ 파일 매핑 (V12, 함정 주의)
G2 슬라이드를 V11에서 삭제하면서 **파일명과 페이지 위치가 어긋남**. “페이지 N”을 편집하려면 아래 **파일**을 열 것(순서는 `ppt/presentation.xml`의 sldIdLst가 결정).

| 페이지 | 파일 | 제목 |
|---:|---|---|
| 1 | `slide1.xml` | 표지 |
| 2 | `slide2.xml` | 분석 개요 및 목적 |
| 3 | `slide3.xml` | 대상 데이터 구성 |
| 4 | `slide4.xml` | 분석 방법론 파이프라인 |
| 5 | `slide5.xml` | 공통 준비 ① 적재·인코딩 |
| 6 | `slide6.xml` | 공통 준비 ② 1D 이산화 |
| 7 | `slide7.xml` | 공통 준비 ③ AOI 롤업 |
| 8 | `slide8.xml` | 공통 준비 ④ 트랜잭션(장바구니 예시) |
| 9 | `slide9.xml` | 분석 결과 종합 |
| 10 | `slide10.xml` | 거버넌스 G1 출처 결손 |
| **11** | **`slide12.xml`** | 거버넌스 G3 컬럼 표준·품질 *(slide11.xml은 삭제됨!)* |
| 12 | `slide13.xml` | 인텔리전스 스키마 패밀리 |
| 13 | `slide14.xml` | 연관규칙 상세 지표 |
| 14 | `slide15.xml` | 연관규칙 시각화 |
| 15 | `slide16.xml` | 교차도메인 재식별 위험 |
| 16 | `slide17.xml` | 규칙#4 검증 |
| 17 | `slide18.xml` | Appendix A |
| 18 | `slide19.xml` | Appendix B |
| 19 | `slide20.xml` | 종합 시사점·로드맵 |
| 20 | `slide21.xml` | 다음 단계·한계 |

> 매핑은 다음으로 항상 재확인 가능:
> ```bash
> PYTHONIOENCODING=utf-8 "/c/ProgramData/anaconda3/python.exe" - <<'PY' 2>&1 | tr -d '\0'
> import zipfile,re
> z=zipfile.ZipFile(r'd:\DataInt_2\report\데이터지도_분석보고서_V12.pptx')
> rels=z.read('ppt/_rels/presentation.xml.rels').decode()
> rid2f={m[0]:m[1] for m in re.findall(r'Id="(rId\d+)"[^>]*Target="(slides/slide\d+\.xml)"',rels)}
> pres=z.read('ppt/presentation.xml').decode()
> for i,rid in enumerate(re.findall(r'<p:sldId id="\d+" r:id="(rId\d+)"/>',pres),1): print(i,rid2f[rid])
> PY
> ```

---

## 5. 버전 히스토리 & 변환 스크립트
각 변환 스크립트가 “무엇을 바꿨는지”의 1차 기록. 상세 변경노트는 [`report/PPT_슬라이드_구성안.md`](report/PPT_슬라이드_구성안.md) 상단.

| 버전 | 스크립트 | 변경 요지 |
|---|---|---|
| V1 | `08_build_pptx.py` | 원본 21장 생성(python-pptx, 다른 PC에서 만든 것) |
| V2 | `10_make_v2_headfont.py` | 헤드(핵심)메시지 15pt→18pt |
| V3 | `11_v3_slide6_explain.py` | 6p 멱법칙 평이화+콜아웃 *(⚠️ ET 재직렬화로 PowerPoint repair 유발 — 반면교사)* |
| V4 | `12_v4_disc_charts.py` | 6p 이산화 3변수 막대차트 + **전 슬라이드 한글잘림 방지** |
| V5 | `13_v5_slide7_rollup.py` | 7p AOI 롤업 3단계 과정 도표 |
| V6 | `14_v6_slide8_basket_and_noclip.py` | 8p 트랜잭션=장바구니 예시 도표 |
| V7 | `15_v7_g1_improve.py` | 10p G1: placeholder·A1 용어, 예시데이터, 의미분석 |
| V8 | `16_v8_metric_banners.py` | 이슈/인텔 슬라이드 상단 ‘발견 패턴+핵심수치(36pt)’ 배너 |
| V9 | `17_v9_a1_explain.py` | 10p A1 규칙 = 어떤 아이템 간 트랜잭션(도식) |
| V10 | `18_v10_g2_improve.py` | (구)11p G2 개선 *(이후 V11에서 G2 삭제)* |
| V11 | `19_v11_delete_g2.py` | **G2 슬라이드 삭제**(21→20장)+재번호+교차참조 정리 |
| V12 | `20_v12_g3_improve.py` | 11p G3: ‘동일 개념’ 수치적 의미·절차·예시·의미분석 |

분석 단계 스크립트: `01_profile`~`05_scale_and_crossdomain`(프로파일/기초/규모), `06_i1_assoc_rules`(FP-Growth·연관규칙), `07_i1_synthesis`, `09_metrics_and_figs`(지표/차트). **06·09는 numpy/sklearn/mlxtend/matplotlib 필요 → 이 PC에선 재실행 불가**(산출물 `analysis/*.md`·`I1_metrics.json` 사용).

---

## 6. 데이터 & 분석 산출물
- 원천: `data/` (UTF-8 CSV) — `table_map.csv`(999), `column_map.csv`(9,840), `distinct_value.csv`(803만), 대용량 pair맵 등. 컬럼사전은 [`analysis/headers.md`](analysis/headers.md).
  - ⚠️ **`data/`는 `.gitignore`로 제외**되어 git에 없음 → 다른 PC에서 **원천 데이터로 재계산하려면 `data/` 폴더를 별도 전송**(USB/클라우드). 단, PPT 편집·수치 인용은 대개 **`analysis/*.md`·`I1_metrics.json`(커밋됨)** 으로 충분.
- 분석 리포트: [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md)(⭐메인), `CANDIDATES.md`, `basic_analysis.md`, `scale_crossdomain.md`, `I1_metrics_detail.md`, `I1_synthesis.md`, `THEORY_이론배경.md`.
- 핵심 수치(자주 씀): 도메인 교육400·보건300·공공행정299 / 이산화 3변수(레코드5·컬럼폭4·수치비3) / AOI 7유형(scripts/06 규칙: 기초404·공공245·미상145·광역83·교육49·중앙39·기타34) / 상수컬럼 1,372개(13.9%) / 스키마 불일치 56 / 연관규칙 lift 최대 124.9(교과목코드↔분반). **수치 인용 전 데이터/리포트로 재확인 권장**(특히 scripts/09와 06의 AOI 정규식이 달라 그림과 §0.2 수치가 어긋났던 전례 있음 → 분석 근거는 **06 기준**).

---

## 7. 새 버전(V13) 만드는 표준 절차
1. **현재 구조 점검**: 대상 페이지의 파일(§4 매핑!)을 ET로 읽어 도형 name/offset/텍스트 확인. 사용자가 PowerPoint로 수정해 런이 쪼개졌는지도 확인.
2. **필요 수치 데이터로 계산**(stdlib csv) 또는 `analysis/`에서 인용.
3. **변환 스크립트 작성** `scripts/21_v13_*.py`: `SRC=V12`, `DST=V13`. §3 규칙 준수(문자열 편집·검증 도형구조·좌표 식별·esc·fix_clipping). 헬퍼는 `scripts/20`에서 복붙.
4. **생성 후 검증**(아래 체크리스트).
5. **문서 갱신**: `report/PPT_슬라이드_구성안.md`에 V13 생성물·변환기·변경노트 추가. 필요시 이 `CLAUDE.md`의 ‘현재 상태/매핑/히스토리’ 갱신.
6. 사용자에게 **해당 페이지를 PowerPoint로 열어 (a)복구 프롬프트 없이 열리는지 (b)레이아웃** 확인 요청.

### 검증 체크리스트 (편집 후 매번)
```python
# 1) 대상 슬라이드만 변경됐나 (직전 버전과 diff)
# 2) ET.fromstring 으로 well-formed
# 3) cNvPr id 유일
# 4) <a:noAutofit/> 0개 & 모든 <p:txBody> bodyPr에 wrap="square"+autofit
# 5) 모든 도형이 슬라이드 경계 내(13.333in×7.5in = 12192000×6858000 EMU)
# 6) 핵심 텍스트가 들어갔는지(부분문자열 검사)
# 7) (삭제 작업이면) 슬라이드 수·sldId·Content_Types·rels 정합(전부 동일 개수) + 페이지번호 연속
```

---

## 8. 한계·주의
- **시각 검증을 이 PC에서 못 함** → 좌표로만 확인. 카드/도형이 빽빽하면 normAutofit으로 글자가 줄어들 수 있음(잘림은 없음). 필요시 사용자 확인 후 좌표·폰트 조정.
- **이전 버전 pptx는 보존**(롤백용). 임시 점검 스크립트는 `_`로 시작하게 만들고 작업 후 삭제(저장소 정리).
- `proj1.PNG`/`proj2.PNG` = 과제 지침 이미지(정상 판독됨; 과거 일부 문서엔 ‘손상’으로 잘못 기록). 지침 요약은 `프로젝트_핵심정리.md` §0.
