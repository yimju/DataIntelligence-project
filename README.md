# 공공데이터 "데이터 지도" 거버넌스·인텔리전스 탐사

> 공공데이터포털 **999개 데이터셋**(교육·보건·공공행정)의 프로파일링 산출물("데이터 지도")을
> **비지도 데이터마이닝**으로 분석하여, 데이터 거버넌스 이슈와 비즈니스 인텔리전스("주말 > 기저귀 > 맥주"형 규칙)를 발굴하는 프로젝트.

분석 결과는 모두 **설명 가능한 알고리즘 실행 결과**이며(LLM 추론 배제), 스크립트로 100% 재현됩니다.

> 📊 **최신 보고서**: [`report/데이터지도_분석보고서_V21.pptx`](report/데이터지도_분석보고서_V21.pptx) (**15장**, 16:9). 슬라이드별 레이아웃·버전 히스토리(V1→V21)는 [`report/PPT_슬라이드_구성안.md`](report/PPT_슬라이드_구성안.md).
>
> 🤝 **다른 PC/세션에서 작업을 이어받는다면** → [`CLAUDE.md`](CLAUDE.md)부터 읽으세요(환경 제약·PPT 안전 편집 방법론·슬라이드↔파일 매핑·버전 히스토리). ⚠️ PPT 버전은 PowerPoint 재저장 시 파트가 재정규화되므로, 편집 전 항상 **ElementTree로 position↔file을 재확인**.

---

## 📌 목차
- [배경 & 목표](#배경--목표)
- [분석 원칙](#분석-원칙)
- [데이터 구성](#데이터-구성)
- [분석 방법론](#분석-방법론)
- [주요 발견](#주요-발견)
- [저장소 구조](#저장소-구조)
- [실행 방법](#실행-방법)
- [한계 & 주의](#한계--주의)
- [이론 배경 자료](#이론-배경-자료)

---

## 배경 & 목표
대상 데이터는 원본 트랜잭션이 아니라, 1,000개 공공데이터셋을 프로파일링한 **메타데이터(데이터 지도)** 입니다.
- 컬럼/테이블별 통계(고유값수·키품질·결측 등)
- 쌍별 값·도메인 중첩(컬럼쌍 524만·테이블쌍 25만·연결쌍 74만)
- 값/이름 유사도 기반 계층 클러스터(덴드로그램)

**목표**: 이 데이터 지도에서 ① 데이터 거버넌스 이슈와 ② 흥미로운 비즈니스 인텔리전스(메타 수준의 연관규칙)를 도출.

---

## 분석 원칙
1. **LLM 능력으로 분석 금지** — 모든 결론은 비지도 알고리즘 실행 결과. 과정이 **설명 가능**해야 함.
2. **1D 클러스터링 · AOI · 빈발패턴 · 연관규칙 · 서브스페이스 클러스터링** 활용. 수치형은 **이산화/클러스터링**.
3. **데이터 롤업** 적극 활용.
4. **국소적으로 찾은 규칙은 전체에서 검증.**
5. 과제 지침은 `proj1.PNG`(비지도 only — AOI·1D/Subspace 군집·빈발패턴·연관·순차패턴, 최소 2기법 연계+다계층), `proj2.PNG`(컬럼 도메인/값 거버넌스 + 단일/연계 테이블 인사이트) 참조. *(현재 두 이미지는 정상 판독됨)*

---

## 데이터 구성
`data/` 폴더(CSV, UTF-8). 대용량 파일은 저장소에 포함되지 않을 수 있음(`.gitignore` 권장).

| 파일 | 역할 | 규모 |
|---|---|---|
| `table_map.csv` | 데이터셋 1건당 메타데이터 + 분류(cate1/cate2) + 통계 | 999행 |
| `column_map.csv` | 컬럼 1건당 프로파일(num_cat_flag, distinct_cnt, pk_ratio, same_ratio …) | 9,840행 |
| `distinct_value.csv` | 컬럼별 distinct 값 + 빈도 (값 수준) | 803만행 · 635MB |
| `table_summary_map.csv` / `column_summary_map.csv` | 테이블/컬럼별 타 객체와의 중첩 통계 롤업 | 996 / 8,360행 |
| `combined_pair_map.csv` | 컬럼 쌍의 조인 가능성(linking_ratio) | 74만행 · 300MB |
| `table_pair_map.csv` | 테이블 쌍 유사도(sim_col_ratio, col_val_sim …) | 25만행 · 296MB |
| `column_pair_map.csv` | 컬럼 쌍 상세 값/도메인 중첩·중복 | 524만행 · 3.4GB |
| `*_value_tree.csv` / `*_str_tree.csv` | 값/이름 유사도 계층 클러스터(덴드로그램) | 331~4,549행 |

도메인 분포: **교육 400 · 보건 300 · 공공행정 299**.

---

## 분석 방법론
```
원자료 ─► [이산화] ─► [AOI 일반화/롤업] ─► [빈발패턴] ─► [연관규칙] ─► [국소→전체 검증]
        (1D KMeans)   (개념 계층)        (FP-Growth)   (lift 등)     (규칙 #4)
```
- **이산화**: 멱법칙 수치(rec_cnt 등)에 `log1p` 후 **1D KMeans** → 순서형 구간(극소~극대).
- **AOI 롤업**: `table_source`(999개 기관명) → 7개 기관유형, `col_nm` → 정규화 컬럼개념.
- **빈발패턴/연관규칙**: 테이블=트랜잭션, 속성·컬럼=아이템. `mlxtend` **FP-Growth** + `association_rules`(support·confidence·**lift**·leverage·conviction).
- **검증**: 도메인별(국소) 채굴 → 전체 999건에서 재계산.

자세한 단계별 설명은 [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md) §0.1, 이론은 [`analysis/THEORY_이론배경.md`](analysis/THEORY_이론배경.md).

---

## 주요 발견 (최신 보고서 V21 기준)

### 🛡️ 거버넌스 — 출처 결손의 계통성 (보고서 11·12p)
- `table_source` placeholder('소스') 145건(14.5%) 중 **‘유아및초·중등교육’ 분류 100건이 100%** (연관규칙 conf 1.00 · lift 6.89).
- **단일 적재 배치 확증**: regi_date 적재 순서 **499~598번(연속 100건)이 전부 유아및초중등**(타 분류 0건 혼입) → 한 수집 배치에서 출처가 통째로 소실. (`url`도 999건 전부 `'url'` placeholder = 카탈로그 전반 별도 이슈)
- **전체 도메인 비교**: 교육 33%(131) ≫ 보건 3% ≫ 공공행정 1% — 출처 누락은 교육(특히 유아및초중등)에 계통 집중(전체 평균 대비 극단 outlier).

### 💡 인텔리전스
- **흥미로운 연관규칙 — 상세 지표 (9p)** · "주말 > 기저귀 > 맥주" 대응:

  | 규칙 | conf | lift |
  |---|---:|---:|
  | 경도 → 위도 (POI 좌표쌍) | 1.00 | 22.7 |
  | 교과목코드 ↔ 분반 | 1.00 | 124.9 |
  | 의료기관전화번호 → {의료기관명, 주소} | 0.88 | 46.4 |
  | 시도명 → 시군구명 (행정구역 계층) | 0.94 | 39.0 |
  | 연령대 → 성별 | 0.85 | 35.2 |

- **탐사 대상 규칙 — 적절한 lift 구간 (10p)**: lift가 매우 높은 자명 규칙(예 교과목코드↔분반 125)은 단일 템플릿 내 정의적 동반 → **적절 구간(lift≈4~50)의 비자명 규칙**(출판사→저자·확진자수→사망자수·유아및초중등→출처미상 등)을 탐사 대상으로 선정.
- **교차도메인 데이터 연계 기회 (13p)**: 의료기관 패밀리(명·주소·전화 항상 동반)는 **기관 정보(개인정보 아님)** → 공통 키로 데이터셋을 **연계(JOIN)** 할 자산. (재식별 위험이 아닌 연계 기회로 해석)
- **보건 협소 테이블 → 데이터 병합 기회 (14p)**: ‘작고 키 있는’ 보건 데이터셋이 동일 스키마로 산재 — 컬럼 집합 **Jaccard≥0.8 동형 41쌍** → 지자체별 조각을 **UNION 병합**해 전국 마스터로 통합(예: 의료기관 현황 4개 지자체 → 1,450행).

> 전체 근거: 보고서 [`report/데이터지도_분석보고서_V21.pptx`](report/데이터지도_분석보고서_V21.pptx) · [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md) · [`analysis/CANDIDATES.md`](analysis/CANDIDATES.md)
>
> 참고: 상수 컬럼 1,372개(13.9%)·스키마 회계 불일치 56개·`same_ratio` 네이밍 오류 등 컬럼 품질 이슈는 [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md)에 별도 정리(현재 보고서에서는 비핵심으로 분리).

---

## 저장소 구조
```
.
├── README.md · CLAUDE.md · 프로젝트_핵심정리.md
├── proj1.PNG / proj2.PNG          # 과제 지침 이미지
├── data/                          # 원천 CSV (대용량, gitignore)
├── scripts/                       # 재현용 스크립트 (번호 순)
│   ├── 01~05_*.py                 # 프로파일·인코딩·헤더·기초분석(1D 이산화)·규모/교차도메인
│   ├── 06_i1_assoc_rules.py       # I1: FP-Growth + 연관규칙 + 국소/전체 검증
│   ├── 07_i1_synthesis.py         # I1: 보강 근거(스키마 패밀리·추적성)
│   ├── 08_build_pptx.py           # 보고서 PPT 최초 생성(그리드/색상 코드화)
│   ├── 09_metrics_and_figs.py     # 상세 지표 + 차트(figs)
│   └── 10~29_*.py                 # PPT 버전 변환기(V2→V21, ZIP/XML 안전 편집)
├── report/                        # ⭐ 보고서 산출물
│   ├── 데이터지도_분석보고서_V21.pptx  # 최신본(15장)  · V12~V20 = 버전 보존
│   ├── PPT_슬라이드_구성안.md          # 슬라이드 레이아웃 주석 + 버전 히스토리
│   └── figs/                       # 차트 PNG(도메인·이산화·AOI·lift·산점·규칙#4)
└── analysis/                      # 근거 리포트 (UTF-8 Markdown)
    ├── I1_REPORT.md               # ⭐ I1 근거중심 리포트 (방법론 §0.1 포함)
    ├── I1_metrics_detail.md / .json  # 규칙 상세 지표 + Appendix
    ├── I1_association_rules.md · I1_synthesis.md  # 원시 규칙표·패밀리
    ├── CANDIDATES.md · basic_analysis.md · scale_crossdomain.md · headers.md
    └── THEORY_이론배경.md          # 이론 학습 자료(이산화·AOI·FP-Growth·연관규칙·Jaccard 등)
```

---

## 실행 방법

### 요구 환경
- Python 3.10+ (개발: Anaconda Python 3.13)
- 패키지: `pandas`, `numpy`, `scikit-learn`, `scipy`, `mlxtend`

```bash
pip install pandas numpy scikit-learn scipy mlxtend
```

### 실행
스크립트는 번호 순서대로 실행합니다. (Windows에서 KMeans 메모리 경고 억제를 위해 `OMP_NUM_THREADS=4` 권장)

```bash
# 데이터는 ./data 에 위치
python scripts/01_profile.py
python scripts/04_basic_analysis.py
python scripts/05_scale_and_crossdomain.py
OMP_NUM_THREADS=4 python scripts/06_i1_assoc_rules.py
OMP_NUM_THREADS=4 python scripts/07_i1_synthesis.py
# 보고서 PPT (선택): python-pptx 필요
OMP_NUM_THREADS=4 python scripts/09_metrics_and_figs.py   # 상세 지표 + figs
python scripts/08_build_pptx.py                            # V1 PPT 생성
```
01~07은 `analysis/`에 Markdown 리포트를, 08·09는 `report/`에 PPT/차트를 생성합니다.
보고서 버전 변환기(`10~29_*.py`)는 **직전 버전 .pptx를 base로** ZIP/XML을 안전 편집(문자열·도형 복제)해 다음 버전을 만듭니다 — 상세는 [`report/PPT_슬라이드_구성안.md`](report/PPT_슬라이드_구성안.md).

> **인코딩 주의**: 데이터는 UTF-8입니다. Windows 콘솔(CP949)에서 한글이 깨져 보일 수 있으므로 결과는 항상 UTF-8 파일로 출력해 확인합니다.

---

## 한계 & 주의
- 컬럼개념은 결정적 정규화만 적용 — **동의어 병합(연번≈순번)은 미적용**(BI 규칙 오염 방지; 별도 표준어 군집으로 처리 권장).
- 장바구니 규칙의 support는 낮음(8~44 테이블) → lift가 높아도 희소하므로 운영 적용 전 도메인 확인 필요.
- `table_qty_index·cum_search_num·null_ratio`가 전부 0(미충전)이라 활용/품질 축 분석 불가.
- `distinct_value.csv`는 정렬되어 있어 앞부분 샘플이 편향 → 전체검증 필수.
- `regi_date`가 999건 전부 `2026-03-30`(일괄 적재)이라 시계열·이력 분석 불가. `url`은 전건 placeholder.

---

## 이론 배경 자료
데이터마이닝 입문자를 위한 친절한 설명: [`analysis/THEORY_이론배경.md`](analysis/THEORY_이론배경.md)
(이산화 · AOI · 빈발패턴/FP-Growth · 연관규칙/lift · 국소-전체 검증 · 서브스페이스 클러스터링)

### 다음 단계 후보
- **연계(JOIN)** — 의료기관·POI 패밀리를 `combined_pair_map.linking_ratio`로 실제 조인키 검증
- **병합(UNION)** — 동형(Jaccard≥0.8) 그룹을 ‘출처’ 컬럼 추가 후 UNION → 전국 마스터 자동 생성
- **G3 표준어 사전** — 동의어 컬럼(연번/순번/번호 …) 문자유사도 군집 → 표준 컬럼명
- **I5 서브스페이스 클러스터링** — 컬럼 프로파일 다차원에서 컬럼 원형 정교화
