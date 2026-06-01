# 공공데이터 "데이터 지도" 거버넌스·인텔리전스 탐사

> 공공데이터포털 **999개 데이터셋**(교육·보건·공공행정)의 프로파일링 산출물("데이터 지도")을
> **비지도 데이터마이닝**으로 분석하여, 데이터 거버넌스 이슈와 비즈니스 인텔리전스("주말 > 기저귀 > 맥주"형 규칙)를 발굴하는 프로젝트.

분석 결과는 모두 **설명 가능한 알고리즘 실행 결과**이며(LLM 추론 배제), 스크립트로 100% 재현됩니다.

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
5. 나머지 규칙은 `proj1.png`, `proj2.png` 참조. *(⚠️ 현재 두 파일은 내용이 전부 0인 손상 파일 — 재전달 필요)*

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

## 주요 발견

### 🛡️ 거버넌스
- **출처 결손의 계통성** — `유아및초·중등교육` 데이터셋 **100건이 100% 출처 placeholder('소스')** (lift 6.89, conf 1.00). 무작위가 아닌 단일 수집배치 손실.
- **죽은 메트릭** — `null_ratio·table_qty_index·cum_search_num`이 9,840개 컬럼 전부 0(미충전).
- **품질 플래그** — 상수 컬럼(distinct=1) 1,372개(13.9%), 스키마 회계 불일치 테이블 56개.
- **메타 네이밍 오류** — `same_ratio`는 비율이 아니라 평균중복도(rec/distinct).

### 💡 인텔리전스 ("주말 > 기저귀 > 맥주" 대응 규칙)
| 규칙 | conf | lift |
|---|---:|---:|
| 경도 → 위도 (3개 도메인 공통 POI 지문) | 1.00 | 22.7 |
| 의료기관전화번호 → {의료기관명, 주소} (재식별 위험 묶음) | 0.88 | 46.4 |
| 교과목코드 ↔ 분반 | 1.00 | 124.9 |
| 시도명 → 시군구명 (행정구역 계층) | 0.94 | 39.0 |
| 연령대 → 성별 | 0.85 | 35.2 |

- **스키마 패밀리(최대 빈발 컬럼군)**: 사학재정군·의료기관군·POI군·국가시험군 등 = 데이터셋 원형.
- **규칙 #4 실증**: `소재지도로명→지번`이 공공행정 국소 conf 1.00 → 전체 0.44로 약화(비일반성 포착).

> 전체 근거: [`analysis/I1_REPORT.md`](analysis/I1_REPORT.md) · [`analysis/CANDIDATES.md`](analysis/CANDIDATES.md)

---

## 저장소 구조
```
.
├── README.md
├── data/                         # 원천 CSV (대용량, gitignore 권장)
├── scripts/                      # 재현용 분석 스크립트 (실행 순서대로 번호)
│   ├── 01_profile.py             # 스키마·인코딩·행수 프로파일
│   ├── 02_encoding_probe.py      # 인코딩 확정(UTF-8)
│   ├── 03_dump_headers.py        # 전체 헤더 덤프
│   ├── 04_basic_analysis.py      # 기초 분석 + 1D 이산화 + 거버넌스 퀵스캔
│   ├── 05_scale_and_crossdomain.py  # 대용량 규모 + 클러스터 트리 + 교차도메인 값
│   ├── 06_i1_assoc_rules.py      # I1: FP-Growth + 연관규칙 + 국소/전체 검증
│   └── 07_i1_synthesis.py        # I1: 보강 근거(패밀리·추적성)
└── analysis/                     # 산출 리포트 (UTF-8 Markdown)
    ├── CANDIDATES.md             # 거버넌스/인텔리전스 이슈 후보 1차
    ├── basic_analysis.md         # 기초 분석 리포트
    ├── scale_crossdomain.md      # 규모 & 교차도메인 스캔
    ├── headers.md                # 전체 컬럼 사전
    ├── I1_REPORT.md              # ⭐ I1 근거중심 리포트 (메인 산출물)
    ├── I1_association_rules.md   # I1 원시 규칙표
    ├── I1_synthesis.md           # I1 보강 근거
    └── THEORY_이론배경.md         # 이론 학습 자료
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
```
각 스크립트는 `analysis/`에 Markdown 리포트를 생성합니다.

> **인코딩 주의**: 데이터는 UTF-8입니다. Windows 콘솔(CP949)에서 한글이 깨져 보일 수 있으므로 결과는 항상 UTF-8 파일로 출력해 확인합니다.

---

## 한계 & 주의
- 컬럼개념은 결정적 정규화만 적용 — **동의어 병합(연번≈순번)은 미적용**(BI 규칙 오염 방지; 별도 표준어 군집으로 처리 권장).
- 장바구니 규칙의 support는 낮음(8~44 테이블) → lift가 높아도 희소하므로 운영 적용 전 도메인 확인 필요.
- `table_qty_index·cum_search_num·null_ratio`가 전부 0(미충전)이라 활용/품질 축 분석 불가.
- `distinct_value.csv`는 정렬되어 있어 앞부분 샘플이 편향 → 전체검증 필수.
- `proj1.png`/`proj2.png`는 손상(전체 0바이트 내용) → 원칙 #5의 "나머지 규칙" 미반영.

---

## 이론 배경 자료
데이터마이닝 입문자를 위한 친절한 설명: [`analysis/THEORY_이론배경.md`](analysis/THEORY_이론배경.md)
(이산화 · AOI · 빈발패턴/FP-Growth · 연관규칙/lift · 국소-전체 검증 · 서브스페이스 클러스터링)

### 다음 단계 후보
- **I3** — POI/의료기관 패밀리를 `combined_pair_map.linking_ratio`로 실제 조인키 검증
- **G3** — 동의어(연번/순번/번호 …) 문자유사도 군집 → 표준어 사전
- **I5** — 컬럼 프로파일 서브스페이스 클러스터링으로 컬럼 원형 정교화
