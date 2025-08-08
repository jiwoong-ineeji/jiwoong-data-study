# 중경2공장 새 프로젝트

## 프로젝트 개요
중경2공장의 품질 데이터를 분석하고 시각화하는 새로운 프로젝트입니다.

## 프로젝트 구조
```
중경2공장_새프로젝트/
├── analysis.py          # 메인 분석 스크립트
├── README.md           # 프로젝트 설명서
├── output/             # 분석 결과 저장 폴더
└── data/               # 데이터 파일 폴더 (필요시 생성)
```

## 주요 기능

### JungKyung2FactoryAnalyzer 클래스
- **데이터 로드**: Excel, CSV 파일 지원
- **데이터 개요**: 기본 통계, 결측값 확인
- **품질별 분석**: 품질 카테고리별 통계 분석
- **시각화**: 다양한 플롯 타입 지원 (scatter, box, violin, strip)
- **상관관계 분석**: 상관계수 히트맵 생성

### 지원하는 시각화 타입
1. **Scatter Plot**: 두 연속형 변수 간의 관계
2. **Box Plot**: 카테고리별 분포 비교
3. **Violin Plot**: 분포의 밀도까지 확인
4. **Strip Plot**: 개별 데이터 포인트 표시

## 사용 방법

### 1. 기본 실행
```bash
uv run analysis.py
```

### 2. 대화형 분석
```python
from analysis import JungKyung2FactoryAnalyzer

# 분석기 초기화
analyzer = JungKyung2FactoryAnalyzer()

# 데이터 로드
data = analyzer.load_data("your_data_file.xlsx")

# 데이터 개요 확인
analyzer.data_overview()

# 품질별 분석
analyzer.quality_analysis(quality_col='품질')

# 시각화 생성
analyzer.create_visualization('두께', '인장강도', hue_col='품질', plot_type='scatter')

# 상관관계 분석
corr_matrix = analyzer.correlation_analysis()
```

## 데이터 요구사항
- Excel (.xlsx, .xls) 또는 CSV 파일
- 한글 컬럼명 지원
- 품질, 두께, 인장강도, 항복강도 등의 컬럼 포함 권장

## 출력 결과
모든 분석 결과와 시각화는 `output/` 폴더에 저장됩니다:
- PNG 형식의 그래프 파일
- 고해상도 (300 DPI) 이미지
- 한글 폰트 지원

## 다음 단계
1. 실제 데이터 파일 경로 확인
2. 컬럼명에 맞게 분석 함수 호출
3. 필요한 추가 분석 기능 구현
4. 결과 해석 및 보고서 작성
