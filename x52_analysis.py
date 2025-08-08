#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 X52 계열 세아제강 제품 규격 내 품질별 항복강도 분포 분석
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# 한글 폰트 설정
import matplotlib.font_manager as fm
import matplotlib
font_path = "C:/Windows/Fonts/malgun.ttf"
font_prop = fm.FontProperties(fname=font_path)
matplotlib.rcParams['font.family'] = font_prop.get_name()
matplotlib.rcParams['axes.unicode_minus'] = False

# 시각화 스타일 설정
plt.style.use('default')
sns.set_palette("husl")

def load_and_analyze_data():
    """중경1공장 데이터 로드 및 기본 분석"""
    print("📊 중경1공장 데이터 로드 중...")
    
    # 필터링된 데이터 로드
    try:
        jg1_data = pd.read_excel('./첫시도/중경1공장_데이터_필터링.xlsx')
        print(f"✅ 필터링된 데이터 로드 성공: {jg1_data.shape}")
    except FileNotFoundError:
        # 원본 데이터에서 중경1공장 데이터 추출
        print("필터링된 데이터가 없어 원본 데이터에서 추출합니다...")
        data = pd.read_excel('./첫시도/joined_coil_jiwoong.xlsx')
        jg1_data = data[data['wc_desc'] == '중경1공장 20" 조관'].copy()
        print(f"✅ 원본에서 중경1공장 데이터 추출: {jg1_data.shape}")
    
    return jg1_data

def filter_x52_data(data):
    """X52 계열 데이터 필터링"""
    print("\n🔍 X52 계열 데이터 필터링 중...")
    
    # 품질 컬럼 확인
    if 'quality' in data.columns:
        quality_col = 'quality'
    elif 'grade' in data.columns:
        quality_col = 'grade'
    else:
        # 품질 관련 컬럼 찾기
        quality_cols = [col for col in data.columns if any(word in col.lower() for word in ['quality', 'grade', 'spec'])]
        if quality_cols:
            quality_col = quality_cols[0]
            print(f"품질 컬럼으로 '{quality_col}' 사용")
        else:
            print("❌ 품질 컬럼을 찾을 수 없습니다.")
            return None
    
    # 전체 품질 유형 확인
    print(f"\n📋 전체 품질 유형:")
    quality_counts = data[quality_col].value_counts()
    for i, (quality, count) in enumerate(quality_counts.items(), 1):
        print(f"{i:2d}. {quality}: {count:,}개")
    
    # X52 계열 데이터 필터링
    x52_mask = data[quality_col].str.contains('X52', case=False, na=False)
    x52_data = data[x52_mask].copy()
    
    print(f"\n✅ X52 계열 데이터 필터링 완료:")
    print(f"   전체 데이터: {len(data):,}개")
    print(f"   X52 계열: {len(x52_data):,}개 ({len(x52_data)/len(data)*100:.1f}%)")
    
    if len(x52_data) > 0:
        print(f"\n📊 X52 계열 품질 분포:")
        x52_qualities = x52_data[quality_col].value_counts()
        for quality, count in x52_qualities.items():
            print(f"   {quality}: {count:,}개")
    
    return x52_data, quality_col

def apply_seah_steel_specs(data, quality_col):
    """세아제강 X52 제품 규격 적용"""
    print("\n🎯 세아제강 X52 제품 규격 적용 중...")
    
    # X52 규격 (API 5L 기준)
    # 항복강도(YS): 최소 359 MPa (52,000 psi)
    # 인장강도(TS): 최소 455 MPa
    # 항복강도/인장강도 비율: 최대 0.93
    
    x52_specs = {
        'min_yield_strength': 359,  # MPa
        'min_tensile_strength': 455,  # MPa
        'max_ys_ts_ratio': 0.93
    }
    
    print(f"📋 X52 규격 기준:")
    print(f"   항복강도(YS): ≥ {x52_specs['min_yield_strength']} MPa")
    print(f"   인장강도(TS): ≥ {x52_specs['min_tensile_strength']} MPa")
    print(f"   YS/TS 비율: ≤ {x52_specs['max_ys_ts_ratio']}")
    
    # 항복강도 컬럼 찾기
    ys_cols = [col for col in data.columns if any(word in col.lower() for word in ['ys', 'yield', 'i_ys'])]
    ts_cols = [col for col in data.columns if any(word in col.lower() for word in ['ts', 'tensile', 'i_ts'])]
    
    print(f"\n🔍 발견된 강도 컬럼들:")
    print(f"   항복강도 관련: {ys_cols}")
    print(f"   인장강도 관련: {ts_cols}")
    
    if not ys_cols:
        print("❌ 항복강도 컬럼을 찾을 수 없습니다.")
        return data
    
    # 주요 항복강도 컬럼 선택 (i_ys 우선, 없으면 첫 번째)
    ys_col = 'i_ys' if 'i_ys' in ys_cols else ys_cols[0]
    print(f"✅ 항복강도 컬럼으로 '{ys_col}' 사용")
    
    # 규격 내 데이터 필터링
    original_count = len(data)
    
    # 항복강도 기준 적용
    spec_data = data[data[ys_col] >= x52_specs['min_yield_strength']].copy()
    
    # 인장강도 기준 적용 (컬럼이 있는 경우)
    if ts_cols:
        ts_col = 'i_ts' if 'i_ts' in ts_cols else ts_cols[0]
        print(f"✅ 인장강도 컬럼으로 '{ts_col}' 사용")
        spec_data = spec_data[spec_data[ts_col] >= x52_specs['min_tensile_strength']]
        
        # YS/TS 비율 확인
        spec_data['ys_ts_ratio'] = spec_data[ys_col] / spec_data[ts_col]
        spec_data = spec_data[spec_data['ys_ts_ratio'] <= x52_specs['max_ys_ts_ratio']]
    
    print(f"\n✅ 세아제강 규격 적용 완료:")
    print(f"   필터링 전: {original_count:,}개")
    print(f"   규격 내: {len(spec_data):,}개 ({len(spec_data)/original_count*100:.1f}%)")
    
    if len(spec_data) > 0:
        print(f"\n📊 규격 내 품질 분포:")
        spec_qualities = spec_data[quality_col].value_counts()
        for quality, count in spec_qualities.items():
            print(f"   {quality}: {count:,}개")
            
        # 항복강도 통계
        print(f"\n📈 규격 내 항복강도({ys_col}) 통계:")
        print(spec_data[ys_col].describe().round(1))
    
    return spec_data, ys_col

def create_scatterplot(data, quality_col, ys_col):
    """품질별 항복강도 분포 scatterplot 생성"""
    print(f"\n🎨 품질별 항복강도 분포 Scatterplot 생성 중...")
    
    if len(data) == 0:
        print("❌ 데이터가 없어 그래프를 생성할 수 없습니다.")
        return
    
    # 그래프 크기 설정
    plt.figure(figsize=(14, 8))
    
    # 품질별로 다른 색상과 마커 사용
    qualities = data[quality_col].unique()
    colors = plt.cm.Set1(np.linspace(0, 1, len(qualities)))
    markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h']
    
    for i, quality in enumerate(qualities):
        quality_data = data[data[quality_col] == quality]
        
        # 산점도 그리기
        plt.scatter(
            range(len(quality_data)), 
            quality_data[ys_col],
            c=[colors[i]], 
            marker=markers[i % len(markers)],
            s=60, 
            alpha=0.7, 
            label=f'{quality} (n={len(quality_data)})',
            edgecolors='black',
            linewidth=0.5
        )
        
        # 품질별 평균선 추가
        mean_ys = quality_data[ys_col].mean()
        plt.axhline(y=mean_ys, color=colors[i], linestyle='--', alpha=0.8, linewidth=1.5)
    
    # X52 최소 규격선 추가
    plt.axhline(y=359, color='red', linestyle='-', linewidth=2, alpha=0.8, label='X52 최소 규격 (359 MPa)')
    
    # 그래프 꾸미기
    plt.title('중경1공장: X52 계열 세아제강 제품 규격 내\n품질별 항복강도 분포 (Scatterplot)', 
              fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('데이터 순서', fontsize=12, fontweight='bold')
    plt.ylabel('항복강도 (MPa)', fontsize=12, fontweight='bold')
    
    # 격자 추가
    plt.grid(True, alpha=0.3)
    
    # 범례 설정
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    
    # Y축 범위 설정 (최소값보다 약간 아래부터)
    y_min = data[ys_col].min() - 10
    y_max = data[ys_col].max() + 10
    plt.ylim(y_min, y_max)
    
    # 레이아웃 조정
    plt.tight_layout()
    
    # 저장
    filename = '중경1공장_X52계열_세아제강규격내_품질별_항복강도분포_scatterplot.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    print(f"💾 그래프를 '{filename}'로 저장했습니다.")
    
    plt.show()
    
    # 통계 정보 출력
    print(f"\n📊 품질별 항복강도 통계:")
    stats = data.groupby(quality_col)[ys_col].agg(['count', 'mean', 'std', 'min', 'max']).round(1)
    print(stats)

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 X52 계열 세아제강 제품 분석 시작")
    print("=" * 80)
    
    # 1. 데이터 로드
    jg1_data = load_and_analyze_data()
    
    # 2. X52 계열 데이터 필터링
    x52_data, quality_col = filter_x52_data(jg1_data)
    
    if x52_data is None or len(x52_data) == 0:
        print("❌ X52 계열 데이터가 없습니다.")
        return
    
    # 3. 세아제강 규격 적용
    spec_data, ys_col = apply_seah_steel_specs(x52_data, quality_col)
    
    if len(spec_data) == 0:
        print("❌ 세아제강 규격을 만족하는 데이터가 없습니다.")
        return
    
    # 4. Scatterplot 생성
    create_scatterplot(spec_data, quality_col, ys_col)
    
    print("\n✅ 분석 완료!")
    print("=" * 80)

if __name__ == "__main__":
    main()
