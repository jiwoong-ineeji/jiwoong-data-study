#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 필터링된 데이터: 상위 5개 품질별 두께 stripplot 생성
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# 한글 폰트 설정 (확실한 방법)
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import matplotlib
import os

def setup_korean_font():
    """한글 폰트 설정 - 더 확실한 방법"""
    
    # matplotlib 폰트 캐시 초기화
    try:
        fm._rebuild()
        print("📝 matplotlib 폰트 캐시 초기화 완료")
    except:
        pass
    
    # 시도할 한글 폰트들 (영문명 사용)
    korean_fonts = [
        'Malgun Gothic',  # 맑은 고딕
        'NanumGothic',    # 나눔고딕
        'Gulim',          # 굴림
        'Batang',         # 바탕
        'HCR Dotum'       # HCR 돋움
    ]
    
    font_set = False
    for font_name in korean_fonts:
        try:
            # 폰트 설정 시도
            plt.rcParams['font.family'] = font_name
            plt.rcParams['axes.unicode_minus'] = False
            
            # 한글 테스트
            fig, ax = plt.subplots(figsize=(1, 1))
            ax.text(0.5, 0.5, '테스트', fontsize=10)
            fig.canvas.draw()
            plt.close(fig)
            
            print(f"✅ 한글 폰트 설정 성공: {font_name}")
            font_set = True
            break
            
        except Exception as e:
            print(f"❌ {font_name} 설정 실패: {str(e)}")
            continue
    
    if not font_set:
        # 직접 폰트 파일 경로 지정
        font_paths = [
            "C:/Windows/Fonts/malgun.ttf",
            "C:/Windows/Fonts/NanumGothic.otf", 
            "C:/Windows/Fonts/gulim.ttc",
            "C:/Windows/Fonts/batang.ttc"
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    font_prop = fm.FontProperties(fname=font_path)
                    plt.rcParams['font.family'] = font_prop.get_name()
                    plt.rcParams['axes.unicode_minus'] = False
                    print(f"✅ 직접 경로로 폰트 설정 성공: {font_path}")
                    font_set = True
                    break
                except Exception as e:
                    print(f"❌ {font_path} 직접 설정 실패: {str(e)}")
                    continue
    
    if not font_set:
        print("⚠️ 모든 한글 폰트 설정 실패. 기본 설정을 사용합니다.")
        plt.rcParams['axes.unicode_minus'] = False
    
    return font_set

# 한글 폰트 설정 실행
setup_korean_font()

# 시각화 스타일 설정
plt.style.use('default')
sns.set_palette("husl")

def load_filtered_data():
    """필터링된 중경1공장 데이터 로드"""
    print("📂 필터링된 중경1공장 데이터 로드 중...")
    
    input_file = '중경1공장_데이터_필터링완료.xlsx'
    
    try:
        data = pd.read_excel(input_file)
        print(f"✅ 데이터 로드 완료: {data.shape}")
        return data
    except FileNotFoundError:
        print(f"❌ 파일을 찾을 수 없습니다: {input_file}")
        return None
    except Exception as e:
        print(f"❌ 데이터 로드 오류: {str(e)}")
        return None

def identify_thickness_column(data):
    """두께 관련 컬럼 식별"""
    print("\n🔍 두께 관련 컬럼 찾는 중...")
    
    # 두께 관련 키워드
    thickness_keywords = ['thick', '두께', 'thk', 'p_thick']
    thickness_cols = []
    
    for col in data.columns:
        if any(keyword in col.lower() for keyword in thickness_keywords):
            thickness_cols.append(col)
    
    print(f"📋 발견된 두께 관련 컬럼: {thickness_cols}")
    
    if not thickness_cols:
        print("❌ 두께 관련 컬럼을 찾을 수 없습니다.")
        return None
    
    # p_thick_mm 우선, 없으면 첫 번째 컬럼 사용
    thickness_col = 'p_thick_mm' if 'p_thick_mm' in thickness_cols else thickness_cols[0]
    print(f"✅ 두께 컬럼으로 '{thickness_col}' 사용")
    
    return thickness_col

def get_top_qualities(data, quality_col, top_n=5):
    """상위 N개 품질 식별"""
    print(f"\n📊 상위 {top_n}개 품질 식별 중...")
    
    quality_counts = data[quality_col].value_counts()
    top_qualities = quality_counts.head(top_n)
    
    print(f"✅ 상위 {top_n}개 품질:")
    for i, (quality, count) in enumerate(top_qualities.items(), 1):
        percentage = (count / len(data)) * 100
        print(f"   {i}. {quality}: {count:,}개 ({percentage:.1f}%)")
    
    return top_qualities

def create_thickness_stripplot(data, quality_col, thickness_col, top_qualities):
    """품질별 두께 stripplot 생성"""
    print(f"\n🎨 품질별 두께 stripplot 생성 중...")
    
    # 상위 품질만 필터링
    top_quality_names = top_qualities.index.tolist()
    filtered_data = data[data[quality_col].isin(top_quality_names)].copy()
    
    print(f"   필터링된 데이터: {len(filtered_data):,}개")
    
    if len(filtered_data) == 0:
        print("❌ 그래프를 그릴 데이터가 없습니다.")
        return
    
    # 품질 순서를 데이터 개수 순으로 정렬
    quality_order = top_quality_names
    
    # 그래프 크기 설정
    plt.figure(figsize=(14, 8))
    
    # stripplot 생성
    ax = sns.stripplot(
        data=filtered_data,
        x=quality_col,
        y=thickness_col,
        order=quality_order,
        size=8,
        alpha=0.7,
        jitter=True,
        linewidth=0.5,
        edgecolor='black'
    )
    
    # 각 품질별 평균선 추가
    for i, quality in enumerate(quality_order):
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        mean_thickness = quality_data[thickness_col].mean()
        
        # 평균선 그리기
        ax.hlines(
            y=mean_thickness,
            xmin=i-0.4,
            xmax=i+0.4,
            colors='red',
            linestyles='--',
            linewidth=2,
            alpha=0.8
        )
        
        # 평균값 텍스트 표시
        ax.text(
            i, mean_thickness + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.02,
            f'평균: {mean_thickness:.2f}',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
            color='red'
        )
    
    # 제목 및 축 레이블
    plt.title(
        '중경1공장 필터링된 데이터: 상위 5개 품질별 두께 분포 (Stripplot)',
        fontsize=16,
        fontweight='bold',
        pad=20
    )
    plt.xlabel('품질 (Quality)', fontsize=12, fontweight='bold')
    plt.ylabel('두께 (mm)', fontsize=12, fontweight='bold')
    
    # X축 레이블 회전
    plt.xticks(rotation=45, ha='right')
    
    # 격자 추가
    plt.grid(True, alpha=0.3, axis='y')
    
    # 품질별 개수 정보를 오른쪽 위에 표시
    info_text = "품질별 데이터 개수:\n"
    for quality, count in top_qualities.items():
        info_text += f"{quality}: {count:,}개\n"
    
    # 총 개수도 추가
    total_count = sum(top_qualities.values)
    info_text += f"\n총계: {total_count:,}개"
    
    # 텍스트 박스 스타일
    bbox_props = dict(
        boxstyle="round,pad=0.5",
        facecolor="lightblue",
        alpha=0.8,
        edgecolor="navy",
        linewidth=1
    )
    
    # 오른쪽 위에 정보 표시
    plt.text(
        0.98, 0.98,
        info_text,
        transform=ax.transAxes,
        fontsize=10,
        verticalalignment='top',
        horizontalalignment='right',
        bbox=bbox_props,
        fontweight='bold'
    )
    
    # 레이아웃 조정
    plt.tight_layout()
    
    # 저장
    filename = '중경1공장_상위5개품질_두께분포_stripplot.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight')
    print(f"💾 그래프를 '{filename}'로 저장했습니다.")
    
    plt.show()
    
    # 통계 정보 출력
    print(f"\n📊 품질별 두께 통계:")
    thickness_stats = filtered_data.groupby(quality_col)[thickness_col].agg([
        'count', 'mean', 'std', 'min', 'max'
    ]).round(3)
    print(thickness_stats)
    
    return filename

def analyze_thickness_distribution(data, quality_col, thickness_col, top_qualities):
    """두께 분포 상세 분석"""
    print(f"\n📈 두께 분포 상세 분석:")
    
    top_quality_names = top_qualities.index.tolist()
    filtered_data = data[data[quality_col].isin(top_quality_names)].copy()
    
    print(f"\n1️⃣ 전체 두께 분포:")
    overall_stats = filtered_data[thickness_col].describe()
    print(f"   평균: {overall_stats['mean']:.3f} mm")
    print(f"   표준편차: {overall_stats['std']:.3f} mm")
    print(f"   최소값: {overall_stats['min']:.3f} mm")
    print(f"   최대값: {overall_stats['max']:.3f} mm")
    print(f"   범위: {overall_stats['max'] - overall_stats['min']:.3f} mm")
    
    print(f"\n2️⃣ 품질별 두께 특성:")
    for quality in top_quality_names:
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        stats = quality_data[thickness_col].describe()
        
        print(f"\n   📋 {quality} ({len(quality_data)}개):")
        print(f"      평균: {stats['mean']:.3f} mm")
        print(f"      표준편차: {stats['std']:.3f} mm")
        print(f"      범위: {stats['min']:.3f} ~ {stats['max']:.3f} mm")
        print(f"      변동계수: {(stats['std']/stats['mean']*100):.1f}%")

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 상위 5개 품질별 두께 stripplot 생성 시작")
    print("=" * 80)
    
    # 1. 데이터 로드
    data = load_filtered_data()
    if data is None:
        return
    
    # 2. 두께 컬럼 식별
    thickness_col = identify_thickness_column(data)
    if thickness_col is None:
        return
    
    # 3. 품질 컬럼 확인
    quality_col = 'p_spec'  # 이전 분석에서 확인된 품질 컬럼
    if quality_col not in data.columns:
        print(f"❌ 품질 컬럼 '{quality_col}'을 찾을 수 없습니다.")
        return
    
    # 4. 상위 5개 품질 식별
    top_qualities = get_top_qualities(data, quality_col, top_n=5)
    
    # 5. stripplot 생성
    filename = create_thickness_stripplot(data, quality_col, thickness_col, top_qualities)
    
    # 6. 상세 분석
    analyze_thickness_distribution(data, quality_col, thickness_col, top_qualities)
    
    print(f"\n✅ stripplot 생성 완료!")
    print(f"   저장된 파일: {filename}")
    print("=" * 80)

if __name__ == "__main__":
    main()
