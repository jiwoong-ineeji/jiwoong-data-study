#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 품질별 YS2_STRESS 분포 stripplot 생성
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import warnings
import matplotlib.font_manager as fm
import platform
import os

warnings.filterwarnings('ignore')

def clear_matplotlib_cache():
    """matplotlib 캐시 완전 초기화"""
    try:
        import matplotlib
        matplotlib.rcdefaults()  # 기본 설정으로 리셋
        
        # 폰트 캐시 삭제
        cache_dir = matplotlib.get_cachedir()
        if os.path.exists(cache_dir):
            import shutil
            try:
                shutil.rmtree(cache_dir)
                print("🗑️ matplotlib 캐시 디렉토리 삭제 완료")
            except:
                pass
        
        # 폰트 매니저 재구축
        fm._rebuild()
        print("🔄 matplotlib 폰트 매니저 재구축 완료")
        
    except Exception as e:
        print(f"⚠️ 캐시 초기화 중 오류: {e}")

def setup_korean_font_robust():
    """강력한 한글 폰트 설정"""
    
    print("🔧 한글 폰트 설정 시작...")
    
    # 1. 캐시 초기화
    clear_matplotlib_cache()
    
    # 2. 시스템별 폰트 경로 설정
    if platform.system() == "Windows":
        font_candidates = [
            ("Malgun Gothic", "C:/Windows/Fonts/malgun.ttf"),
            ("NanumGothic", "C:/Windows/Fonts/NanumGothic.otf"),
            ("Gulim", "C:/Windows/Fonts/gulim.ttc"),
            ("Batang", "C:/Windows/Fonts/batang.ttc"),
            ("HCR Dotum", "C:/Windows/Fonts/HANDotum.ttf")
        ]
    else:
        font_candidates = [
            ("NanumGothic", "/usr/share/fonts/truetype/nanum/NanumGothic.ttf"),
            ("DejaVu Sans", None)
        ]
    
    # 3. 폰트 설정 시도
    font_success = False
    
    for font_name, font_path in font_candidates:
        try:
            if font_path and os.path.exists(font_path):
                # 직접 파일 경로로 폰트 등록
                font_prop = fm.FontProperties(fname=font_path)
                plt.rcParams['font.family'] = font_prop.get_name()
                print(f"✅ 폰트 파일로 설정 성공: {font_name} ({font_path})")
                font_success = True
                break
            else:
                # 시스템 폰트명으로 설정
                plt.rcParams['font.family'] = font_name
                print(f"✅ 시스템 폰트로 설정 성공: {font_name}")
                font_success = True
                break
                
        except Exception as e:
            print(f"❌ {font_name} 설정 실패: {e}")
            continue
    
    # 4. 기본 설정
    plt.rcParams['axes.unicode_minus'] = False
    plt.rcParams['font.size'] = 10
    
    if not font_success:
        print("⚠️ 모든 한글 폰트 설정 실패. 기본 폰트 사용")
    
    return font_success

def load_data():
    """데이터 로드"""
    print("\n📂 데이터 로드 중...")
    try:
        data = pd.read_excel('중경1공장_데이터_필터링완료.xlsx')
        print(f"✅ 데이터 로드 성공: {data.shape}")
        return data
    except Exception as e:
        print(f"❌ 데이터 로드 실패: {e}")
        return None

def create_ys2_stress_stripplot(data):
    """품질별 YS2_STRESS 분포 stripplot 생성"""
    
    if data is None or len(data) == 0:
        print("❌ 데이터가 없습니다.")
        return None
    
    print("\n🎨 YS2_STRESS stripplot 생성 중...")
    
    # 컬럼 확인
    quality_col = 'p_spec'
    ys2_stress_col = 'ys2_stress'
    
    if ys2_stress_col not in data.columns:
        print(f"❌ {ys2_stress_col} 컬럼을 찾을 수 없습니다.")
        print(f"사용 가능한 컬럼: {list(data.columns)}")
        return None
    
    # YS2_STRESS 데이터 확인
    print(f"\n📊 YS2_STRESS 데이터 분석:")
    print(f"   총 데이터: {len(data):,}개")
    print(f"   유효한 YS2_STRESS 값: {data[ys2_stress_col].notna().sum():,}개")
    print(f"   YS2_STRESS 범위: {data[ys2_stress_col].min():.2f} ~ {data[ys2_stress_col].max():.2f} MPa")
    print(f"   YS2_STRESS 평균: {data[ys2_stress_col].mean():.2f} MPa")
    
    # 상위 5개 품질 선정
    top_qualities = data[quality_col].value_counts().head(5)
    filtered_data = data[data[quality_col].isin(top_qualities.index)].copy()
    
    print(f"\n📋 상위 5개 품질별 YS2_STRESS 분포:")
    for i, (quality, count) in enumerate(top_qualities.items(), 1):
        quality_ys2_data = filtered_data[filtered_data[quality_col] == quality][ys2_stress_col]
        mean_ys2 = quality_ys2_data.mean()
        std_ys2 = quality_ys2_data.std()
        print(f"   {i}. {quality}: {count:,}개 (평균: {mean_ys2:.2f}±{std_ys2:.2f} MPa)")
    
    # 그래프 생성
    plt.figure(figsize=(14, 8))
    
    # stripplot 생성
    ax = sns.stripplot(
        data=filtered_data,
        x=quality_col,
        y=ys2_stress_col,
        order=top_qualities.index,
        size=8,
        alpha=0.7,
        jitter=True,
        linewidth=0.5,
        edgecolor='black'
    )
    
    # 각 품질별 평균선 추가
    for i, quality in enumerate(top_qualities.index):
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        mean_ys2_stress = quality_data[ys2_stress_col].mean()
        
        # 평균선
        ax.hlines(
            y=mean_ys2_stress,
            xmin=i-0.4,
            xmax=i+0.4,
            colors='red',
            linestyles='--',
            linewidth=2,
            alpha=0.8
        )
        
        # 평균값 텍스트
        ax.text(
            i, mean_ys2_stress + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.02,
            f'평균: {mean_ys2_stress:.1f}',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
            color='red'
        )
    
    # 제목 및 레이블 (한글 사용)
    plt.title(
        '중경1공장 필터링된 데이터: 상위 5개 품질별 YS2_STRESS 분포 (Stripplot)',
        fontsize=16,
        fontweight='bold',
        pad=20
    )
    plt.xlabel('품질 (Quality)', fontsize=12, fontweight='bold')
    plt.ylabel('YS2_STRESS (MPa)', fontsize=12, fontweight='bold')
    
    # X축 레이블 회전
    plt.xticks(rotation=45, ha='right')
    
    # 격자
    plt.grid(True, alpha=0.3, axis='y')
    
    # 품질별 개수 및 통계 정보 박스 (오른쪽 위)
    info_text = "품질별 YS2_STRESS 정보:\n"
    for quality in top_qualities.index:
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        count = len(quality_data)
        mean_val = quality_data[ys2_stress_col].mean()
        info_text += f"{quality}: {count:,}개 (평균: {mean_val:.1f})\n"
    
    total_count = sum(top_qualities.values)
    overall_mean = filtered_data[ys2_stress_col].mean()
    info_text += f"\n총계: {total_count:,}개\n전체 평균: {overall_mean:.1f} MPa"
    
    # 정보 박스 스타일
    bbox_props = dict(
        boxstyle="round,pad=0.5",
        facecolor="lightgreen",
        alpha=0.8,
        edgecolor="darkgreen",
        linewidth=1
    )
    
    # 오른쪽 위에 정보 표시
    plt.text(
        0.98, 0.98,
        info_text,
        transform=ax.transAxes,
        fontsize=9,
        verticalalignment='top',
        horizontalalignment='right',
        bbox=bbox_props,
        fontweight='bold'
    )
    
    # 레이아웃 조정
    plt.tight_layout()
    
    # 저장
    filename = '중경1공장_상위5개품질_YS2_STRESS분포_stripplot.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"💾 그래프 저장 완료: {filename}")
    
    plt.show()
    
    # 상세 통계 출력
    print(f"\n📊 품질별 YS2_STRESS 상세 통계:")
    stats = filtered_data.groupby(quality_col)[ys2_stress_col].agg([
        'count', 'mean', 'std', 'min', 'max', 
        lambda x: x.quantile(0.25), lambda x: x.quantile(0.75)
    ]).round(2)
    stats.columns = ['개수', '평균', '표준편차', '최소값', '최대값', '25%', '75%']
    print(stats)
    
    return filename

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 품질별 YS2_STRESS 분포 stripplot 생성")
    print("=" * 80)
    
    # 1. 한글 폰트 설정
    font_success = setup_korean_font_robust()
    
    # 2. 데이터 로드
    data = load_data()
    
    # 3. YS2_STRESS stripplot 생성
    if data is not None:
        filename = create_ys2_stress_stripplot(data)
        if filename:
            print(f"\n✅ 작업 완료! 생성된 파일: {filename}")
        else:
            print("\n❌ 그래프 생성 실패")
    else:
        print("\n❌ 데이터 로드 실패")
    
    print("=" * 80)

if __name__ == "__main__":
    main()
