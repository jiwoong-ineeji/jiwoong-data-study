#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 X52 계열 YS2_STRESS 360~530 MPa 필터링된 데이터 stripplot
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
        try:
            fm._rebuild()
            print("🔄 matplotlib 폰트 매니저 재구축 완료")
        except:
            print("⚠️ 폰트 매니저 재구축 스킵")
        
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

def filter_ys2_stress_range(data, min_ys2=360, max_ys2=530):
    """YS2_STRESS 범위로 데이터 필터링"""
    
    print(f"\n🔍 YS2_STRESS {min_ys2}~{max_ys2} MPa 범위로 필터링 중...")
    
    ys2_col = 'ys2_stress'
    
    if ys2_col not in data.columns:
        print(f"❌ {ys2_col} 컬럼을 찾을 수 없습니다.")
        return None
    
    # 필터링 전 현황
    print(f"📊 필터링 전 YS2_STRESS 현황:")
    print(f"   전체 데이터: {len(data):,}개")
    print(f"   YS2_STRESS 범위: {data[ys2_col].min():.2f} ~ {data[ys2_col].max():.2f} MPa")
    print(f"   YS2_STRESS 평균: {data[ys2_col].mean():.2f} MPa")
    
    # 범위 필터링
    filtered_data = data[(data[ys2_col] >= min_ys2) & (data[ys2_col] <= max_ys2)].copy()
    
    # 필터링 후 현황
    print(f"\n✅ YS2_STRESS 범위 필터링 완료:")
    print(f"   필터링 전: {len(data):,}개")
    print(f"   필터링 후: {len(filtered_data):,}개")
    print(f"   제거된 데이터: {len(data) - len(filtered_data):,}개 ({(len(data) - len(filtered_data))/len(data)*100:.1f}%)")
    print(f"   남은 데이터: {len(filtered_data)/len(data)*100:.1f}%")
    
    if len(filtered_data) > 0:
        print(f"   필터링 후 YS2_STRESS 범위: {filtered_data[ys2_col].min():.2f} ~ {filtered_data[ys2_col].max():.2f} MPa")
        print(f"   필터링 후 YS2_STRESS 평균: {filtered_data[ys2_col].mean():.2f} MPa")
    
    return filtered_data

def filter_x52_grades(data):
    """X52 계열 품질만 추출"""
    
    print(f"\n🔍 X52 계열 품질 필터링 중...")
    
    quality_col = 'p_spec'
    
    if quality_col not in data.columns:
        print(f"❌ {quality_col} 컬럼을 찾을 수 없습니다.")
        return None
    
    # 전체 품질 확인
    print(f"📋 전체 품질 분포:")
    all_qualities = data[quality_col].value_counts()
    for quality, count in all_qualities.items():
        print(f"   {quality}: {count:,}개")
    
    # X52 계열 필터링 (대소문자 무관, 부분 문자열 포함)
    x52_mask = data[quality_col].str.contains('X52', case=False, na=False)
    x52_data = data[x52_mask].copy()
    
    print(f"\n✅ X52 계열 필터링 완료:")
    print(f"   전체 데이터: {len(data):,}개")
    print(f"   X52 계열: {len(x52_data):,}개 ({len(x52_data)/len(data)*100:.1f}%)")
    
    if len(x52_data) > 0:
        print(f"\n📊 X52 계열 품질 분포:")
        x52_qualities = x52_data[quality_col].value_counts()
        for quality, count in x52_qualities.items():
            print(f"   {quality}: {count:,}개")
    else:
        print("❌ X52 계열 데이터가 없습니다.")
    
    return x52_data

def create_x52_ys2_stress_stripplot(data):
    """X52 계열 상위 5개 품질별 YS2_STRESS stripplot 생성"""
    
    if data is None or len(data) == 0:
        print("❌ 데이터가 없습니다.")
        return None
    
    print("\n🎨 X52 계열 YS2_STRESS stripplot 생성 중...")
    
    # 컬럼 설정
    quality_col = 'p_spec'
    ys2_col = 'ys2_stress'
    
    # 상위 5개 X52 품질 선정
    top_x52_qualities = data[quality_col].value_counts().head(5)
    
    if len(top_x52_qualities) == 0:
        print("❌ X52 계열 품질이 없습니다.")
        return None
    
    print(f"\n📋 상위 5개 X52 품질별 YS2_STRESS 분포:")
    for i, (quality, count) in enumerate(top_x52_qualities.items(), 1):
        quality_ys2_data = data[data[quality_col] == quality][ys2_col]
        mean_ys2 = quality_ys2_data.mean()
        std_ys2 = quality_ys2_data.std()
        median_ys2 = quality_ys2_data.median()
        print(f"   {i}. {quality}: {count:,}개")
        print(f"      평균: {mean_ys2:.2f}±{std_ys2:.2f} MPa, 중앙값: {median_ys2:.2f} MPa")
    
    # 그래프 생성
    plt.figure(figsize=(14, 8))
    
    # stripplot 생성
    ax = sns.stripplot(
        data=data,
        x=quality_col,
        y=ys2_col,
        order=top_x52_qualities.index,
        size=8,
        alpha=0.7,
        jitter=True,
        linewidth=0.5,
        edgecolor='black'
    )
    
    # X52 최소 규격선 추가 (359 MPa)
    ax.axhline(y=359, color='orange', linestyle=':', linewidth=2, alpha=0.8, 
               label='X52 최소 규격 (359 MPa)')
    
    # 필터링 범위 표시
    ax.axhline(y=360, color='green', linestyle='--', linewidth=1.5, alpha=0.8, 
               label='필터링 하한 (360 MPa)')
    ax.axhline(y=530, color='green', linestyle='--', linewidth=1.5, alpha=0.8, 
               label='필터링 상한 (530 MPa)')
    
    # 각 품질별 평균선 추가
    for i, quality in enumerate(top_x52_qualities.index):
        quality_data = data[data[quality_col] == quality]
        mean_ys2_stress = quality_data[ys2_col].mean()
        
        # 평균선
        ax.hlines(
            y=mean_ys2_stress,
            xmin=i-0.4,
            xmax=i+0.4,
            colors='red',
            linestyles='-',
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
        '중경1공장 X52 계열: YS2_STRESS 360~530 MPa 필터링된 상위 5개 품질별 분포',
        fontsize=16,
        fontweight='bold',
        pad=20
    )
    plt.xlabel('X52 품질 (Quality)', fontsize=12, fontweight='bold')
    plt.ylabel('YS2_STRESS (MPa)', fontsize=12, fontweight='bold')
    
    # X축 레이블 회전
    plt.xticks(rotation=45, ha='right')
    
    # Y축 범위 설정 (350~540 MPa)
    plt.ylim(350, 540)
    
    # 격자
    plt.grid(True, alpha=0.3, axis='y')
    
    # X52 품질별 통계 정보 박스 (오른쪽 위)
    info_text = "X52 계열 YS2_STRESS 정보:\n"
    info_text += f"필터링 범위: 360~530 MPa\n\n"
    
    for quality in top_x52_qualities.index:
        quality_data = data[data[quality_col] == quality]
        count = len(quality_data)
        mean_val = quality_data[ys2_col].mean()
        median_val = quality_data[ys2_col].median()
        info_text += f"{quality}: {count:,}개\n"
        info_text += f"  평균: {mean_val:.1f}, 중앙값: {median_val:.1f}\n"
    
    total_count = sum(top_x52_qualities.values)
    overall_mean = data[ys2_col].mean()
    overall_median = data[ys2_col].median()
    info_text += f"\n총계: {total_count:,}개\n"
    info_text += f"전체 평균: {overall_mean:.1f} MPa\n"
    info_text += f"전체 중앙값: {overall_median:.1f} MPa"
    
    # 정보 박스 스타일
    bbox_props = dict(
        boxstyle="round,pad=0.5",
        facecolor="lightyellow",
        alpha=0.9,
        edgecolor="darkorange",
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
    
    # 범례 추가
    plt.legend(loc='upper left', fontsize=10)
    
    # 레이아웃 조정
    plt.tight_layout()
    
    # 저장
    filename = '중경1공장_X52계열_YS2_STRESS_360-530MPa_필터링_stripplot.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"💾 그래프 저장 완료: {filename}")
    
    plt.show()
    
    # 상세 통계 출력
    print(f"\n📊 X52 계열 품질별 YS2_STRESS 상세 통계:")
    stats = data.groupby(quality_col)[ys2_col].agg([
        'count', 'mean', 'std', 'min', 'max', 'median',
        lambda x: x.quantile(0.25), lambda x: x.quantile(0.75)
    ]).round(2)
    stats.columns = ['개수', '평균', '표준편차', '최소값', '최대값', '중앙값', '25%', '75%']
    print(stats)
    
    return filename

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 X52 계열 YS2_STRESS 360~530 MPa 필터링 stripplot 생성")
    print("=" * 80)
    
    # 1. 한글 폰트 설정
    font_success = setup_korean_font_robust()
    
    # 2. 데이터 로드
    data = load_data()
    
    if data is not None:
        # 3. YS2_STRESS 범위 필터링
        filtered_data = filter_ys2_stress_range(data, min_ys2=360, max_ys2=530)
        
        if filtered_data is not None and len(filtered_data) > 0:
            # 4. X52 계열 필터링
            x52_data = filter_x52_grades(filtered_data)
            
            if x52_data is not None and len(x52_data) > 0:
                # 5. stripplot 생성
                filename = create_x52_ys2_stress_stripplot(x52_data)
                if filename:
                    print(f"\n✅ 작업 완료! 생성된 파일: {filename}")
                else:
                    print("\n❌ 그래프 생성 실패")
            else:
                print("\n❌ X52 계열 데이터가 없습니다.")
        else:
            print("\n❌ YS2_STRESS 범위 필터링 실패")
    else:
        print("\n❌ 데이터 로드 실패")
    
    print("=" * 80)

if __name__ == "__main__":
    main()
