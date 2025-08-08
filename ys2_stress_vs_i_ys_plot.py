#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 YS2_STRESS vs I_YS 관계 분석 차트
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import warnings
import matplotlib.font_manager as fm
import platform
import os
from scipy import stats

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

def analyze_ys2_vs_iys_correlation(data):
    """YS2_STRESS와 I_YS 상관관계 분석"""
    
    print("\n📊 YS2_STRESS vs I_YS 상관관계 분석:")
    
    # 컬럼 확인
    ys2_col = 'ys2_stress'
    iys_col = 'i_ys'
    
    if ys2_col not in data.columns or iys_col not in data.columns:
        print(f"❌ 필요한 컬럼을 찾을 수 없습니다.")
        return None
    
    # 유효한 데이터만 선택
    valid_data = data[[ys2_col, iys_col]].dropna()
    
    # 상관계수 계산
    correlation = valid_data[ys2_col].corr(valid_data[iys_col])
    
    # 선형 회귀 분석
    slope, intercept, r_value, p_value, std_err = stats.linregress(
        valid_data[iys_col], valid_data[ys2_col]
    )
    
    print(f"   📈 상관계수 (Pearson): {correlation:.4f}")
    print(f"   📈 결정계수 (R²): {r_value**2:.4f}")
    print(f"   📈 회귀식: YS2_STRESS = {slope:.3f} × I_YS + {intercept:.3f}")
    print(f"   📈 p-value: {p_value:.6f}")
    print(f"   📈 표준오차: {std_err:.4f}")
    
    return {
        'correlation': correlation,
        'r_squared': r_value**2,
        'slope': slope,
        'intercept': intercept,
        'p_value': p_value,
        'valid_data': valid_data
    }

def create_ys2_vs_iys_plot(data):
    """YS2_STRESS vs I_YS 관계 차트 생성"""
    
    if data is None or len(data) == 0:
        print("❌ 데이터가 없습니다.")
        return None
    
    print("\n🎨 YS2_STRESS vs I_YS 관계 차트 생성 중...")
    
    # 컬럼 설정
    quality_col = 'p_spec'
    ys2_col = 'ys2_stress'
    iys_col = 'i_ys'
    
    # 상관관계 분석
    correlation_info = analyze_ys2_vs_iys_correlation(data)
    if correlation_info is None:
        return None
    
    # 상위 5개 품질 선정
    top_qualities = data[quality_col].value_counts().head(5)
    filtered_data = data[data[quality_col].isin(top_qualities.index)].copy()
    
    # 그래프 생성
    plt.figure(figsize=(14, 10))
    
    # 품질별로 다른 색상과 마커 사용
    colors = plt.cm.Set1(np.linspace(0, 1, len(top_qualities)))
    markers = ['o', 's', '^', 'D', 'v']
    
    # 각 품질별 산점도
    for i, (quality, count) in enumerate(top_qualities.items()):
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        
        plt.scatter(
            quality_data[iys_col],
            quality_data[ys2_col],
            c=[colors[i]],
            marker=markers[i % len(markers)],
            s=60,
            alpha=0.7,
            label=f'{quality} (n={count})',
            edgecolors='black',
            linewidth=0.5
        )
    
    # 전체 데이터에 대한 회귀선 추가
    valid_data = correlation_info['valid_data']
    x_range = np.linspace(valid_data[iys_col].min(), valid_data[iys_col].max(), 100)
    y_pred = correlation_info['slope'] * x_range + correlation_info['intercept']
    
    plt.plot(x_range, y_pred, 'r-', linewidth=2, alpha=0.8, 
             label=f'회귀선 (R² = {correlation_info["r_squared"]:.3f})')
    
    # 신뢰구간 추가 (선택사항)
    # 95% 신뢰구간 계산
    residuals = valid_data[ys2_col] - (correlation_info['slope'] * valid_data[iys_col] + correlation_info['intercept'])
    mse = np.mean(residuals**2)
    std_error = np.sqrt(mse)
    
    plt.fill_between(x_range, y_pred - 1.96*std_error, y_pred + 1.96*std_error, 
                     alpha=0.2, color='red', label='95% 신뢰구간')
    
    # 제목 및 레이블
    plt.title(
        '중경1공장 필터링된 데이터: YS2_STRESS vs I_YS 관계 분석',
        fontsize=16,
        fontweight='bold',
        pad=20
    )
    plt.xlabel('I_YS (MPa)', fontsize=12, fontweight='bold')
    plt.ylabel('YS2_STRESS (MPa)', fontsize=12, fontweight='bold')
    
    # 격자
    plt.grid(True, alpha=0.3)
    
    # 범례
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=10)
    
    # 통계 정보 박스 (왼쪽 아래)
    stats_text = f"상관관계 분석 결과:\n"
    stats_text += f"• 상관계수: {correlation_info['correlation']:.4f}\n"
    stats_text += f"• 결정계수 (R²): {correlation_info['r_squared']:.4f}\n"
    stats_text += f"• 회귀식: Y = {correlation_info['slope']:.3f}X + {correlation_info['intercept']:.1f}\n"
    
    if correlation_info['p_value'] < 0.001:
        stats_text += f"• p-value: < 0.001 (매우 유의함)"
    else:
        stats_text += f"• p-value: {correlation_info['p_value']:.4f}"
    
    # 통계 박스 스타일
    bbox_props = dict(
        boxstyle="round,pad=0.5",
        facecolor="lightyellow",
        alpha=0.9,
        edgecolor="orange",
        linewidth=1
    )
    
    # 왼쪽 아래에 통계 정보 표시
    plt.text(
        0.02, 0.02,
        stats_text,
        transform=plt.gca().transAxes,
        fontsize=10,
        verticalalignment='bottom',
        horizontalalignment='left',
        bbox=bbox_props,
        fontweight='bold'
    )
    
    # 레이아웃 조정
    plt.tight_layout()
    
    # 저장
    filename = '중경1공장_YS2_STRESS_vs_I_YS_관계분석.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"💾 그래프 저장 완료: {filename}")
    
    plt.show()
    
    # 품질별 상관관계 분석
    print(f"\n📊 품질별 YS2_STRESS vs I_YS 상관관계:")
    for quality in top_qualities.index:
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        if len(quality_data) >= 3:  # 최소 3개 이상의 데이터가 있을 때만
            corr = quality_data[ys2_col].corr(quality_data[iys_col])
            print(f"   {quality}: 상관계수 = {corr:.4f} (n={len(quality_data)})")
    
    return filename

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 YS2_STRESS vs I_YS 관계 분석")
    print("=" * 80)
    
    # 1. 한글 폰트 설정
    font_success = setup_korean_font_robust()
    
    # 2. 데이터 로드
    data = load_data()
    
    # 3. YS2_STRESS vs I_YS 관계 차트 생성
    if data is not None:
        filename = create_ys2_vs_iys_plot(data)
        if filename:
            print(f"\n✅ 작업 완료! 생성된 파일: {filename}")
        else:
            print("\n❌ 그래프 생성 실패")
    else:
        print("\n❌ 데이터 로드 실패")
    
    print("=" * 80)

if __name__ == "__main__":
    main()
