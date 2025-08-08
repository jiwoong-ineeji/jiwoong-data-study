#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 품질별 (TS_STRESS - I_TS) 차이값 분포 stripplot
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

def calculate_ts_minus_its_difference(data):
    """TS_STRESS - I_TS 차이값 계산 및 분석"""
    
    print("\n🔢 TS_STRESS - I_TS 차이값 계산 중...")
    
    # 컬럼 확인
    ts_stress_col = 'ts_stress'
    its_col = 'i_ts'
    quality_col = 'p_spec'
    
    if ts_stress_col not in data.columns or its_col not in data.columns:
        print(f"❌ 필요한 컬럼을 찾을 수 없습니다.")
        print(f"사용 가능한 컬럼: {list(data.columns)}")
        return None
    
    # 차이값 계산
    data['ts_minus_its'] = data[ts_stress_col] - data[its_col]
    
    print(f"📊 (TS_STRESS - I_TS) 차이값 분석:")
    print(f"   전체 데이터: {len(data):,}개")
    print(f"   차이값 범위: {data['ts_minus_its'].min():.2f} ~ {data['ts_minus_its'].max():.2f} MPa")
    print(f"   차이값 평균: {data['ts_minus_its'].mean():.2f} MPa")
    print(f"   차이값 표준편차: {data['ts_minus_its'].std():.2f} MPa")
    
    # 양수/음수 분포
    positive_count = (data['ts_minus_its'] > 0).sum()
    negative_count = (data['ts_minus_its'] < 0).sum()
    zero_count = (data['ts_minus_its'] == 0).sum()
    
    print(f"\n📈 차이값 분포:")
    print(f"   양수 (TS_STRESS > I_TS): {positive_count:,}개 ({positive_count/len(data)*100:.1f}%)")
    print(f"   음수 (TS_STRESS < I_TS): {negative_count:,}개 ({negative_count/len(data)*100:.1f}%)")
    print(f"   0 (TS_STRESS = I_TS): {zero_count:,}개 ({zero_count/len(data)*100:.1f}%)")
    
    return data

def create_ts_minus_its_stripplot(data):
    """품질별 (TS_STRESS - I_TS) 차이값 stripplot 생성"""
    
    if data is None or len(data) == 0:
        print("❌ 데이터가 없습니다.")
        return None
    
    print("\n🎨 (TS_STRESS - I_TS) 차이값 stripplot 생성 중...")
    
    # 컬럼 설정
    quality_col = 'p_spec'
    diff_col = 'ts_minus_its'
    
    # 상위 5개 품질 선정
    top_qualities = data[quality_col].value_counts().head(5)
    filtered_data = data[data[quality_col].isin(top_qualities.index)].copy()
    
    print(f"\n📋 상위 5개 품질별 (TS_STRESS - I_TS) 차이값 분포:")
    for i, (quality, count) in enumerate(top_qualities.items(), 1):
        quality_diff_data = filtered_data[filtered_data[quality_col] == quality][diff_col]
        mean_diff = quality_diff_data.mean()
        std_diff = quality_diff_data.std()
        median_diff = quality_diff_data.median()
        print(f"   {i}. {quality}: {count:,}개")
        print(f"      평균: {mean_diff:.2f}±{std_diff:.2f} MPa, 중앙값: {median_diff:.2f} MPa")
    
    # 그래프 생성
    plt.figure(figsize=(14, 8))
    
    # stripplot 생성
    ax = sns.stripplot(
        data=filtered_data,
        x=quality_col,
        y=diff_col,
        order=top_qualities.index,
        size=8,
        alpha=0.7,
        jitter=True,
        linewidth=0.5,
        edgecolor='black'
    )
    
    # 0선 추가 (기준선)
    ax.axhline(y=0, color='black', linestyle='-', linewidth=1, alpha=0.8, label='기준선 (차이=0)')
    
    # 각 품질별 평균선 추가
    for i, quality in enumerate(top_qualities.index):
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        mean_diff = quality_data[diff_col].mean()
        
        # 평균선
        ax.hlines(
            y=mean_diff,
            xmin=i-0.4,
            xmax=i+0.4,
            colors='red',
            linestyles='--',
            linewidth=2,
            alpha=0.8
        )
        
        # 평균값 텍스트
        ax.text(
            i, mean_diff + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.02,
            f'평균: {mean_diff:.1f}',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
            color='red'
        )
    
    # 제목 및 레이블 (한글 사용)
    plt.title(
        '중경1공장 필터링된 데이터: 상위 5개 품질별 (TS_STRESS - I_TS) 차이값 분포',
        fontsize=16,
        fontweight='bold',
        pad=20
    )
    plt.xlabel('품질 (Quality)', fontsize=12, fontweight='bold')
    plt.ylabel('TS_STRESS - I_TS (MPa)', fontsize=12, fontweight='bold')
    
    # X축 레이블 회전
    plt.xticks(rotation=45, ha='right')
    
    # 격자
    plt.grid(True, alpha=0.3, axis='y')
    
    # 품질별 차이값 통계 정보 박스 (오른쪽 위)
    info_text = "품질별 차이값 (TS_STRESS - I_TS):\n"
    for quality in top_qualities.index:
        quality_data = filtered_data[filtered_data[quality_col] == quality]
        count = len(quality_data)
        mean_val = quality_data[diff_col].mean()
        median_val = quality_data[diff_col].median()
        info_text += f"{quality}: {count:,}개\n"
        info_text += f"  평균: {mean_val:.1f}, 중앙값: {median_val:.1f}\n"
    
    total_count = sum(top_qualities.values)
    overall_mean = filtered_data[diff_col].mean()
    overall_median = filtered_data[diff_col].median()
    info_text += f"\n총계: {total_count:,}개\n"
    info_text += f"전체 평균: {overall_mean:.1f} MPa\n"
    info_text += f"전체 중앙값: {overall_median:.1f} MPa"
    
    # 정보 박스 스타일
    bbox_props = dict(
        boxstyle="round,pad=0.5",
        facecolor="lavender",
        alpha=0.9,
        edgecolor="purple",
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
    filename = '중경1공장_상위5개품질_TS_STRESS_minus_I_TS_차이값_stripplot.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"💾 그래프 저장 완료: {filename}")
    
    plt.show()
    
    # 상세 통계 출력
    print(f"\n📊 품질별 (TS_STRESS - I_TS) 차이값 상세 통계:")
    stats = filtered_data.groupby(quality_col)[diff_col].agg([
        'count', 'mean', 'std', 'min', 'max', 'median',
        lambda x: x.quantile(0.25), lambda x: x.quantile(0.75)
    ]).round(2)
    stats.columns = ['개수', '평균', '표준편차', '최소값', '최대값', '중앙값', '25%', '75%']
    print(stats)
    
    return filename

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 품질별 (TS_STRESS - I_TS) 차이값 분석")
    print("=" * 80)
    
    # 1. 한글 폰트 설정
    font_success = setup_korean_font_robust()
    
    # 2. 데이터 로드
    data = load_data()
    
    # 3. 차이값 계산
    if data is not None:
        data_with_diff = calculate_ts_minus_its_difference(data)
        
        # 4. stripplot 생성
        if data_with_diff is not None:
            filename = create_ts_minus_its_stripplot(data_with_diff)
            if filename:
                print(f"\n✅ 작업 완료! 생성된 파일: {filename}")
            else:
                print("\n❌ 그래프 생성 실패")
        else:
            print("\n❌ 차이값 계산 실패")
    else:
        print("\n❌ 데이터 로드 실패")
    
    print("=" * 80)

if __name__ == "__main__":
    main()
