#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 품질별 두께 stripplot - 한글 폰트 완전 해결 버전
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
        # 마지막 수단: 환경변수 설정
        os.environ['MPLCONFIGDIR'] = '/tmp/matplotlib_config'
    
    # 5. 폰트 설정 확인 테스트
    try:
        fig, ax = plt.subplots(figsize=(2, 1))
        ax.text(0.5, 0.5, '한글테스트', ha='center', va='center', fontsize=12)
        ax.set_title('폰트 테스트')
        plt.savefig('font_test_check.png', dpi=100, bbox_inches='tight')
        plt.close(fig)
        print("✅ 한글 폰트 테스트 완료")
        
        # 테스트 파일 삭제
        if os.path.exists('font_test_check.png'):
            os.remove('font_test_check.png')
            
    except Exception as e:
        print(f"❌ 폰트 테스트 실패: {e}")
    
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

def create_quality_thickness_stripplot(data):
    """품질별 두께 stripplot 생성"""
    
    if data is None or len(data) == 0:
        print("❌ 데이터가 없습니다.")
        return None
    
    print("\n🎨 stripplot 생성 중...")
    
    # 상위 5개 품질 선정
    quality_col = 'p_spec'
    thickness_col = 'p_thick_mm'
    
    top_qualities = data[quality_col].value_counts().head(5)
    filtered_data = data[data[quality_col].isin(top_qualities.index)].copy()
    
    print(f"📊 상위 5개 품질:")
    for i, (quality, count) in enumerate(top_qualities.items(), 1):
        print(f"   {i}. {quality}: {count:,}개")
    
    # 그래프 생성
    plt.figure(figsize=(14, 8))
    
    # stripplot 생성
    ax = sns.stripplot(
        data=filtered_data,
        x=quality_col,
        y=thickness_col,
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
        mean_thickness = quality_data[thickness_col].mean()
        
        # 평균선
        ax.hlines(
            y=mean_thickness,
            xmin=i-0.4,
            xmax=i+0.4,
            colors='red',
            linestyles='--',
            linewidth=2,
            alpha=0.8
        )
        
        # 평균값 텍스트
        ax.text(
            i, mean_thickness + (ax.get_ylim()[1] - ax.get_ylim()[0]) * 0.02,
            f'평균: {mean_thickness:.2f}',
            ha='center',
            va='bottom',
            fontsize=9,
            fontweight='bold',
            color='red'
        )
    
    # 제목 및 레이블 (한글 사용)
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
    
    # 격자
    plt.grid(True, alpha=0.3, axis='y')
    
    # 품질별 개수 정보 박스 (오른쪽 위)
    info_text = "품질별 데이터 개수:\n"
    for quality, count in top_qualities.items():
        info_text += f"{quality}: {count:,}개\n"
    info_text += f"\n총계: {sum(top_qualities.values):,}개"
    
    # 정보 박스 스타일
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
    filename = '중경1공장_상위5개품질_두께분포_stripplot_한글수정.png'
    plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"💾 그래프 저장 완료: {filename}")
    
    plt.show()
    
    # 통계 출력
    print(f"\n📊 품질별 두께 통계:")
    stats = filtered_data.groupby(quality_col)[thickness_col].agg([
        'count', 'mean', 'std', 'min', 'max'
    ]).round(3)
    print(stats)
    
    return filename

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 품질별 두께 stripplot 생성 (한글 폰트 완전 수정 버전)")
    print("=" * 80)
    
    # 1. 한글 폰트 설정
    font_success = setup_korean_font_robust()
    
    # 2. 데이터 로드
    data = load_data()
    
    # 3. stripplot 생성
    if data is not None:
        filename = create_quality_thickness_stripplot(data)
        if filename:
            print(f"\n✅ 작업 완료! 생성된 파일: {filename}")
        else:
            print("\n❌ 그래프 생성 실패")
    else:
        print("\n❌ 데이터 로드 실패")
    
    print("=" * 80)

if __name__ == "__main__":
    main()
