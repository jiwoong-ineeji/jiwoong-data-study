#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 데이터 추출 스크립트
"""

import pandas as pd
import os

def extract_jg1_data():
    """중경1공장 데이터를 원본 파일에서 추출하여 별도 파일로 저장"""
    
    print("📊 중경1공장 데이터 추출 시작...")
    
    # 원본 데이터 파일 경로
    input_file = './첫시도/joined_coil_jiwoong.xlsx'
    
    if not os.path.exists(input_file):
        print(f"❌ 원본 파일을 찾을 수 없습니다: {input_file}")
        return None
    
    try:
        # 원본 데이터 로드
        print("📂 원본 데이터 로드 중...")
        data = pd.read_excel(input_file)
        print(f"✅ 원본 데이터 로드 완료: {data.shape}")
        
        # 컬럼 정보 확인
        print(f"\n📋 전체 컬럼 목록 ({len(data.columns)}개):")
        for i, col in enumerate(data.columns, 1):
            print(f"{i:2d}. {col}")
        
        # 공장 관련 컬럼 찾기
        factory_cols = [col for col in data.columns if any(word in col.lower() for word in ['wc', 'factory', '공장', 'plant'])]
        print(f"\n🏭 공장 관련 컬럼: {factory_cols}")
        
        if not factory_cols:
            print("❌ 공장을 구분할 수 있는 컬럼을 찾을 수 없습니다.")
            return None
        
        # 주요 공장 컬럼 선택 (wc_desc 우선)
        factory_col = 'wc_desc' if 'wc_desc' in factory_cols else factory_cols[0]
        print(f"✅ 공장 구분 컬럼으로 '{factory_col}' 사용")
        
        # 공장별 데이터 분포 확인
        print(f"\n📊 {factory_col} 분포:")
        factory_counts = data[factory_col].value_counts()
        for factory, count in factory_counts.items():
            print(f"   {factory}: {count:,}개")
        
        # 중경1공장 데이터 추출
        jg1_keywords = ['중경1공장', '중경1', 'JG1']
        jg1_mask = data[factory_col].str.contains('|'.join(jg1_keywords), case=False, na=False)
        jg1_data = data[jg1_mask].copy()
        
        print(f"\n✅ 중경1공장 데이터 추출 완료:")
        print(f"   전체 데이터: {len(data):,}개")
        print(f"   중경1공장: {len(jg1_data):,}개 ({len(jg1_data)/len(data)*100:.1f}%)")
        
        if len(jg1_data) == 0:
            print("❌ 중경1공장 데이터가 없습니다.")
            return None
        
        # 중경1공장 데이터의 기본 정보
        print(f"\n📈 중경1공장 데이터 기본 정보:")
        print(f"   데이터 크기: {jg1_data.shape}")
        print(f"   메모리 사용량: {jg1_data.memory_usage(deep=True).sum() / 1024**2:.1f} MB")
        
        # 주요 컬럼 통계 (숫자형 컬럼만)
        numeric_cols = jg1_data.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            print(f"\n📊 주요 숫자형 컬럼 통계:")
            print(jg1_data[numeric_cols].describe().round(2))
        
        # 저장할 파일명 생성
        output_file = '중경1공장_데이터.xlsx'
        
        # 데이터 저장
        print(f"\n💾 중경1공장 데이터 저장 중: {output_file}")
        jg1_data.to_excel(output_file, index=False)
        
        # 저장된 파일 크기 확인
        file_size = os.path.getsize(output_file) / 1024**2
        print(f"✅ 저장 완료! 파일 크기: {file_size:.1f} MB")
        
        return jg1_data
        
    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")
        return None

def analyze_jg1_data(jg1_data):
    """추출된 중경1공장 데이터의 상세 분석"""
    
    if jg1_data is None or len(jg1_data) == 0:
        print("❌ 분석할 데이터가 없습니다.")
        return
    
    print("\n" + "="*60)
    print("📊 중경1공장 데이터 상세 분석")
    print("="*60)
    
    # 1. 기본 정보
    print(f"\n1️⃣ 기본 정보:")
    print(f"   - 총 레코드 수: {len(jg1_data):,}개")
    print(f"   - 총 컬럼 수: {len(jg1_data.columns)}개")
    print(f"   - 결측값 있는 컬럼: {jg1_data.isnull().any().sum()}개")
    
    # 2. 품질 관련 분석
    quality_cols = [col for col in jg1_data.columns if any(word in col.lower() for word in ['quality', 'grade', 'spec'])]
    if quality_cols:
        print(f"\n2️⃣ 품질 분포 ({quality_cols[0]}):")
        quality_counts = jg1_data[quality_cols[0]].value_counts()
        for quality, count in quality_counts.head(10).items():
            print(f"   - {quality}: {count:,}개")
    
    # 3. 시간 관련 분석
    date_cols = [col for col in jg1_data.columns if any(word in col.lower() for word in ['date', 'time', '일자', '시간'])]
    if date_cols:
        print(f"\n3️⃣ 시간 범위 ({date_cols[0]}):")
        try:
            date_col = date_cols[0]
            if jg1_data[date_col].dtype == 'object':
                jg1_data[date_col] = pd.to_datetime(jg1_data[date_col], errors='coerce')
            
            min_date = jg1_data[date_col].min()
            max_date = jg1_data[date_col].max()
            print(f"   - 시작일: {min_date}")
            print(f"   - 종료일: {max_date}")
            print(f"   - 기간: {(max_date - min_date).days}일")
        except:
            print("   - 날짜 정보 분석 실패")
    
    # 4. 주요 측정값 분석
    measurement_cols = [col for col in jg1_data.columns if any(word in col.lower() for word in ['ys', 'ts', 'thickness', '두께', '강도'])]
    if measurement_cols:
        print(f"\n4️⃣ 주요 측정값 통계:")
        for col in measurement_cols[:5]:  # 상위 5개만
            if jg1_data[col].dtype in ['float64', 'int64']:
                stats = jg1_data[col].describe()
                print(f"   - {col}: 평균 {stats['mean']:.1f}, 표준편차 {stats['std']:.1f}")

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 데이터 추출 및 분석 시작")
    print("=" * 80)
    
    # 1. 중경1공장 데이터 추출
    jg1_data = extract_jg1_data()
    
    # 2. 추출된 데이터 분석
    if jg1_data is not None:
        analyze_jg1_data(jg1_data)
    
    print("\n✅ 작업 완료!")
    print("=" * 80)

if __name__ == "__main__":
    main()
