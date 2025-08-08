#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
중경1공장 데이터에서 특정 컬럼의 0값 제거 스크립트
"""

import pandas as pd
import numpy as np

def filter_zero_values():
    """지정된 컬럼들에서 0값을 제거하여 필터링"""
    
    print("🔍 중경1공장 데이터 0값 필터링 시작...")
    
    # 입력 파일
    input_file = '중경1공장_데이터.xlsx'
    
    try:
        # 데이터 로드
        print("📂 중경1공장 데이터 로드 중...")
        data = pd.read_excel(input_file)
        print(f"✅ 데이터 로드 완료: {data.shape}")
        
        # 필터링할 컬럼들 (대소문자 구분 없이 찾기)
        target_columns = ['PCM', 'CEQ', 'Hardness', 'i_YS', 'YS2_STRESS', 'i_TS', 'TS_STRESS']
        
        # 실제 데이터에서 컬럼명 매칭
        actual_columns = []
        column_mapping = {}
        
        print(f"\n🎯 필터링 대상 컬럼 매칭:")
        for target_col in target_columns:
            # 정확한 매칭 먼저 시도
            if target_col in data.columns:
                actual_columns.append(target_col)
                column_mapping[target_col] = target_col
                print(f"   ✅ {target_col} → {target_col}")
            else:
                # 대소문자 무시하고 매칭
                found = False
                for col in data.columns:
                    if col.lower() == target_col.lower():
                        actual_columns.append(col)
                        column_mapping[target_col] = col
                        print(f"   ✅ {target_col} → {col}")
                        found = True
                        break
                
                if not found:
                    print(f"   ❌ {target_col} 컬럼을 찾을 수 없습니다.")
        
        if not actual_columns:
            print("❌ 필터링할 컬럼이 없습니다.")
            return None
        
        print(f"\n📊 필터링 전 데이터 분석:")
        print(f"   총 레코드 수: {len(data):,}개")
        
        # 각 컬럼별 0값 개수 확인
        zero_counts = {}
        for col in actual_columns:
            zero_count = (data[col] == 0).sum()
            total_count = len(data)
            zero_percentage = (zero_count / total_count) * 100
            zero_counts[col] = zero_count
            print(f"   {col}: 0값 {zero_count:,}개 ({zero_percentage:.1f}%)")
        
        # 필터링 적용 (모든 지정 컬럼에서 0이 아닌 값만 유지)
        print(f"\n🔧 필터링 적용 중...")
        filtered_data = data.copy()
        
        for col in actual_columns:
            before_count = len(filtered_data)
            filtered_data = filtered_data[filtered_data[col] != 0]
            after_count = len(filtered_data)
            removed_count = before_count - after_count
            print(f"   {col} 필터링: {removed_count:,}개 제거 → {after_count:,}개 남음")
        
        print(f"\n✅ 필터링 완료:")
        print(f"   필터링 전: {len(data):,}개")
        print(f"   필터링 후: {len(filtered_data):,}개")
        print(f"   제거된 데이터: {len(data) - len(filtered_data):,}개 ({(len(data) - len(filtered_data))/len(data)*100:.1f}%)")
        print(f"   남은 데이터: {len(filtered_data)/len(data)*100:.1f}%")
        
        # 필터링 후 각 컬럼의 기본 통계
        print(f"\n📈 필터링 후 주요 컬럼 통계:")
        for col in actual_columns:
            if filtered_data[col].dtype in ['float64', 'int64']:
                stats = filtered_data[col].describe()
                print(f"   {col}: 평균 {stats['mean']:.2f}, 최소 {stats['min']:.2f}, 최대 {stats['max']:.2f}")
        
        # 결과 저장
        output_file = '중경1공장_데이터_필터링완료.xlsx'
        print(f"\n💾 필터링된 데이터 저장 중: {output_file}")
        filtered_data.to_excel(output_file, index=False)
        
        # 파일 크기 확인
        import os
        file_size = os.path.getsize(output_file) / 1024**2
        print(f"✅ 저장 완료! 파일 크기: {file_size:.1f} MB")
        
        return filtered_data
        
    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")
        return None

def analyze_filtered_data(filtered_data):
    """필터링된 데이터의 상세 분석"""
    
    if filtered_data is None or len(filtered_data) == 0:
        print("❌ 분석할 데이터가 없습니다.")
        return
    
    print("\n" + "="*60)
    print("📊 필터링된 중경1공장 데이터 분석")
    print("="*60)
    
    # 1. 품질별 분포 (필터링 후)
    if 'p_spec' in filtered_data.columns:
        print(f"\n1️⃣ 필터링 후 품질 분포:")
        quality_counts = filtered_data['p_spec'].value_counts()
        for quality, count in quality_counts.head(10).items():
            percentage = (count / len(filtered_data)) * 100
            print(f"   - {quality}: {count:,}개 ({percentage:.1f}%)")
    
    # 2. 주요 측정값 분포
    measurement_cols = ['pcm', 'ceq', 'hardness', 'i_ys', 'ys2_stress', 'i_ts', 'ts_stress']
    available_cols = [col for col in measurement_cols if col in filtered_data.columns]
    
    if available_cols:
        print(f"\n2️⃣ 주요 측정값 분포:")
        for col in available_cols:
            if filtered_data[col].dtype in ['float64', 'int64']:
                stats = filtered_data[col].describe()
                print(f"   - {col.upper()}:")
                print(f"     평균: {stats['mean']:.3f}, 표준편차: {stats['std']:.3f}")
                print(f"     최소: {stats['min']:.3f}, 최대: {stats['max']:.3f}")
                print(f"     25%: {stats['25%']:.3f}, 75%: {stats['75%']:.3f}")
    
    # 3. 데이터 품질 확인
    print(f"\n3️⃣ 데이터 품질 확인:")
    print(f"   - 총 레코드 수: {len(filtered_data):,}개")
    print(f"   - 결측값 있는 컬럼: {filtered_data.isnull().any().sum()}개")
    
    # 결측값이 있는 컬럼 상세 정보
    null_cols = filtered_data.columns[filtered_data.isnull().any()].tolist()
    if null_cols:
        print(f"   - 결측값 있는 컬럼 목록:")
        for col in null_cols:
            null_count = filtered_data[col].isnull().sum()
            null_percentage = (null_count / len(filtered_data)) * 100
            print(f"     {col}: {null_count:,}개 ({null_percentage:.1f}%)")

def main():
    """메인 실행 함수"""
    print("🚀 중경1공장 데이터 0값 필터링 시작")
    print("=" * 80)
    
    # 1. 0값 필터링
    filtered_data = filter_zero_values()
    
    # 2. 필터링된 데이터 분석
    if filtered_data is not None:
        analyze_filtered_data(filtered_data)
    
    print("\n✅ 필터링 작업 완료!")
    print("=" * 80)

if __name__ == "__main__":
    main()
