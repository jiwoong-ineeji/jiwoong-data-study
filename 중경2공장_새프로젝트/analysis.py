"""
중경2공장 새 프로젝트 - 데이터 분석 스크립트
작성일: 2024년
목적: 중경2공장 품질 데이터 분석 및 시각화
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import warnings

# 한글 폰트 설정
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
warnings.filterwarnings('ignore')

class JungKyung2FactoryAnalyzer:
    """중경2공장 데이터 분석 클래스"""
    
    def __init__(self, data_path=None):
        """
        초기화
        Args:
            data_path: 데이터 파일 경로
        """
        self.data_path = data_path
        self.data = None
        self.output_dir = Path("output")
        self.output_dir.mkdir(exist_ok=True)
        
    def load_data(self, file_path):
        """
        데이터 로드
        Args:
            file_path: Excel 또는 CSV 파일 경로
        """
        try:
            if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
                self.data = pd.read_excel(file_path)
            elif file_path.endswith('.csv'):
                self.data = pd.read_csv(file_path, encoding='utf-8')
            else:
                raise ValueError("지원하지 않는 파일 형식입니다.")
                
            print(f"데이터 로드 완료: {self.data.shape}")
            print(f"컬럼: {list(self.data.columns)}")
            return self.data
            
        except Exception as e:
            print(f"데이터 로드 실패: {e}")
            return None
    
    def data_overview(self):
        """데이터 개요 출력"""
        if self.data is None:
            print("데이터가 로드되지 않았습니다.")
            return
            
        print("=" * 50)
        print("데이터 개요")
        print("=" * 50)
        print(f"데이터 크기: {self.data.shape}")
        print(f"컬럼 수: {len(self.data.columns)}")
        print(f"행 수: {len(self.data)}")
        print("\n컬럼 정보:")
        print(self.data.info())
        print("\n기본 통계:")
        print(self.data.describe())
        print("\n결측값:")
        print(self.data.isnull().sum())
        
    def quality_analysis(self, quality_col='품질', target_cols=None):
        """
        품질별 분석
        Args:
            quality_col: 품질 컬럼명
            target_cols: 분석할 대상 컬럼들
        """
        if self.data is None:
            print("데이터가 로드되지 않았습니다.")
            return
            
        if quality_col not in self.data.columns:
            print(f"'{quality_col}' 컬럼을 찾을 수 없습니다.")
            return
            
        print("=" * 50)
        print("품질별 분석")
        print("=" * 50)
        
        # 품질별 개수
        quality_counts = self.data[quality_col].value_counts()
        print("품질별 데이터 개수:")
        print(quality_counts)
        
        # 품질별 통계 (숫자형 컬럼만)
        numeric_cols = self.data.select_dtypes(include=[np.number]).columns
        if target_cols:
            numeric_cols = [col for col in target_cols if col in numeric_cols]
            
        quality_stats = self.data.groupby(quality_col)[numeric_cols].agg(['mean', 'std', 'min', 'max'])
        print("\n품질별 통계:")
        print(quality_stats)
        
        return quality_stats
    
    def create_visualization(self, x_col, y_col, hue_col=None, plot_type='scatter'):
        """
        시각화 생성
        Args:
            x_col: X축 컬럼
            y_col: Y축 컬럼  
            hue_col: 색상 구분 컬럼
            plot_type: 플롯 타입 ('scatter', 'box', 'violin', 'strip')
        """
        if self.data is None:
            print("데이터가 로드되지 않았습니다.")
            return
            
        plt.figure(figsize=(12, 8))
        
        if plot_type == 'scatter':
            sns.scatterplot(data=self.data, x=x_col, y=y_col, hue=hue_col, s=100, alpha=0.7)
        elif plot_type == 'box':
            sns.boxplot(data=self.data, x=x_col, y=y_col, hue=hue_col)
        elif plot_type == 'violin':
            sns.violinplot(data=self.data, x=x_col, y=y_col, hue=hue_col)
        elif plot_type == 'strip':
            sns.stripplot(data=self.data, x=x_col, y=y_col, hue=hue_col, size=8, alpha=0.7)
            
        plt.title(f'{plot_type.title()} Plot: {x_col} vs {y_col}', fontsize=16, fontweight='bold')
        plt.xlabel(x_col, fontsize=12)
        plt.ylabel(y_col, fontsize=12)
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        # 저장
        filename = f"{plot_type}_{x_col}_{y_col}.png"
        if hue_col:
            filename = f"{plot_type}_{x_col}_{y_col}_{hue_col}.png"
        filepath = self.output_dir / filename
        plt.savefig(filepath, dpi=300, bbox_inches='tight')
        plt.show()
        
        print(f"그래프 저장됨: {filepath}")
        
    def correlation_analysis(self, target_cols=None):
        """상관관계 분석"""
        if self.data is None:
            print("데이터가 로드되지 않았습니다.")
            return
            
        numeric_data = self.data.select_dtypes(include=[np.number])
        if target_cols:
            numeric_data = numeric_data[target_cols]
            
        # 상관계수 계산
        corr_matrix = numeric_data.corr()
        
        # 히트맵 생성
        plt.figure(figsize=(12, 10))
        sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', center=0, 
                   square=True, fmt='.3f', cbar_kws={'shrink': 0.8})
        plt.title('상관관계 히트맵', fontsize=16, fontweight='bold')
        plt.tight_layout()
        
        # 저장
        filepath = self.output_dir / "correlation_heatmap.png"
        plt.savefig(filepath, dpi=300, bbox_inches='tight')
        plt.show()
        
        print(f"상관관계 히트맵 저장됨: {filepath}")
        return corr_matrix

def main():
    """메인 실행 함수"""
    print("중경2공장 새 프로젝트 분석 시작")
    print("=" * 60)
    
    # 분석기 초기화
    analyzer = JungKyung2FactoryAnalyzer()
    
    # 데이터 파일 경로 설정 (필요시 수정)
    data_file = "../첫시도/joined_coil_jiwoong.xlsx"
    
    # 데이터 로드
    data = analyzer.load_data(data_file)
    
    if data is not None:
        # 데이터 개요
        analyzer.data_overview()
        
        # 품질별 분석 (컬럼명은 실제 데이터에 맞게 조정 필요)
        # analyzer.quality_analysis()
        
        print("\n분석 준비 완료!")
        print("실제 데이터 컬럼명을 확인한 후 분석을 진행하세요.")
    
if __name__ == "__main__":
    main()
