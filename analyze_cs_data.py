import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns
from collections import Counter

def analyze_cs_data():
    """
    CS_자동답변_데이터.xlsx 파일을 분석하여 일자별 자동답변 건수를 확인
    """
    try:
        # 엑셀 파일 읽기
        print("엑셀 파일을 읽는 중...")
        df = pd.read_excel('shevano7/CS_자동답변_데이터.xlsx')
        
        print(f"데이터 형태: {df.shape}")
        print(f"컬럼명: {list(df.columns)}")
        
        # 데이터 미리보기
        print("\n=== 데이터 미리보기 ===")
        print(df.head())
        
        # Advice ID에서 날짜 추출 (YYYYMMDD 형식)
        print("\n=== Advice ID 분석 ===")
        if 'Advice ID' in df.columns:
            df['date'] = df['Advice ID'].astype(str).str[:8]
            print("Advice ID에서 날짜 추출 완료")
        else:
            print("Advice ID 컬럼을 찾을 수 없습니다.")
            return
        
        # I열 분석 (자동답변 내용)
        print("\n=== I열 자동답변 분석 ===")
        if len(df.columns) >= 9:  # I열은 9번째 컬럼 (0부터 시작하면 8)
            i_column = df.columns[8]
            print(f"I열 컬럼명: {i_column}")
            
            # 자동답변 키워드 정의
            auto_response_keywords = [
                '자동답변',
                '자동답변 :',
                '욕설 및 비속어 표현이 포함된 문의',
                '게임 결과 불만 문의',
                '스펨처리',
                '짜고 치기 신고 답변'
            ]
            
            # I열에서 자동답변 키워드 검색
            df['is_auto_response'] = df[i_column].astype(str).apply(
                lambda x: any(keyword in x for keyword in auto_response_keywords)
            )
            
            # 일자별 자동답변 건수 집계
            daily_auto_response = df[df['is_auto_response'] == True].groupby('date').size().reset_index()
            daily_auto_response.columns = ['날짜', '자동답변_건수']
            
            # 전체 일자별 건수
            daily_total = df.groupby('date').size().reset_index()
            daily_total.columns = ['날짜', '전체_건수']
            
            # 결과 병합
            result = pd.merge(daily_total, daily_auto_response, on='날짜', how='left')
            result['자동답변_건수'] = result['자동답변_건수'].fillna(0).astype(int)
            result['자동답변_비율'] = (result['자동답변_건수'] / result['전체_건수'] * 100).round(2)
            
            # 날짜 정렬
            result = result.sort_values('날짜')
            
            print("\n=== 일자별 자동답변 분석 결과 ===")
            print(result.to_string(index=False))
            
            # 결과를 엑셀로 저장
            output_file = '일자별_자동답변_분석결과.xlsx'
            result.to_excel(output_file, index=False)
            print(f"\n분석 결과가 '{output_file}'에 저장되었습니다.")
            
            # 시각화
            create_visualization(result)
            
            # 상세 분석
            detailed_analysis(df, i_column)
            
        else:
            print("I열을 찾을 수 없습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")

def create_visualization(result):
    """
    분석 결과를 시각화
    """
    plt.figure(figsize=(15, 10))
    
    # 서브플롯 1: 일자별 전체 건수 vs 자동답변 건수
    plt.subplot(2, 2, 1)
    x = range(len(result))
    plt.bar(x, result['전체_건수'], alpha=0.7, label='전체 건수', color='skyblue')
    plt.bar(x, result['자동답변_건수'], alpha=0.9, label='자동답변 건수', color='red')
    plt.xlabel('날짜')
    plt.ylabel('건수')
    plt.title('일자별 전체 건수 vs 자동답변 건수')
    plt.legend()
    plt.xticks(x, result['날짜'], rotation=45)
    
    # 서브플롯 2: 자동답변 비율
    plt.subplot(2, 2, 2)
    plt.plot(range(len(result)), result['자동답변_비율'], marker='o', color='green', linewidth=2)
    plt.xlabel('날짜')
    plt.ylabel('자동답변 비율 (%)')
    plt.title('일자별 자동답변 비율')
    plt.xticks(range(len(result)), result['날짜'], rotation=45)
    plt.grid(True, alpha=0.3)
    
    # 서브플롯 3: 자동답변 건수만
    plt.subplot(2, 2, 3)
    plt.bar(range(len(result)), result['자동답변_건수'], color='orange', alpha=0.8)
    plt.xlabel('날짜')
    plt.ylabel('자동답변 건수')
    plt.title('일자별 자동답변 건수')
    plt.xticks(range(len(result)), result['날짜'], rotation=45)
    
    # 서브플롯 4: 전체 통계
    plt.subplot(2, 2, 4)
    total_auto = result['자동답변_건수'].sum()
    total_all = result['전체_건수'].sum()
    labels = ['자동답변', '일반답변']
    sizes = [total_auto, total_all - total_auto]
    colors = ['red', 'lightblue']
    plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.title('전체 자동답변 비율')
    
    plt.tight_layout()
    plt.savefig('자동답변_분석_차트.png', dpi=300, bbox_inches='tight')
    plt.show()
    print("차트가 '자동답변_분석_차트.png'로 저장되었습니다.")

def detailed_analysis(df, i_column):
    """
    자동답변 상세 분석
    """
    print("\n=== 자동답변 상세 분석 ===")
    
    # 자동답변 키워드별 분석
    keywords = [
        '자동답변',
        '자동답변 :',
        '욕설 및 비속어 표현이 포함된 문의',
        '게임 결과 불만 문의',
        '스펨처리',
        '짜고 치기 신고 답변'
    ]
    
    keyword_counts = {}
    for keyword in keywords:
        count = df[i_column].astype(str).str.contains(keyword, na=False).sum()
        keyword_counts[keyword] = count
    
    print("\n키워드별 자동답변 건수:")
    for keyword, count in keyword_counts.items():
        print(f"{keyword}: {count}건")
    
    # 가장 많이 사용된 자동답변 패턴 찾기
    auto_responses = df[df['is_auto_response'] == True][i_column].astype(str)
    if len(auto_responses) > 0:
        print(f"\n총 자동답변 건수: {len(auto_responses)}건")
        
        # 자동답변 내용 샘플
        print("\n자동답변 내용 샘플 (처음 5개):")
        for i, response in enumerate(auto_responses.head()):
            print(f"{i+1}. {response[:100]}...")

if __name__ == "__main__":
    print("CS 자동답변 데이터 분석을 시작합니다...")
    analyze_cs_data()
    print("\n분석이 완료되었습니다.")
