import pandas as pd
import numpy as np
from datetime import datetime
import json

def extract_auto_response_data():
    """
    CS_자동답변_데이터.xlsx에서 I열이 '자동답변:' 또는 '자동답변 :'으로 시작하는
    데이터의 E열과 F열만 추출하는 함수
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
        
        # I열 확인 (요약 결과 컬럼)
        if len(df.columns) >= 9:
            i_column = df.columns[8]  # I열 (9번째 컬럼)
            print(f"\nI열 컬럼명: {i_column}")
            
            # E열과 F열 확인
            e_column = df.columns[4]  # E열 (5번째 컬럼)
            f_column = df.columns[5]  # F열 (6번째 컬럼)
            print(f"E열 컬럼명: {e_column}")
            print(f"F열 컬럼명: {f_column}")
            
            # I열에서 자동답변으로 시작하는 데이터 필터링
            print("\n=== 자동답변 데이터 필터링 ===")
            
            # 자동답변 시작 패턴들
            auto_response_patterns = [
                '자동답변:',
                '자동답변 :',
                '자동답변 : ',
                '자동답변: '
            ]
            
            # I열에서 자동답변으로 시작하는 데이터 찾기 (JSON 파싱 포함)
            def is_auto_response(text):
                try:
                    # JSON 파싱 시도
                    if isinstance(text, str) and text.strip().startswith('{'):
                        data = json.loads(text)
                        if 'value' in data and 'result' in data['value']:
                            result_text = data['value']['result']
                            return any(result_text.startswith(pattern) for pattern in auto_response_patterns)
                    # 일반 텍스트로 시작하는 경우
                    return any(text.strip().startswith(pattern) for pattern in auto_response_patterns)
                except:
                    # JSON 파싱 실패 시 일반 텍스트로 처리
                    return any(text.strip().startswith(pattern) for pattern in auto_response_patterns)
            
            mask = df[i_column].astype(str).apply(is_auto_response)
            
            # 자동답변 데이터만 추출
            auto_response_df = df[mask].copy()
            
            print(f"전체 데이터 건수: {len(df)}")
            print(f"자동답변 데이터 건수: {len(auto_response_df)}")
            print(f"자동답변 비율: {len(auto_response_df)/len(df)*100:.2f}%")
            
            if len(auto_response_df) > 0:
                # E열과 F열만 선택
                result_df = auto_response_df[[e_column, f_column]].copy()
                
                # 컬럼명 변경
                result_df.columns = ['질문내용', '답변내용']
                
                print("\n=== 추출된 자동답변 데이터 미리보기 ===")
                print(result_df.head(10))
                
                # 결과를 엑셀로 저장
                output_file = '자동답변_E열_F열_추출결과.xlsx'
                result_df.to_excel(output_file, index=False)
                print(f"\n추출 결과가 '{output_file}'에 저장되었습니다.")
                
                # 상세 분석
                detailed_analysis(auto_response_df, i_column, e_column, f_column)
                
                # Advice ID도 포함한 버전 저장
                if 'Advice ID' in auto_response_df.columns:
                    result_with_id = auto_response_df[['Advice ID', e_column, f_column]].copy()
                    result_with_id.columns = ['Advice ID', '질문내용', '답변내용']
                    result_with_id.to_excel('자동답변_AdviceID_포함.xlsx', index=False)
                    print("Advice ID가 포함된 결과가 '자동답변_AdviceID_포함.xlsx'에 저장되었습니다.")
                
            else:
                print("자동답변으로 시작하는 데이터를 찾을 수 없습니다.")
                
        else:
            print("I열을 찾을 수 없습니다.")
            
    except Exception as e:
        print(f"오류 발생: {e}")

def detailed_analysis(auto_response_df, i_column, e_column, f_column):
    """
    자동답변 데이터 상세 분석
    """
    print("\n=== 자동답변 데이터 상세 분석 ===")
    
    # 자동답변 패턴별 분석
    patterns = [
        '자동답변:',
        '자동답변 :',
        '자동답변 : ',
        '자동답변: '
    ]
    
    pattern_counts = {}
    for pattern in patterns:
        count = 0
        for text in auto_response_df[i_column].astype(str):
            try:
                if text.strip().startswith('{'):
                    data = json.loads(text)
                    if 'value' in data and 'result' in data['value']:
                        result_text = data['value']['result']
                        if result_text.startswith(pattern):
                            count += 1
                elif text.strip().startswith(pattern):
                    count += 1
            except:
                if text.strip().startswith(pattern):
                    count += 1
        if count > 0:
            pattern_counts[pattern] = count
    
    print("\n자동답변 패턴별 건수:")
    for pattern, count in pattern_counts.items():
        print(f"{pattern}: {count}건")
    
    # E열과 F열의 텍스트 길이 분석
    print("\n=== 텍스트 길이 분석 ===")
    
    # E열(질문내용) 길이 분석
    e_lengths = auto_response_df[e_column].astype(str).str.len()
    print(f"질문내용 평균 길이: {e_lengths.mean():.1f}자")
    print(f"질문내용 최대 길이: {e_lengths.max()}자")
    print(f"질문내용 최소 길이: {e_lengths.min()}자")
    
    # F열(답변내용) 길이 분석
    f_lengths = auto_response_df[f_column].astype(str).str.len()
    print(f"답변내용 평균 길이: {f_lengths.mean():.1f}자")
    print(f"답변내용 최대 길이: {f_lengths.max()}자")
    print(f"답변내용 최소 길이: {f_lengths.min()}자")
    
    # 샘플 데이터 출력
    print("\n=== 자동답변 샘플 데이터 (처음 5개) ===")
    for i, (idx, row) in enumerate(auto_response_df.head().iterrows()):
        print(f"\n--- 샘플 {i+1} ---")
        print(f"Advice ID: {row.get('Advice ID', 'N/A')}")
        print(f"질문내용: {row[e_column][:100]}...")
        print(f"답변내용: {row[f_column][:100]}...")
        print(f"I열 내용: {row[i_column][:100]}...")

def create_summary_report(auto_response_df, e_column, f_column):
    """
    요약 리포트 생성
    """
    print("\n=== 요약 리포트 ===")
    
    # 기본 통계
    total_count = len(auto_response_df)
    print(f"총 자동답변 건수: {total_count}건")
    
    # 날짜별 분석 (Advice ID에서 날짜 추출)
    if 'Advice ID' in auto_response_df.columns:
        auto_response_df['date'] = auto_response_df['Advice ID'].astype(str).str[:8]
        daily_counts = auto_response_df.groupby('date').size()
        print(f"\n날짜별 자동답변 건수:")
        for date, count in daily_counts.items():
            print(f"{date}: {count}건")
    
    # 카테고리별 분석
    if 'Category' in auto_response_df.columns:
        category_counts = auto_response_df['Category'].value_counts()
        print(f"\n카테고리별 자동답변 건수:")
        for category, count in category_counts.head(10).items():
            print(f"{category}: {count}건")

if __name__ == "__main__":
    print("자동답변 데이터 추출을 시작합니다...")
    extract_auto_response_data()
    print("\n추출이 완료되었습니다.")
