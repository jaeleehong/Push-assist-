import pandas as pd

# 1. 엑셀 파일의 시트명 확인
excel_file = './cs자동답변_3차.xlsx'
xl = pd.ExcelFile(excel_file)
print(f"엑셀 파일의 시트명들: {xl.sheet_names}")

# 2. 추출할 자동답변 항목 리스트 정의
target_list = [
    '자동답변: 욕설 및 비속어 표현이 90% 이상 포함된 문의',
    '자동답변 : 욕설 및 비속어 표현이 90% 이상 포함된 문의',
    '자동답변: 광고와 홍보 스팸 문의 처리',
    '자동답변 : 광고와 홍보 스팸 문의 처리'
]

# 3. 모든 시트의 데이터를 하나로 합치기
all_data = []
sheet_summary = []
sheet_data = {}  # 시트별 데이터 저장

for sheet_name in xl.sheet_names:
    print(f"\n=== {sheet_name} 시트 처리 중 ===")
    
    # 시트별 데이터 읽기
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    print(f"전체 데이터 행 수: {len(df)}")
    
    # '요약 결과' 정리 - value.result 부분 추출
    df['요약 결과_정리'] = df['요약 결과'].str.extract(r'"value":\{"result":"([^"]*)"')
    
    # 대상 항목 필터링
    filtered_df = df[df['요약 결과_정리'].isin(target_list)]
    print(f"추출된 데이터 행 수: {len(filtered_df)}")
    
    # 시트별 요약 정보 저장
    sheet_summary.append({
        '시트명': sheet_name,
        '전체 데이터': len(df),
        '추출된 데이터': len(filtered_df)
    })
    
    # 필요한 컬럼만 선택하고 시트명 추가 (집계용)
    result_df = filtered_df[['Advice ID', '질문내용', '요약 결과', '요약 결과_정리']].copy()
    result_df['시트명'] = sheet_name
    
    # 날짜 추출 (Advice ID에서)
    result_df['날짜'] = result_df['Advice ID'].str[:8]
    
    all_data.append(result_df)
    
    # 시트별 상세 데이터 저장 (추출용) - I열(요약 결과) 사용
    if len(filtered_df) > 0:
        sheet_data[sheet_name] = filtered_df[['Advice ID', 'Title', '질문내용', '요약 결과']]

# 4. 모든 데이터 합치기
combined_df = pd.concat(all_data, ignore_index=True)
print(f"\n전체 합쳐진 데이터: {len(combined_df)}개")

# 5. 시트별 집계
sheet_summary_df = pd.DataFrame(sheet_summary)
print("\n=== 시트별 집계 결과 ===")
print(sheet_summary_df)

# 6. 일자별 집계
# 일자별 전체 건수
total_by_date = combined_df.groupby('날짜').size().reset_index(name='전체 건수')

# 일자별, 항목별 건수
summary_by_date = combined_df.groupby(['날짜', '요약 결과_정리']).size().unstack(fill_value=0)

# 7. 결과 합치기
final_table = pd.merge(total_by_date, summary_by_date, on='날짜', how='outer')

# 8. 결과 저장
output_file = 'CS_자동답변_집계결과.xlsx'
with pd.ExcelWriter(output_file) as writer:
    # 시트별 집계
    sheet_summary_df.to_excel(writer, sheet_name='시트별 집계', index=False)
    
    # 일자별 집계
    final_table.to_excel(writer, sheet_name='일자별 집계', index=False)
    
    # 전체 요약 시트 추가
    summary_stats = combined_df['요약 결과_정리'].value_counts().reset_index()
    summary_stats.columns = ['자동답변 항목', '총 건수']
    summary_stats.to_excel(writer, sheet_name='전체 요약', index=False)
    
    # 시트별 상세 데이터 저장
    for sheet_name, data in sheet_data.items():
        data.to_excel(writer, sheet_name=f'{sheet_name}_추출', index=False)
        print(f"{sheet_name} 시트 상세 데이터 저장: {len(data)}개")

print(f"\n처리 완료! {output_file} 파일이 생성되었습니다.")

# 9. 결과 출력
print("\n=== 일자별 집계 결과 ===")
print(final_table)

print("\n=== 전체 요약 ===")
print(summary_stats)
print(f"\n총 추출된 데이터: {len(combined_df)}개")
