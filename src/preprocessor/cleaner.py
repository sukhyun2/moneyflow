"""
DataFrame 데이터 정제 및 전처리 모듈
"""

import pandas as pd
from datetime import datetime
from typing import Dict, Any
import os


def convert_datetime64_to_datetime(df: pd.DataFrame) -> pd.DataFrame:
    """
    DataFrame에서 datetime64[ns] 타입 컬럼들을 자동으로 찾아서 datetime.date으로 변환합니다.

    Args:
        df (pd.DataFrame): 변환할 DataFrame

    Returns:
        pd.DataFrame: datetime 컬럼이 date로 변환된 DataFrame
    """
    result_df = df.copy()
    datetime_columns = []

    # datetime64 타입의 컬럼들을 찾기
    for col in result_df.columns:
        if pd.api.types.is_datetime64_any_dtype(result_df[col]):
            datetime_columns.append(col)

    # 각 datetime 컬럼을 변환
    for col in datetime_columns:
        # 변환 전 타입 확인
        original_dtype = result_df[col].dtype
        print(f"convert_datetime64_to_datetime: 변환 전 = {original_dtype}")

        # datetime64[ns]를 datetime.date로 변환
        converted_values = []
        for dt_val in result_df[col]:
            if pd.isna(dt_val):
                converted_values.append(None)
            elif isinstance(dt_val, pd.Timestamp):
                converted_values.append(dt_val.to_pydatetime().date())
            else:
                converted_values.append(dt_val)

        # 명시적으로 object 타입으로 Series 생성
        result_df[col] = pd.Series(converted_values, index=result_df.index, dtype='object')

        # 변환 후 타입 확인
        new_dtype = result_df[col].dtype
        sample_type = type(result_df[col].iloc[0]) if len(result_df) > 0 and not pd.isna(result_df[col].iloc[0]) else None
        print(f"convert_datetime64_to_datetime: 변환 후 = {new_dtype}, 실제 값 타입 = {sample_type}")

    if not datetime_columns:
        print("convert_datetime64_to_datetime: datetime64 타입의 컬럼이 없습니다.")

    return result_df


def clean_data(config: Dict[str, Any]) -> pd.DataFrame:
    """
    Excel 파일을 읽어서 가계부 데이터를 정제하고 필터링한 후 CSV로 저장합니다.

    Args:
        config (Dict[str, Any]): 설정 딕셔너리 (config.yaml에서 로드)
            - input_path: 입력 파일 경로
            - input_file_name: 입력 파일명
            - sheet_name: Excel 시트명
            - preprocessed_path: 전처리 파일 저장 경로
            - preprocessed_file_name: 전처리 파일명 (날짜 치환 지원)
            - output_path: 최종 출력 파일 저장 경로
            - output_file_name: 최종 출력 파일명 (날짜 치환 지원)
            - column_names: 사용할 컬럼명 리스트
            - income_sources: 포함할 수입 대분류 리스트
            - payment_methods: 포함할 결제수단 리스트
            - exclude_large_cat: 제외할 대분류 카테고리 리스트
            - target_month: 분석 시작 날짜 (YYYY-MM-DD 문자열)

    Returns:
        pd.DataFrame: 정제되고 필터링된 가계부 데이터

    Process:
        0. 파일 경로 생성 및 Excel 파일 읽기
        1. config에서 설정값들 추출 및 날짜 변환
        2. 필요한 컬럼만 추출
        3. datetime 컬럼을 date 타입으로 변환
        4. 지정된 기간의 데이터만 필터링
        5. 제외할 대분류 카테고리 제거
        6. 수입 데이터 필터링 (타입='수입', 대분류 in income_sources)
        7. 지출 데이터 필터링 (타입 in ['지출','이체'], 결제수단 in payment_methods)
        8. 수입과 지출 데이터 합치기
        9. 전처리된 데이터를 CSV 파일로 저장
    """

    # Step 0: 파일 경로 생성 및 Excel 파일 읽기
    print('clean_data: Excel 파일을 읽어옵니다.')
    input_file_path = os.path.join(config['input_path'], config['input_file_name'])
    sheet_name = config['sheet_name']

    print(f'  - 파일 경로: {input_file_path}')
    print(f'  - 시트명: {sheet_name}')

    try:
        df = pd.read_excel(input_file_path, sheet_name=sheet_name)
        print(f'  - 파일 읽기 성공: {df.shape}')
    except Exception as e:
        print(f'  X 파일 읽기 실패: {e}')
        raise

    # Step 1: config에서 설정값들 추출
    print('clean_data: config에서 설정값들을 추출합니다.')
    column_names = config['column_names']
    income_sources = config['income_sources']
    payment_methods = config['payment_methods']
    exclude_large_cat = config['exclude_large_cat']

    # target_month 문자열을 date로 변환
    target_month_str = config['target_month']
    if isinstance(target_month_str, str):
        target_month = datetime.strptime(target_month_str, '%Y-%m-%d').date()
    else:
        target_month = target_month_str

    # target_month_next 계산 (다음 달 1일)
    target_month_next = target_month.replace(day=1)
    if target_month_next.month == 12:
        target_month_next = target_month_next.replace(year=target_month_next.year + 1, month=1)
    else:
        target_month_next = target_month_next.replace(month=target_month_next.month + 1)

    print(f'  - 분석 기간: {target_month} ~ {target_month_next}')
    print(f'  - 컬럼: {len(column_names)}개')
    print(f'  - 수입원: {income_sources}')
    print(f'  - 결제수단: {payment_methods}')
    print(f'  - 제외 카테고리: {exclude_large_cat}')

    # Step 1: 컬럼 서브셋 생성 - 분석에 필요한 컬럼만 추출
    print(f'clean_data: raw xlsx을 {column_names} 으로 subset 합니다.')
    df_clnd = df[column_names]
    print(f'  - 원본 {df.shape} - 서브셋 {df_clnd.shape}')

    # Step 2: datetime 컬럼을 date 타입으로 변환
    print('clean_data: datetime64[ns] - datetime.date 변환을 수행합니다.')
    df_clnd = convert_datetime64_to_datetime(df_clnd)

    # Step 3: 대상 기간 필터링 - 지정된 월 범위의 데이터만 유지
    print(f'clean_data: target date만 남깁니다. {target_month} ~ {target_month_next}')
    original_count = len(df_clnd)
    df_clnd = df_clnd[
        (df_clnd['날짜'] >= target_month) &
        (df_clnd['날짜'] < target_month_next)
    ].copy()
    print(f'  - 기간 필터링: {original_count}건 - {len(df_clnd)}건')

    # Step 4: 제외 대분류 카테고리 필터링
    print(f'clean_data: 대분류 카테고리가 {exclude_large_cat}인 케이스를 제거합니다.')
    category_filtered_count = len(df_clnd)
    df_clnd = df_clnd[
        (df_clnd['대분류'].isin(exclude_large_cat) == False)
    ].copy()
    print(f'  - 카테고리 필터링: {category_filtered_count}건 - {len(df_clnd)}건')

    # Step 5: 수입 데이터 필터링 - 지정된 수입원만 포함
    print(f'clean_data: 수입이 {income_sources}인 케이스만 남깁니다.')
    df_in = df_clnd[
        (df_clnd['타입'] == '수입') &
        (df_clnd['대분류'].isin(income_sources))
    ].copy()
    print(f'  - 수입 데이터: {len(df_in)}건')

    # Step 6: 지출/이체 데이터 필터링 - 지정된 결제수단만 포함
    print(f'clean_data: {payment_methods}로 지출한 내역과 이체만 남깁니다.')
    df_out = df_clnd[
        (df_clnd['타입'].isin(['지출', '이체'])) &
        (df_clnd['결제수단'].isin(payment_methods))
    ].copy()
    print(f'  - 지출/이체 데이터: {len(df_out)}건')

    # Step 7: 수입과 지출 데이터 합치기
    print('clean_data: 수입과 지출 데이터를 합칩니다.')
    df_concat = pd.concat([df_in, df_out], axis=0)
    df_concat = df_concat.reset_index(drop=True)

    print(f'clean_data: 최종 정제 완료. 총 {len(df_concat)}건의 데이터')

    # Step 8: 전처리된 데이터를 CSV 파일로 저장
    print('clean_data: 전처리된 데이터를 CSV 파일로 저장합니다.')

    # preprocessed 파일명에서 {date} 치환
    current_date = datetime.now().strftime("%Y%m%d_%H%M")
    preprocessed_file_name = config['preprocessed_file_name'].replace('{date}', current_date)
    preprocessed_file_path = os.path.join(config['preprocessed_path'], preprocessed_file_name)

    # 디렉터리가 없으면 생성
    os.makedirs(config['preprocessed_path'], exist_ok=True)

    try:
        df_concat.to_csv(preprocessed_file_path, index=False, encoding='utf-8-sig')
        print(f'  - 저장 완료: {preprocessed_file_path}')
        print(f'  - 저장된 데이터: {len(df_concat)}행 x {len(df_concat.columns)}열')
    except Exception as e:
        print(f'  X 파일 저장 실패: {e}')
        raise

    return df_concat