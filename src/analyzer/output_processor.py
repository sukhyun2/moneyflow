import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

def add_target_month(df: pd.DataFrame, config) -> pd.DataFrame:
    """config의 target_month에 해당하는 데이터를 필터링하고 target_month_summary 시트로 저장하는 함수"""
    # config의 target_month(YYYY-MM-DD)를 YYYY-MM 형식으로 변환하여 필터링
    target_month_str = config['target_month']
    if isinstance(target_month_str, str):
        target_month = datetime.strptime(target_month_str, '%Y-%m-%d').date()
    else:
        target_month = target_month_str

    target_month_yyyy_mm = target_month.strftime('%Y-%m')

    # df에서 해당 월 데이터만 필터링 (month가 MultiIndex의 첫 번째 레벨이라고 가정)
    df_filtered = df[df.index.get_level_values('month') == target_month_yyyy_mm]

    file_path = config['output_path'] + '/' + config['output_file_name']
    target_sheet = 'target_month_summary'

    # --- 엑셀 파일을 불러오고 target_month 존재 시 삭제 ---
    try:
        wb = load_workbook(file_path)
        if target_sheet in wb.sheetnames:
            del wb[target_sheet]  # 기존 시트 삭제
        wb.save(file_path)
    except FileNotFoundError:
        # 파일이 없으면 새로 생성될 것이므로 그냥 통과
        pass

    with pd.ExcelWriter(
            file_path,
            engine='openpyxl', mode='a', if_sheet_exists='new'
    ) as writer:
        # 첫 번째 데이터프레임 (대분류가 빈 값인 것들)
        df1 = df_filtered[df_filtered.index.get_level_values('대분류') == ''].reset_index()

        # 두 번째 데이터프레임 (대분류는 있고 소분류가 빈 값인 것들)
        df2 = df_filtered[
            (df_filtered.index.get_level_values('대분류') != '') &
            (df_filtered.index.get_level_values('소분류') == '')
        ].reset_index()
        df2 = df2.sort_values(by=['타입', '금액합계'], ascending=[True, True])

        # 여백을 위한 빈 행 DataFrame 생성 (첫 번째 DF와 같은 컬럼 구조)
        if len(df1) > 0:
            empty_rows = pd.DataFrame([[None] * len(df1.columns)], columns=df1.columns)
        elif len(df2) > 0:
            empty_rows = pd.DataFrame([[None] * len(df2.columns)], columns=df2.columns)
        else:
            empty_rows = pd.DataFrame()

        # 데이터프레임들을 여백과 함께 결합
        df_tmp = pd.concat([df1, empty_rows, df2], ignore_index=True)
        df_tmp.to_excel(writer, sheet_name=target_sheet, index=False)

    return df_filtered


def process_asset_data(config) -> pd.DataFrame:
    """asset.xlsx 파일을 읽어와서 pivot table로 변환하고 엑셀 파일에 저장하는 함수"""
    # asset.xlsx 파일 읽기
    asset_file_path = config['input_path'] + '/' + config['asset_file_name']

    try:
        pdf = pd.read_excel(asset_file_path)
    except FileNotFoundError:
        print(f"Asset file not found: {asset_file_path}")
        return pd.DataFrame()

    # 필수 컬럼 확인
    required_columns = ['날짜', '카테고리', '세부항목', '금액']
    missing_columns = [col for col in required_columns if col not in pdf.columns]
    if missing_columns:
        print(f"Missing required columns in asset.xlsx: {missing_columns}")
        return pd.DataFrame()

    # 날짜 컬럼을 datetime으로 변환 후 yyyy-mm 형식으로 변경
    pdf['날짜'] = pd.to_datetime(pdf['날짜']).dt.strftime('%Y-%m')

    # pivot table 생성
    pdf_pivot = pdf.pivot_table(
        index=['카테고리', '세부항목'],
        columns='날짜',
        values='금액',
        aggfunc='sum'
    )

    # 출력 파일 경로 설정
    output_file_path = config['output_path'] + '/' + config['output_file_name']
    asset_sheet = 'asset_summary'

    # 기존 asset_summary 시트가 있으면 삭제
    try:
        wb = load_workbook(output_file_path)
        if asset_sheet in wb.sheetnames:
            del wb[asset_sheet]
        wb.save(output_file_path)
    except FileNotFoundError:
        # 파일이 없으면 새로 생성될 것이므로 통과
        pass

    # 엑셀 파일에 새 시트로 저장
    with pd.ExcelWriter(
        output_file_path,
        engine='openpyxl',
        mode='a',
        if_sheet_exists='new'
    ) as writer:
        pdf_pivot.to_excel(writer, sheet_name=asset_sheet)

    print(f"Asset data processed and saved to sheet: {asset_sheet}")
    return pdf_pivot


def auto_adjust_column_width(file_path: str):
    """엑셀 파일의 모든 컬럼 너비를 자동으로 조정하는 함수"""
    try:
        wb = load_workbook(file_path)

        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]

            # 각 컬럼의 너비를 자동 조정
            for column in sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter

                for cell in column:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))

                # 너비 조정 (최소 10, 최대 50)
                adjusted_width = min(max(max_length + 2, 10), 50)
                sheet.column_dimensions[column_letter].width = adjusted_width

        wb.save(file_path)
        print(f"Column width adjusted for: {file_path}")
        return True

    except Exception as e:
        print(f"Error adjusting column width: {e}")
        return False


def auto_adjust_column_width_to_output(config):
    """config에서 지정된 출력 파일에 컬럼 너비 자동 조정을 적용하는 함수"""
    output_file_path = config['output_path'] + '/' + config['output_file_name']
    return auto_adjust_column_width(output_file_path)


def apply_accounting_format(file_path: str):
    """지정된 엑셀 파일의 모든 시트에서 숫자값을 회계 형식으로 변경하는 함수"""
    try:
        # 엑셀 파일 로드
        wb = load_workbook(file_path)

        # 회계 형식 정의 (천 단위 구분자, 음수는 괄호 표시)
        accounting_format = "#,##0_);(#,##0)"

        # 모든 시트 순회
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            print(f"Processing accounting format for sheet: {sheet_name}")

            # 모든 셀 순회하여 숫자 포맷 적용
            for row in sheet.iter_rows():
                for cell in row:
                    # 셀에 값이 있고 숫자인 경우에만 포맷 적용
                    if cell.value is not None and isinstance(cell.value, (int, float)):
                        cell.number_format = accounting_format

        # 파일 저장
        wb.save(file_path)
        print(f"Accounting format applied successfully to: {file_path}")
        return True

    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return False
    except Exception as e:
        print(f"Error applying accounting format: {e}")
        return False


def apply_accounting_format_to_output(config):
    """config에서 지정된 출력 파일에 회계 형식을 적용하는 함수
    """
    output_file_path = config['output_path'] + '/' + config['output_file_name']
    return apply_accounting_format(output_file_path)
