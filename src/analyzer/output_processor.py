import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

def add_target_month(df: pd.DataFrame, config) -> pd.DataFrame:
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
    target_sheet = 'target_month'

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
