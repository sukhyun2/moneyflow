import pandas as pd

def create_income_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    수입 데이터에 대한 계층적 집계 (월별 + 대분류까지만)

    Args:
        df: target_data DataFrame (month 컬럼 포함)

    Returns:
        pd.DataFrame: 수입 데이터 계층적 집계 결과 (월별)
    """
    # 수입 데이터만 필터링
    income_df = df[df['타입'] == '수입'].copy()

    if len(income_df) == 0:
        return pd.DataFrame()

    all_levels = []

    # Level 1: 월별 수입 총계
    level1 = income_df.groupby('month')['금액'].agg(['sum', 'count', 'mean'])
    level1.index = pd.MultiIndex.from_tuples(
        [(x, '수입', '', '', '') for x in level1.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level1)

    # Level 2: 월별 + 대분류별 수입 집계
    level2 = income_df.groupby(['month', '대분류'])['금액'].agg(['sum', 'count', 'mean'])
    level2.index = pd.MultiIndex.from_tuples(
        [(x[0], '수입', x[1], '', '') for x in level2.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level2)

    # 모든 레벨 합치기
    combined = pd.concat(all_levels)
    combined.columns = ['금액합계', '거래건수', '평균금액']

    return combined


def create_expense_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    지출 데이터에 대한 계층적 집계 (월별 + 대분류-소분류-내용)

    Args:
        df: target_data DataFrame (month 컬럼 포함)

    Returns:
        pd.DataFrame: 지출 데이터 계층적 집계 결과 (월별)
    """
    # 지출/이체 데이터만 필터링
    expense_df = df[df['타입'].isin(['지출', '이체'])].copy()

    if len(expense_df) == 0:
        return pd.DataFrame()

    all_levels = []

    # Level 1: 월별 지출 총계
    level1 = expense_df.groupby('month')['금액'].agg(['sum', 'count', 'mean'])
    level1.index = pd.MultiIndex.from_tuples(
        [(x, '지출', '', '', '') for x in level1.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level1)

    # Level 2: 월별 + 대분류별 지출
    level2 = expense_df.groupby(['month', '대분류'])['금액'].agg(['sum', 'count', 'mean'])
    level2.index = pd.MultiIndex.from_tuples(
        [(x[0], '지출', x[1], '', '') for x in level2.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level2)

    # Level 3: 월별 + 대분류 + 소분류별 지출
    level3 = expense_df.groupby(['month', '대분류', '소분류'])['금액'].agg(['sum', 'count', 'mean'])
    level3.index = pd.MultiIndex.from_tuples(
        [(x[0], '지출', x[1], x[2], '') for x in level3.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level3)

    # Level 4: 월별 + 대분류 + 소분류 + 내용별 지출 (최상세 레벨)
    level4 = expense_df.groupby(['month', '대분류', '소분류', '내용'])['금액'].agg(['sum', 'count', 'mean'])
    level4.index = pd.MultiIndex.from_tuples(
        [(x[0], '지출', x[1], x[2], x[3]) for x in level4.index],
        names=['month', '타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level4)

    # 모든 레벨 합치기
    combined = pd.concat(all_levels)
    combined.columns = ['금액합계', '거래건수', '평균금액']

    return combined


def create_hierarchical_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    MultiIndex를 사용한 계층적 집계 (월별로 수입과 지출을 분리하여 분석)

    Args:
        df: target_data DataFrame (month 컬럼 포함)

    Returns:
        pd.DataFrame: MultiIndex로 계층화된 집계 결과 (월별)
    """
    # 수입 요약
    income_summary = create_income_summary(df)

    # 지출 요약
    expense_summary = create_expense_summary(df)

    # 결과 합치기
    if len(income_summary) > 0 and len(expense_summary) > 0:
        combined = pd.concat([income_summary, expense_summary])
    elif len(income_summary) > 0:
        combined = income_summary
    elif len(expense_summary) > 0:
        combined = expense_summary
    else:
        return pd.DataFrame()

    # 정렬: month는 내림차순, 나머지는 오름차순
    combined = combined.sort_index(
        level=['month', '타입', '대분류', '소분류', '내용'],
        ascending=[False, True, True, True, True]
    )

    # 숫자 포맷팅
    combined['금액합계'] = combined['금액합계'].apply(lambda x: f"{int(x):,}")
    combined['거래건수'] = combined['거래건수'].apply(lambda x: f"{x:.1f}")
    combined['평균금액'] = combined['평균금액'].apply(lambda x: f"{int(x):,}")

    return combined