import pandas as pd

def create_income_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    수입 데이터에 대한 계층적 집계 (대분류까지만)

    Args:
        df: target_data DataFrame

    Returns:
        pd.DataFrame: 수입 데이터 계층적 집계 결과
    """
    # 수입 데이터만 필터링
    income_df = df[df['타입'] == '수입'].copy()

    if len(income_df) == 0:
        return pd.DataFrame()

    all_levels = []

    # Level 1: 수입 총계
    level1 = income_df['금액'].agg(['sum', 'count', 'mean'])
    level1_df = pd.DataFrame([level1],
                            index=pd.MultiIndex.from_tuples([('수입', '', '', '')],
                                                           names=['타입', '대분류', '소분류', '내용']))
    all_levels.append(level1_df)

    # Level 2: 수입 + 대분류 (소분류, 내용은 의미없으므로 여기서 종료)
    level2 = income_df.groupby('대분류')['금액'].agg(['sum', 'count', 'mean'])
    level2.index = pd.MultiIndex.from_tuples(
        [('수입', x, '', '') for x in level2.index],
        names=['타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level2)

    # 모든 레벨 합치기
    combined = pd.concat(all_levels)
    combined.columns = ['금액합계', '거래건수', '평균금액']

    return combined


def create_expense_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    지출 데이터에 대한 계층적 집계 (대분류-소분류-내용)

    Args:
        df: target_data DataFrame

    Returns:
        pd.DataFrame: 지출 데이터 계층적 집계 결과
    """
    # 지출/이체 데이터만 필터링
    expense_df = df[df['타입'].isin(['지출', '이체'])].copy()

    if len(expense_df) == 0:
        return pd.DataFrame()

    all_levels = []

    # Level 1: 지출 총계
    level1 = expense_df['금액'].agg(['sum', 'count', 'mean'])
    level1_df = pd.DataFrame([level1],
                            index=pd.MultiIndex.from_tuples([('지출', '', '', '')],
                                                           names=['타입', '대분류', '소분류', '내용']))
    all_levels.append(level1_df)

    # Level 2: 지출 + 대분류
    level2 = expense_df.groupby('대분류')['금액'].agg(['sum', 'count', 'mean'])
    level2.index = pd.MultiIndex.from_tuples(
        [('지출', x, '', '') for x in level2.index],
        names=['타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level2)

    # Level 3: 지출 + 대분류 + 소분류
    level3 = expense_df.groupby(['대분류', '소분류'])['금액'].agg(['sum', 'count', 'mean'])
    level3.index = pd.MultiIndex.from_tuples(
        [('지출', x[0], x[1], '') for x in level3.index],
        names=['타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level3)

    # Level 4: 지출 + 대분류 + 소분류 + 내용 (최상세 레벨)
    level4 = expense_df.groupby(['대분류', '소분류', '내용'])['금액'].agg(['sum', 'count', 'mean'])
    level4.index = pd.MultiIndex.from_tuples(
        [('지출', x[0], x[1], x[2]) for x in level4.index],
        names=['타입', '대분류', '소분류', '내용']
    )
    all_levels.append(level4)

    # 모든 레벨 합치기
    combined = pd.concat(all_levels)
    combined.columns = ['금액합계', '거래건수', '평균금액']

    return combined


def create_hierarchical_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    MultiIndex를 사용한 계층적 집계 (수입과 지출을 분리하여 분석)

    Args:
        df: target_data DataFrame

    Returns:
        pd.DataFrame: MultiIndex로 계층화된 집계 결과
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

    # 정렬
    combined = combined.sort_index()

    # 숫자 포맷팅
    combined['금액합계'] = combined['금액합계'].apply(lambda x: f"{int(x):,}")
    combined['거래건수'] = combined['거래건수'].apply(lambda x: f"{x:.1f}")
    combined['평균금액'] = combined['평균금액'].apply(lambda x: f"{int(x):,}")

    return combined