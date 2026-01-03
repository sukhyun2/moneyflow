from src.utils.utils import read_yaml
from src.preprocessor.cleaner import clean_data, save_file, read_prepro
from src.analyzer.aggregator import create_hierarchical_summary
from src.analyzer.output_processor import (
    create_summary_by_month, filter_target_month_summary,
    create_dataframes_with_separators, add_dataframe_to_excel,
    process_asset_data,
    auto_adjust_column_width_to_output, apply_accounting_format_to_output, set_font_size_for_output
)

# Read config
# 계산할 일자, 원본 데이터 위치 등등 각종 설정을 config 파일로 제어
config_path = 'config/config.yaml'
config = read_yaml(config_path)

# Cleaning data & Save
# 입력 엑셀 파일을 읽어 분석 가능한 형태로 정제 & prepro 경로에 이력 저장, 경로에 동일 파일 존재시 overwrite됨.
pdf_prepro_target_date = clean_data(config)
save_file(pdf_prepro_target_date, config, 'prepro')

# Load all data (with past data)
# target date 뿐만 아니고 그 이전 파일까지 한꺼번에 불러옴
pdf_prepro = read_prepro(config)

# Aggregation data & create output data
# 최종 output 계산, 경로에 동일 파일 존재시 overwrite됨.
pdf_agg = create_hierarchical_summary(pdf_prepro) # final output
save_file(pdf_agg.reset_index(), config, 'output')

# Output data processing (append)
# target_month와 전월 데이터를 필터링하고 최종 파일에 별도 시트로 추가
pdf_summ_type_all, pdf_summ_small_all = create_summary_by_month(pdf_agg) # all date
pdf_summ_type_tar, pdf_summ_small_tar = filter_target_month_summary(pdf_summ_type_all, pdf_summ_small_all, config)
pdf_tar = create_dataframes_with_separators([pdf_summ_type_tar, pdf_summ_small_tar])
add_dataframe_to_excel(
    pdf_tar,
    config['output_path'] + '/' + config['output_file_name'],
    sheet_name='target_month_summary'
)

# Asset data processing (append)
# 자산 데이터를 불러와 피벗테이블로 변환하고 최종 파일에 별도 시트로 추가
process_asset_data(config)

# Cleaning format
# 읽기 편한 형식으로 엑셀 파일 수정
set_font_size_for_output(config, 15)
apply_accounting_format_to_output(config)
auto_adjust_column_width_to_output(config)