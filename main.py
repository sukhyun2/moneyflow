from src.utils.utils import read_yaml
from src.preprocessor.cleaner import clean_data, save_file, read_prepro
from src.analyzer.aggregator import create_hierarchical_summary
from src.analyzer.output_processor import add_target_month


# read config
config_path = 'config/config.yaml'
config = read_yaml(config_path)

# cleaning data
pdf_prepro_temp = clean_data(config)

# save prepro
if config['save_temp_file']:
    save_file(pdf_prepro_temp, config, 'temp')
save_file(pdf_prepro_temp, config, 'prepro')

# load all data (with past data)
pdf_prepro = read_prepro(config)

# aggregation data & create output data
pdf_agg = create_hierarchical_summary(pdf_prepro)
save_file(pdf_agg.reset_index(), config, 'output')

# output data processing
add_target_month(pdf_agg, config)

print(pdf_agg)