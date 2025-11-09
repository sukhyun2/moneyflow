from src.utils.utils import read_yaml
from src.preprocessor.cleaner import clean_data
from src.analyzer.aggregator import create_hierarchical_summary


# read config
config_path = 'config/config.yaml'
config = read_yaml(config_path)

# cleaning data
pdf_prepro = clean_data(config)

# aggregation data
pdf_result = create_hierarchical_summary(pdf_prepro)

print(pdf_result)