import yaml


def read_yaml(file_path):
    """
    YAML 파일을 읽어서 파이썬 객체로 반환합니다.

    Args:
        file_path (str): YAML 파일 경로

    Returns:
        dict or None: YAML 데이터 또는 에러 시 None
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            data = yaml.safe_load(file)
        return data

    except FileNotFoundError:
        print(f"파일을 찾을 수 없습니다: {file_path}")
        return None
    except yaml.YAMLError as e:
        print(f"YAML 파싱 오류: {e}")
        return None