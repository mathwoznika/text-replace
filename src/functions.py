from pathlib import Path

from unidecode import unidecode
import pandas as pd
import re

def subistitute_words(text: str, maping: dict[str, str]) -> str:
    """_summary_

    Args:
        text (str): _description_
        maping (dict[str, str]): _description_

    Returns:
        str: _description_
    """
    padrao = r'\b(?:{})\b'.format('|'.join(map(re.escape, maping.keys())))
    substituido = re.sub(padrao, lambda match: maping[match.group(0)], text)
    return substituido

def normalized(text: str) -> str:
    """_summary_

    Args:
        text (str): _description_

    Returns:
        str: _description_
    """
    text_normalized = unidecode(text)
    return text_normalized

def get_map(file_path: Path) -> dict[str, str]:
    """_summary_

    Args:
        file_path (Path): _description_

    Returns:
        dict[str, str]: _description_
    """
    document = pd.read_excel(file_path)
    file_dic = dict(zip(document["Current translation"], document["Correct translation"]))
    return file_dic