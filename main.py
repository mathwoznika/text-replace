from pathlib import Path
import os
from src.functions import excel_substitution


def main(data_path: Path, column_name: str, from_to_path: Path, from_column: str,
         to_column: str, normalize: bool = False):
    """Applying, from an Excel FROM-TO, the substitution of the words contained in 
    the spreadsheet.

    Args:
        data_path (Path): Path of the data folder.
        column_name (str): Name of the column you need to change.
        from_to_path (Path): Path of the FROM-TO file.
        from_column (str): Name of the FROM column, this column contains the word will be
        changed.
        to_column (str): Name of the TO column, this column contains the word that 
        replaced the old one.
        normalize (bool, optional): If you need to normalize the words, normalize will be 
        True, else will be False. Defaults to False.
    """

    for file_path in data_path.rglob("*xlsx"):
        print(file_path)
        excel_substitution(file_path, column_name, from_to_path, from_column, 
                            to_column, normalize=normalize)
                

if __name__ == "__main__":

    data_path = Path(__file__).parent / "data"
    column_name = "Translated"
    from_to_path = Path(__file__).parent / "data" /"Interface translations Chinese - document control, training management and archive.xlsx"
    from_column = "Current translation"
    to_column = "Correct translation"

    main(data_path, column_name, from_to_path, from_column, to_column, normalize=False)

