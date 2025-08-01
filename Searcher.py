import argparse
from pathlib import Path
import sys
from pyxlsb import open_workbook
import string

class Searcher:

    @staticmethod
    def normalize(text):
        """
        Uppercases text and removes punctuation; keeps spaces
        """

        if not isinstance(text, str):
            text=str(text)
        return text.upper().translate(str.maketrans("", "", string.punctuation))

    @staticmethod
    def readExcelFile(file_path, keyword):
        """
        Searches an Excel file from the given path for a keyword
        """
        
        norm_key=Searcher.normalize(keyword)
        found=False

        with open_workbook(file_path) as wb:
            for sheet in wb.sheets:
                with wb.get_sheet(sheet) as s:
                    for row_idx, row in enumerate(s.rows(), 1):
                        for col_idx, cell in enumerate(row, 1):
                            if cell.v is not None:
                                cell_string=Searcher.normalize(cell.v)
                                if norm_key in cell_string:
                                    print(f"Keyword found in '{file_path}', sheet '{sheet}', row {row_idx}, col {col_idx}: {cell.v}")
                                    found=True

        if not found:
            print(f"Keyword not found in '{file_path}'")

    def main():
   
        parser=argparse.ArgumentParser(description="Searches the input file for the input keyword.")
 
        # Syntax of the file path
        parser.add_argument("-f", "--filePath", type=str, help="The file to search from")
        # Syntax of the keyword
        parser.add_argument("-k", "--keyword", nargs='+', help="The keyword to search for (can be multiple words)")
    
        args = parser.parse_args()
        
        filePath = Path(args.filePath)
       
        try: 
            if filePath.exists():

                keyword = ' '.join(args.keyword).upper()
                print(f"You entered the keyword: {keyword}")

                Searcher.readExcelFile(filePath, keyword)
            else:
                raise FileNotFoundError(f"The file was not found at the specified path: {filePath}")

        except Exception as e:
            print(f"Error: {e}")

        

if __name__ == "__main__":
    Searcher.main()


