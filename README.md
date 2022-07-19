# ExcelComparator
Compares items in excel and finds elements in excel A that not exists in B and viceversa, and finds elements existents in both excel files

usage: excel_inspect.py [-h] [-a BOOK_A] [-b BOOK_B] [-v] [-o OUTPUT_FILE] [-f]

optional arguments:
  -h, --help            show this help message and exit
  
  -a BOOK_A, --book-a BOOK_A
                        Excel A (default: None)
                        
  -b BOOK_B, --book-b BOOK_B
                        Excel B (default: "")
                        
  -v, --verbose         Show results in console too (default: False)
  
  -o OUTPUT_FILE, --output-file OUTPUT_FILE
                        Output filename with results (default: "")
                        
  -f, --force-overwrite
                        Force overwrite output file if already exists (default: False)
                        
