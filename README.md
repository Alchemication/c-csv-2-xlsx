# c-csv-2-xlsx
Efficient script to convert csv files to xlsx


This script was created to convert csv files to xlsx files with high speed.
It's nice as it will parse the csv files correctly, respecting comas, double
quotes, etc.
  
Feel free to use and abuse!
  
Kudos to following articles for help with the beautiful (:D) C language:
- Process csv: https://codingboost.com/parsing-csv-files-in-c
- Populate xlsx: http://libxlsxwriter.github.io/getting_started.html
  
Usage:
```bash
./csvtoxlsx test.csv test10.xlsx
```
  
Compile C file (need to install libxlsxwriter first):
```bash
cc csvtoxlsx.c -o csvtoxlsx -lxlsxwriter
```
 