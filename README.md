# c-csv-2-xlsx
Efficient script to convert csv files to xlsx written in C.

This script was created to convert csv files to xlsx files with high speed.
It's nice as it will parse the csv files correctly, respecting commas, double
quotes, etc.

Feel free to use and abuse!

Kudos to following articles for help with the beautiful (:D) C language:
- Process csv: https://codingboost.com/parsing-csv-files-in-c
- Populate xlsx: http://libxlsxwriter.github.io/getting_started.html
  
Usage:
```bash
./csvtoxlsx ./input-file.csv ./output-file.xlsx my-sheet-name 
```
  
Compile C file (need to install libxlsxwriter first):
```bash
cc csvtoxlsx.c -o csvtoxlsx -lxlsxwriter
```
 
 Troubleshooting:

 If the script is bombing out with the "Can't find shared library", create env variable like this (or find where your libxlsxwriter is located):
 ```bash
 export LD_LIBRARY_PATH="/usr/local/lib:$LD_LIBRARY_PATH"
 ```

Improvements:

This script could be improved a lot, for example there could be another argument to provide a delimiter, but time is the enemy tonight, so I'm leaving it for future iterations.

Please feel free to contribute by creating a GitHub Issue or a Pull Request.

Acknowledgments:
- P.U. - for finding an issue with the bloody csv parser (do not copy code blindy from articles!)
- A.P. - for idea to use libxlsxwriter
