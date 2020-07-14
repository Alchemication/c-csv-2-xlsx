#include <stdio.h>
#include <string.h>
#include "xlsxwriter.h"

/**
 * This script was created to convert csv files to xlsx files with high speed.
 * It's nice as it will parse the csv files correctly, respecting comas, double
 * quotes, etc.
 * 
 * Feel free to use and abuse,
 * 
 * Kudos to following articles for help with the beautiful (:D) C language:
 * - Process csv: https://codingboost.com/parsing-csv-files-in-c
 * - Populate xlsx: http://libxlsxwriter.github.io/getting_started.html
 * 
 * Usage:
 * ./csvtoxlsx test.csv test10.xlsx
 * 
 * Compile C file (need to install libxlsxwriter first):
 * cc csvtoxlsx.c -o csvtoxlsx -lxlsxwriter
 **/

void add_to_sheet(lxw_worksheet *worksheet, int row_count, int col_count, char *value) {
    // printf("Row number: %d, col number: %d, value: %s", row_count, col_count, value);
    // printf("\n");
    worksheet_write_string(worksheet, row_count, col_count, value, NULL);
}

int main(int argc, char *argv[]) {

    // check params
    if( argc < 2 || argc > 4 ) {
        printf("Need to provide 2 arguments: input file-path-name and output file-path-name\n");
        return 1;
    }

    // set up initial variables
    char buf[1024];
    char token[1024];

    int row_count = 0;
    int field_count = 0;
    int in_double_quotes = 0;
    int token_pos = 0;
    int i = 0;

    // open csv file
    FILE *fp = fopen(argv[1], "r");
    if (!fp) {
        printf("Can't open file\n");
        return 1;
    }

    // create new workbook (and worksheet)
    lxw_workbook  *workbook = workbook_new(argv[2]);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    // iterate through csv file
    while (fgets(buf, 1024, fp)) {

        // cycle through each line
        field_count = 0;
        i = 0;
        do {
            token[token_pos++] = buf[i];

            if (!in_double_quotes && (buf[i] == ',' || buf[i] == '\n')) {
                token[token_pos - 1] = 0;
                token_pos = 0;
                // add value to worksheet
                add_to_sheet(worksheet, row_count, field_count++, token);
            }

            if (buf[i] == '"' && buf[i + 1] != '"') {
                token_pos--;
                in_double_quotes = !in_double_quotes;
            }

            if (buf[i] == '"' && buf[i + 1] == '"')
                i++;

        } while (buf[++i]);

        row_count++;
    }

    // close csv file
    fclose(fp);

    // close excel file
    workbook_close(workbook);

    // return success
    return 0;
}