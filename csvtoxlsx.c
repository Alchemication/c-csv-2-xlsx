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

int main(int argc, char *argv[]) {

    // check params
    if( argc != 4) {
        fprintf(stderr, "Need to provide 3 arguments: csv-filepath, xlsx-filepath and sheetname\n");
        return 1;
    }

    // set up initial variables, max characters we are expecting on a
    // single line is 4096, this should be probably made dynamic (or even a parameter), but
    // it will be easy to adjust if required
    char buf[4096];
    char token[4096];

    int row_count = 0;
    int field_count = 0;
    int in_double_quotes = 0;
    int token_pos = 0;
    int i = 0;

    // open csv file
    FILE *fp = fopen(argv[1], "r");
    if (!fp) {
        fprintf(stderr, "Can't open file\n");
        return 1;
    }

    // create new workbook (and worksheet)
    lxw_workbook  *workbook = workbook_new(argv[2]);
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, argv[3]);

    // set format for headers
    lxw_format *format_header = workbook_add_format(workbook);
    format_set_bg_color(format_header, 0x54a0ff);
    format_set_font_color(format_header, 0xfefefe);
    format_set_align(format_header, LXW_ALIGN_CENTER);

    // set width for columns
    int col_width = 30;

    // iterate through csv file
    while (fgets(buf, 4096, fp)) {

        // reset variables for each row
        field_count = 0;
        i = 0;
        in_double_quotes = 0;

        // cycle through each line
        do {
            token[token_pos++] = buf[i];

            if (!in_double_quotes && (buf[i] == ',' || buf[i] == '\n')) {
                token[token_pos - 1] = 0;
                token_pos = 0;
                // add value to worksheet, on a first row - add format for headers
                if (row_count == 0) {
                    worksheet_set_column(worksheet, 0, field_count, col_width, NULL);
                    worksheet_write_string(worksheet, row_count, field_count++, token, format_header);
                } else {
                    worksheet_write_string(worksheet, row_count, field_count++, token, NULL);
                }
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