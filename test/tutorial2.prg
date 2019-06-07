/*
 * A simple program to write some data to an Excel file using the
 * libxlsxwriter library.
 *
 * This program is shown, with explanations, in Tutorial 2 of the
 * libxlsxwriter documentation.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 */

#include "hbxlsxwriter.ch"

#define ITEM   1
#define COST   2

function main() 

    local workbook, worksheet, row, col, i, bold, money, expenses

    expenses := {;
	          {"Rent", 1000},;
		  {"Gas",   100},;
		  {"Food",  300},;
		  {"Gym",    50} ;
                }

    /* Create a workbook and add a worksheet. */
    workbook  := workbook_new("tutorial02.xlsx")
    worksheet := workbook_add_worksheet(workbook, NIL)
    row := 0
    col := 0

    /* Add a bold format to use to highlight cells. */
    bold = workbook_add_format(workbook)
    format_set_bold(bold)

    /* Add a number format for cells with money. */
    money = workbook_add_format(workbook)
    format_set_num_format(money, "$#,##0")

    /* Write some data header. */
    worksheet_write_string(worksheet, row, col,     "Item", bold)
    worksheet_write_string(worksheet, row, col + 1, "Cost", bold)

    /* Iterate over the data and write it out element by element. */
    for i:=1 to 4 //(i = 0; i < 4; i++) {
        /* Write from the first cell below the headers. */
        row := i 
        worksheet_write_string(worksheet, row, col,     expenses[i][ITEM], NIL)
        worksheet_write_number(worksheet, row, col + 1, expenses[i][COST], money)
    next i

    /* Write a total using a formula. */
    worksheet_write_string (worksheet, row + 1, col,     "Total",       bold)
    worksheet_write_formula(worksheet, row + 1, col + 1, "=SUM(B2:B5)", money)

    /* Save the workbook and free any allocated memory. */
    return workbook_close(workbook)

