/*
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * Copyright 2014-2022, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 * 
 */

#include "hbxlsxwriter.ch"

procedure main()

    local workbook

    workbook  = workbook_new("test_protect07.xlsx")
    workbook_add_worksheet(workbook, NIL)

    workbook_read_only_recommended(workbook)

    ? "Status ", workbook_close(workbook)

//eof
