/*
 * A simple example of some of the features of the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

PROCEDURE Main()

   LOCAL workbook, worksheet, format

   /* Create a new workbook and add a worksheet. */
   workbook  := workbook_new( "demo.xlsx" )
   worksheet := workbook_add_worksheet( workbook )

   /* Add a format. */
   format := workbook_add_format( workbook )

   /* Set the bold property for the format */
   format_set_bold( format )

   /* Change the column width for clarity. */
   worksheet_set_column( worksheet, 0, 0, 20 )

   /* Write some simple text. */
   worksheet_write_string( worksheet, 0, 0, "Hello" )

   /* Text with formatting. */
   worksheet_write_string( worksheet, 1, 0, "World", format )

   /* Write some numbers. */
   worksheet_write_number( worksheet, 2, 0, 123 )
   worksheet_write_number( worksheet, 3, 0, 123.456 )

   /* Insert an image. */
   worksheet_insert_image( worksheet, 1, 2, "hb_logo.png" )

   workbook_close( workbook )

//eof
