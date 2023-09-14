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

   LOCAL workbook, worksheet, format, options

   /* Create a new workbook and add a worksheet. */
   workbook  := workbook_new( "demo.xlsx" )
   worksheet := workbook_add_worksheet( workbook )

   /* Add a format. */
   format := workbook_add_format( workbook )

   /* Set the bold property for the format */
   format_set_bold( format )

   format_set_pattern( format, LXW_PATTERN_SOLID )
   format_set_font_color( format, LXW_COLOR_CYAN )
   format_set_bg_color( format, LXW_COLOR_BLUE )


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

   /* Insert an image with options. */
   options := {"x_scale" => .5, "y_scale" => .5 }
   worksheet_insert_image_opt( worksheet, CELL("B15"), "hb_logo.png", options )

   options := {"x_offset"=> 10 , "y_offset" => 5 }
   worksheet_insert_image_opt( worksheet, CELL("G2"), "hb_logo.png", options )

   workbook_close( workbook )

//eof
