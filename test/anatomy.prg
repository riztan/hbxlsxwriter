/*
 * Anatomy of a simple libxlsxwriter program.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

PROCEDURE Main()

   LOCAL workbook, worksheet1, worksheet2, myformat1, myformat2, error

   /* Create a new workbook. */
   workbook   := workbook_new( "anatomy.xlsx" )

   /* Add a worksheet with a user defined sheet name. */
   worksheet1 := workbook_add_worksheet( workbook, "Demo" )

   /* Add a worksheet with Excel's default sheet name: Sheet2. */
   worksheet2 := workbook_add_worksheet( workbook )

   /* Add some cell formats. */
   myformat1   := workbook_add_format( workbook )
   myformat2   := workbook_add_format( workbook )

   /* Set the bold property for the first format. */
   format_set_bold( myformat1 )

   /* Set a number format for the second format. */
   format_set_num_format( myformat2, "$#,##0.00" )

   /* Widen the first column to make the text clearer. */
   worksheet_set_column( worksheet1, 0, 0, 20 )

   /* Write some unformatted data. */
   worksheet_write_string( worksheet1, 0, 0, "Peach" )
   worksheet_write_string( worksheet1, 1, 0, "Plum" )

   /* Write formatted data. */
   worksheet_write_string( worksheet1, 2, 0, "Pear",  myformat1 )

   /* Formats can be reused. */
   worksheet_write_string( worksheet1, 3, 0, "Persimmon",  myformat1 )


   /* Write some numbers. */
   worksheet_write_number( worksheet1, 5, 0, 123 )
   worksheet_write_number( worksheet1, 6, 0, 4567.555,  myformat2 )


   /* Write to the second worksheet. */
   worksheet_write_string( worksheet2, 0, 0, "Some text", myformat1 )


   /* Close the workbook, save the file and free any memory. */
   error := workbook_close( workbook )

   /* Check if there was any error creating the xlsx file. */
   IF error > 0
      Alert( "Error in workbook_close(). " + hb_eol() + ;
         "Error " + AllTrim( Str( error ) ) + " = " + lxw_strerror( error ) )

   ENDIF

   RETURN

//eof
