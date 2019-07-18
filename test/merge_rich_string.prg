/*
 * An example of merging cells containing a rich string using libxlsxwriter.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

procedure main() 

    local workbook, worksheet, merge_format, red, blue, aRich_strings
    local hFragment1, hFragment2, hFragment3, hFragment4

    workbook  := workbook_new("merge_rich_string.xlsx")
    worksheet := workbook_add_worksheet( workbook )

    /* Configure a format for the merged range. */
    merge_format := workbook_add_format(workbook)

    format_set_align(merge_format, LXW_ALIGN_CENTER)
    format_set_align(merge_format, LXW_ALIGN_VERTICAL_CENTER)
    format_set_border(merge_format, LXW_BORDER_THIN)

    /* Configure formats for the rich string. */
    red := workbook_add_format(workbook)
    format_set_font_color(red, LXW_COLOR_RED)

    blue := workbook_add_format(workbook)
    format_set_font_color(blue, LXW_COLOR_BLUE)

    /* Create the fragments for the rich string. */
    hFragment1 := { "format" => NIL,  "string" => "This is"      }
    hFragment2 := { "format" => red,  "string" => "red"          }
    hFragment3 := { "format" => NIL,  "string" => " and this is" }
    hFragment4 := { "format" => blue, "string" => "blue"         }

    aRich_strings := { hFragment1, hFragment2, hFragment3, hFragment4 }

    /* Write an empty string to the merged range. */
    worksheet_merge_range(worksheet, 1, 1, 4, 3, "", merge_format)

    /* We then overwrite the first merged cell with a rich string. Note that
     * we must also pass the cell format used in the merged cells format at
     * the end. */
    worksheet_write_rich_string(worksheet, 1, 1, aRich_strings, merge_format)

    workbook_close(workbook)

    
//eof
