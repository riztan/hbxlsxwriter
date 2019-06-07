/*
 * An example of a simple Excel chart with user defined fonts using the
 * libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"


/* Create a worksheet with a chart. */
function main() 

    local workbook, worksheet, chart, aFont, y_axis, x_axis

    workbook  := new_workbook("chart_fonts.xlsx")
    worksheet := workbook_add_worksheet(workbook,  NIL)

    /* Write some data for the chart. */
    worksheet_write_number(worksheet, 0, 0, 10,  NIL)
    worksheet_write_number(worksheet, 1, 0, 40,  NIL)
    worksheet_write_number(worksheet, 2, 0, 50,  NIL)
    worksheet_write_number(worksheet, 3, 0, 20,  NIL)
    worksheet_write_number(worksheet, 4, 0, 10,  NIL)
    worksheet_write_number(worksheet, 5, 0, 50,  NIL)

    /* Create a chart object. */
    chart := workbook_add_chart(workbook, LXW_CHART_LINE)

    /* Configure the chart. */
    chart_add_series(chart,  NIL, "Sheet1!$A$1:$A$6")


    /* Create some fonts to use in the chart.  */
    aFont := ARRAY( 6 )

    /* font 1 */
    aFont[1] := hb_lxw_font_new()
    hb_lxw_font_set_name(  aFont[1], "Calibri" )
    hb_lxw_font_set_color( aFont[1], LXW_COLOR_BLUE )

    /* font 2 */
    aFont[2] := hb_lxw_font_new()
    hb_lxw_font_set_name(  aFont[2], "Courier" )
    hb_lxw_font_set_color( aFont[2], 0x92D050 )

    /* font 3 */
    aFont[3] := hb_lxw_font_new()
    hb_lxw_font_set_name(  aFont[3], "Arial" )
    hb_lxw_font_set_color( aFont[3], 0x00B0F0 )

    /* font 4 */
    aFont[4] := hb_lxw_font_new() 
    hb_lxw_font_set_name(  aFont[4], "Century" )
    hb_lxw_font_set_color( aFont[4], LXW_COLOR_RED )

    /* font 5 */
    aFont[5] := hb_lxw_font_new() 
    hb_lxw_font_set_rotation( aFont[5], -30 )

    /* font 6 */
    aFont[6] := hb_lxw_font_new() 
    hb_lxw_font_set_bold(      aFont[6], .T. )
    hb_lxw_font_set_italic(    aFont[6], .T. )
    hb_lxw_font_set_underline( aFont[6], .T. )
    hb_lxw_font_set_color(     aFont[6], 0x7030A0 )


    /* Write the chart title with a font. */
    chart_title_set_name(chart, "Test Results")
    chart_title_set_name_font(chart, aFont[1])


    /* Write the Y axis with a font. */
    y_axis := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_Y )
    chart_axis_set_name( y_axis, "Units" )
    chart_axis_set_name_font( y_axis, aFont[2])
    chart_axis_set_num_font(  y_axis, aFont[3])


    /* Write the X axis with a font. */
    x_axis := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_X )
    chart_axis_set_name(      x_axis, "Month"  )
    chart_axis_set_name_font( x_axis, aFont[4] )
    chart_axis_set_num_font(  x_axis, aFont[5] )


    /* Display the chart legend at the bottom of the chart. */
    chart_legend_set_position(chart, LXW_CHART_LEGEND_BOTTOM)
    chart_legend_set_font(chart, @aFont[6])


    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("C1"), chart)

    return workbook_close(workbook)

//eof
