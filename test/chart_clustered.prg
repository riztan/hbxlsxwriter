/*
 * An example of a clustered category chart using the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

/*
 * Write some data to the worksheet.
 */
procedure write_worksheet_data( worksheet, bold ) 

    worksheet_write_string(worksheet, 0, 0, "Types",      bold)
    worksheet_write_string(worksheet, 1, 0, "Type 1",      NIL)
    worksheet_write_string(worksheet, 4, 0, "Type 2",      NIL)

    worksheet_write_string(worksheet, 0, 1, "Sub Type",   bold)
    worksheet_write_string(worksheet, 1, 1, "Sub Type A",  NIL)
    worksheet_write_string(worksheet, 2, 1, "Sub Type B",  NIL)
    worksheet_write_string(worksheet, 3, 1, "Sub Type C",  NIL)
    worksheet_write_string(worksheet, 4, 1, "Sub Type D",  NIL)
    worksheet_write_string(worksheet, 5, 1, "Sub Type E",  NIL)

    worksheet_write_string(worksheet, 0, 2, "Value 1",    bold)
    worksheet_write_number(worksheet, 1, 2, 5000,          NIL)
    worksheet_write_number(worksheet, 2, 2, 2000,          NIL)
    worksheet_write_number(worksheet, 3, 2, 250,           NIL)
    worksheet_write_number(worksheet, 4, 2, 6000,          NIL)
    worksheet_write_number(worksheet, 5, 2, 500,           NIL)

    worksheet_write_string(worksheet, 0, 3, "Value 2",    bold)
    worksheet_write_number(worksheet, 1, 3, 8000,          NIL)
    worksheet_write_number(worksheet, 2, 3, 3000,          NIL)
    worksheet_write_number(worksheet, 3, 3, 1000,          NIL)
    worksheet_write_number(worksheet, 4, 3, 6000,          NIL)
    worksheet_write_number(worksheet, 5, 3, 300,           NIL)

    worksheet_write_string(worksheet, 0, 4, "Value 3",    bold)
    worksheet_write_number(worksheet, 1, 4, 6000,          NIL)
    worksheet_write_number(worksheet, 2, 4, 4000,          NIL)
    worksheet_write_number(worksheet, 3, 4, 2000,          NIL)
    worksheet_write_number(worksheet, 4, 4, 6500,          NIL)
    worksheet_write_number(worksheet, 5, 4, 200,           NIL)



/*
 * Create a worksheet with examples charts.
 */
function main() 

    local workbook, worksheet, chart, bold

    workbook  := new_workbook("chart_clustered2.xlsx")
    worksheet := workbook_add_worksheet(workbook,  NIL)
    chart     := workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /* Add a bold format to use to highlight the header cells. */
    bold := workbook_add_format(workbook)
    format_set_bold(bold)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet, bold)

    /*
     * Configure the series. Note, that the categories are 2D ranges (from
     * column A to column B). This creates the clusters. The series are shown
     * as formula strings for clarity but you can also use variables with the
     * chart_series_set_categories() and chart_series_set_values()
     * functions. See the docs.
     */
    chart_add_series(chart,               ;
                     "=Sheet1!$A$2:$B$6", ;
                     "=Sheet1!$C$2:$C$6")

    chart_add_series(chart,               ;
                     "=Sheet1!$A$2:$B$6", ;
                     "=Sheet1!$D$2:$D$6")

    chart_add_series(chart,               ;
                     "=Sheet1!$A$2:$B$6", ;
                     "=Sheet1!$E$2:$E$6")

    /* Set an Excel chart style. */
    chart_set_style(chart, 37)

    /* Turn off the legend. */
    chart_legend_set_position(chart, LXW_CHART_LEGEND_NONE)

    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("G3"), chart)

    return workbook_close(workbook)

