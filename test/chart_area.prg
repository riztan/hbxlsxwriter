/*
 * An example of creating Excel area charts using the libxlsxwriter library.
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
procedure write_worksheet_data( worksheet, bold) 

    local row, col, data

    /* Three columns of data. */
    data := {  ;
        {2, 40, 30}, ;
        {3, 40, 25}, ;
        {4, 50, 30}, ;
        {5, 30, 10}, ;
        {6, 25,  5}, ;
        {7, 50, 10}  ;
    }

    worksheet_write_string(worksheet, CELL("A1"), "Number",  bold)
    worksheet_write_string(worksheet, CELL("B1"), "Batch 1", bold)
    worksheet_write_string(worksheet, CELL("C1"), "Batch 2", bold)

    for row := 1 to 6 //(row = 0; row < 6; row++)
        for col := 1 to 3  //(col = 0; col < 3; col++)
            worksheet_write_number(worksheet, row, col-1, data[row][col] , NIL)
        next col
    next row


/*
 * Create a worksheet with examples charts.
 */
function main() 
    local workbook, worksheet, series, chart, bold
    local chart_axis_x, chart_axis_y

    workbook  := new_workbook("chart_area.xlsx")
    worksheet := workbook_add_worksheet(workbook, NIL)

    /* Add a bold format to use to highlight the header cells. */
    bold := workbook_add_format(workbook)
    format_set_bold(bold)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet, bold)


    /*
     * Chart 1. Create a area chart.
     */
    chart := workbook_add_chart(workbook, LXW_CHART_AREA)

    /* Add the first series to the chart. */
    series := chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /* Set the name for the series instead of the default "Series 1". */
    chart_series_set_name(series, "=Sheet1!$B$1")

    /* Add a second series but leave the categories and values undefined. They
     * can be defined later using the alternative syntax shown below.  */
    series := chart_add_series(chart, NIL, NIL)

    /* Configure the series using a syntax that is easier to define programmatically. */
    chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0) /* "=Sheet1!$A$2:$A$7" */
    chart_series_set_values(series,     "Sheet1", 1, 2, 6, 2) /* "=Sheet1!$C$2:$C$7" */
    chart_series_set_name_range(series, "Sheet1", 0, 2)       /* "=Sheet1!$C$1"      */

    /* Add a chart title and some axis labels. */
    chart_axis_x := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_X )
    chart_axis_y := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_Y )
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart_axis_x, "Test number")
    chart_axis_set_name(chart_axis_y, "Sample length (mm)")

    /* Set an Excel chart style. */
    chart_set_style(chart, 11)

    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("E2"), chart)


    /*
     * Chart 2. Create a stacked area chart.
     */
    chart := workbook_add_chart(workbook, LXW_CHART_AREA_STACKED)

    /* Add the first series to the chart. */
    series := chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /* Set the name for the series instead of the default "Series 1". */
    chart_series_set_name(series, "=Sheet1!$B$1")

    /* Add the second series to the chart. */
    series := chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7")

    /* Set the name for the series instead of the default "Series 2". */
    chart_series_set_name(series, "=Sheet1!$C$1")

    /* Add a chart title and some axis labels. */
    chart_axis_x := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_X )
    chart_axis_y := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_Y )
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart_axis_x, "Test number")
    chart_axis_set_name(chart_axis_y, "Sample length (mm)")

    /* Set an Excel chart style. */
    chart_set_style(chart, 12)

    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("E18"), chart)


    /*
     * Chart 3. Create a percent stacked area chart.
     */
    chart := workbook_add_chart(workbook, LXW_CHART_AREA_STACKED_PERCENT)

    /* Add the first series to the chart. */
    series := chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7")

    /* Set the name for the series instead of the default "Series 1". */
    chart_series_set_name(series, "=Sheet1!$B$1")

    /* Add the second series to the chart. */
    series := chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7")

    /* Set the name for the series instead of the default "Series 2". */
    chart_series_set_name(series, "=Sheet1!$C$1")

    /* Add a chart title and some axis labels. */
    chart_axis_x := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_X )
    chart_axis_y := chart_axis_get( chart, LXW_CHART_AXIS_TYPE_Y )
    chart_title_set_name(chart,        "Results of sample analysis")
    chart_axis_set_name(chart_axis_x, "Test number")
    chart_axis_set_name(chart_axis_y, "Sample length (mm)")

    /* Set an Excel chart style. */
    chart_set_style(chart, 13)

    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, CELL("E34"), chart)


    return workbook_close(workbook)

//eof
