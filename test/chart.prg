/*
 * An example of a simple Excel chart using the libxlsxwriter library.
 *
 * Copyright 2014-2018, John McNamara, jmcnamara@cpan.org
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "hbxlsxwriter.ch"

/* Write some data to the worksheet. */
procedure write_worksheet_data( worksheet )

    local aData, row, col

    aData := { ;
	       { 1,  2,  3 }, ;
	       { 2,  4,  6 }, ;
	       { 3,  6,  9 }, ;
	       { 4,  8, 12 }, ;
	       { 5, 10, 15 }  ;
	     }

    for row = 1 to 5 
        for col = 1 to 3 
            worksheet_write_number(worksheet, row, col, aData[row][col], NIL)
	next col
    next row

/* Create a worksheet with a chart. */
function main() 
    local workbook, worksheet, chart

    workbook  = new_workbook("chart.xlsx")
    worksheet = workbook_add_worksheet(workbook, NIL)

    /* Write some data for the chart. */
    write_worksheet_data(worksheet)

    /* Create a chart object. */
    chart = workbook_add_chart(workbook, LXW_CHART_COLUMN)

    /* Configure the chart. In simplest case we just add some value data
     * series. The NULL categories will default to 1 to 5 like in Excel.
     */
    chart_add_series(chart, NIL, "Sheet1!$A$1:$A$5")
    chart_add_series(chart, NIL, "Sheet1!$B$1:$B$5")
    chart_add_series(chart, NIL, "Sheet1!$C$1:$C$5")

    /* Insert the chart into the worksheet. */
    worksheet_insert_chart(worksheet, LXW_CELL("B7"), chart)

    return workbook_close(workbook)

