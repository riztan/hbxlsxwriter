/*
 * Examples of how to add data validation and dropdown lists using the
 * libxlsxwriter library.
 *
 * Data validation is a feature of Excel which allows you to restrict the data
 * that a user enters in a cell and to display help and warning messages. It
 * also allows you to restrict input to values in a drop down list.
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
procedure write_worksheet_data( worksheet, format)

    worksheet_write_string(worksheet, CELL("A1"),                               ;
                           "Some examples of data validation in libxlsxwriter", ;
                           format)

    worksheet_write_string(worksheet, CELL("B1"), "Enter values in this column", format)
    worksheet_write_string(worksheet, CELL("D1"), "Sample Data", format)

    worksheet_write_string(worksheet, CELL("D3"), "Integers",  NIL)
    worksheet_write_number(worksheet, CELL("E3"), 1,  NIL)
    worksheet_write_number(worksheet, CELL("F3"), 10,  NIL)

    worksheet_write_string(worksheet, CELL("D4"), "List Data",  NIL)
    worksheet_write_string(worksheet, CELL("E4"), "open",  NIL)
    worksheet_write_string(worksheet, CELL("F4"), "high",  NIL)
    worksheet_write_string(worksheet, CELL("G4"), "close",  NIL)

    worksheet_write_string(worksheet,  CELL("D5"), "Formula",  NIL)
    worksheet_write_formula(worksheet, CELL("E5"), "=AND(F5=50,G5=60)",  NIL)
    worksheet_write_number(worksheet,  CELL("F5"), 50,  NIL)
    worksheet_write_number(worksheet,  CELL("G5"), 60,  NIL)



/*
 * Create a worksheet with data validations.
 */
function main() 
    local workbook, worksheet, data_validation, format, list

    workbook  = new_workbook("data_validate1.xlsx")
    worksheet = workbook_add_worksheet(workbook,  NIL)
    data_validation = hb_lxw_data_validation_new() //calloc(1, sizeof(lxw_data_validation))

    /* Add a format to use to highlight the header cells. */
    format = workbook_add_format(workbook)
    format_set_border(format, LXW_BORDER_THIN)
    format_set_fg_color(format, 0xC6EFCE)
    format_set_bold(format)
    format_set_text_wrap(format)
    format_set_align(format, LXW_ALIGN_VERTICAL_CENTER)
    format_set_indent(format, 1)

    /* Write some data for the validations. */
    write_worksheet_data(worksheet, format)

    /* Set up layout of the worksheet. */
    worksheet_set_column(worksheet, 0, 0, 55,  NIL)
    worksheet_set_column(worksheet, 1, 1, 15,  NIL)
    worksheet_set_column(worksheet, 3, 3, 15,  NIL)
    worksheet_set_row(worksheet, 0, 36,  NIL)


    /*
     * Example 1. Limiting input to an integer in a fixed range.
     */
    worksheet_write_string(worksheet,                          ;
                           CELL("A3"),                         ;
                           "Enter an integer between 1 and 10",;
                            NIL)

    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_INTEGER )
    hb_lxw_dv_criteria( data_validation, LXW_VALIDATION_CRITERIA_BETWEEN )
    hb_lxw_dv_minimum_number( data_validation,  1 )
    hb_lxw_dv_maximum_number( data_validation, 10 )

/*
    hb_lxw_data_validation_set( data_validation, ;
                             LXW_DV_VALIDATE, ;
                             LXW_VALIDATION_TYPE_INTEGER )

    hb_lxw_data_validation_set( data_validation, ;
                             LXW_DV_CRITERIA, ;
                             LXW_VALIDATION_CRITERIA_BETWEEN )

    hb_lxw_data_validation_set( data_validation, ;
                             LXW_DV_MINIMUM_NUMBER, ;
                             1 )

    hb_lxw_data_validation_set( data_validation, ;
                             LXW_DV_MAXIMUM_NUMBER, ;
                             10 )
*/
/*
    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_number = 1
    data_validation->maximum_number = 10
*/
    worksheet_data_validation_cell(worksheet, CELL("B3"), data_validation)


    /*
     * Example 2. Limiting input to an integer outside a fixed range.
     */

    worksheet_write_string(worksheet,   ;
                           CELL("A5"),  ;
                           "Enter an integer not between 1 and 10 (using cell references)",;
                            NIL)

/*
    data_validation->validate        = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria        = LXW_VALIDATION_CRITERIA_NOT_BETWEEN
    data_validation->minimum_formula = "=E3"
    data_validation->maximum_formula = "=F3"
*/

    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_INTEGER )
    hb_lxw_dv_criteria( data_validation, LXW_VALIDATION_CRITERIA_BETWEEN )
    hb_lxw_dv_minimum_formula( data_validation, "=E3" )
    hb_lxw_dv_maximum_formula( data_validation, "=F3" )

    worksheet_data_validation_cell(worksheet, CELL("B5"), data_validation)


    /*
     * Example 3. Limiting input to an integer greater than a fixed value.
     */

    worksheet_write_string(worksheet,   ;
                           CELL("A7"),  ;
                           "Enter an integer greater than 0", ;
                            NIL)
/*
    data_validation->validate     = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria     = LXW_VALIDATION_CRITERIA_GREATER_THAN
    data_validation->value_number = 0
*/
    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_INTEGER )
    hb_lxw_dv_criteria( data_validation, LXW_VALIDATION_CRITERIA_GREATER_THAN )
    hb_lxw_dv_value_number( data_validation, 0 )

    

    worksheet_data_validation_cell(worksheet, CELL("B7"), data_validation)


    /*
     * Example 4. Limiting input to an integer less than a fixed value.
     */

    worksheet_write_string(worksheet,   ;
                           CELL("A9"),  ;
                           "Enter an integer less than 10", ;
                            NIL)

    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_INTEGER )
    hb_lxw_dv_criteria( data_validation, LXW_VALIDATION_CRITERIA_LESS_THAN )
    hb_lxw_dv_value_number( data_validation, 10 )

/*
    data_validation->validate     = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria     = LXW_VALIDATION_CRITERIA_LESS_THAN
    data_validation->value_number = 10
*/
    worksheet_data_validation_cell(worksheet, CELL("B9"), data_validation)


    /*
     * Example 5. Limiting input to a decimal in a fixed range.
     */

    worksheet_write_string(worksheet,   ;
                           CELL("A11"), ;
                           "Enter a decimal between 0.1 and 0.5", ;
                            NIL)

    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_DECIMAL )
    hb_lxw_dv_criteria( data_validation, LXW_VALIDATION_CRITERIA_BETWEEN )
    hb_lxw_dv_minimum_number( data_validation, 0.1 )
    hb_lxw_dv_maximum_number( data_validation, 0.5 )
/*
    data_validation->validate       = LXW_VALIDATION_TYPE_DECIMAL
    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_number = 0.1
    data_validation->maximum_number = 0.5
*/

    worksheet_data_validation_cell(worksheet, CELL("B11"), data_validation)



    /*
     * Example 6. Limiting input to a value in a dropdown list.
     */

    worksheet_write_string(worksheet,    ;
                           CELL("A13"),  ;
                           "Select a value from a drop down list", ;
                            NIL)

//    char *list[] = {"open", "high", "close",  NIL}
    list := { "open", "high", "close" }

    hb_lxw_dv_validate( data_validation, LXW_VALIDATION_TYPE_LIST )
    hb_lxw_dv_value_list( data_validation, list) //, worksheet )

/*
    data_validation->validate   = LXW_VALIDATION_TYPE_LIST
    data_validation->value_list = list
*/

    worksheet_data_validation_cell(worksheet, CELL("B13"), data_validation)



    /*
     * Example 7. Limiting input to a value in a dropdown list.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A15"),
                           "Select a value from a drop down list (using a cell range)",
                            NIL)

    data_validation->validate      = LXW_VALIDATION_TYPE_LIST
    data_validation->value_formula = "=$E$4:$G$4"

    worksheet_data_validation_cell(worksheet, CELL("B15"), data_validation)


*/
    /*
     * Example 8. Limiting input to a date in a fixed range.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A17"),
                           "Enter a date between 1/1/2008 and 12/12/2008",
                            NIL)

    lxw_datetime datetime1 = {2008,  1,  1, 0, 0, 0}
    lxw_datetime datetime2 = {2008, 12, 12, 0, 0, 0}

    data_validation->validate         = LXW_VALIDATION_TYPE_DATE
    data_validation->criteria         = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_datetime = datetime1
    data_validation->maximum_datetime = datetime2

    worksheet_data_validation_cell(worksheet, CELL("B17"), data_validation)


*/
    /*
     * Example 9. Limiting input to a time in a fixed range.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A19"),
                           "Enter a time between 6:00 and 12:00",
                            NIL)

    lxw_datetime datetime3 = {0, 0, 0,  6, 0, 0}
    lxw_datetime datetime4 = {0, 0, 0, 12, 0, 0}

    data_validation->validate         = LXW_VALIDATION_TYPE_DATE
    data_validation->criteria         = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_datetime = datetime3
    data_validation->maximum_datetime = datetime4

    worksheet_data_validation_cell(worksheet, CELL("B19"), data_validation)


*/
    /*
     * Example 10. Limiting input to a string greater than a fixed length.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A21"),
                           "Enter a string longer than 3 characters",
                            NIL)

    data_validation->validate     = LXW_VALIDATION_TYPE_LENGTH
    data_validation->criteria     = LXW_VALIDATION_CRITERIA_GREATER_THAN
    data_validation->value_number = 3

    worksheet_data_validation_cell(worksheet, CELL("B21"), data_validation)


*/
    /*
     * Example 11. Limiting input based on a formula.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A23"),
                           "Enter a value if the following is true \"=AND(F5=50,G5=60)\"",
                            NIL)

    data_validation->validate      = LXW_VALIDATION_TYPE_CUSTOM_FORMULA
    data_validation->value_formula = "=AND(F5=50,G5=60)"

    worksheet_data_validation_cell(worksheet, CELL("B23"), data_validation)


*/
    /*
     * Example 12. Displaying and modifying data validation messages.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A25"),
                           "Displays a message when you select the cell",
                            NIL)

    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_number = 1
    data_validation->maximum_number = 100
    data_validation->input_title    = "Enter an integer:"
    data_validation->input_message  = "between 1 and 100"

    worksheet_data_validation_cell(worksheet, CELL("B25"), data_validation)


*/
    /*
     * Example 13. Displaying and modifying data validation messages.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A27"),
                           "Display a custom error message when integer isn't between 1 and 100",
                            NIL)

    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_number = 1
    data_validation->maximum_number = 100
    data_validation->input_title    = "Enter an integer:"
    data_validation->input_message  = "between 1 and 100"
    data_validation->error_title    = "Input value is not valid!"
    data_validation->error_message  = "It should be an integer between 1 and 100"

    worksheet_data_validation_cell(worksheet, CELL("B27"), data_validation)


*/
    /*
     * Example 14. Displaying and modifying data validation messages.
     */
/*
    worksheet_write_string(worksheet,
                           CELL("A29"),
                           "Display a custom info message when integer isn't between 1 and 100",
                            NIL)

    data_validation->validate       = LXW_VALIDATION_TYPE_INTEGER
    data_validation->criteria       = LXW_VALIDATION_CRITERIA_BETWEEN
    data_validation->minimum_number = 1
    data_validation->maximum_number = 100
    data_validation->input_title    = "Enter an integer:"
    data_validation->input_message  = "between 1 and 100"
    data_validation->error_title    = "Input value is not valid!"
    data_validation->error_message  = "It should be an integer between 1 and 100"
    data_validation->error_type     = LXW_VALIDATION_ERROR_TYPE_INFORMATION

    worksheet_data_validation_cell(worksheet, CELL("B29"), data_validation)

*/
    /* Cleanup. */
    lxw_free(data_validation)

    return workbook_close(workbook)

