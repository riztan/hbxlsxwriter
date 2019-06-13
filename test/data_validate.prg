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
#xtranslate <dv>.<key> := <value>  =>  hb_lxw_dv( @<dv>, #<key>, <value> )


/*
 * Write some data to the worksheet.
 */
procedure write_worksheet_data( worksheet, format)

    worksheet_write_string(worksheet, CELL("A1"),                              ;
                           "Some examples of data validation in libxlsxwriter",;
                           format)

    worksheet_write_string(worksheet, CELL("B1"), "Enter values in this column", format)
    worksheet_write_string(worksheet, CELL("D1"), "Sample Data", format)

    worksheet_write_string(worksheet, CELL("D3"), "Integers"  )
    worksheet_write_number(worksheet, CELL("E3"), 1           )
    worksheet_write_number(worksheet, CELL("F3"), 10          )

    worksheet_write_string(worksheet, CELL("D4"), "List Data" )
    worksheet_write_string(worksheet, CELL("E4"), "open"      )
    worksheet_write_string(worksheet, CELL("F4"), "high"      )
    worksheet_write_string(worksheet, CELL("G4"), "close"     )

    worksheet_write_string(worksheet,  CELL("D5"), "Formula"  )
    worksheet_write_formula(worksheet, CELL("E5"), "=AND(F5=50,G5=60)"  )
    worksheet_write_number(worksheet,  CELL("F5"), 50         )
    worksheet_write_number(worksheet,  CELL("G5"), 60         )



/*
 * Create a worksheet with data validations.
 */
function main() 
    local workbook, worksheet, format, list
    local hValidation

    workbook  = new_workbook("data_validate1.xlsx")
    worksheet = workbook_add_worksheet(workbook,  NIL)

    hValidation := hb_Hash()


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
    worksheet_set_column(worksheet, 0,  0, 55,  NIL)
    worksheet_set_column(worksheet, 1,  1, 15,  NIL)
    worksheet_set_column(worksheet, 3,  3, 15,  NIL)
    worksheet_set_row(worksheet,  0, 36,  NIL)
    

    /*
     * Example 1. Limiting input to an integer in a fixed range.
     */
    worksheet_write_string(worksheet,                          ;
                           CELL("A3"),                         ;
                           "Enter an integer between 1 and 10")



    hValidation.validate       := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria       := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_number := 1
    hValidation.maximum_number := 10

    worksheet_data_validation_cell(worksheet, CELL("B3"), hValidation )

    /*
     * Example 2. Limiting input to an integer outside a fixed range.
     */
    worksheet_write_string(worksheet,                               ;
                           CELL("A5"),                              ;
                           "Enter an integer not between 1 and 10 "+;
                           "(using cell references)" )

    hValidation.validate        := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria        := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_formula := "=E3"
    hValidation.maximum_formula := "=F3"

    worksheet_data_validation_cell(worksheet, CELL("B5"), hValidation )

    /*
     * Example 3. Limiting input to an integer greater than a fixed value.
     */
    worksheet_write_string(worksheet,                         ;
                           CELL("A7"),                        ;
                           "Enter an integer greater than 0" )

    hValidation.validate     := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria     := LXW_VALIDATION_CRITERIA_GREATER_THAN
    hValidation.value_number := 0
    

    worksheet_data_validation_cell(worksheet, CELL("B7"), hValidation )

    /*
     * Example 4. Limiting input to an integer less than a fixed value.
     */
    worksheet_write_string(worksheet,                       ;
                           CELL("A9"),                      ;
                           "Enter an integer less than 10" )
                            

    hValidation.validate     := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria     := LXW_VALIDATION_CRITERIA_LESS_THAN
    hValidation.value_number := 10

    worksheet_data_validation_cell(worksheet, CELL("B9"), hValidation )


    /*
     * Example 5. Limiting input to a decimal in a fixed range.
     */

    worksheet_write_string(worksheet,                             ;
                           CELL("A11"),                           ;
                           "Enter a decimal between 0.1 and 0.5" )

    hValidation.validate       := LXW_VALIDATION_TYPE_DECIMAL
    hValidation.criteria       := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_number := 0.1
    hValidation.maximum_number := 0.5

    worksheet_data_validation_cell(worksheet, CELL("B11"), hValidation )


    /*
     * Example 6. Limiting input to a value in a dropdown list.
     */

    worksheet_write_string(worksheet,                              ;
                           CELL("A13"),                            ;
                           "Select a value from a drop down list" )

    list := { "open", "high", "close" }

    hValidation.validate   := LXW_VALIDATION_TYPE_LIST
    hValidation.value_list := list //{"open", "high", "close"}

    worksheet_data_validation_cell(worksheet, CELL("B13"), hValidation )


    /*
     * Example 7. Limiting input to a value in a dropdown list.
     */
    worksheet_write_string(worksheet,                              ;
                           CELL("A15"),                            ;
                           "Select a value from a drop down list "+;
                           "(using a cell range)" )

    hValidation.validate      := LXW_VALIDATION_TYPE_LIST
    hValidation.value_formula := "=$E$4:$G$4"

    worksheet_data_validation_cell(worksheet, CELL("B15"), hValidation )


    /*
     * Example 8. Limiting input to a date in a fixed range.
     */

    worksheet_write_string(worksheet,                                      ;
                           CELL("A17"),                                    ;
                           "Enter a date between 1/1/2008 and 12/12/2008" )

    hValidation.validate         := LXW_VALIDATION_TYPE_DATE
    hValidation.criteria         := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_datetime := CTOD( "01/01/2008" )
    hValidation.maximum_datetime := CTOD( "12/12/2008" )

    worksheet_data_validation_cell(worksheet, CELL("B17"), hValidation )


    /*
     * Example 9. Limiting input to a time in a fixed range.
     */
    worksheet_write_string(worksheet,                             ;
                           CELL("A19"),                           ;
                           "Enter a time between 6:00 and 12:00" )

    hValidation.validate         := LXW_VALIDATION_TYPE_DATE
    hValidation.criteria         := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_datetime := hb_DateTime( 0, 0, 0, 6, 0, 0 )
    hValidation.maximum_datetime := hb_DateTime( 0, 0, 0,12, 0, 0 )

    worksheet_data_validation_cell(worksheet, CELL("B19"), hValidation )


    /*
     * Example 10. Limiting input to a string greater than a fixed length.
     */
    worksheet_write_string(worksheet,                                 ;
                           CELL("A21"),                               ;
                           "Enter a string longer than 3 characters" )

    hValidation.validate     := LXW_VALIDATION_TYPE_LENGTH
    hValidation.criteria     := LXW_VALIDATION_CRITERIA_GREATER_THAN
    hValidation.value_number := 3

    worksheet_data_validation_cell(worksheet, CELL("B21"), hValidation )


    /*
     * Example 11. Limiting input based on a formula.
     */
    worksheet_write_string(worksheet,                                 ;
                           CELL("A23"),                               ;
                           'Enter a value if the following is true '+ ;
                           '"=AND(F5=50,G5=60)"' )

    hValidation.validate      := LXW_VALIDATION_TYPE_CUSTOM_FORMULA
    hValidation.value_formula := "=AND(F5=50,G5=60)"

    worksheet_data_validation_cell(worksheet, CELL("B23"), hValidation )


    /*
     * Example 12. Displaying and modifying data validation messages.
     */
    worksheet_write_string(worksheet,                                     ;
                           CELL("A25"),                                   ;
                           "Displays a message when you select the cell" )

    hValidation.validate       := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria       := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_number := 1
    hValidation.maximum_number := 100
    hValidation.input_title    := "Enter an integer:"
    hValidation.input_message  := "between 1 and 100"

    worksheet_data_validation_cell(worksheet, CELL("B25"), hValidation )


    /*
     * Example 13. Displaying and modifying data validation messages.
     */
    worksheet_write_string(worksheet,                                     ;
                           CELL("A27"),                                   ;
                           "Display a custom error message when integer "+;
                           "isn't between 1 and 100" )

    hValidation.validate       := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria       := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_number := 1
    hValidation.maximum_number := 100
    hValidation.input_title    := "Enter an integer:"
    hValidation.input_message  := "between 1 and 100"
    hValidation.error_title    := "Input value is not valid!"
    hValidation.error_message  := "It should be an integer between 1 and 100"

    worksheet_data_validation_cell(worksheet, CELL("B27"), hValidation )


    /*
     * Example 14. Displaying and modifying data validation messages.
     */
    worksheet_write_string(worksheet,                                          ;
                           CELL("A29"),                                        ;
                           "Display a custom info message when integer isn't "+;
                           "between 1 and 100" )

    hValidation.validate       := LXW_VALIDATION_TYPE_INTEGER
    hValidation.criteria       := LXW_VALIDATION_CRITERIA_BETWEEN
    hValidation.minimum_number := 1
    hValidation.maximum_number := 100
    hValidation.input_title    := "Enter an integer:"
    hValidation.input_message  := "between 1 and 100"
    hValidation.error_title    := "Input value is not valid!"
    hValidation.error_message  := "It should be an integer between 1 and 100"
    hValidation.error_type     := LXW_VALIDATION_ERROR_TYPE_INFORMATION

    worksheet_data_validation_cell(worksheet, CELL("B29"), hValidation )


    return workbook_close(workbook)

//eof
