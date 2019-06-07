/*****************************************************************************
 * hb_misc - A library for creating Excel XLSX format files.
 *
 * Complementary functions for use with xlsxwriter on Harbour
 *
 * Copyright 2019, Riztan Gutierrez, riztan@gmail.com. See LICENSE.txt.
 *
 */

#include "xlsxwriter/utility.h"
#include "xlsxwriter/chart.h"

#include "hbapi.h"
#include "hbapiitm.h"


HB_FUNC( LXW_FREE )
{
   free( hb_parptr( 1 ) );
}


/***************************
 * HARBOUR  FONT UTILITIES
 ***************************/

/*
//ToDo:  No funciona correctamente, se debe continuar (RIGC-20190602)
HB_FUNC( HB_LXW_CHART_FONT_NEW )
{
   lxw_chart_font font;

   font.name = "Arial";
   font.size = 11;
   font.bold = LXW_TRUE;
   font.italic = LXW_TRUE;
   font.underline = LXW_TRUE;
   font.rotation = 0;
   font.color = LXW_COLOR_BLUE;
   font.pitch_family = LXW_COLOR_BLUE;

   hb_retptr( &font );
}
*/

HB_FUNC( HB_LXW_FONT_NEW )
{
   lxw_format *format = lxw_format_new();
   hb_retptr( lxw_format_get_font_key( format ) );
}


HB_FUNC( HB_LXW_FONT_SET_NAME )
{
   lxw_chart_font *font = hb_parptr( 1 );
   const char *name = hb_parcx( 2 );
   font->name = lxw_strdup( name );
}


HB_FUNC( HB_LXW_FONT_SET_COLOR )
{
   lxw_chart_font *font = hb_parptr( 1 );
   lxw_color_t color = hb_parnl( 2 );
   font->color = color;
}


HB_FUNC( HB_LXW_FONT_SET_BOLD )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
//printf( "BOLD \n" );
      font->bold = LXW_TRUE;
   }
   else
   {
//printf( "NOT BOLD \n" );
      font->bold = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_ITALIC )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
      font->italic = LXW_TRUE;
   }
   else
   {
      font->italic = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_UNDERLINE )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
      font->underline = LXW_TRUE;
   }
   else
   {
      font->underline = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_ROTATION )
{
   lxw_chart_font *font = hb_parptr( 1 );
   int rotation = hb_parni( 2 );

   font->rotation = rotation;
}



/*
 *   DATA VALIDATION API
 *
 */

#include "xlsxwriter/worksheet.h"

HB_FUNC( HB_LXW_DATA_VALIDATION_NEW )
{
   //lxw_data_validation dv;
   //lxw_data_validation dv;
   lxw_data_validation *dv = calloc(1 , sizeof(lxw_data_validation));
   hb_retptr( dv ); 
}


#define LXW_DV_VALIDATE			0
#define LXW_DV_CRITERIA			1
#define LXW_DV_IGNORE_BLANK		2
#define LXW_DV_SHOW_INPUT		3
#define LXW_DV_SHOW_ERROR		4
#define LXW_DV_DROPDOWN			5
#define LXW_DV_IS_BETWEEN		6
#define LXW_DV_ERROR_TYPE		7
#define LXW_DV_VALUE_NUMBER		8
#define LXW_DV_VALUE_FORMULA		9
#define LXW_DV_VALUE_LIST		10
#define LXW_DV_VALUE_DATETIME		11
#define LXW_DV_MINIMUM_NUMBER		12
#define LXW_DV_MINIMUM_FORMULA		13
#define LXW_DV_MINIMUM_DATETIME		14
#define LXW_DV_MAXIMUM_NUMBER		15
#define LXW_DV_MAXIMUM_FORMULA		16
#define LXW_DV_MAXIMUM_DATETIME		17
#define LXW_DV_INPUT_TITLE		18
#define LXW_DV_INPUT_MESSAGE		19
#define LXW_DV_ERROR_TITLE		20
#define LXW_DV_ERROR_MESSAGE		21


HB_FUNC( HB_LXW_DATA_VALIDATION_SET )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t key = hb_parni( 2 );
   switch( key )
   {
      case LXW_DV_VALIDATE: 
         //uint8_t value = hb_parni( 3 );
         //dv->validate = validate;
         dv->validate = hb_parni( 3 );
	 return;
            case 1: // criteria
               dv->criteria = hb_parni( 3 );
	       return;
            case 2: // ignore_blank
               dv->ignore_blank = hb_parni( 3 );
	       return;
            case 3: // show_input
               dv->show_input = hb_parni( 3 );
	       return;
            case 4: // show_error
               dv->show_error = hb_parni( 3 );
	       return;
            case 5: // dropdown
               dv->dropdown = hb_parni( 3 );
	       return;
            case 6: // is_between
               dv->is_between = hb_parni( 3 );
	       return;
            case 7: // error_type
               dv->show_error = hb_parni( 3 );
	       return;
            case 8: // value_number
               dv->value_number = hb_parnd( 3 );
	       return;
/*
            case 9: // value_formula
	       char *value = hb_parcx( 3 );
               dv->value_formula = value;
	       return;
//            case 10: // value_list
//               dv->value_list = hb_parcx( 3 );
//	       return;
            //case 11: // value_datetime
            //   dv->value_datetime = hb_pard( 3 );
	    //   return;
//    lxw_datetime value_datetime;
            case 12: // minimum_number
               dv->minimum_number = hb_pard( 3 );
	       return;
            case 13: // minimum_formula
               dv->minimum_formula = hb_parcx( 3 );
	       return;
//            case 14: // minimum_datetime
//               dv->minimum_datetime = hb_pard( 3 );
//	       return;
//    lxw_datetime minimum_datetime;
            case 15: // maximum_number
               dv->maximum_number = hb_parnd( 3 );
	       return;
            case 16: // maximum_formula
               dv->maximum_formula = hb_parcx( 3 );
	       return;
//            case 17: // maximum_datetime
//               dv->maximum_datetime = hb_pard( 3 );
//	       return;
//    lxw_datetime maximum_datetime;
            case 18: // input_title
               dv->input_title = hb_parcx( 3 );
	       return;
            case 19: // input_message
               dv->input_message = hb_parcx( 3 );
	       return;
            case 20: // error_title
               dv->error_title = hb_parcx( 3 );
	       return;
            case 21: // error_message
               dv->error_message = hb_parcx( 3 );
	       return;
*/
   }
}

    /**
     * Set the validation type. Should be a #lxw_validation_types value.
     *
     * uint8_t validate;
     *
     */
HB_FUNC( HB_LXW_DV_VALIDATE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t validate = hb_parni( 2 );
   dv->validate = validate;
}

    /**
     * Set the validation criteria type to select the data. Should be a
     * #lxw_validation_criteria value.
     *
     * uint8_t criteria;
     *
     */
HB_FUNC( HB_LXW_DV_CRITERIA )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t criteria = hb_parni( 2 );
   dv->criteria = criteria;
}

    /** Controls whether a data validation is not applied to blank data in the
     * cell. Should be a #lxw_validation_boolean value. It is on by
     * default.
     *
     * uint8_t ignore_blank;
     *
     */
HB_FUNC( HB_LXW_DV_IGNORE_BLANK )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t ignore_blank = hb_parni( 2 );
   dv->ignore_blank = ignore_blank;
}

    /**
     * This parameter is used to toggle on and off the 'Show input message
     * when cell is selected' option in the Excel data validation dialog. When
     * the option is off an input message is not displayed even if it has been
     * set using input_message. Should be a #lxw_validation_boolean value. It
     * is on by default.
     *
     * uint8_t show_input;
     *
     */
HB_FUNC( HB_LXW_DV_SHOW_INPUT )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t show_input = hb_parni( 2 );
   dv->show_input = show_input;
}

    /**
     * This parameter is used to toggle on and off the 'Show error alert
     * after invalid data is entered' option in the Excel data validation
     * dialog. When the option is off an error message is not displayed even
     * if it has been set using error_message. Should be a
     * #lxw_validation_boolean value. It is on by default.
     */
//    uint8_t show_error;
HB_FUNC( HB_LXW_DV_SHOW_ERROR )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t show_error = hb_parni( 2 );
   dv->show_error = show_error;
}

    /**
     * This parameter is used to specify the type of error dialog that is
     * displayed. Should be a #lxw_validation_error_types value.
     */
//    uint8_t error_type;
HB_FUNC( HB_LXW_DV_ERROR_TYPE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t error_type = hb_parni( 2 );
   dv->error_type = error_type;
}

    /**
     * This parameter is used to toggle on and off the 'In-cell dropdown'
     * option in the Excel data validation dialog. When the option is on a
     * dropdown list will be shown for list validations. Should be a
     * #lxw_validation_boolean value. It is on by default.
     */
//    uint8_t dropdown;
HB_FUNC( HB_LXW_DV_DROPDOWN )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t dropdown = hb_parni( 2 );
   dv->dropdown = dropdown;
}


//    uint8_t is_between;
HB_FUNC( HB_LXW_DV__IS_BETWEEN )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   uint8_t is_between = hb_parni( 2 );
   dv->is_between = is_between;
}


    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a whole or decimal number.
     */
//    double value_number;
HB_FUNC( HB_LXW_DV_VALUE_NUMBER )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   double value_number = hb_parnd( 2 );
   dv->value_number = value_number;
}


    /**
     * This parameter is used to set the limiting value to which the criteria
     * is applied using a cell reference. It is valid for any of the
     * `_FORMULA` validation types.
     */
//    char *value_formula;
HB_FUNC( HB_LXW_DV_VALUE_FORMULA )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *formula = hb_param( 2, HB_IT_STRING );
   dv->value_formula = formula;
}


    /**
     * This parameter is used to set a list of strings for a drop down list.
     * The list should be a `NULL` terminated array of char* strings:
     *
     * @code
     *    char *list[] = {"open", "high", "close", NULL};
     *
     *    data_validation->validate   = LXW_VALIDATION_TYPE_LIST;
     *    data_validation->value_list = list;
     * @endcode
     *
     * The `value_formula` parameter can also be used to specify a list from
     * an Excel cell range.
     *
     * Note, the string list is restricted by Excel to 255 characters,
     * including comma separators.
     */
//    char **value_list;
/**  Funcion temporal para luego borrar */
HB_FUNC( HB_SET_LIST )
{
   PHB_ITEM pArray = hb_param( 1, HB_IT_ARRAY );
   if( pArray )
   {
      HB_SIZE nLen = hb_itemSize( pArray );
      if( nLen )
      {
         HB_SIZE nIndex;
         char *list[ nLen + 1 ];

	 for( nIndex = 1; nIndex<=nLen; nIndex++ )
	 {
	    list[ nIndex - 1 ] = hb_arrayGetC( pArray, nIndex );
	 }
	 list[ nLen ] = NULL;
	 hb_retptr( list );
      }
   }

}

HB_FUNC( HB_LXW_DV_VALUE_LIST )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   PHB_ITEM pArray = hb_param( 2, HB_IT_ARRAY );
   //lxw_worksheet *worksheet = hb_parptr( 3 );

   if( pArray )
   {
      HB_SIZE nLen = hb_itemSize( pArray );
      if( nLen )
      {
         HB_SIZE nIndex;
         char *list[ nLen + 1 ];

	 for( nIndex = 1; nIndex<=nLen; nIndex++ )
	 {
//            printf( "Elemento %i = %s \n", (int)nIndex, hb_arrayGetC( pArray, nIndex ) );
	    list[ nIndex - 1 ] = hb_arrayGetC( pArray, nIndex );
	 }
	 list[ nLen ] = NULL;
         dv->value_list = list;
/*
	 printf( "Lista definitiva:\n" );
	 for( nIndex=0; nIndex<=nLen; nIndex++ )
	 	printf( "  %s\n", dv->value_list[nIndex] );	 
         printf("fin\n");
*/
//         worksheet_data_validation_cell(worksheet, CELL("B13"), dv);
      }
   }
}


HB_FUNC( HB_LXW_DV_LIST )
{
printf("\nHB_LXW_DV_LIST\n");
   lxw_data_validation *dv = hb_parptr( 1 );
   char *list[] = {"ABC","DEF","GEH",NULL}; //hb_parptr(2);

   lxw_worksheet *worksheet = hb_parptr( 3 );
   uint8_t i = 0;
   dv->value_list = list; //hb_parptr( 2 );

while (list[i])//(list[i])
{
	printf( "<%s>\n", list[i] );
	//printf("yo");
	i++;
}
i = 0;

   worksheet_data_validation_cell(worksheet, CELL("B13"), dv);
   
}


    /**
     * This parameter is used to set the limiting value to which the date or
     * time criteria is applied using a #lxw_datetime struct.
     */
//    lxw_datetime value_datetime;
#include "hbdate.h"
HB_FUNC( HB_LXW_DV_VALUE_DATETIME )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   lxw_datetime datetime;
   long lDate, lTime;

   if( hb_partdt( &lDate, &lTime, 2 ) )
   {
       int iYear, iMonth, iDay ;
       int iHour, iMin, iSec, iMSec ;

       hb_timeDecode( lTime, &iHour, &iMin, &iSec, &iMSec );
       hb_dateDecode( lDate, &iYear, &iMonth, &iDay );

       datetime.year = iYear;
       datetime.month = iMonth;
       datetime.day = iDay;
       datetime.hour = iHour;
       datetime.min = iMin;
       datetime.sec = iSec;

       dv->value_datetime = datetime;
   }
}

    /**
     * This parameter is the same as `value_number` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
//    double minimum_number;
HB_FUNC( HB_LXW_DV_MINIMUM_NUMBER )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   double minimum_number = hb_parnd( 2 );
   dv->minimum_number = minimum_number;
}


    /**
     * This parameter is the same as `value_formula` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
//    char *minimum_formula;
HB_FUNC( HB_LXW_DV_MINIMUM_FORMULA )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *formula = hb_param( 2, HB_IT_STRING );
   dv->minimum_formula = formula;
}


    /**
     * This parameter is the same as `value_datetime` but for the minimum value
     * when a `BETWEEN` criteria is used.
     */
//    lxw_datetime minimum_datetime;
#include "hbdate.h"
HB_FUNC( HB_LXW_DV_MINIMUM_DATETIME )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   lxw_datetime datetime;
   long lDate, lTime;

   if( hb_partdt( &lDate, &lTime, 2 ) )
   {
       int iYear, iMonth, iDay ;
       int iHour, iMin, iSec, iMSec ;

       hb_timeDecode( lTime, &iHour, &iMin, &iSec, &iMSec );
       hb_dateDecode( lDate, &iYear, &iMonth, &iDay );

       datetime.year = iYear;
       datetime.month = iMonth;
       datetime.day = iDay;
       datetime.hour = iHour;
       datetime.min = iMin;
       datetime.sec = iSec;

       dv->minimum_datetime = datetime;
   }
}

    /**
     * This parameter is the same as `value_number` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
//    double maximum_number;
HB_FUNC( HB_LXW_DV_MAXIMUM_NUMBER )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   double maximum_number = hb_parnd( 2 );
   dv->maximum_number = maximum_number;
}


    /**
     * This parameter is the same as `value_formula` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
//    char *maximum_formula;
HB_FUNC( HB_LXW_DV_MAXIMUM_FORMULA )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *formula = hb_param( 2, HB_IT_STRING );
   dv->maximum_formula = formula;
}


    /**
     * This parameter is the same as `value_datetime` but for the maximum value
     * when a `BETWEEN` criteria is used.
     */
//    lxw_datetime maximum_datetime;
#include "hbdate.h"
HB_FUNC( HB_LXW_DV_MAXIMUM_DATETIME )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   lxw_datetime datetime;
   long lDate, lTime;

   if( hb_partdt( &lDate, &lTime, 2 ) )
   {
       int iYear, iMonth, iDay ;
       int iHour, iMin, iSec, iMSec ;

       hb_timeDecode( lTime, &iHour, &iMin, &iSec, &iMSec );
       hb_dateDecode( lDate, &iYear, &iMonth, &iDay );

       datetime.year = iYear;
       datetime.month = iMonth;
       datetime.day = iDay;
       datetime.hour = iHour;
       datetime.min = iMin;
       datetime.sec = iSec;

       dv->maximum_datetime = datetime;
   }
}

    /**
     * The input_title parameter is used to set the title of the input message
     * that is displayed when a cell is entered. It has no default value and
     * is only displayed if the input message is displayed. See the
     * `input_message` parameter below.
     *
     * The maximum title length is 32 characters.
     */
//    char *input_title;
HB_FUNC( HB_LXW_DV_INPUT_TVALUEITLE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *input_title = hb_param( 2, HB_IT_STRING );
   dv->input_title = input_title;
}


    /**
     * The input_message parameter is used to set the input message that is
     * displayed when a cell is entered. It has no default value.
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
//    char *input_message;
HB_FUNC( HB_LXW_DV_INPUT_MESSAGE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *input_message = hb_param( 2, HB_IT_STRING );
   dv->input_message = input_message;
}


    /**
     * The error_title parameter is used to set the title of the error message
     * that is displayed when the data validation criteria is not met. The
     * default error title is 'Microsoft Excel'. The maximum title length is
     * 32 characters.
     */
//    char *error_title;
HB_FUNC( HB_LXW_DV_ERROR_TITLE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *error_title = hb_param( 2, HB_IT_STRING );
   dv->error_title = error_title;
}

    /**
     * The error_message parameter is used to set the error message that is
     * displayed when a cell is entered. The default error message is "The
     * value you entered is not valid. A user has restricted values that can
     * be entered into the cell".
     *
     * The message can be split over several lines using newlines. The maximum
     * message length is 255 characters.
     */
//    char *error_message;
HB_FUNC( HB_LXW_DV_ERROR_MESSAGE )
{
   lxw_data_validation *dv = hb_parptr( 1 );
   char *error_message = hb_param( 2, HB_IT_STRING );
   dv->error_message = error_message;
}

//    char sqref[LXW_MAX_CELL_RANGE_LENGTH];

//    STAILQ_ENTRY (lxw_data_validation) list_pointers;



//eof
