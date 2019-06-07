

/*****************************************************************************
 * chartsheet - A library for creating Excel XLSX chartsheet files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2019, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */
/*
 * Wrapped for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/chartsheet.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"




/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/


/*
 * Create a new chartsheet object.
 */
/*
lxw_chartsheet *
lxw_chartsheet_new(lxw_worksheet_init_data *init_data)
*/
HB_FUNC( LXW_CHARTSHEET_NEW )
{ 
   lxw_worksheet_init_data *init_data = hb_parptr( 1 ) ;

   hb_retptr( lxw_chartsheet_new(init_data) ); 
}




/*
 * Free a chartsheet object.
 */
/*
void
lxw_chartsheet_free(lxw_chartsheet *chartsheet)
*/
HB_FUNC( LXW_CHARTSHEET_FREE )
{ 
   lxw_chartsheet *chartsheet = hb_parptr( 1 ) ;

   lxw_chartsheet_free(chartsheet) ; 
}


/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/


/*
 * Assemble and write the XML file.
 */
/*
void
lxw_chartsheet_assemble_xml_file(lxw_chartsheet *self)
*/
HB_FUNC( LXW_CHARTSHEET_ASSEMBLE_XML_FILE )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   lxw_chartsheet_assemble_xml_file(self) ; 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Set a chartsheet chart, with options.
 */
/*
lxw_error
chartsheet_set_chart_opt(lxw_chartsheet *self,
   lxw_chart *chart, lxw_image_options *user_options)
*/
HB_FUNC( CHARTSHEET_SET_CHART_OPT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   lxw_chart *chart = hb_parptr( 2 ) ;
   lxw_image_options *user_options = hb_parptr( 3 ) ;

   hb_retni( chartsheet_set_chart_opt(self, chart, user_options) ); 
}





/*
 * Set a chartsheet charts.
 */
/*
lxw_error
chartsheet_set_chart(lxw_chartsheet *self, lxw_chart *chart)
*/
HB_FUNC( CHARTSHEET_SET_CHART )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   lxw_chart *chart = hb_parptr( 2 ) ;

   hb_retni( chartsheet_set_chart(self, chart) ); 
}





/*
 * Set this chartsheet as a selected worksheet, i.e. the worksheet has its tab
 * highlighted.
 */
/*
void
chartsheet_select(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_SELECT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_select(self) ; 
}





/*
 * Set this chartsheet as the active worksheet, i.e. the worksheet that is
 * displayed when the workbook is opened. Also set it as selected.
 */
/*
void
chartsheet_activate(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_ACTIVATE )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_activate(self) ; 
}





/*
 * Set this chartsheet as the first visible sheet. This is necessary
 * when there are a large number of worksheets and the activated
 * worksheet is not visible on the screen.
 */
/*
void
chartsheet_set_first_sheet(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_SET_FIRST_SHEET )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_set_first_sheet(self) ; 
}





/*
 * Hide this chartsheet.
 */
/*
void
chartsheet_hide(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_HIDE )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_hide(self) ; 
}





/*
 * Set the color of the chartsheet tab.
 */
/*
void
chartsheet_set_tab_color(lxw_chartsheet *self, lxw_color_t color)
*/
HB_FUNC( CHARTSHEET_SET_TAB_COLOR )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   chartsheet_set_tab_color(self, color) ; 
}





/*
 * Set the chartsheet protection flags to prevent modification of chartsheet
 * objects.
 */
/*
void
chartsheet_protect(lxw_chartsheet *self, const char *password,
   lxw_protection *options)
*/
HB_FUNC( CHARTSHEET_PROTECT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   const char *password = hb_parcx( 2 ) ;
   lxw_protection *options = hb_parptr( 3 ) ;

   chartsheet_protect(self, password, options) ; 
}





/*
 * Set the chartsheet zoom factor.
 */
/*
void
chartsheet_set_zoom(lxw_chartsheet *self, uint16_t scale)
*/
HB_FUNC( CHARTSHEET_SET_ZOOM )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   uint16_t scale = hb_parnl( 2 ) ;

   chartsheet_set_zoom(self, scale) ; 
}





/*
 * Set the page orientation as portrait.
 */
/*
void
chartsheet_set_portrait(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_SET_PORTRAIT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_set_portrait(self) ; 
}





/*
 * Set the page orientation as landscape.
 */
/*
void
chartsheet_set_landscape(lxw_chartsheet *self)
*/
HB_FUNC( CHARTSHEET_SET_LANDSCAPE )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;

   chartsheet_set_landscape(self) ; 
}





/*
 * Set the paper type. Example. 1 = US Letter, 9 = A4
 */
/*
void
chartsheet_set_paper(lxw_chartsheet *self, uint8_t paper_size)
*/
HB_FUNC( CHARTSHEET_SET_PAPER )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   uint8_t paper_size = hb_parni( 2 ) ;

   chartsheet_set_paper(self, paper_size) ; 
}





/*
 * Set all the page margins in inches.
 */
/*
void
chartsheet_set_margins(lxw_chartsheet *self, double left, double right,
       double top, double bottom)
*/
HB_FUNC( CHARTSHEET_SET_MARGINS )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   double left = hb_parnd( 2 ) ;
   double right = hb_parnd( 3 ) ;
   double top = hb_parnd( 4 ) ;
   double bottom = hb_parnd( 5 ) ;

   chartsheet_set_margins(self, left, right, top, bottom) ; 
}





/*
 * Set the page header caption and options.
 */
/*
lxw_error
chartsheet_set_header_opt(lxw_chartsheet *self, const char *string,
         lxw_header_footer_options *options)
*/
HB_FUNC( CHARTSHEET_SET_HEADER_OPT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   const char *string = hb_parcx( 2 ) ;
   lxw_header_footer_options *options = hb_parptr( 3 ) ;

   hb_retni( chartsheet_set_header_opt(self, string, options) ); 
}





/*
 * Set the page footer caption and options.
 */
/*
lxw_error
chartsheet_set_footer_opt(lxw_chartsheet *self, const char *string,
            lxw_header_footer_options *options)
*/
HB_FUNC( CHARTSHEET_SET_FOOTER_OPT )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   const char *string = hb_parcx( 2 ) ;
   lxw_header_footer_options *options = hb_parptr( 3 ) ;

   hb_retni( chartsheet_set_footer_opt(self, string, options) ); 
}





/*
 * Set the page header caption.
 */
/*
lxw_error
chartsheet_set_header(lxw_chartsheet *self, const char *string)
*/
HB_FUNC( CHARTSHEET_SET_HEADER )
{ 
   lxw_chartsheet *self = hb_parptr( 1 ) ;
   const char *string = hb_parcx( 2 ) ;

   hb_retni( chartsheet_set_header(self, string) ); 
}


//eof
