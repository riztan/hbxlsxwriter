/*****************************************************************************
 * app - A library for creating Excel XLSX app files.
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
#include "xlsxwriter/app.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/


/*
 * Create a new app object.
 */
/*
lxw_app *
lxw_app_new(void)
*/
HB_FUNC( LXW_APP_NEW )
{ 
   hb_retptr( lxw_app_new() ); 
}





/*
 * Free a app object.
 */
/*
void
lxw_app_free(lxw_app *app)
*/
HB_FUNC( LXW_APP_FREE )
{ 
   lxw_app *app = hb_parptr( 1 ) ;

   lxw_app_free(app) ; 
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
lxw_app_assemble_xml_file(lxw_app *self)
*/
HB_FUNC( LXW_APP_ASSEMBLE_XML_FILE )
{ 
   lxw_app *self = hb_parptr( 1 ) ;

   lxw_app_assemble_xml_file(self) ; 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/


/*
 * Add the name of a workbook Part such as 'Sheet1' or 'Print_Titles'.
 */
/*
void
lxw_app_add_part_name(lxw_app *self, const char *name)
*/
HB_FUNC( LXW_APP_ADD_PART_NAME )
{ 
   lxw_app *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   lxw_app_add_part_name(self, name) ; 
}


//eof()
