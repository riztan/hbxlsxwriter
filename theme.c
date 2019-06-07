/*****************************************************************************
 * theme - A library for creating Excel XLSX theme files.
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

#include <string.h>

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/theme.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new theme object.
 */
/*
lxw_theme *
lxw_theme_new(void)
*/
HB_FUNC( LXW_THEME_NEW )
{ 
   hb_retptr( lxw_theme_new() ); 
}





/*
 * Free a theme object.
 */
/*
void
lxw_theme_free(lxw_theme *theme)
*/
HB_FUNC( LXW_THEME_FREE )
{ 
   lxw_theme *theme = hb_parptr( 1 ) ;

   lxw_theme_free(theme) ; 
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
lxw_theme_assemble_xml_file(lxw_theme *self)
*/
HB_FUNC( LXW_THEME_ASSEMBLE_XML_FILE )
{ 
   lxw_theme *self = hb_parptr( 1 ) ;

   lxw_theme_assemble_xml_file(self) ; 
}


//eof
