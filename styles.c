/*****************************************************************************
 * styles - A library for creating Excel XLSX styles files.
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
#include "xlsxwriter/styles.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"



/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new styles object.
 */
/*
lxw_styles *
lxw_styles_new(void)
*/
HB_FUNC( LXW_STYLES_NEW )
{ 
   hb_retptr( lxw_styles_new() ); 
}





/*
 * Free a styles object.
 */
/*
void
lxw_styles_free(lxw_styles *styles)
*/
HB_FUNC( LXW_STYLES_FREE )
{ 
   lxw_styles *styles = hb_parptr( 1 ) ;

   lxw_styles_free(styles) ; 
}





/*
 * Write the <t> element for rich strings.
 */
/*
void
lxw_styles_write_string_fragment(lxw_styles *self, char *string)
*/
/*
HB_FUNC( LXW_STYLES_WRITE_STRING_FRAGMENT )
{ 
   lxw_styles *self = hb_parptr( 1 ) ;
   char *string = hb_parcx( 2 ) ; // No pasa.. (RIGC - 2019/05/28)

   lxw_styles_write_string_fragment(self, string) ; 
}
*/



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
lxw_styles_assemble_xml_file(lxw_styles *self)
*/
HB_FUNC( LXW_STYLES_ASSEMBLE_XML_FILE )
{ 
   lxw_styles *self = hb_parptr( 1 ) ;

   lxw_styles_assemble_xml_file(self) ; 
}


//eof
