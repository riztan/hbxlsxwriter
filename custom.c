/*****************************************************************************
 * custom - A library for creating Excel custom property files.
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
#include "xlsxwriter/custom.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new custom object.
 */
/*
lxw_custom *
lxw_custom_new(void)
*/
HB_FUNC( LXW_CUSTOM_NEW )
{ 
   hb_retptr( lxw_custom_new() ); 
}





/*
 * Free a custom object.
 */
/*
void
lxw_custom_free(lxw_custom *custom)
*/
HB_FUNC( LXW_CUSTOM_FREE )
{ 
   lxw_custom *custom = hb_parptr( 1 ) ;

   lxw_custom_free(custom) ; 
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
lxw_custom_assemble_xml_file(lxw_custom *self)
*/
HB_FUNC( LXW_CUSTOM_ASSEMBLE_XML_FILE )
{ 
   lxw_custom *self = hb_parptr( 1 ) ;

   lxw_custom_assemble_xml_file(self) ; 
}


//eof
