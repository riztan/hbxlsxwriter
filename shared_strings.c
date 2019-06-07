/*****************************************************************************
 * shared_strings - A library for creating Excel XLSX sst files.
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
#include "xlsxwriter/shared_strings.h"
#include "xlsxwriter/utility.h"
#include <ctype.h>

#include "hbapi.h"




/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new SST SharedString object.
 */
/*
lxw_sst *
lxw_sst_new(void)
*/
HB_FUNC( LXW_SST_NEW )
{ 
   hb_retptr( lxw_sst_new() ); 
}





/*
 * Free a SST SharedString table object.
 */
/*
void
lxw_sst_free(lxw_sst *sst)
*/
HB_FUNC( LXW_SST_FREE )
{ 
   lxw_sst *sst = hb_parptr( 1 ) ;

   lxw_sst_free(sst) ; 
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
lxw_sst_assemble_xml_file(lxw_sst *self)
*/
HB_FUNC( LXW_SST_ASSEMBLE_XML_FILE )
{ 
   lxw_sst *self = hb_parptr( 1 ) ;

   lxw_sst_assemble_xml_file(self) ; 
}



//eof
