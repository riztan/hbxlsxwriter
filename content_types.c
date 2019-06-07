

/*****************************************************************************
 * content_types - A library for creating Excel XLSX content_types files.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2019, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 */

#include "xlsxwriter/xmlwriter.h"
#include "xlsxwriter/content_types.h"
#include "xlsxwriter/utility.h"


#include "hbapi.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/


/*
 * Create a new content_types object.
 */
/*
lxw_content_types *
lxw_content_types_new(void)
*/
HB_FUNC( LXW_CONTENT_TYPES_NEW )
{ 
   hb_retptr( lxw_content_types_new() ); 
}





/*
 * Free a content_types object.
 */
/*
void
lxw_content_types_free(lxw_content_types *content_types)
*/
HB_FUNC( LXW_CONTENT_TYPES_FREE )
{ 
   lxw_content_types *content_types = hb_parptr( 1 ) ;

   lxw_content_types_free(content_types) ; 
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
lxw_content_types_assemble_xml_file(lxw_content_types *self)
*/
HB_FUNC( LXW_CONTENT_TYPES_ASSEMBLE_XML_FILE )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;

   lxw_content_types_assemble_xml_file(self) ; 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/


/*
 * Add elements to the ContentTypes defaults.
 */
/*
void
lxw_ct_add_default(lxw_content_types *self, const char *key,
   const char *value)
*/
HB_FUNC( LXW_CT_ADD_DEFAULT )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *key = hb_parcx( 2 ) ;
   const char *value = hb_parcx( 3 ) ;

   lxw_ct_add_default(self, key, value) ; 
}





/*
 * Add elements to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_override(lxw_content_types *self, const char *key,
   const char *value)
*/
HB_FUNC( LXW_CT_ADD_OVERRIDE )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *key = hb_parcx( 2 ) ;
   const char *value = hb_parcx( 3 ) ;

   lxw_ct_add_override(self, key, value) ; 
}





/*
 * Add the name of a worksheet to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_worksheet_name(lxw_content_types *self, const char *name)
*/
HB_FUNC( LXW_CT_ADD_WORKSHEET_NAME )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   lxw_ct_add_worksheet_name(self, name) ; 
}





/*
 * Add the name of a chartsheet to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_chartsheet_name(lxw_content_types *self, const char *name)
*/
HB_FUNC( LXW_CT_ADD_CHARTSHEET_NAME )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   lxw_ct_add_chartsheet_name(self, name) ; 
}





/*
 * Add the name of a chart to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_chart_name(lxw_content_types *self, const char *name)
*/
HB_FUNC( LXW_CT_ADD_CHART_NAME )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   lxw_ct_add_chart_name(self, name) ; 
}





/*
 * Add the name of a drawing to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_drawing_name(lxw_content_types *self, const char *name)
*/
HB_FUNC( LXW_CT_ADD_DRAWING_NAME )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   lxw_ct_add_drawing_name(self, name) ; 
}





/*
 * Add the sharedStrings link to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_shared_strings(lxw_content_types *self)
*/
HB_FUNC( LXW_CT_ADD_SHARED_STRINGS )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;

   lxw_ct_add_shared_strings(self) ; 
}





/*
 * Add the calcChain link to the ContentTypes overrides.
 */
/*
void
lxw_ct_add_calc_chain(lxw_content_types *self)
*/
HB_FUNC( LXW_CT_ADD_CALC_CHAIN )
{ 
   lxw_content_types *self = hb_parptr( 1 ) ;

   lxw_ct_add_calc_chain(self) ; 
}


//eof()
