/*****************************************************************************
 * relationships - A library for creating Excel XLSX relationships files.
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
#include "xlsxwriter/relationships.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"



/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Create a new relationships object.
 */
/*
lxw_relationships *
lxw_relationships_new(void)
*/
HB_FUNC( LXW_RELATIONSHIPS_NEW )
{ 
   hb_retptr( lxw_relationships_new() ); 
}





/*
 * Free a relationships object.
 */
/*
void
lxw_free_relationships(lxw_relationships *rels)
*/
HB_FUNC( LXW_FREE_RELATIONSHIPS )
{ 
   lxw_relationships *rels = hb_parptr( 1 ) ;

   lxw_free_relationships(rels) ; 
}





/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/






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
lxw_relationships_assemble_xml_file(lxw_relationships *self)
*/
HB_FUNC( LXW_RELATIONSHIPS_ASSEMBLE_XML_FILE )
{ 
   lxw_relationships *self = hb_parptr( 1 ) ;

   lxw_relationships_assemble_xml_file(self) ; 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/






/*
 * Add a document relationship to XLSX .rels xml files.
 */
/*
void
lxw_add_document_relationship(lxw_relationships *self, const char *type,
           const char *target)
*/
HB_FUNC( LXW_ADD_DOCUMENT_RELATIONSHIP )
{ 
   lxw_relationships *self = hb_parptr( 1 ) ;
   const char *type = hb_parcx( 2 ) ;
   const char *target = hb_parcx( 3 ) ;

   lxw_add_document_relationship(self, type, target) ; 
}





/*
 * Add a package relationship to XLSX .rels xml files.
 */
/*
void
lxw_add_package_relationship(lxw_relationships *self, const char *type,
             const char *target)
*/
HB_FUNC( LXW_ADD_PACKAGE_RELATIONSHIP )
{ 
   lxw_relationships *self = hb_parptr( 1 ) ;
   const char *type = hb_parcx( 2 ) ;
   const char *target = hb_parcx( 3 ) ;

   lxw_add_package_relationship(self, type, target) ; 
}





/*
 * Add a MS schema package relationship to XLSX .rels xml files.
 */
/*
void
lxw_add_ms_package_relationship(lxw_relationships *self, const char *type,
                  const char *target)
*/
HB_FUNC( LXW_ADD_MS_PACKAGE_RELATIONSHIP )
{ 
   lxw_relationships *self = hb_parptr( 1 ) ;
   const char *type = hb_parcx( 2 ) ;
   const char *target = hb_parcx( 3 ) ;

   lxw_add_ms_package_relationship(self, type, target) ; 
}


//eof
