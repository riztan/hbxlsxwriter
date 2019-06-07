/*****************************************************************************
 * drawing - A library for creating Excel XLSX drawing files.
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
#include "xlsxwriter/common.h"
#include "xlsxwriter/drawing.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/


/*
 * Create a new drawing collection.
 */
/*
lxw_drawing *
lxw_drawing_new(void)
*/
HB_FUNC( LXW_DRAWING_NEW )
{ 
   hb_retptr( lxw_drawing_new() ); 
}





/*
 * Free a drawing object.
 */
/*
void
lxw_free_drawing_object(lxw_drawing_object *drawing_object)
*/
HB_FUNC( LXW_FREE_DRAWING_OBJECT )
{ 
   lxw_drawing_object *drawing_object = hb_parptr( 1 ) ;

   lxw_free_drawing_object(drawing_object) ; 
}





/*
 * Free a drawing collection.
 */
/*
void
lxw_drawing_free(lxw_drawing *drawing)
*/
HB_FUNC( LXW_DRAWING_FREE )
{ 
   lxw_drawing *drawing = hb_parptr( 1 ) ;

   lxw_drawing_free(drawing) ; 
}





/*
 * Add a drawing object to the drawing collection.
 */
/*
void
lxw_add_drawing_object(lxw_drawing *drawing,
              lxw_drawing_object *drawing_object)
*/
HB_FUNC( LXW_ADD_DRAWING_OBJECT )
{ 
   lxw_drawing *drawing = hb_parptr( 1 ) ;
   lxw_drawing_object *drawing_object = hb_parptr( 2 ) ;

   lxw_add_drawing_object(drawing, drawing_object) ; 
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
lxw_drawing_assemble_xml_file(lxw_drawing *self)
*/
HB_FUNC( LXW_DRAWING_ASSEMBLE_XML_FILE )
{ 
   lxw_drawing *self = hb_parptr( 1 ) ;

   lxw_drawing_assemble_xml_file(self) ; 
}


//eof
