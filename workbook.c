/*****************************************************************************
 * workbook - A library for creating Excel XLSX workbook files.
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
#include "xlsxwriter/workbook.h"
#include "xlsxwriter/utility.h"
#include "xlsxwriter/packager.h"
#include "xlsxwriter/hash_table.h"


#include "hbapierr.h"
#include "hbapiitm.h"


/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Free a workbook object.
 *
 * void
 * lxw_workbook_free(lxw_workbook *workbook)
 *
 */
HB_FUNC( LXW_WORKBOOK_FREE )
{ 
   lxw_workbook *workbook = hb_parptr( 1 ) ;

   lxw_workbook_free( workbook ); 
}




/*
 * Set the default index for each format. This is only used for testing.
 *
 * void
 * lxw_workbook_set_default_xf_indices(lxw_workbook *self)
 *
 */
HB_FUNC( LXW_WORKBOOK_SET_DEFAULT_XF_INDICES )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;

   lxw_workbook_set_default_xf_indices( self ); 
}





/*****************************************************************************
 *
 * XML file assembly functions.
 *
 ****************************************************************************/


/*
 * Assemble and write the XML file.
 *
 * void
 * lxw_workbook_assemble_xml_file(lxw_workbook *self)
 *
 */
HB_FUNC( LXW_WORKBOOK_ASSEMBLE_XML_FILE )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;

   lxw_workbook_assemble_xml_file( self ); 
}




/*
 *
 * Public functions.
 *
 ****************************************************************************/





/*
 * Create a new workbook object.
 *
 * lxw_workbook *
 * workbook_new(const char *filename)
 *
 */
HB_FUNC( WORKBOOK_NEW )
{ 
   const char *filename = hb_parcx( 1 ) ;

   hb_retptr( workbook_new( filename ) ); 
}

/* Deprecated function name for backwards compatibility. */
/*
lxw_workbook *
new_workbook(const char *filename)
*/
HB_FUNC( NEW_WORKBOOK )
{
   const char *filename = hb_parcx( 1 ) ;
   hb_retptr( workbook_new_opt(filename, NULL) );
}


/*
 * Create a new workbook object with options.
 *
 * lxw_workbook *
 * workbook_new_opt(const char *filename, lxw_workbook_options *options)
 *
 */
HB_FUNC( WORKBOOK_NEW_OPT )
{
   const char *filename = hb_parcx( 1 );
   lxw_workbook_options *options = hb_param( 2, HB_IT_ANY );
   if HB_ISNIL( 2 )
   {
      workbook_new_opt(filename, NULL);
   }
   else
   {
      workbook_new_opt(filename, options);
   }
}




/*
 * Add a new worksheet to the Excel workbook.
 *
 * lxw_worksheet *
 * workbook_add_worksheet(lxw_workbook *self, const char *sheetname)
 *
 */
HB_FUNC( WORKBOOK_ADD_WORKSHEET )
{ 
   lxw_workbook *self = hb_parptr(1);
   const char *sheetname = hb_parcx( 2 );
   if HB_ISNIL( 2 ) 
   {
      hb_retptr( workbook_add_worksheet( self, NULL ) );
   }
   else
   {
      hb_retptr( workbook_add_worksheet( self, sheetname ) );
   }
}




/*
 * Add a new chartsheet to the Excel workbook.
 *
 * lxw_chartsheet *
 * workbook_add_chartsheet(lxw_workbook *self, const char *sheetname)
 *
 */
HB_FUNC( WORKBOOK_ADD_CHARTSHEET )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;

   hb_retptr( workbook_add_chartsheet( self, sheetname ) ); 
}




/*
 * Add a new chart to the Excel workbook.
 *
 * lxw_chart *
 * workbook_add_chart(lxw_workbook *self, uint8_t type)
 *
 */
HB_FUNC( WORKBOOK_ADD_CHART )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   hb_retptr( workbook_add_chart( self, type ) ); 
}




/*
 * Add a new format to the Excel workbook.
 *
 * lxw_format *
 * workbook_add_format(lxw_workbook *self)
 *
 */
HB_FUNC( WORKBOOK_ADD_FORMAT )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;

   hb_retptr( workbook_add_format( self ) ); 
}




/*
 * Call finalization code and close file.
 *
 * lxw_error
 * workbook_close(lxw_workbook *self)
 *
 */
HB_FUNC( WORKBOOK_CLOSE )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;

   hb_retni( workbook_close( self ) ); 
}




/*
 * Create a defined name in Excel. We handle global/workbook level names and
 * local/worksheet names.
 *
 * lxw_error
 * workbook_define_name(lxw_workbook *self, const char *name,
 *    const char *formula)
 *
 */
HB_FUNC( WORKBOOK_DEFINE_NAME )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   const char *formula = hb_parcx( 3 ) ;

   hb_retni( workbook_define_name( self, name, formula ) ); 
}




/*
 * Set the document properties such as Title, Author etc.
 *
 * lxw_error
 * workbook_set_properties(lxw_workbook *self, lxw_doc_properties *user_props)
 *
 */
HB_FUNC( WORKBOOK_SET_PROPERTIES )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   lxw_doc_properties *user_props = hb_parptr(2 ) ;

   hb_retni( workbook_set_properties( self, user_props ) ); 
}




/*
 * Set a string custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_string(lxw_workbook *self, const char *name,
 *      const char *value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_STRING )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   const char *value = hb_parcx( 3 ) ;

   hb_retni( workbook_set_custom_property_string( self, name, value ) ); 
}




/*
 * Set a double number custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_number(lxw_workbook *self, const char *name,
 *       double value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_NUMBER )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   double value = hb_parnd( 3 ) ;

   hb_retni( workbook_set_custom_property_number( self, name, value ) ); 
}




/*
 * Set a integer number custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_integer(lxw_workbook *self, const char *name,
 *        int32_t value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_INTEGER )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   int32_t value = hb_parnl(3 ) ;

   hb_retni( workbook_set_custom_property_integer( self, name, value ) ); 
}




/*
 * Set a boolean custom document property.
 *
 * lxw_error
 * workbook_set_custom_property_boolean(lxw_workbook *self, const char *name,
 *          uint8_t value)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_BOOLEAN )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   uint8_t value = hb_parni( 3 ) ;

   hb_retni( workbook_set_custom_property_boolean( self, name, value ) ); 
}




/*
 * Set a datetime custom document property.
 *
 * lxw_error 
 * workbook_set_custom_property_datetime(lxw_workbook *self, const char *name,
 *           lxw_datetime *datetime)
 *
 */
HB_FUNC( WORKBOOK_SET_CUSTOM_PROPERTY_DATETIME )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;
   lxw_datetime *datetime = hb_parptr(3 ) ;

   hb_retni( workbook_set_custom_property_datetime( self, name, datetime ) ); 
}




/*
 * Get a worksheet object from its name.
 *
 * lxw_worksheet *
 * workbook_get_worksheet_by_name(lxw_workbook *self, const char *name)
 *
 */
HB_FUNC( WORKBOOK_GET_WORKSHEET_BY_NAME )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   hb_retptr( workbook_get_worksheet_by_name( self, name ) ); 
}




/*
 * Get a chartsheet object from its name.
 *
 * lxw_chartsheet *
 * workbook_get_chartsheet_by_name(lxw_workbook *self, const char *name)
 *
 */
HB_FUNC( WORKBOOK_GET_CHARTSHEET_BY_NAME )
{ 
   lxw_workbook *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   hb_retptr( workbook_get_chartsheet_by_name( self, name ) ); 
}




//eof
