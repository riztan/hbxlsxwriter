/*****************************************************************************
 * hb_misc - A library for creating Excel XLSX format files.
 *
 * Complementary functions for use with xlsxwriter on Harbour
 *
 * Copyright 2019, Riztan Gutierrez, riztan@gmail.com. See LICENSE.txt.
 *
 */

#include "xlsxwriter/utility.h"
#include "xlsxwriter/chart.h"

#include "hbapi.h"
#include "hbapiitm.h"


/***************************
 * HARBOUR  FONT UTILITIES
 ***************************/

/*
//ToDo:  No funciona correctamente, se debe continuar (RIGC-20190602)
HB_FUNC( HB_LXW_CHART_FONT_NEW )
{
   lxw_chart_font font;

   font.name = "Arial";
   font.size = 11;
   font.bold = LXW_TRUE;
   font.italic = LXW_TRUE;
   font.underline = LXW_TRUE;
   font.rotation = 0;
   font.color = LXW_COLOR_BLUE;
   font.pitch_family = LXW_COLOR_BLUE;

   hb_retptr( &font );
}
*/

HB_FUNC( HB_LXW_FONT_NEW )
{
   lxw_format *format = lxw_format_new();
   hb_retptr( lxw_format_get_font_key( format ) );
}


HB_FUNC( HB_LXW_FONT_SET_NAME )
{
   lxw_chart_font *font = hb_parptr( 1 );
   const char *name = hb_parcx( 2 );
   font->name = lxw_strdup( name );
}


HB_FUNC( HB_LXW_FONT_SET_COLOR )
{
   lxw_chart_font *font = hb_parptr( 1 );
   lxw_color_t color = hb_parnl( 2 );
   font->color = color;
}


HB_FUNC( HB_LXW_FONT_SET_BOLD )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
//printf( "BOLD \n" );
      font->bold = LXW_TRUE;
   }
   else
   {
//printf( "NOT BOLD \n" );
      font->bold = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_ITALIC )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
      font->italic = LXW_TRUE;
   }
   else
   {
      font->italic = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_UNDERLINE )
{
   lxw_chart_font *font = hb_parptr( 1 );
   if( hb_parl( 2 ) == 1 )
   {
      font->underline = LXW_TRUE;
   }
   else
   {
      font->underline = LXW_FALSE;
   }
}


HB_FUNC( HB_LXW_FONT_SET_ROTATION )
{
   lxw_chart_font *font = hb_parptr( 1 );
   int rotation = hb_parni( 2 );

   font->rotation = rotation;
}



//eof
