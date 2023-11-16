/*****************************************************************************
 * format - A library for creating Excel XLSX format files.
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
#include "xlsxwriter/format.h"
#include "xlsxwriter/utility.h"

#include "hbapi.h"




/*
 * Create a new format object.
 *
 * lxw_format *
 * lxw_format_new(void)
 *
 */
HB_FUNC( LXW_FORMAT_NEW )
{
   lxw_format *format = lxw_format_new();
   hb_retptr( format ); 
}




/*
 * Free a format object.
 *
 * void
 * lxw_format_free(lxw_format *format)
 * 
 */
HB_FUNC( LXW_FORMAT_FREE )
{ 
   lxw_format *format = hb_parptr( 1 ) ;

   lxw_format_free( format ); 
}




/*
 * Check a user input color.
 *
 * lxw_color_t
 * lxw_format_check_color(lxw_color_t color)
 *
 */
HB_FUNC( LXW_FORMAT_CHECK_COLOR )
{ 
   lxw_color_t color = hb_parnl( 1 ) ;

   hb_retnl( lxw_format_check_color( color ) ); 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/






/*
 * Returns a font struct suitable for hashing as a lookup key.
 *
 * lxw_font *
 * lxw_format_get_font_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_FONT_KEY )
{ 
   lxw_format *self = hb_parptr(1 ) ;

   hb_retptr( lxw_format_get_font_key( self ) ); 
}




/*
 * Returns a border struct suitable for hashing as a lookup key.
 *
 * lxw_border *
 * lxw_format_get_border_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_BORDER_KEY )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   hb_retptr( lxw_format_get_border_key( self ) ); 
}




/*
 * Returns a pattern fill struct suitable for hashing as a lookup key.
 *
 * lxw_fill *
 * lxw_format_get_fill_key(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_FILL_KEY )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   hb_retptr( lxw_format_get_fill_key( self ) ); 
}




/*
 * Returns the XF index number used by Excel to identify a format.
 *
 * int32_t
 * lxw_format_get_xf_index(lxw_format *self)
 *
 */
HB_FUNC( LXW_FORMAT_GET_XF_INDEX )
{ 
   lxw_format *self = hb_parptr(1 ) ;

   hb_retnl( lxw_format_get_xf_index( self ) ); 
}



/*
 * Set the font_name property.
 *
 * void
 * format_set_font_name(lxw_format *self, const char *font_name)
 *
 */
HB_FUNC( FORMAT_SET_FONT_NAME )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   const char *font_name = hb_parcx( 2 ) ;

   format_set_font_name( self, font_name ); 
}




/*
 * Set the font_size property.
 *
 * void
 * format_set_font_size(lxw_format *self, double size)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SIZE )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   double size = hb_parnd( 2 ) ;

   format_set_font_size( self, size ); 
}




/*
 * Set the font_color property.
 *
 * void
 * format_set_font_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_FONT_COLOR )
{ 
   lxw_format *self = hb_parptr(1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   self->font_color = lxw_format_check_color(color);
}




/*
 * Set the bold property.
 *
 * void
 * format_set_bold(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_BOLD )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   self->bold = LXW_TRUE;
}




/*
 * Set the italic property.
 *
 * void
 * format_set_italic(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_ITALIC )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   self->italic = LXW_TRUE;
}




/*
 * Set the underline property.
 *
 * void
 * format_set_underline(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_UNDERLINE )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_underline( self, style ); 
}




/*
 * Set the font_strikeout property.
 *
 * void
 * format_set_font_strikeout(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_STRIKEOUT )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_font_strikeout( self ); 
}




/*
 * Set the font_script property.
 *
 * void
 * format_set_font_script(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SCRIPT )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_font_script( self, style ); 
}




/*
 * Set the font_outline property.
 *
 * void
 * format_set_font_outline(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_OUTLINE )
{ 
   lxw_format *self = hb_parptr(1 ) ;

   format_set_font_outline( self ); 
}




/*
 * Set the font_shadow property.
 *
 * void
 * format_set_font_shadow(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SHADOW )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_font_shadow( self ); 
}




/*
 * Set the num_format property.
 *
 * void
 * format_set_num_format(lxw_format *self, const char *num_format)
 *
 */
HB_FUNC( FORMAT_SET_NUM_FORMAT )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   const char *num_format = hb_parcx( 2 ) ;

   format_set_num_format( self, num_format ); 
}




/*
 * Set the unlocked property.
 *
 * void
 * format_set_unlocked(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_UNLOCKED )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_unlocked( self ); 
}




/*
 * Set the hidden property.
 *
 * void
 * format_set_hidden(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_HIDDEN )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_hidden( self ); 
}




/*
 * Set the align property.
 *
 * void
 * format_set_align(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_ALIGN )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_align( self, value ); 
}




/*
 * Set the text_wrap property.
 *
 * void
 * format_set_text_wrap(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_TEXT_WRAP )
{ 
   lxw_format *self = hb_parptr(1 ) ;

   format_set_text_wrap( self ); 
}




/*
 * Set the rotation property.
 *
 * void
 * format_set_rotation(lxw_format *self, int16_t angle)
 *
 */
HB_FUNC( FORMAT_SET_ROTATION )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   int16_t angle = hb_parnl( 2 ) ;

   format_set_rotation( self, angle ); 
}




/*
 * Set the indent property.
 *
 * void
 * format_set_indent(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_INDENT )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_indent( self, value ); 
}




/*
 * Set the shrink property.
 *
 * void
 * format_set_shrink(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_SHRINK )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_shrink( self ); 
}




/*
 * Set the text_justlast property.
 *
 * void
 * format_set_text_justlast(lxw_format *self)
 *
 */
/*
HB_FUNC( FORMAT_SET_TEXT_JUSTLAST )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_text_justlast( self ); 
}
*/



/*
 * Set the pattern property.
 *
 * void
 * format_set_pattern(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_PATTERN )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_pattern( self, value ); 
}




/*
 * Set the bg_color property.
 *
 * void
 * format_set_bg_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BG_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   format_set_bg_color( self, color ); 
}




/*
 * Set the fg_color property.
 *
 * void
 * format_set_fg_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_FG_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   format_set_fg_color( self, color ); 
}




/*
 * Set the border property.
 *
 * void
 * format_set_border(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_BORDER )
{ 
   lxw_format *self = hb_parptr(1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_border( self, style ); 
}




/*
 * Set the border_color property.
 *
 * void
 * format_set_border_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BORDER_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   format_set_border_color( self, color ); 
}




/*
 * Set the bottom property.
 *
 * void
 * format_set_bottom(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_BOTTOM )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_bottom( self, style ); 
}




/*
 * Set the bottom_color property.
 *
 * void
 * format_set_bottom_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_BOTTOM_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   format_set_bottom_color( self, color) ; 
}




/*
 * Set the left property.
 *
 * void
 * format_set_left(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_LEFT )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_left( self, style ); 
}




/*
 * Set the left_color property.
 *
 * void
 * format_set_left_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_LEFT_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl( 2 ) ;

   format_set_left_color( self, color ); 
}




/*
 * Set the right property.
 *
 * void
 * format_set_right(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_RIGHT )
{ 
   lxw_format *self = hb_parptr(1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_right( self, style ); 
}




/*
 * Set the right_color property.
 *
 * void
 * format_set_right_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_RIGHT_COLOR )
{ 
   lxw_format *self = hb_parptr(1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   format_set_right_color( self, color ); 
}




/*
 * Set the top property.
 *
 * void
 * format_set_top(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_TOP )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_top( self, style ); 
}




/*
 * Set the top_color property.
 *
 * void
 * format_set_top_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_TOP_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   format_set_top_color( self, color ); 
}




/*
 * Set the diag_type property.
 *
 * void
 * format_set_diag_type(lxw_format *self, uint8_t type)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_TYPE )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   format_set_diag_type( self, type ); 
}




/*
 * Set the diag_color property.
 *
 * void
 * format_set_diag_color(lxw_format *self, lxw_color_t color)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_COLOR )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   lxw_color_t color = hb_parnl(2 ) ;

   format_set_diag_color( self, color ); 
}




/*
 * Set the diag_border property.
 *
 * void
 * format_set_diag_border(lxw_format *self, uint8_t style)
 *
 */
HB_FUNC( FORMAT_SET_DIAG_BORDER )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t style = hb_parni( 2 ) ;

   format_set_diag_border( self, style ); 
}




/*
 * Set the num_format_index property.
 *
 * void
 * format_set_num_format_index(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_NUM_FORMAT_INDEX )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_num_format_index( self, value ); 
}




/*
 * Set the valign property.
 *
 * void
 * format_set_valign(lxw_format *self, uint8_t value)
 *
 */
/*
HB_FUNC( FORMAT_SET_VALIGN )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_valign( self, value ); 
}
*/



/*
 * Set the reading_order property.
 *
 * void
 * format_set_reading_order(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_READING_ORDER )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_reading_order( self, value ); 
}




/*
 * Set the font_family property.
 *
 * void
 * format_set_font_family(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_FONT_FAMILY )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_font_family( self, value ); 
}




/*
 * Set the font_charset property.
 *
 * void
 * format_set_font_charset(lxw_format *self, uint8_t value)
 *
 */
HB_FUNC( FORMAT_SET_FONT_CHARSET )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   uint8_t value = hb_parni( 2 ) ;

   format_set_font_charset( self, value ); 
}




/*
 * Set the font_scheme property.
 *
 * void
 * format_set_font_scheme(lxw_format *self, const char *font_scheme)
 *
 */
HB_FUNC( FORMAT_SET_FONT_SCHEME )
{ 
   lxw_format *self = hb_parptr( 1 ) ;
   const char *font_scheme = hb_parcx( 2 ) ;

   format_set_font_scheme( self, font_scheme ); 
}




/*
 * Set the font_condense property.
 *
 * void
 * format_set_font_condense(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_CONDENSE )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_font_condense( self ); 
}




/*
 * Set the font_extend property.
 *
 * void
 * format_set_font_extend(lxw_format *self)
 *
 */
HB_FUNC( FORMAT_SET_FONT_EXTEND )
{ 
   lxw_format *self = hb_parptr( 1 ) ;

   format_set_font_extend( self ); 
}


//eof
