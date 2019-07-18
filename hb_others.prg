/*
 * hb_others.prg  A library for creating Excel XLSX worksheet files.
 *
 * Used in conjunction with the libxlsxwriter library of
 * John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 */
/*
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

FUNCTION hb_lxw_dv( hDV, cKey, uValue )

   LOCAL aKey

   aKey := { "validate", "criteria", "ignore_blank", ;
      "show_input", "show_error", "error_type", ;
      "dropdown", "is_between", "value_number", ;
      "value_formula", "value_list", "value_datetime", ;
      "minimum_number", "minimum_formula", "minimum_datetime", ;
      "maximum_number", "maximum_formula", "maximum_datetime", ;
      "input_title", "input_message", "error_title", ;
      "error_message" }

   IF AScan( aKey, {| key| key == cKey } ) = 0
      RETURN .F.
   ENDIF

   hb_HSet( hDV, cKey, uValue )

   RETURN .T.



// eof
