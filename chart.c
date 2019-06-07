/*****************************************************************************
 * chart - A library for creating Excel XLSX chart files.
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
#include "xlsxwriter/chart.h"
#include "xlsxwriter/utility.h"


#include "hbapi.h"



/*****************************************************************************
 *
 * Private functions.
 *
 ****************************************************************************/

/*
 * Free a chart object.
 */
/*
void
lxw_chart_free(lxw_chart *chart)
*/
HB_FUNC( LXW_CHART_FREE )
{ 
   lxw_chart *chart = hb_parptr( 1 ) ;
   lxw_chart_free(chart) ; 
}





/*
 * Create a new chart object.
 */
/*
lxw_chart *
lxw_chart_new(uint8_t type)
*/
HB_FUNC( LXW_CHART_NEW )
{ 
   uint8_t type = hb_parni( 1 ) ;

   hb_retptr( lxw_chart_new(type) ); 
}





/*
 * Set an axis number format.
 */
/*
void
_chart_axis_set_default_num_format(lxw_chart_axis *axis, char *num_format)
*/
/*
HB_FUNC( _CHART_AXIS_SET_DEFAULT_NUM_FORMAT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   char *num_format = hb_parcx( 2 ) ;

   _chart_axis_set_default_num_format(axis, num_format) ; 
}
*/




/*
 * Verify that a X/Y error bar property is support for the chart type.
 * All chart types, except Bar have Y error bars. Only Bar and Scatter
 * support X error bars.
 */
/*
lxw_error
_chart_check_error_bars(lxw_series_error_bars *error_bars, char *property)
*/
/*
HB_FUNC( _CHART_CHECK_ERROR_BARS )
{ 
   lxw_series_error_bars *error_bars = hb_parptr( 1 ) ;
   char *property = hb_parcx( 2 ) ;

   hb_retni( _chart_check_error_bars(error_bars, property) ); 
}
*/



/*****************************************************************************
 *
 * XML functions.
 *
 ****************************************************************************/


/*
 * Assemble and write the XML file.
 */
/*
void
lxw_chart_assemble_xml_file(lxw_chart *self)
*/
HB_FUNC( LXW_CHART_ASSEMBLE_XML_FILE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;

   lxw_chart_assemble_xml_file(self) ; 
}





/*****************************************************************************
 *
 * Public functions.
 *
 ****************************************************************************/

/*
 * Add data to a data cache in a range object, for testing only.
 */
/*
lxw_error
lxw_chart_add_data_cache(lxw_series_range *range, uint8_t *data,
   uint16_t rows, uint8_t cols, uint8_t col)
*/
HB_FUNC( LXW_CHART_ADD_DATA_CACHE )
{ 
   lxw_series_range *range = hb_parptr( 1 ) ;
   uint8_t *data = hb_parptr( 2 ) ;
   uint16_t rows = hb_parnl( 3 ) ;
   uint8_t cols = hb_parni( 4 ) ;
   uint8_t col = hb_parni( 5 ) ;

   hb_retni( lxw_chart_add_data_cache(range, data, rows, cols, col) ); 
}





/*
 * Insert an image into the worksheet.
 */
/*
lxw_chart_series *
chart_add_series(lxw_chart *self, const char *categories, const char *values)
*/
HB_FUNC( CHART_ADD_SERIES )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   const char *categories = hb_parcx( 2 ) ;
   const char *values = hb_parcx( 3 ) ;

   hb_retptr( chart_add_series(self, categories, values) ); 
}





/*
 * Set on of the 48 built-in Excel chart styles.
 */
/*
void
chart_set_style(lxw_chart *self, uint8_t style_id)
*/
HB_FUNC( CHART_SET_STYLE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint8_t style_id = hb_parni( 2 ) ;

   chart_set_style(self, style_id) ; 
}





/*
 * Set a user defined name for a series.
 */
/*
void
chart_series_set_name(lxw_chart_series *series, const char *name)
*/
HB_FUNC( CHART_SERIES_SET_NAME )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   chart_series_set_name(series, name) ; 
}





/*
 * Set an axis caption, with a range instead or a formula..
 */
/*
void
chart_series_set_name_range(lxw_chart_series *series, const char *sheetname,
   lxw_row_t row, lxw_col_t col)
*/
HB_FUNC( CHART_SERIES_SET_NAME_RANGE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;
   lxw_row_t row = hb_parni( 3 ) ;
   lxw_col_t col = hb_parni( 4 ) ;

   chart_series_set_name_range(series, sheetname, row, col) ; 
}





/*
 * Set the categories range for a series.
 */
/*
void
chart_series_set_categories(lxw_chart_series *series, const char *sheetname,
   lxw_row_t first_row, lxw_col_t first_col,
   lxw_row_t last_row, lxw_col_t last_col)
*/
HB_FUNC( CHART_SERIES_SET_CATEGORIES )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;
   lxw_row_t first_row = hb_parni( 3 ) ;
   lxw_col_t first_col = hb_parni( 4 ) ;
   lxw_row_t last_row = hb_parni( 5 ) ;
   lxw_col_t last_col = hb_parni( 6 ) ;

   chart_series_set_categories(series, sheetname, first_row, first_col, last_row, last_col) ; 
}





/*
 * Set the values range for a series.
 */
/*
void
chart_series_set_values(lxw_chart_series *series, const char *sheetname,
   lxw_row_t first_row, lxw_col_t first_col,
   lxw_row_t last_row, lxw_col_t last_col)
*/
HB_FUNC( CHART_SERIES_SET_VALUES )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;
   lxw_row_t first_row = hb_parni( 3 ) ;
   lxw_col_t first_col = hb_parni( 4 ) ;
   lxw_row_t last_row = hb_parni( 5 ) ;
   lxw_col_t last_col = hb_parni( 6 ) ;

   chart_series_set_values(series, sheetname, first_row, first_col, last_row, last_col) ; 
}





/*
 * Set a line type for a series.
 */
/*
void
chart_series_set_line(lxw_chart_series *series, lxw_chart_line *line)
*/
HB_FUNC( CHART_SERIES_SET_LINE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_series_set_line(series, line) ; 
}





/*
 * Set a fill type for a series.
 */
/*
void
chart_series_set_fill(lxw_chart_series *series, lxw_chart_fill *fill)
*/
HB_FUNC( CHART_SERIES_SET_FILL )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_fill *fill = hb_parptr( 2 ) ;

   chart_series_set_fill(series, fill) ; 
}





/*
 * Invert the colors of a fill for a series.
 */
/*
void
chart_series_set_invert_if_negative(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_INVERT_IF_NEGATIVE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_invert_if_negative(series) ; 
}





/*
 * Set a pattern type for a series.
 */
/*
void
chart_series_set_pattern(lxw_chart_series *series, lxw_chart_pattern *pattern)
*/
HB_FUNC( CHART_SERIES_SET_PATTERN )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_pattern *pattern = hb_parptr( 2 ) ;

   chart_series_set_pattern(series, pattern) ; 
}





/*
 * Set a marker type for a series.
 */
/*
void
chart_series_set_marker_type(lxw_chart_series *series, uint8_t type)
*/
HB_FUNC( CHART_SERIES_SET_MARKER_TYPE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   chart_series_set_marker_type(series, type) ; 
}





/*
 * Set a marker size for a series.
 */
/*
void
chart_series_set_marker_size(lxw_chart_series *series, uint8_t size)
*/
HB_FUNC( CHART_SERIES_SET_MARKER_SIZE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t size = hb_parni( 2 ) ;

   chart_series_set_marker_size(series, size) ; 
}





/*
 * Set a line type for a series marker.
 */
/*
void
chart_series_set_marker_line(lxw_chart_series *series, lxw_chart_line *line)
*/
HB_FUNC( CHART_SERIES_SET_MARKER_LINE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_series_set_marker_line(series, line) ; 
}





/*
 * Set a fill type for a series marker.
 */
/*
void
chart_series_set_marker_fill(lxw_chart_series *series, lxw_chart_fill *fill)
*/
HB_FUNC( CHART_SERIES_SET_MARKER_FILL )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_fill *fill = hb_parptr( 2 ) ;

   chart_series_set_marker_fill(series, fill) ; 
}





/*
 * Set a pattern type for a series.
 */
/*
void
chart_series_set_marker_pattern(lxw_chart_series *series,
   lxw_chart_pattern *pattern)
*/
HB_FUNC( CHART_SERIES_SET_MARKER_PATTERN )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_pattern *pattern = hb_parptr( 2 ) ;

   chart_series_set_marker_pattern(series, pattern) ; 
}





/*
 * Store the horizontal page breaks on a worksheet.
 */
/*
lxw_error
chart_series_set_points(lxw_chart_series *series, lxw_chart_point *points[])
*/
/*
HB_FUNC( CHART_SERIES_SET_POINTS )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_point *points[] = hb_parptr( 2 ) ;

   hb_retni( chart_series_set_points(series, points[]) ); 
}
*/




/*
 * Set the smooth property for a line or scatter series.
 */
/*
void
chart_series_set_smooth(lxw_chart_series *series, uint8_t smooth)
*/
HB_FUNC( CHART_SERIES_SET_SMOOTH )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t smooth = hb_parni( 2 ) ;

   chart_series_set_smooth(series, smooth) ; 
}





/*
 * Turn on default data labels for a series.
 */
/*
void
chart_series_set_labels(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_LABELS )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_labels(series) ; 
}





/*
 * Set the data labels options for a series.
 */
/*
void
chart_series_set_labels_options(lxw_chart_series *series, uint8_t show_name,
   uint8_t show_category, uint8_t show_value)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_OPTIONS )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t show_name = hb_parni( 2 ) ;
   uint8_t show_category = hb_parni( 3 ) ;
   uint8_t show_value = hb_parni( 4 ) ;

   chart_series_set_labels_options(series, show_name, show_category, show_value) ; 
}





/*
 * Set the data labels separator for a series.
 */
/*
void
chart_series_set_labels_separator(lxw_chart_series *series, uint8_t separator)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_SEPARATOR )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t separator = hb_parni( 2 ) ;

   chart_series_set_labels_separator(series, separator) ; 
}





/*
 * Set the data labels position for a series.
 */
/*
void
chart_series_set_labels_position(lxw_chart_series *series, uint8_t position)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_POSITION )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t position = hb_parni( 2 ) ;

   chart_series_set_labels_position(series, position) ; 
}





/*
 * Set the data labels position for a series.
 */
/*
void
chart_series_set_labels_leader_line(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_LEADER_LINE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_labels_leader_line(series) ; 
}





/*
 * Turn on the data labels legend for a series.
 */
/*
void
chart_series_set_labels_legend(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_LEGEND )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_labels_legend(series) ; 
}





/*
 * Turn on the data labels percentage for a series.
 */
/*
void
chart_series_set_labels_percentage(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_PERCENTAGE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_labels_percentage(series) ; 
}





/*
 * Set an data labels number format.
 */
/*
void
chart_series_set_labels_num_format(lxw_chart_series *series,
   const char *num_format)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_NUM_FORMAT )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *num_format = hb_parcx( 2 ) ;

   chart_series_set_labels_num_format(series, num_format) ; 
}





/*
 * Set an data labels font.
 */
/*
void
chart_series_set_labels_font(lxw_chart_series *series, lxw_chart_font *font)
*/
HB_FUNC( CHART_SERIES_SET_LABELS_FONT )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;

   chart_series_set_labels_font(series, font) ; 
}





/*
 * Set the trendline for a chart series.
 */
/*
void
chart_series_set_trendline(lxw_chart_series *series, uint8_t type,
   uint8_t value)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;
   uint8_t value = hb_parni( 3 ) ;

   chart_series_set_trendline(series, type, value) ; 
}





/*
 * Set the trendline forecast for a chart series.
 */
/*
void
chart_series_set_trendline_forecast(lxw_chart_series *series, double forward,
   double backward)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_FORECAST )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   double forward = hb_parnd( 2 ) ;
   double backward = hb_parnd( 3 ) ;

   chart_series_set_trendline_forecast(series, forward, backward) ; 
}





/*
 * Display the equation for a series trendline.
 */
/*
void
chart_series_set_trendline_equation(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_EQUATION )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_trendline_equation(series) ; 
}





/*
 * Display the R squared value for a series trendline.
 */
/*
void
chart_series_set_trendline_r_squared(lxw_chart_series *series)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_R_SQUARED )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;

   chart_series_set_trendline_r_squared(series) ; 
}





/*
 * Set the trendline intercept for a chart series.
 */
/*
void
chart_series_set_trendline_intercept(lxw_chart_series *series,
   double intercept)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_INTERCEPT )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   double intercept = hb_parnd( 2 ) ;

   chart_series_set_trendline_intercept(series, intercept) ; 
}





/*
 * Set a line type for a series trendline.
 */
/*
void
chart_series_set_trendline_name(lxw_chart_series *series, const char *name)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_NAME )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   chart_series_set_trendline_name(series, name) ; 
}





/*
 * Set a line type for a series trendline.
 */
/*
void
chart_series_set_trendline_line(lxw_chart_series *series,
   lxw_chart_line *line)
*/
HB_FUNC( CHART_SERIES_SET_TRENDLINE_LINE )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_series_set_trendline_line(series, line) ; 
}





/*
 * Set the X or Y error bars from a chart series.
 */
/*
lxw_series_error_bars *
chart_series_get_error_bars(lxw_chart_series *series,
   lxw_chart_error_bar_axis axis_type)
*/
/*
HB_FUNC( CHART_SERIES_GET_ERROR_BARS )
{ 
   lxw_chart_series *series = hb_parptr( 1 ) ;
   lxw_chart_error_bar_axis axis_type = hb_parptr( 2 ) ;

   hb_retptr( chart_series_get_error_bars(series, axis_type) ); 
}
*/




/*
 * Set the error bars and type for a chart series.
 */
/*
void
chart_series_set_error_bars(lxw_series_error_bars *error_bars,
   uint8_t type, double value)
*/
HB_FUNC( CHART_SERIES_SET_ERROR_BARS )
{ 
   lxw_series_error_bars *error_bars = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;
   double value = hb_parnd( 3 ) ;

   chart_series_set_error_bars(error_bars, type, value) ; 
}





/*
 * Set the error bars direction for a chart series.
 */
/*
void
chart_series_set_error_bars_direction(lxw_series_error_bars *error_bars,
   uint8_t direction)
*/
HB_FUNC( CHART_SERIES_SET_ERROR_BARS_DIRECTION )
{ 
   lxw_series_error_bars *error_bars = hb_parptr( 1 ) ;
   uint8_t direction = hb_parni( 2 ) ;

   chart_series_set_error_bars_direction(error_bars, direction) ; 
}





/*
 * Set the error bars end cap type for a chart series.
 */
/*
void
chart_series_set_error_bars_endcap(lxw_series_error_bars *error_bars,
   uint8_t endcap)
*/
HB_FUNC( CHART_SERIES_SET_ERROR_BARS_ENDCAP )
{ 
   lxw_series_error_bars *error_bars = hb_parptr( 1 ) ;
   uint8_t endcap = hb_parni( 2 ) ;

   chart_series_set_error_bars_endcap(error_bars, endcap) ; 
}





/*
 * Set a line type for a series error bars.
 */
/*
void
chart_series_set_error_bars_line(lxw_series_error_bars *error_bars,
   lxw_chart_line *line)
*/
HB_FUNC( CHART_SERIES_SET_ERROR_BARS_LINE )
{ 
   lxw_series_error_bars *error_bars = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_series_set_error_bars_line(error_bars, line) ; 
}





/*
 * Get an axis pointer from a chart.
 */
/*
lxw_chart_axis *
chart_axis_get(lxw_chart *self, lxw_chart_axis_type axis_type)
*/
HB_FUNC( CHART_AXIS_GET )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_axis_type axis_type = hb_parni( 2 ) ;

   hb_retptr( chart_axis_get(self, axis_type) ); 
}





/*
 * Set an axis caption.
 */
/*
void
chart_axis_set_name(lxw_chart_axis *axis, const char *name)
*/
HB_FUNC( CHART_AXIS_SET_NAME )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   chart_axis_set_name(axis, name) ; 
}





/*
 * Set an axis caption, with a range instead or a formula.
 */
/*
void
chart_axis_set_name_range(lxw_chart_axis *axis, const char *sheetname,
   lxw_row_t row, lxw_col_t col)
*/
HB_FUNC( CHART_AXIS_SET_NAME_RANGE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;
   lxw_row_t row = hb_parni( 3 ) ;
   lxw_col_t col = hb_parni( 4 ) ;

   chart_axis_set_name_range(axis, sheetname, row, col) ; 
}





/*
 * Set an axis title/name font.
 */
/*
void
chart_axis_set_name_font(lxw_chart_axis *axis, lxw_chart_font *font)
*/
HB_FUNC( CHART_AXIS_SET_NAME_FONT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;

   chart_axis_set_name_font(axis, font) ; 
}





/*
 * Set an axis number font.
 */
/*
void
chart_axis_set_num_font(lxw_chart_axis *axis, lxw_chart_font *font)
*/
HB_FUNC( CHART_AXIS_SET_NUM_FONT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;

   chart_axis_set_num_font(axis, font) ; 
}





/*
 * Set an axis number format.
 */
/*
void
chart_axis_set_num_format(lxw_chart_axis *axis, const char *num_format)
*/
HB_FUNC( CHART_AXIS_SET_NUM_FORMAT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   const char *num_format = hb_parcx( 2 ) ;

   chart_axis_set_num_format(axis, num_format) ; 
}





/*
 * Set a line type for an axis.
 */
/*
void
chart_axis_set_line(lxw_chart_axis *axis, lxw_chart_line *line)
*/
HB_FUNC( CHART_AXIS_SET_LINE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_axis_set_line(axis, line) ; 
}





/*
 * Set a fill type for an axis.
 */
/*
void
chart_axis_set_fill(lxw_chart_axis *axis, lxw_chart_fill *fill)
*/
HB_FUNC( CHART_AXIS_SET_FILL )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_fill *fill = hb_parptr( 2 ) ;

   chart_axis_set_fill(axis, fill) ; 
}





/*
 * Set a pattern type for an axis.
 */
/*
void
chart_axis_set_pattern(lxw_chart_axis *axis, lxw_chart_pattern *pattern)
*/
HB_FUNC( CHART_AXIS_SET_PATTERN )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_pattern *pattern = hb_parptr( 2 ) ;

   chart_axis_set_pattern(axis, pattern) ; 
}





/*
 * Reverse the direction of an axis.
 */
/*
void
chart_axis_set_reverse(lxw_chart_axis *axis)
*/
HB_FUNC( CHART_AXIS_SET_REVERSE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;

   chart_axis_set_reverse(axis) ; 
}





/*
 * Set the axis crossing position.
 */
/*
void
chart_axis_set_crossing(lxw_chart_axis *axis, double value)
*/
HB_FUNC( CHART_AXIS_SET_CROSSING )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   double value = hb_parnd( 2 ) ;

   chart_axis_set_crossing(axis, value) ; 
}





/*
 * Set the axis crossing position as the max possible value.
 */
/*
void
chart_axis_set_crossing_max(lxw_chart_axis *axis)
*/
HB_FUNC( CHART_AXIS_SET_CROSSING_MAX )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;

   chart_axis_set_crossing_max(axis) ; 
}





/*
 * Turn off/hide the axis.
 */
/*
void
chart_axis_off(lxw_chart_axis *axis)
*/
HB_FUNC( CHART_AXIS_OFF )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;

   chart_axis_off(axis) ; 
}





/*
 * Set the category axis position.
 */
/*
void
chart_axis_set_position(lxw_chart_axis *axis, uint8_t position)
*/
HB_FUNC( CHART_AXIS_SET_POSITION )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t position = hb_parni( 2 ) ;

   chart_axis_set_position(axis, position) ; 
}





/*
 * Set the axis label position.
 */
/*
void
chart_axis_set_label_position(lxw_chart_axis *axis, uint8_t position)
*/
HB_FUNC( CHART_AXIS_SET_LABEL_POSITION )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t position = hb_parni( 2 ) ;

   chart_axis_set_label_position(axis, position) ; 
}





/*
 * Set the minimum value for an axis.
 */
/*
void
chart_axis_set_min(lxw_chart_axis *axis, double min)
*/
HB_FUNC( CHART_AXIS_SET_MIN )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   double min = hb_parnd( 2 ) ;

   chart_axis_set_min(axis, min) ; 
}





/*
 * Set the maximum value for an axis.
 */
/*
void
chart_axis_set_max(lxw_chart_axis *axis, double max)
*/
HB_FUNC( CHART_AXIS_SET_MAX )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   double max = hb_parnd( 2 ) ;

   chart_axis_set_max(axis, max) ; 
}





/*
 * Set the log base for an axis.
 */
/*
void
chart_axis_set_log_base(lxw_chart_axis *axis, uint16_t log_base)
*/
HB_FUNC( CHART_AXIS_SET_LOG_BASE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint16_t log_base = hb_parnl( 2 ) ;

   chart_axis_set_log_base(axis, log_base) ; 
}





/*
 * Set the major mark for an axis.
 */
/*
void
chart_axis_set_major_tick_mark(lxw_chart_axis *axis, uint8_t type)
*/
HB_FUNC( CHART_AXIS_SET_MAJOR_TICK_MARK )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   chart_axis_set_major_tick_mark(axis, type) ; 
}





/*
 * Set the minor mark for an axis.
 */
/*
void
chart_axis_set_minor_tick_mark(lxw_chart_axis *axis, uint8_t type)
*/
HB_FUNC( CHART_AXIS_SET_MINOR_TICK_MARK )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t type = hb_parni( 2 ) ;

   chart_axis_set_minor_tick_mark(axis, type) ; 
}





/*
 * Set interval unit for a category axis.
 */
/*
void
chart_axis_set_interval_unit(lxw_chart_axis *axis, uint16_t unit)
*/
HB_FUNC( CHART_AXIS_SET_INTERVAL_UNIT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint16_t unit = hb_parnl( 2 ) ;

   chart_axis_set_interval_unit(axis, unit) ; 
}





/*
 * Set tick interval for a category axis.
 */
/*
void
chart_axis_set_interval_tick(lxw_chart_axis *axis, uint16_t unit)
*/
HB_FUNC( CHART_AXIS_SET_INTERVAL_TICK )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint16_t unit = hb_parnl( 2 ) ;

   chart_axis_set_interval_tick(axis, unit) ; 
}





/*
 * Set major unit for a value axis.
 */
/*
void
chart_axis_set_major_unit(lxw_chart_axis *axis, double unit)
*/
HB_FUNC( CHART_AXIS_SET_MAJOR_UNIT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   double unit = hb_parnd( 2 ) ;

   chart_axis_set_major_unit(axis, unit) ; 
}





/*
 * Set minor unit for a value axis.
 */
/*
void
chart_axis_set_minor_unit(lxw_chart_axis *axis, double unit)
*/
HB_FUNC( CHART_AXIS_SET_MINOR_UNIT )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   double unit = hb_parnd( 2 ) ;

   chart_axis_set_minor_unit(axis, unit) ; 
}





/*
 * Set the display units for a value axis.
 */
/*
void
chart_axis_set_display_units(lxw_chart_axis *axis, uint8_t units)
*/
HB_FUNC( CHART_AXIS_SET_DISPLAY_UNITS )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t units = hb_parni( 2 ) ;

   chart_axis_set_display_units(axis, units) ; 
}





/*
 * Turn on/off the display units for a value axis.
 */
/*
void
chart_axis_set_display_units_visible(lxw_chart_axis *axis, uint8_t visible)
*/
HB_FUNC( CHART_AXIS_SET_DISPLAY_UNITS_VISIBLE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t visible = hb_parni( 2 ) ;

   chart_axis_set_display_units_visible(axis, visible) ; 
}





/*
 * Set the axis major gridlines on/off.
 */
/*
void
chart_axis_major_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible)
*/
HB_FUNC( CHART_AXIS_MAJOR_GRIDLINES_SET_VISIBLE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t visible = hb_parni( 2 ) ;

   chart_axis_major_gridlines_set_visible(axis, visible) ; 
}





/*
 * Set a line type for the major gridlines.
 */
/*
void
chart_axis_major_gridlines_set_line(lxw_chart_axis *axis,
   lxw_chart_line *line)
*/
HB_FUNC( CHART_AXIS_MAJOR_GRIDLINES_SET_LINE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_axis_major_gridlines_set_line(axis, line) ; 
}





/*
 * Set the axis minor gridlines on/off.
 */
/*
void
chart_axis_minor_gridlines_set_visible(lxw_chart_axis *axis, uint8_t visible)
*/
HB_FUNC( CHART_AXIS_MINOR_GRIDLINES_SET_VISIBLE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t visible = hb_parni( 2 ) ;

   chart_axis_minor_gridlines_set_visible(axis, visible) ; 
}





/*
 * Set a line type for the minor gridlines.
 */
/*
void
chart_axis_minor_gridlines_set_line(lxw_chart_axis *axis,
   lxw_chart_line *line)
*/
HB_FUNC( CHART_AXIS_MINOR_GRIDLINES_SET_LINE )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_axis_minor_gridlines_set_line(axis, line) ; 
}





/*
 * Set the chart axis label alignment.
 */
/*
void
chart_axis_set_label_align(lxw_chart_axis *axis, uint8_t align)
*/
HB_FUNC( CHART_AXIS_SET_LABEL_ALIGN )
{ 
   lxw_chart_axis *axis = hb_parptr( 1 ) ;
   uint8_t align = hb_parni( 2 ) ;

   chart_axis_set_label_align(axis, align) ; 
}





/*
 * Set the chart title.
 */
/*
void
chart_title_set_name(lxw_chart *self, const char *name)
*/
HB_FUNC( CHART_TITLE_SET_NAME )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   const char *name = hb_parcx( 2 ) ;

   chart_title_set_name(self, name) ; 
}





/*
 * Set the chart title, with a range instead or a formula.
 */
/*
void
chart_title_set_name_range(lxw_chart *self, const char *sheetname,
   lxw_row_t row, lxw_col_t col)
*/
HB_FUNC( CHART_TITLE_SET_NAME_RANGE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   const char *sheetname = hb_parcx( 2 ) ;
   lxw_row_t row = hb_parni( 3 ) ;
   lxw_col_t col = hb_parni( 4 ) ;

   chart_title_set_name_range(self, sheetname, row, col) ; 
}





/*
 * Set the chart title font.
 */
/*
void
chart_title_set_name_font(lxw_chart *self, lxw_chart_font *font)
*/
HB_FUNC( CHART_TITLE_SET_NAME_FONT )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;
   chart_title_set_name_font(self, font) ; 
}





/*
 * Turn off the chart title.
 */
/*
void
chart_title_off(lxw_chart *self)
*/
HB_FUNC( CHART_TITLE_OFF )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;

   chart_title_off(self) ; 
}





/*
 * Set the chart legend position.
 */
/*
void
chart_legend_set_position(lxw_chart *self, uint8_t position)
*/
HB_FUNC( CHART_LEGEND_SET_POSITION )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint8_t position = hb_parni( 2 ) ;

   chart_legend_set_position(self, position) ; 
}





/*
 * Set the legend font.
 */
/*
void
chart_legend_set_font(lxw_chart *self, lxw_chart_font *font)
*/
HB_FUNC( CHART_LEGEND_SET_FONT )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;

   chart_legend_set_font(self, font) ; 
}





/*
 * Remove one or more series from the the legend.
 */
/*
lxw_error
chart_legend_delete_series(lxw_chart *self, int16_t delete_series[])
*/
/*
HB_FUNC( CHART_LEGEND_DELETE_SERIES )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   int16_t delete_series[] = hb_parnl( 2 ) ;

   hb_retni( chart_legend_delete_series(self, delete_series[]) ); 
}
*/




/*
 * Set a line type for the chartarea.
 */
/*
void
chart_chartarea_set_line(lxw_chart *self, lxw_chart_line *line)
*/
HB_FUNC( CHART_CHARTAREA_SET_LINE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_chartarea_set_line(self, line) ; 
}





/*
 * Set a fill type for the chartarea.
 */
/*
void
chart_chartarea_set_fill(lxw_chart *self, lxw_chart_fill *fill)
*/
HB_FUNC( CHART_CHARTAREA_SET_FILL )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_fill *fill = hb_parptr( 2 ) ;

   chart_chartarea_set_fill(self, fill) ; 
}





/*
 * Set a pattern type for the chartarea.
 */
/*
void
chart_chartarea_set_pattern(lxw_chart *self, lxw_chart_pattern *pattern)
*/
HB_FUNC( CHART_CHARTAREA_SET_PATTERN )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_pattern *pattern = hb_parptr( 2 ) ;

   chart_chartarea_set_pattern(self, pattern) ; 
}





/*
 * Set a line type for the plotarea.
 */
/*
void
chart_plotarea_set_line(lxw_chart *self, lxw_chart_line *line)
*/
HB_FUNC( CHART_PLOTAREA_SET_LINE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_plotarea_set_line(self, line) ; 
}





/*
 * Set a fill type for the plotarea.
 */
/*
void
chart_plotarea_set_fill(lxw_chart *self, lxw_chart_fill *fill)
*/
HB_FUNC( CHART_PLOTAREA_SET_FILL )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_fill *fill = hb_parptr( 2 ) ;

   chart_plotarea_set_fill(self, fill) ; 
}





/*
 * Set a pattern type for the plotarea.
 */
/*
void
chart_plotarea_set_pattern(lxw_chart *self, lxw_chart_pattern *pattern)
*/
HB_FUNC( CHART_PLOTAREA_SET_PATTERN )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_pattern *pattern = hb_parptr( 2 ) ;

   chart_plotarea_set_pattern(self, pattern) ; 
}





/*
 * Turn on the chart data table.
 */
/*
void
chart_set_table(lxw_chart *self)
*/
HB_FUNC( CHART_SET_TABLE )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;

   chart_set_table(self) ; 
}





/*
 * Set the options for the chart data table grid.
 */
/*
void
chart_set_table_grid(lxw_chart *self, uint8_t horizontal, uint8_t vertical,
   uint8_t outline, uint8_t legend_keys)
*/
HB_FUNC( CHART_SET_TABLE_GRID )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint8_t horizontal = hb_parni( 2 ) ;
   uint8_t vertical = hb_parni( 3 ) ;
   uint8_t outline = hb_parni( 4 ) ;
   uint8_t legend_keys = hb_parni( 5 ) ;

   chart_set_table_grid(self, horizontal, vertical, outline, legend_keys) ; 
}





/*
 * Set the font for the chart data table grid.
 */
/*
void
chart_set_table_font(lxw_chart *self, lxw_chart_font *font)
*/
HB_FUNC( CHART_SET_TABLE_FONT )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_font *font = hb_parptr( 2 ) ;

   chart_set_table_font(self, font) ; 
}





/*
 * Turn on up-down bars for the chart.
 */
/*
void
chart_set_up_down_bars(lxw_chart *self)
*/
HB_FUNC( CHART_SET_UP_DOWN_BARS )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;

   chart_set_up_down_bars(self) ; 
}





/*
 * Turn on up-down bars for the chart, with formatting.
 */
/*
void
chart_set_up_down_bars_format(lxw_chart *self, lxw_chart_line *up_bar_line,
    lxw_chart_fill *up_bar_fill,
    lxw_chart_line *down_bar_line,
    lxw_chart_fill *down_bar_fill)
*/
HB_FUNC( CHART_SET_UP_DOWN_BARS_FORMAT )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_line *up_bar_line = hb_parptr( 2 ) ;
   lxw_chart_fill *up_bar_fill = hb_parptr( 3 ) ;
   lxw_chart_line *down_bar_line = hb_parptr( 4 ) ;
   lxw_chart_fill *down_bar_fill = hb_parptr( 5 ) ;

   chart_set_up_down_bars_format(self, up_bar_line, up_bar_fill, down_bar_line, down_bar_fill) ; 
}





/*
 * Turn on drop lines for the chart.
 */
/*
void
chart_set_drop_lines(lxw_chart *self, lxw_chart_line *line)
*/
HB_FUNC( CHART_SET_DROP_LINES )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_set_drop_lines(self, line) ; 
}





/*
 * Turn on high_low lines for the chart.
 */
/*
void
chart_set_high_low_lines(lxw_chart *self, lxw_chart_line *line)
*/
HB_FUNC( CHART_SET_HIGH_LOW_LINES )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   lxw_chart_line *line = hb_parptr( 2 ) ;

   chart_set_high_low_lines(self, line) ; 
}





/*
 * Set the Bar/Column overlap for all data series.
 */
/*
void
chart_set_series_overlap(lxw_chart *self, int8_t overlap)
*/
HB_FUNC( CHART_SET_SERIES_OVERLAP )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   int8_t overlap = hb_parni( 2 ) ;

   chart_set_series_overlap(self, overlap) ; 
}





/*
 * Set the option for displaying blank data in a chart.
 */
/*
void
chart_show_blanks_as(lxw_chart *self, uint8_t option)
*/
HB_FUNC( CHART_SHOW_BLANKS_AS )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint8_t option = hb_parni( 2 ) ;

   chart_show_blanks_as(self, option) ; 
}





/*
 * Display data on charts from hidden rows or columns.
 */
/*
void
chart_show_hidden_data(lxw_chart *self)
*/
HB_FUNC( CHART_SHOW_HIDDEN_DATA )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;

   chart_show_hidden_data(self) ; 
}





/*
 * Set the Bar/Column gap for all data series.
 */
/*
void
chart_set_series_gap(lxw_chart *self, uint16_t gap)
*/
HB_FUNC( CHART_SET_SERIES_GAP )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint16_t gap = hb_parnl( 2 ) ;

   chart_set_series_gap(self, gap) ; 
}





/*
 * Set the Pie/Doughnut chart rotation: the angle of the first slice.
 */
/*
void
chart_set_rotation(lxw_chart *self, uint16_t rotation)
*/
HB_FUNC( CHART_SET_ROTATION )
{ 
   lxw_chart *self = hb_parptr( 1 ) ;
   uint16_t rotation = hb_parnl( 2 ) ;

   chart_set_rotation(self, rotation) ; 
}


//eof
