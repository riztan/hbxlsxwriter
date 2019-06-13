/**********************************************
 *  
 * hbxlsxwriter.ch - A library for creating Excel XLSX worksheet files for Harbour.
 *
 * Used in conjunction with the libxlsxwriter library.
 *
 * Copyright 2014-2019, John McNamara, jmcnamara@cpan.org. See LICENSE.txt.
 *
 * Adapted for Harbour by Riztan Gutierrez, riztan@gmail.com
 *
 */

//#xtranslate <dv>.<key> := <value>  =>  hb_lxw_dv( @<dv>, #<key>, <value> )




/**
 * @brief Convert an Excel `A1:B2` range into a `(first_row, first_col,
 *        last_row, last_col)` sequence.
 *
 * Convert an Excel `A1:B2` range into a `(first_row, first_col, last_row,
 * last_col)` sequence.
 *
 * This is a little syntactic shortcut to help with worksheet layout.
 *
 * @code
 *     worksheet_print_area(worksheet, 0, 0, 41, 10); // A1:K42.
 *
 *     // Same as:
 *     worksheet_print_area(worksheet, RANGE("A1:K42"));
 * @endcode
 */
#define RANGE(range) ;
    lxw_name_to_row(range), lxw_name_to_col(range), ;
    lxw_name_to_row_2(range), lxw_name_to_col_2(range) 


/**
 * @brief Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * Convert an Excel `A1` cell string into a `(row, col)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *      worksheet_write_string(worksheet, CELL("A1"), "Foo", NULL);
 *
 *      //Same as:
 *      worksheet_write_string(worksheet, 0, 0,       "Foo", NULL);
 * @endcode
 *
 * @note
 *
 * This macro shouldn't be used in performance critical situations since it
 * expands to two function calls.
 */
#define CELL(cell) ;
    lxw_name_to_row(cell), lxw_name_to_col(cell)




/**
 * @brief Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * Convert an Excel `A:B` column range into a `(col1, col2)` pair.
 *
 * This is a little syntactic shortcut to help with worksheet layout:
 *
 * @code
 *     worksheet_set_column(worksheet, COLS("B:D"), 20, NULL, NULL);
 *
 *     // Same as:
 *     worksheet_set_column(worksheet, 1, 3,        20, NULL, NULL);
 * @endcode
 *
 */
#define COLS(cols) ;
    lxw_name_to_col(cols), lxw_name_to_col_2(cols)



    /** No error. */
    #define LXW_NO_ERROR					0

    /** Memory error, failed to malloc() required memory. */
    #define LXW_ERROR_MEMORY_MALLOC_FAILED			1

    /** Error creating output xlsx file. Usually a permissions error. */
    #define LXW_ERROR_CREATING_XLSX_FILE			2

    /** Error encountered when creating a tmpfile during file assembly. */
    #define LXW_ERROR_CREATING_TMPFILE				3

    /** Error reading a tmpfile. */
    #define LXW_ERROR_READING_TMPFILE				4

    /** Zlib error with a file operation while creating xlsx file. */
    #define LXW_ERROR_ZIP_FILE_OPERATION			5

    /** Zlib error when adding sub file to xlsx file. */
    #define LXW_ERROR_ZIP_FILE_ADD				6

    /** Zlib error when closing xlsx file. */
    #define LXW_ERROR_ZIP_CLOSE					7

    /** NULL function parameter ignored. */
    #define LXW_ERROR_NULL_PARAMETER_IGNORED			8

    /** Function parameter validation error. */
    #define LXW_ERROR_PARAMETER_VALIDATION			9

    /** Worksheet name exceeds Excel's limit of 31 characters. */
    #define LXW_ERROR_SHEETNAME_LENGTH_EXCEEDED			10

    /** Worksheet name contains invalid Excel character: '[]:*?/\\' */
    #define LXW_ERROR_INVALID_SHEETNAME_CHARACTER		11

    /** Worksheet name is already in use. */
    #define LXW_ERROR_SHEETNAME_ALREADY_USED			12

    /** Parameter exceeds Excel's limit of 32 characters. */
    #define LXW_ERROR_32_STRING_LENGTH_EXCEEDED			13

    /** Parameter exceeds Excel's limit of 128 characters. */
    #define LXW_ERROR_128_STRING_LENGTH_EXCEEDED		14

    /** Parameter exceeds Excel's limit of 255 characters. */
    #define LXW_ERROR_255_STRING_LENGTH_EXCEEDED		15

    /** String exceeds Excel's limit of 32,767 characters. */
    #define LXW_ERROR_MAX_STRING_LENGTH_EXCEEDED		16

    /** Error finding internal string index. */
    #define LXW_ERROR_SHARED_STRING_INDEX_NOT_FOUND		17

    /** Worksheet row or column index out of range. */
    #define LXW_ERROR_WORKSHEET_INDEX_OUT_OF_RANGE		18

    /** Maximum number of worksheet URLs (65530) exceeded. */
    #define LXW_ERROR_WORKSHEET_MAX_NUMBER_URLS_EXCEEDED	19

    /** Couldn't read image dimensions or DPI. */
    #define LXW_ERROR_IMAGE_DIMENSIONS				20

    #define LXW_MAX_ERRNO


/****************************
 *  CHART
 ****************************/
/**
 * @brief Available chart types.
 */

    /** None. */
    #define LXW_CHART_NONE 				0

    /** Area chart. */
    #define LXW_CHART_AREA				1

    /** Area chart - stacked. */
    #define LXW_CHART_AREA_STACKED			2

    /** Area chart - percentage stacked. */
    #define LXW_CHART_AREA_STACKED_PERCENT		3

    /** Bar chart. */
    #define LXW_CHART_BAR				4

    /** Bar chart - stacked. */
    #define LXW_CHART_BAR_STACKED			5

    /** Bar chart - percentage stacked. */
    #define LXW_CHART_BAR_STACKED_PERCENT		6

    /** Column chart. */
    #define LXW_CHART_COLUMN				7

    /** Column chart - stacked. */
    #define LXW_CHART_COLUMN_STACKED			8

    /** Column chart - percentage stacked. */
    #define LXW_CHART_COLUMN_STACKED_PERCENT		9

    /** Doughnut chart. */
    #define LXW_CHART_DOUGHNUT				10

    /** Line chart. */
    #define LXW_CHART_LINE				11

    /** Pie chart. */
    #define LXW_CHART_PIE				12

    /** Scatter chart. */
    #define LXW_CHART_SCATTER				13

    /** Scatter chart - straight. */
    #define LXW_CHART_SCATTER_STRAIGHT			14

    /** Scatter chart - straight with markers. */
    #define LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS	15

    /** Scatter chart - smooth. */
    #define LXW_CHART_SCATTER_SMOOTH			16

    /** Scatter chart - smooth with markers. */
    #define LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS	17

    /** Radar chart. */
    #define LXW_CHART_RADAR				18

    /** Radar chart - with markers. */
    #define LXW_CHART_RADAR_WITH_MARKERS		19

    /** Radar chart - filled. */
    #define LXW_CHART_RADAR_FILLED 			20



/**
 * @brief Chart legend positions.
 */
    /** No chart legend. */
    #define LXW_CHART_LEGEND_NONE 			0

    /** Chart legend positioned at right side. */
    #define LXW_CHART_LEGEND_RIGHT			1

    /** Chart legend positioned at left side. */
    #define LXW_CHART_LEGEND_LEFT			2

    /** Chart legend positioned at top. */
    #define LXW_CHART_LEGEND_TOP			3

    /** Chart legend positioned at bottom. */
    #define LXW_CHART_LEGEND_BOTTOM			4

    /** Chart legend positioned at top right. */
    #define LXW_CHART_LEGEND_TOP_RIGHT			5

    /** Chart legend overlaid at right side. */
    #define LXW_CHART_LEGEND_OVERLAY_RIGHT		6

    /** Chart legend overlaid at left side. */
    #define LXW_CHART_LEGEND_OVERLAY_LEFT		7

    /** Chart legend overlaid at top right. */
    #define LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT		8



/**
 * @brief Chart line dash types.
 *
 * The dash types are shown in the order that they appear in the Excel dialog.
 * See @ref chart_lines.
 */
    /** Solid. */
    #define LXW_CHART_LINE_DASH_SOLID 				0

    /** Round Dot. */
    #define LXW_CHART_LINE_DASH_ROUND_DOT			1

    /** Square Dot. */
    #define LXW_CHART_LINE_DASH_SQUARE_DOT			2

    /** Dash. */
    #define LXW_CHART_LINE_DASH_DASH				3

    /** Dash Dot. */
    #define LXW_CHART_LINE_DASH_DASH_DOT			4

    /** Long Dash. */
    #define LXW_CHART_LINE_DASH_LONG_DASH			5

    /** Long Dash Dot. */
    #define LXW_CHART_LINE_DASH_LONG_DASH_DOT			6

    /** Long Dash Dot Dot. */
    #define LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT		7

    /* These aren't available in the dialog but are used by Excel. */
    #define LXW_CHART_LINE_DASH_DOT				8
    #define LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT			9
    #define LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT		10





/**
 * @brief Chart marker types.
 */
    /** Automatic, series default, marker type. */
    #define LXW_CHART_MARKER_AUTOMATIC				0

    /** No marker type. */
    #define LXW_CHART_MARKER_NONE				1

    /** Square marker type. */
    #define LXW_CHART_MARKER_SQUARE				2

    /** Diamond marker type. */
    #define LXW_CHART_MARKER_DIAMOND				3

    /** Triangle marker type. */
    #define LXW_CHART_MARKER_TRIANGLE				4

    /** X shape marker type. */
    #define LXW_CHART_MARKER_X					5

    /** Star marker type. */
    #define LXW_CHART_MARKER_STAR				6

    /** Short dash marker type. */
    #define LXW_CHART_MARKER_SHORT_DASH				7

    /** Long dash marker type. */
    #define LXW_CHART_MARKER_LONG_DASH				8

    /** Circle marker type. */
    #define LXW_CHART_MARKER_CIRCLE				9

    /** Plus (+) marker type. */
    #define LXW_CHART_MARKER_PLUS				10




/**
 * @brief Chart pattern types.
 */
    /** None pattern. */
    #define LXW_CHART_PATTERN_NONE				0

    /** 5 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_5				1

    /** 10 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_10			2

    /** 20 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_20			3

    /** 25 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_25			4

    /** 30 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_30			5

    /** 40 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_40			6

    /** 50 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_50			7

    /** 60 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_60			8

    /** 70 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_70			9

    /** 75 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_75			10

    /** 80 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_80			11

    /** 90 Percent pattern. */
    #define LXW_CHART_PATTERN_PERCENT_90			12

    /** Light downward diagonal pattern. */
    #define LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL		13

    /** Light upward diagonal pattern. */
    #define LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL		14

    /** Dark downward diagonal pattern. */
    #define LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL		15

    /** Dark upward diagonal pattern. */
    #define LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL		16

    /** Wide downward diagonal pattern. */
    #define LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL		17

    /** Wide upward diagonal pattern. */
    #define LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL		18

    /** Light vertical pattern. */
    #define LXW_CHART_PATTERN_LIGHT_VERTICAL			19

    /** Light horizontal pattern. */
    #define LXW_CHART_PATTERN_LIGHT_HORIZONTAL			20

    /** Narrow vertical pattern. */
    #define LXW_CHART_PATTERN_NARROW_VERTICAL			21

    /** Narrow horizontal pattern. */
    #define LXW_CHART_PATTERN_NARROW_HORIZONTAL			22

    /** Dark vertical pattern. */
    #define LXW_CHART_PATTERN_DARK_VERTICAL			23

    /** Dark horizontal pattern. */
    #define LXW_CHART_PATTERN_DARK_HORIZONTAL			24

    /** Dashed downward diagonal pattern. */
    #define LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL		25

    /** Dashed upward diagonal pattern. */
    #define LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL		26

    /** Dashed horizontal pattern. */
    #define LXW_CHART_PATTERN_DASHED_HORIZONTAL			27

    /** Dashed vertical pattern. */
    #define LXW_CHART_PATTERN_DASHED_VERTICAL			28

    /** Small confetti pattern. */
    #define LXW_CHART_PATTERN_SMALL_CONFETTI			29

    /** Large confetti pattern. */
    #define LXW_CHART_PATTERN_LARGE_CONFETTI			30

    /** Zigzag pattern. */
    #define LXW_CHART_PATTERN_ZIGZAG				31

    /** Wave pattern. */
    #define LXW_CHART_PATTERN_WAVE				32

    /** Diagonal brick pattern. */
    #define LXW_CHART_PATTERN_DIAGONAL_BRICK			33

    /** Horizontal brick pattern. */
    #define LXW_CHART_PATTERN_HORIZONTAL_BRICK			34

    /** Weave pattern. */
    #define LXW_CHART_PATTERN_WEAVE				35

    /** Plaid pattern. */
    #define LXW_CHART_PATTERN_PLAID				36

    /** Divot pattern. */
    #define LXW_CHART_PATTERN_DIVOT				37

    /** Dotted grid pattern. */
    #define LXW_CHART_PATTERN_DOTTED_GRID			38

    /** Dotted diamond pattern. */
    #define LXW_CHART_PATTERN_DOTTED_DIAMOND			39

    /** Shingle pattern. */
    #define LXW_CHART_PATTERN_SHINGLE				40

    /** Trellis pattern. */
    #define LXW_CHART_PATTERN_TRELLIS				41

    /** Sphere pattern. */
    #define LXW_CHART_PATTERN_SPHERE				42

    /** Small grid pattern. */
    #define LXW_CHART_PATTERN_SMALL_GRID			43

    /** Large grid pattern. */
    #define LXW_CHART_PATTERN_LARGE_GRID			44

    /** Small check pattern. */
    #define LXW_CHART_PATTERN_SMALL_CHECK			45

    /** Large check pattern. */
    #define LXW_CHART_PATTERN_LARGE_CHECK			46

    /** Outlined diamond pattern. */
    #define LXW_CHART_PATTERN_OUTLINED_DIAMOND			47

    /** Solid diamond pattern. */
    #define LXW_CHART_PATTERN_SOLID_DIAMOND 			48



/**
 * @brief Chart data label positions.
 */
    /** Series data label position: default position. */
    #define LXW_CHART_LABEL_POSITION_DEFAULT			0

    /** Series data label position: center. */
    #define LXW_CHART_LABEL_POSITION_CENTER			1

    /** Series data label position: right. */
    #define LXW_CHART_LABEL_POSITION_RIGHT			2

    /** Series data label position: left. */
    #define LXW_CHART_LABEL_POSITION_LEFT			3

    /** Series data label position: above. */
    #define LXW_CHART_LABEL_POSITION_ABOVE			4

    /** Series data label position: below. */
    #define LXW_CHART_LABEL_POSITION_BELOW			5

    /** Series data label position: inside base.  */
    #define LXW_CHART_LABEL_POSITION_INSIDE_BASE		6

    /** Series data label position: inside end. */
    #define LXW_CHART_LABEL_POSITION_INSIDE_END			7

    /** Series data label position: outside end. */
    #define LXW_CHART_LABEL_POSITION_OUTSIDE_END		8

    /** Series data label position: best fit. */
    #define LXW_CHART_LABEL_POSITION_BEST_FIT			9




/**
 * @brief Chart data label separator.
 */
    /** Series data label separator: comma (the default). */
    #define LXW_CHART_LABEL_SEPARATOR_COMMA			0

    /** Series data label separator: semicolon. */
    #define LXW_CHART_LABEL_SEPARATOR_SEMICOLON			1

    /** Series data label separator: period. */
    #define LXW_CHART_LABEL_SEPARATOR_PERIOD			2

    /** Series data label separator: newline. */
    #define LXW_CHART_LABEL_SEPARATOR_NEWLINE			3

    /** Series data label separator: space. */
    #define LXW_CHART_LABEL_SEPARATOR_SPACE			4



/**
 * @brief Chart axis types.
 */
    /** Chart X axis. */
    #define LXW_CHART_AXIS_TYPE_X			0

    /** Chart Y axis. */
    #define LXW_CHART_AXIS_TYPE_Y			1



    #define LXW_CHART_SUBTYPE_NONE 			0
    #define LXW_CHART_SUBTYPE_STACKED			1
    #define LXW_CHART_SUBTYPE_STACKED_PERCENT		2

    #define LXW_GROUPING_CLUSTERED			0
    #define LXW_GROUPING_STANDARD			1
    #define LXW_GROUPING_PERCENTSTACKED			2
    #define LXW_GROUPING_STACKED			3



/**
 * @brief Axis positions for category axes.
 */
    #define LXW_CHART_AXIS_POSITION_DEFAULT			0

    /** Position category axis on tick marks. */
    #define LXW_CHART_AXIS_POSITION_ON_TICK			1

    /** Position category axis between tick marks. */
    #define LXW_CHART_AXIS_POSITION_BETWEEN			2



/**
 * @brief Axis label positions.
 */
    /** Position the axis labels next to the axis. The default. */
    #define LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO		0

    /** Position the axis labels at the top of the chart, for horizontal
     * axes, or to the right for vertical axes.*/
    #define LXW_CHART_AXIS_LABEL_POSITION_HIGH			1

    /** Position the axis labels at the bottom of the chart, for horizontal
     * axes, or to the left for vertical axes.*/
    #define LXW_CHART_AXIS_LABEL_POSITION_LOW			2

    /** Turn off the the axis labels. */
    #define LXW_CHART_AXIS_LABEL_POSITION_NONE			3



/**
 * @brief Axis label alignments.
 */
    /** Chart axis label alignment: center. */
    #define LXW_CHART_AXIS_LABEL_ALIGN_CENTER			0

    /** Chart axis label alignment: left. */
    #define LXW_CHART_AXIS_LABEL_ALIGN_LEFT			1

    /** Chart axis label alignment: right. */
    #define LXW_CHART_AXIS_LABEL_ALIGN_RIGHT			2



/**
 * @brief Display units for chart value axis.
 */
    /** Axis display units: None. The default. */
    #define LXW_CHART_AXIS_UNITS_NONE				0

    /** Axis display units: Hundreds. */
    #define LXW_CHART_AXIS_UNITS_HUNDREDS			1

    /** Axis display units: Thousands. */
    #define LXW_CHART_AXIS_UNITS_THOUSANDS			2

    /** Axis display units: Ten thousands. */
    #define LXW_CHART_AXIS_UNITS_TEN_THOUSANDS			3

    /** Axis display units: Hundred thousands. */
    #define LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS		4

    /** Axis display units: Millions. */
    #define LXW_CHART_AXIS_UNITS_MILLIONS			5

    /** Axis display units: Ten millions. */
    #define LXW_CHART_AXIS_UNITS_TEN_MILLIONS			6

    /** Axis display units: Hundred millions. */
    #define LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS		7

    /** Axis display units: Billions. */
    #define LXW_CHART_AXIS_UNITS_BILLIONS			8

    /** Axis display units: Trillions. */
    #define LXW_CHART_AXIS_UNITS_TRILLIONS			9


/**
 * @brief Tick mark types for an axis.
 */
    /** Default tick mark for the chart axis. Usually outside. */
    #define LXW_CHART_AXIS_TICK_MARK_DEFAULT			0

    /** No tick mark for the axis. */
    #define LXW_CHART_AXIS_TICK_MARK_NONE			1

    /** Tick mark inside the axis only. */
    #define LXW_CHART_AXIS_TICK_MARK_INSIDE			2

    /** Tick mark outside the axis only. */
    #define LXW_CHART_AXIS_TICK_MARK_OUTSIDE			3

    /** Tick mark inside and outside the axis. */
    #define LXW_CHART_AXIS_TICK_MARK_CROSSING			4



/**
 * @brief Define how blank values are displayed in a chart.
 */
    /** Show empty chart cells as gaps in the data. The default. */
    #define LXW_CHART_BLANKS_AS_GAP			0

    /** Show empty chart cells as zeros. */
    #define LXW_CHART_BLANKS_AS_ZERO			1

    /** Show empty chart cells as connected. Only for charts with lines. */
    #define LXW_CHART_BLANKS_AS_CONNECTED		2

//enum lxw_chart_position {
    #define LXW_CHART_AXIS_RIGHT			0
    #define LXW_CHART_AXIS_LEFT				1
    #define LXW_CHART_AXIS_TOP				2
    #define LXW_CHART_AXIS_BOTTOM			3

/**
 * @brief Type/amount of data series error bar.
 */
    /** Error bar type: Standard error. */
    #define LXW_CHART_ERROR_BAR_TYPE_STD_ERROR			0

    /** Error bar type: Fixed value. */
    #define LXW_CHART_ERROR_BAR_TYPE_FIXED			1

    /** Error bar type: Percentage. */
    #define LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE			2

    /** Error bar type: Standard deviation(s). */
    #define LXW_CHART_ERROR_BAR_TYPE_STD_DEV			3

/**
 * @brief Direction for a data series error bar.
 */
    /** Error bar extends in both directions. The default. */
    #define LXW_CHART_ERROR_BAR_DIR_BOTH			0

    /** Error bar extends in positive direction. */
    #define LXW_CHART_ERROR_BAR_DIR_PLUS			1

    /** Error bar extends in negative direction. */
    #define LXW_CHART_ERROR_BAR_DIR_MINUS			2

/**
 * @brief Direction for a data series error bar.
 */
    /** X axis error bar. */
    #define LXW_CHART_ERROR_BAR_AXIS_X			0

    /** Y axis error bar. */
    #define LXW_CHART_ERROR_BAR_AXIS_Y			1

/**
 * @brief End cap styles for a data series error bar.
 */
    /** Flat end cap. The default. */
    #define LXW_CHART_ERROR_BAR_END_CAP			0

    /** No end cap. */
    #define LXW_CHART_ERROR_BAR_NO_CAP			1

/**
 * @brief Series trendline/regression types.
 */
    /** Trendline type: Linear. */
    #define LXW_CHART_TRENDLINE_TYPE_LINEAR			0

    /** Trendline type: Logarithm. */
    #define LXW_CHART_TRENDLINE_TYPE_LOG			1

    /** Trendline type: Polynomial. */
    #define LXW_CHART_TRENDLINE_TYPE_POLY			2

    /** Trendline type: Power. */
    #define LXW_CHART_TRENDLINE_TYPE_POWER			3

    /** Trendline type: Exponential. */
    #define LXW_CHART_TRENDLINE_TYPE_EXP			4

    /** Trendline type: Moving Average. */
    #define LXW_CHART_TRENDLINE_TYPE_AVERAGE			5






/***************************************************
 *    UTILITY
 */

#xtranslate LXW_CELL(<cell>) => ;
    lxw_name_to_row(<cell>), lxw_name_to_col(<cell>)





/***************************************************
 *    FORMAT
 */

/** Format underline values for format_set_underline(). */
    /** Single underline */
    #define LXW_UNDERLINE_SINGLE 				1

    /** Double underline */
    #define LXW_UNDERLINE_DOUBLE				2

    /** Single accounting underline */
    #define LXW_UNDERLINE_SINGLE_ACCOUNTING			3

    /** Double accounting underline */
    #define LXW_UNDERLINE_DOUBLE_ACCOUNTING			4


/** Superscript and subscript values for format_set_font_script(). */
    /** Superscript font */
    #define LXW_FONT_SUPERSCRIPT 				1

    /** Subscript font */
    #define LXW_FONT_SUBSCRIPT					2	


/** Alignment values for format_set_align(). */
    /** No alignment. Cell will use Excel's default for the data type */
    #define LXW_ALIGN_NONE  					0

    /** Left horizontal alignment */
    #define LXW_ALIGN_LEFT					1

    /** Center horizontal alignment */
    #define LXW_ALIGN_CENTER					2

    /** Right horizontal alignment */
    #define LXW_ALIGN_RIGHT					3

    /** Cell fill horizontal alignment */
    #define LXW_ALIGN_FILL					4

    /** Justify horizontal alignment */
    #define LXW_ALIGN_JUSTIFY					5

    /** Center Across horizontal alignment */
    #define LXW_ALIGN_CENTER_ACROSS				6

    /** Left horizontal alignment */
    #define LXW_ALIGN_DISTRIBUTED				7

    /** Top vertical alignment */
    #define LXW_ALIGN_VERTICAL_TOP				8

    /** Bottom vertical alignment */
    #define LXW_ALIGN_VERTICAL_BOTTOM				9

    /** Center vertical alignment */
    #define LXW_ALIGN_VERTICAL_CENTER				10

    /** Justify vertical alignment */
    #define LXW_ALIGN_VERTICAL_JUSTIFY				12

    /** Distributed vertical alignment */
    #define LXW_ALIGN_VERTICAL_DISTRIBUTED			13


//enum lxw_format_diagonal_types {
    #define LXW_DIAGONAL_BORDER_UP  				1
    #define LXW_DIAGONAL_BORDER_DOWN				2
    #define LXW_DIAGONAL_BORDER_UP_DOWN				3


/** Pattern value for use with format_set_pattern(). */
    /** Empty pattern */
    #define LXW_PATTERN_NONE 					0

    /** Solid pattern */
    #define LXW_PATTERN_SOLID					1

    /** Medium gray pattern */
    #define LXW_PATTERN_MEDIUM_GRAY				2

    /** Dark gray pattern */
    #define LXW_PATTERN_DARK_GRAY				3

    /** Light gray pattern */
    #define LXW_PATTERN_LIGHT_GRAY				4

    /** Dark horizontal line pattern */
    #define LXW_PATTERN_DARK_HORIZONTAL				5

    /** Dark vertical line pattern */
    #define LXW_PATTERN_DARK_VERTICAL				6

    /** Dark diagonal stripe pattern */
    #define LXW_PATTERN_DARK_DOWN				7

    /** Reverse dark diagonal stripe pattern */
    #define LXW_PATTERN_DARK_UP					8

    /** Dark grid pattern */
    #define LXW_PATTERN_DARK_GRID				9

    /** Dark trellis pattern */
    #define LXW_PATTERN_DARK_TRELLIS				10

    /** Light horizontal Line pattern */
    #define LXW_PATTERN_LIGHT_HORIZONTAL			11

    /** Light vertical line pattern */
    #define LXW_PATTERN_LIGHT_VERTICAL				12

    /** Light diagonal stripe pattern */
    #define LXW_PATTERN_LIGHT_DOWN				13

    /** Reverse light diagonal stripe pattern */
    #define LXW_PATTERN_LIGHT_UP				14

    /** Light grid pattern */
    #define LXW_PATTERN_LIGHT_GRID				15

    /** Light trellis pattern */
    #define LXW_PATTERN_LIGHT_TRELLIS				16

    /** 12.5% gray pattern */
    #define LXW_PATTERN_GRAY_125				17

    /** 6.25% gray pattern */
    #define LXW_PATTERN_GRAY_0625				18


/** Predefined values for common colors. */
    /** Black */
    #define LXW_COLOR_BLACK      0x1000000

    /** Blue */
    #define LXW_COLOR_BLUE       0x0000FF

    /** Brown */
    #define LXW_COLOR_BROWN      0x800000

    /** Cyan */
    #define LXW_COLOR_CYAN       0x00FFFF

    /** Gray */
    #define LXW_COLOR_GRAY       0x808080

    /** Green */
    #define LXW_COLOR_GREEN      0x008000

    /** Lime */
    #define LXW_COLOR_LIME       0x00FF00

    /** Magenta */
    #define LXW_COLOR_MAGENTA    0xFF00FF

    /** Navy */
    #define LXW_COLOR_NAVY       0x000080

    /** Orange */
    #define LXW_COLOR_ORANGE     0xFF6600

    /** Pink */
    #define LXW_COLOR_PINK       0xFF00FF

    /** Purple */
    #define LXW_COLOR_PURPLE     0x800080

    /** Red */
    #define LXW_COLOR_RED        0xFF0000

    /** Silver */
    #define LXW_COLOR_SILVER     0xC0C0C0

    /** White */
    #define LXW_COLOR_WHITE      0xFFFFFF

    /** Yellow */
    #define LXW_COLOR_YELLOW     0xFFFF00


/** Cell border styles for use with format_set_border(). */
    /** No border */
    #define LXW_BORDER_NONE			0

    /** Thin border style */
    #define LXW_BORDER_THIN			1

    /** Medium border style */
    #define LXW_BORDER_MEDIUM			2

    /** Dashed border style */
    #define LXW_BORDER_DASHED			3

    /** Dotted border style */
    #define LXW_BORDER_DOTTED			4

    /** Thick border style */
    #define LXW_BORDER_THICK			5

    /** Double border style */
    #define LXW_BORDER_DOUBLE			6

    /** Hair border style */
    #define LXW_BORDER_HAIR			7

    /** Medium dashed border style */
    #define LXW_BORDER_MEDIUM_DASHED		8

    /** Dash-dot border style */
    #define LXW_BORDER_DASH_DOT			9

    /** Medium dash-dot border style */
    #define LXW_BORDER_MEDIUM_DASH_DOT		10

    /** Dash-dot-dot border style */
    #define LXW_BORDER_DASH_DOT_DOT		11

    /** Medium dash-dot-dot border style */
    #define LXW_BORDER_MEDIUM_DASH_DOT_DOT	12

    /** Slant dash-dot border style */
    #define LXW_BORDER_SLANT_DASH_DOT		13



/********************************************************
 * 
 *  Worksheet
 *
 ********************************************************/
#define LXW_ROW_MAX           1048576
#define LXW_COL_MAX           16384
#define LXW_COL_META_MAX      128
#define LXW_HEADER_FOOTER_MAX 255
#define LXW_MAX_NUMBER_URLS   65530
#define LXW_PANE_NAME_LENGTH  12        /* bottomRight + 1 */

/* The Excel 2007 specification says that the maximum number of page
 * breaks is 1026. However, in practice it is actually 1023. */
#define LXW_BREAKS_MAX        1023

/** Default column width in Excel */
#define LXW_DEF_COL_WIDTH     8.43

/** Default row height in Excel */
#define LXW_DEF_ROW_HEIGHT    15.0

/** Gridline options using in `worksheet_gridlines()`. */
    /** Hide screen and print gridlines. */
    #define LXW_HIDE_ALL_GRIDLINES 				0

    /** Show screen gridlines. */
    #define LXW_SHOW_SCREEN_GRIDLINES				1

    /** Show print gridlines. */
    #define LXW_SHOW_PRINT_GRIDLINES				2

    /** Show screen and print gridlines. */
    #define LXW_SHOW_ALL_GRIDLINES				3


/** Data validation property values. */
    #define LXW_VALIDATION_DEFAULT				0

    /** Turn a data validation property off. */
    #define LXW_VALIDATION_OFF					1

    /** Turn a data validation property on. Data validation properties are
     * generally on by default. */
    #define LXW_VALIDATION_ON					2


/** Data validation types. */
    #define LXW_VALIDATION_TYPE_NONE				0

    /** Restrict cell input to whole/integer numbers only. */
    #define LXW_VALIDATION_TYPE_INTEGER				1

    /** Restrict cell input to whole/integer numbers only, using a cell
     *  reference. */
    #define LXW_VALIDATION_TYPE_INTEGER_FORMULA			2

    /** Restrict cell input to decimal numbers only. */
    #define LXW_VALIDATION_TYPE_DECIMAL				3

    /** Restrict cell input to decimal numbers only, using a cell
     * reference. */
    #define LXW_VALIDATION_TYPE_DECIMAL_FORMULA			4

    /** Restrict cell input to a list of strings in a dropdown. */
    #define LXW_VALIDATION_TYPE_LIST				5

    /** Restrict cell input to a list of strings in a dropdown, using a
     * cell range. */
    #define LXW_VALIDATION_TYPE_LIST_FORMULA			6

    /** Restrict cell input to date values only, using a lxw_datetime type. */
    #define LXW_VALIDATION_TYPE_DATE				7

    /** Restrict cell input to date values only, using a cell reference. */
    #define LXW_VALIDATION_TYPE_DATE_FORMULA			8

    /* Restrict cell input to date values only, as a serial number.
     * Undocumented. */
    #define LXW_VALIDATION_TYPE_DATE_NUMBER			9

    /** Restrict cell input to time values only, using a lxw_datetime type. */
    #define LXW_VALIDATION_TYPE_TIME				10

    /** Restrict cell input to time values only, using a cell reference. */
    #define LXW_VALIDATION_TYPE_TIME_FORMULA			11

    /* Restrict cell input to time values only, as a serial number.
     * Undocumented. */
    #define LXW_VALIDATION_TYPE_TIME_NUMBER			12

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    #define LXW_VALIDATION_TYPE_LENGTH				13

    /** Restrict cell input to strings of defined length, using a cell
     * reference. */
    #define LXW_VALIDATION_TYPE_LENGTH_FORMULA			14

    /** Restrict cell to input controlled by a custom formula that returns
     * `TRUE/FALSE`. */
    #define LXW_VALIDATION_TYPE_CUSTOM_FORMULA			15

    /** Allow any type of input. Mainly only useful for pop-up messages. */
    #define LXW_VALIDATION_TYPE_ANY				16

/** Data validation criteria uses to control the selection of data. */
    #define LXW_VALIDATION_CRITERIA_NONE			0

    /** Select data between two values. */
    #define LXW_VALIDATION_CRITERIA_BETWEEN			1

    /** Select data that is not between two values. */
    #define LXW_VALIDATION_CRITERIA_NOT_BETWEEN			2

    /** Select data equal to a value. */
    #define LXW_VALIDATION_CRITERIA_EQUAL_TO			3

    /** Select data not equal to a value. */
    #define LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO		4

    /** Select data greater than a value. */
    #define LXW_VALIDATION_CRITERIA_GREATER_THAN		5

    /** Select data less than a value. */
    #define LXW_VALIDATION_CRITERIA_LESS_THAN			6

    /** Select data greater than or equal to a value. */
    #define LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO	7

    /** Select data less than or equal to a value. */
    #define LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO	8


/** Data validation error types for pop-up messages. */
    /** Show a "Stop" data validation pop-up message. This is the default. */
    #define LXW_VALIDATION_ERROR_TYPE_STOP			0

    /** Show an "Error" data validation pop-up message. */
    #define LXW_VALIDATION_ERROR_TYPE_WARNING			1

    /** Show an "Information" data validation pop-up message. */
    #define LXW_VALIDATION_ERROR_TYPE_INFORMATION		2


//enum cell_types {
    #define NUMBER_CELL 				1
    #define STRING_CELL					2
    #define INLINE_STRING_CELL				3
    #define INLINE_RICH_STRING_CELL			4
    #define FORMULA_CELL				5
    #define ARRAY_FORMULA_CELL				6
    #define BLANK_CELL					7
    #define BOOLEAN_CELL				8
    #define HYPERLINK_URL				9
    #define HYPERLINK_INTERNAL				10
    #define HYPERLINK_EXTERNAL				11

//enum pane_types {
    #define NO_PANES 					0
    #define FREEZE_PANES				1
    #define SPLIT_PANES					2
    #define FREEZE_SPLIT_PANES				3




/****************************************************
 *   DATA VALIDATION
 ***************************************************/
#define LXW_DV_VALIDATE			0
#define LXW_DV_CRITERIA			1
#define LXW_DV_IGNORE_BLANK		2
#define LXW_DV_SHOW_INPUT		3
#define LXW_DV_SHOW_ERROR		4
#define LXW_DV_DROPDOWN			5
#define LXW_DV_IS_BETWEEN		6
#define LXW_DV_ERROR_TYPE		7
#define LXW_DV_VALUE_NUMBER		8
#define LXW_DV_VALUE_FORMULA		9
#define LXW_DV_VALUE_LIST		10
#define LXW_DV_VALUE_DATETIME		11
#define LXW_DV_MINIMUM_NUMBER		12
#define LXW_DV_MINIMUM_FORMULA		13
#define LXW_DV_MINIMUM_DATETIME		14
#define LXW_DV_MAXIMUM_NUMBER		15
#define LXW_DV_MAXIMUM_FORMULA		16
#define LXW_DV_MAXIMUM_DATETIME		17
#define LXW_DV_INPUT_TITLE		18
#define LXW_DV_INPUT_MESSAGE		19
#define LXW_DV_ERROR_TITLE		20
#define LXW_DV_ERROR_MESSAGE		21

//eof
