/* xls.i EXCEL SPREADSHEET IMPORT/EXPORT PROCEDURES
         KURT GUNDERSON 11/12/2008

   PURPOSE:

   MODIFICATIONS:
      11/12/2008 KAG Pulled out of fix.i
      11/24/2008 KAG Added preDEFINED protection.
      02/19/2009 KAG Added getExcelVersion() and getExcelExtension() functions.
      03/23/2009 KAG Among other things, added merge functionality.
      04/06/2009 KAG Added sheet properties for landscape/portrait and 
                     fit to page.
      04/30/2009 KAG Added extension list to inputXLS().
      09/05/2009 KAG Added header picture functionality for letterhead.
      03/24/2010 KAG Use vchSheet:Columns:Count for column count when sent
                     zero as column parameter the same as rows.
      01/19/2010 KAG Added watermark/header logic.  LeftHeader spelled 
                     incorrectly was not placing header correctly.
      12/23/2011 KAG Added font size.  Added com-handle manager to manage
                     memory used by Excel API.
      06/28/2012 KAG Created second excel.exe process as sacrificial anode to
                     avoid ghost/malformed spreadsheets when user is working in
                     Excel while this process generates a spreadsheet.
      07/17/2012 KAG Microsoft must have changed the API.  Range():FormulaR1C1
                     will now produce invalid component error; changed to use 
                     Range():Value
      08/07/2012 KAG Added Office 2003 XML Spreadsheet functionality in
                     outputXML() procedure.  Significant performance gains and
                     the ability to generate spreadsheets w/o Office installed.
      11/27/2012 KAG Added printcommunication flag to possibly correct errors
                     when default printer does not exist.
      12/13/2012 KAG Remove printcommunication flag; failing at HCC.
      04/03/2013 KAG BUGFIX: Excel spreadsheet format in OE11.1 was failing on
                     HeaderPictures.  File name needed to be absolute pathed.
      10/02/2013 KAG Restructured temp-table/buffer names to allow shared/new shared
                     functionality.
      10/06/2014 KAG Allow string formatting within cells of XML Spreadsheet.

   CERTIFIED:

      OE10.1C
      OE10.2B
      OE11.1

   NOTES:
      XLS REQUIREMENTS:

      XLS UNSUPPORTED FEATURES:
         1. SUBSCRIPT/SUPERSCRIPT not supported (yet)
         2. FORMATTING OF INDIVIDUAL CHARACTERS IN CELLS

      XML REQUIREMENTS:
         0. See document...
            http://msdn.microsoft.com/en-us/library/office/aa140066%28v=office.10%29.aspx
         1. createFormula() must be in R1C1 notation as opposed to A1 notation.
            outputXLS() supports both R1C1 and A1 notations.
         2. createCell() numbers must have NUMBER=TRUE

      XML UNSUPPORTED FEATURES:
         1. PRINTAREA not supported (yet)
         2. PASSWORDS/PROTECTION not supported (ever --Excel 2003 XML
            Spreadsheet architecture does not support)

*/
/******************************************************************************/
&IF DEFINED(XLS_I) = 0 &THEN
&GLOBAL-DEFINE XLS_I                          TRUE
&GLOBAL-DEFINE XLS_ALIGNLEFT                  2
&GLOBAL-DEFINE XLS_ALIGNCENTER                3
&GLOBAL-DEFINE XLS_ALIGNRIGHT                 4
&GLOBAL-DEFINE XLS_BORDERLEFT                 7
&GLOBAL-DEFINE XLS_BORDERTOP                  8
&GLOBAL-DEFINE XLS_BORDERRIGHT                10
&GLOBAL-DEFINE XLS_BORDERBOTTOM               9
&GLOBAL-DEFINE XLS_LINENONE                   0
&GLOBAL-DEFINE XLS_LINECONTINUOUS             1
&GLOBAL-DEFINE XLS_LINEDASH                   2
&GLOBAL-DEFINE XLS_LINEDOT                    3
&GLOBAL-DEFINE XLS_LINEDOUBLE                 9
&GLOBAL-DEFINE XLS_LINELIGHT                  1
&GLOBAL-DEFINE XLS_LINEMEDIUM                 2
&GLOBAL-DEFINE XLS_LINEHEAVY                  3
&GLOBAL-DEFINE XLS_EXCEL_97                   8.0
&GLOBAL-DEFINE XLS_EXCEL_2000                 9.0
&GLOBAL-DEFINE XLS_EXCEL_2003                 11.0
&GLOBAL-DEFINE XLS_EXCEL_2007                 12.0
&GLOBAL-DEFINE XLS_EXCEL_2010                 14.0
&GLOBAL-DEFINE XLS_EXT_LIST                   "xls,xlsx":U
&GLOBAL-DEFINE XLS_LANDSCAPE                  2
&GLOBAL-DEFINE XLS_PORTRAIT                   1
&GLOBAL-DEFINE XML_DEFAULT_FONT               "Calibri":U
&GLOBAL-DEFINE XML_DEFAULT_FONT_FAMILY        "Swiss":U
&GLOBAL-DEFINE XML_DEFAULT_FONT_SIZE          11
&GLOBAL-DEFINE XML_DEFAULT_FONT_COLOR         "#000000":U
&GLOBAL-DEFINE XML_DEFAULT_ROW_HEIGHT         15
&GLOBAL-DEFINE XML_DEFAULT_HEADER_MARGIN      0.3
&GLOBAL-DEFINE XML_DEFAULT_FOOTER_MARGIN      0.3
&GLOBAL-DEFINE XML_DEFAULT_PAGE_MARGIN_BOTTOM 0.75
&GLOBAL-DEFINE XML_DEFAULT_PAGE_MARGIN_LEFT   0.7
&GLOBAL-DEFINE XML_DEFAULT_PAGE_MARGIN_RIGHT  0.7
&GLOBAL-DEFINE XML_DEFAULT_PAGE_MARGIN_TOP    0.75
define variable gxls_no      as integer initial 0 no-undo.
define variable gxls_ref_seq as integer initial 1 no-undo.
define variable gwb_no       as integer initial 1 no-undo.
define variable gws_no       as integer initial 1 no-undo.
define variable gstyle_id    as integer initial 1 no-undo.
define variable gnum_id      as integer initial 1 no-undo.
define variable gint_id      as integer initial 1 no-undo.
define variable gfont_id     as integer initial 1 no-undo.
define variable galign_id    as integer initial 1 no-undo.
define variable gprot_id     as integer initial 1 no-undo.
define variable gbord_id     as integer initial 1 no-undo.
define variable gpane_no     as integer initial 3 no-undo.
define stream sxmlin.
define stream sxmlout.
{sys/sys.i}

define {1} temp-table tt_excel no-undo
   field xls_no          as integer    label "XLS"        format ">>9"
   field sheet_no        as integer    label "SHEET"      format ">>9"
   field row_no          as integer    label "ROW"        format ">>,>>9"
   field column_no       as integer    label "COLUMN"     format ">>,>>9"
   field cell_name       as character  label "NAME"       format "x(10)"
   field cell_value      as character  label "VALUE"      format "x(20)"
   field cell_properties as character  label "PROPERTIES" format "x(10)"
   field is_formula      as logical    label "FUNCTION?"  format "yes/no" initial no

   index idx_prim as primary unique
      xls_no
      sheet_no
      row_no
      column_no

   index idx_name
      xls_no
      sheet_no
      cell_name
   .

define {1} temp-table tt_excel_sheet no-undo
   field xls_no           as integer   label "XLS"   format ">>9"
   field sheet_no         as integer   label "SHEET" format ">>9"
   field sheet_name       as character label "NAME"  format "x(30)"
   field sheet_properties as character label "PROPERTIES" format "x(10)"

   field wb_no            as integer initial ?
   field ws_no            as integer initial ?

   index idx_prim as primary unique
      xls_no
      sheet_no

   index idx_name as unique
      xls_no
      sheet_name
   .

define {1} temp-table tt_excel_merge no-undo
   field xls_no     as integer   label "XLS"   format ">>9"
   field sheet_no   as integer   label "SHEET" format ">>9"
   field merge      as character label "MERGE" format "x(30)"

   index idx_prim as primary unique
      xls_no
      sheet_no
      merge
   .

define {1} temp-table tt_excel_handle no-undo
   field ref_hdl as com-handle
   field ref_seq as integer

   index idx_prim as primary unique
      ref_hdl

   index idx_no as unique
      ref_seq descending
   .

define temp-table tt_xml_workbook no-undo
   serialize-name "Workbook"

   field ns0       as character initial "urn:schemas-microsoft-com:office:spreadsheet" serialize-name "xmlns" xml-node-type "attribute"
   field ns1       as character initial "urn:schemas-microsoft-com:office:office" serialize-name "xmlns:o" xml-node-type "attribute"
   field ns2       as character initial "urn:schemas-microsoft-com:office:excel" serialize-name "xmlns:x" xml-node-type "attribute"
   field ns3       as character initial "urn:schemas-microsoft-com:office:spreadsheet" serialize-name "xmlns:ss" xml-node-type "attribute"
   field ns4       as character initial "http://www.w3.org/TR/REC-html40" serialize-name "xmlns:html" xml-node-type "attribute"

   field wb_no     as integer serialize-hidden

   index idx_prim as primary unique
      wb_no
   .

define temp-table tt_xml_excelworkbook no-undo
   serialize-name "ExcelWorkbook"

   field ns0                 as character initial "urn:schemas-microsoft-com:office:excel" serialize-name "xmlns" xml-node-type "attribute"

   field wb_no               as integer serialize-hidden
   field window_height       as integer serialize-name "WindowHeight"
   field window_width        as integer serialize-name "WindowWidth"
   field window_topx         as integer serialize-name "WindowTopX"
   field window_topy         as integer serialize-name "WindowTopY"
   field active_sheet        as integer initial 0 serialize-name "ActiveSheet"
   field first_visible_sheet as integer initial 0 serialize-name "FirstVisibleSheet"
   field protect_structure   as logical serialize-hidden /* serialize-name "ProtectStructure" */
   field protect_windows     as logical serialize-hidden /* serialize-name "ProtectWindows" */

   index idx_prim as primary unique
      wb_no
   .

define temp-table tt_xml_officedocumentsettings no-undo
   serialize-name "OfficeDocumentSettings"

   field ns0       as character initial "urn:schemas-microsoft-com:office:office" serialize-name "xmlns" xml-node-type "attribute"

   field wb_no     as integer serialize-hidden
   field allow_png as character initial "" serialize-name "AllowPNG"

   index idx_prim as primary unique
      wb_no
   .

define temp-table tt_xml_documentproperties no-undo
   serialize-name "DocumentProperties"

   field ns0         as character initial "urn:schemas-microsoft-com:office:office" serialize-name "xmlns" xml-node-type "attribute"

   field wb_no       as integer serialize-hidden
   field author      as character serialize-name "Author"
   field last_author as character serialize-name "LastAuthor"
   field created     as datetime-tz serialize-name "Created"
   field last_saved  as datetime-tz serialize-name "LastSaved"
   field company     as character serialize-name "Company"
   field version     as character serialize-name "Version"

   index idx_prim as primary unique
      wb_no
   .

define temp-table tt_xml_worksheet no-undo
   serialize-name "Worksheet"   

   field wb_no      as integer serialize-hidden
   field ws_no      as integer serialize-hidden
   field name       as character serialize-name "ss:Name" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_worksheetoptions no-undo
   serialize-name "WorksheetOptions"

   field ns0               as character initial "urn:schemas-microsoft-com:office:excel" serialize-name "xmlns" xml-node-type "attribute"

   field wb_no             as integer serialize-hidden
   field ws_no             as integer serialize-hidden
   field no_grid_lines     as logical initial false serialize-name "DoNotDisplayGridLines"
   field protect_objects   as logical serialize-hidden /* serialize-name "ProtectObjects" */
   field protect_scenarios as logical serialize-hidden /* serialize-name "ProtectScenarios" */
   field fit_to_page       as character initial "" serialize-name "FitToPage"
   field is_selected       as character initial "" serialize-name "Selected"
   field unsynced          as character initial "" serialize-name "Unsynced"
   field enable_selection  as character initial "" serialize-name "EnableSelection"
   field top_row_visible   as integer initial 0 serialize-name "TopRowVisible"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_print no-undo
   serialize-name "Print"

   field wb_no                 as integer serialize-hidden
   field ws_no                 as integer serialize-hidden
   field fit_width             as integer initial 0 serialize-name "FitWidth"
   field fit_height            as integer initial 0 serialize-name "FitHeight"
   field valid_printer_info    as character initial "" serialize-name "ValidPrinterInfo"
   field scale                 as integer initial 0 serialize-name "Scale"
   field vertical_resolution   as integer initial 0 serialize-name "VerticalResolution"
   field horizontal_resolution as integer initial 0 serialize-name "HorizontalResolution"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_pagesetup no-undo
   serialize-name "PageSetup"

   field wb_no as integer serialize-hidden
   field ws_no as integer serialize-hidden

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_panes no-undo
   serialize-name "Panes"

   field wb_no as integer serialize-hidden
   field ws_no as integer serialize-hidden

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_pane no-undo
   serialize-name "Pane"

   field wb_no           as integer serialize-hidden
   field ws_no           as integer serialize-hidden
   field pane_no         as integer serialize-name "Number"
   field active_row      as integer serialize-name "ActiveRow"
   field active_column   as integer serialize-name "ActiveCol"
   field range_Selection as character serialize-name "RangeSelection"

   index idx_prim as primary unique
      wb_no
      ws_no
      pane_no
   .

define temp-table tt_xml_header no-undo
   serialize-name "Header"

   field wb_no  as integer serialize-hidden
   field ws_no  as integer serialize-hidden
   field margin as decimal serialize-name "x:Margin" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_footer no-undo
   serialize-name "Footer"

   field wb_no  as integer serialize-hidden
   field ws_no  as integer serialize-hidden
   field margin as decimal serialize-name "x:Margin" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_pagemargins no-undo
   serialize-name "PageMargins"

   field wb_no  as integer serialize-hidden
   field ws_no  as integer serialize-hidden
   field bottom as decimal serialize-name "x:Bottom" xml-node-type "attribute"
   field left   as decimal serialize-name "x:Left" xml-node-type "attribute"
   field right  as decimal serialize-name "x:Right" xml-node-type "attribute"
   field top    as decimal serialize-name "x:Top" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_layout no-undo
   serialize-name "Layout"

   field wb_no       as integer serialize-hidden
   field ws_no       as integer serialize-hidden
   field orientation as character serialize-name "x:Orientation" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_table no-undo
   serialize-name "Table"

   field wb_no                 as integer serialize-hidden
   field ws_no                 as integer serialize-hidden
   field expanded_column_count as integer initial 0 serialize-name "ss:ExpandedColumnCount" xml-node-type "attribute"
   field expanded_row_count    as integer initial 0 serialize-name "ss:ExpandedRowCount" xml-node-type "attribute"
   field full_columns          as integer initial 1 serialize-name "x:FullColumns" xml-node-type "attribute"
   field full_rows             as integer initial 1 serialize-name "x:FullRows" xml-node-type "attribute"
   field default_row_height    as integer serialize-name "ss:DefaultRowHeight" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_names no-undo
   serialize-name "Names"

   field wb_no as integer serialize-hidden
   field ws_no as integer serialize-hidden

   index idx_prim as primary unique
      wb_no
      ws_no
   .

define temp-table tt_xml_namedrange no-undo
   serialize-name "NamedRange"

   field wb_no      as integer serialize-hidden
   field ws_no      as integer serialize-hidden
   field range_name as character serialize-name "ss:Name" xml-node-type "attribute"
   field refers_to  as character initial ? serialize-name "ss:RefersTo" xml-node-type "attribute"
   field hidden     as integer initial ? serialize-name "ss:Hidden" xml-node-type "attribute" 

   index idx_prim as primary unique
      wb_no
      ws_no
      range_name
   .

define temp-table tt_xml_column no-undo
   serialize-name "Column"

   field wb_no          as integer serialize-hidden
   field ws_no          as integer serialize-hidden
   field co_no          as integer serialize-name "ss:Index" xml-node-type "attribute"
   field column_width   as decimal initial ? serialize-name "ss:Width" xml-node-type "attribute"
   field auto_fit_width as integer initial ? serialize-name "ss:AutoFitWidth" xml-node-type "attribute"
   field span           as integer initial ? serialize-name "ss:Span" xml-node-type "attribute"
   field hidden         as logical initial ? serialize-name "ss:Hidden" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
      co_no
   .

define temp-table tt_xml_row no-undo
   serialize-name "Row"

   field wb_no           as integer serialize-hidden
   field ws_no           as integer serialize-hidden
   field ro_no           as integer serialize-name "ss:Index" xml-node-type "attribute"
   field row_height      as integer initial ? serialize-name "ss:Height" xml-node-type "attribute"
   field auto_fit_height as integer initial ? serialize-name "ss:AutoFitHeight" xml-node-type "attribute"
   field span            as integer initial ? serialize-name "ss:Span" xml-node-type "attribute"
   field hidden          as logical initial ? serialize-name "ss:Hidden" xml-node-type "attribute"

   index idx_prim as primary unique
      wb_no
      ws_no
      ro_no
   .

define temp-table tt_xml_cell no-undo
   serialize-name "Cell"

   field wb_no        as integer serialize-hidden
   field ws_no        as integer serialize-hidden
   field ro_no        as integer serialize-hidden
   field co_no        as integer serialize-name "ss:Index" xml-node-type "attribute"
   field merge_across as integer initial ? serialize-name "ss:MergeAcross" xml-node-type "attribute"
   field merge_down   as integer initial ? serialize-name "ss:MergeDown" xml-node-type "attribute"
   field style_id     as character serialize-name "ss:StyleID" xml-node-type "attribute"
   field name         as character initial ? serialize-name "ss:Name" xml-node-type "attribute"
   field formula      as character initial ? serialize-name "ss:Formula" xml-node-type "attribute"
   field ticked       as integer initial ? serialize-name "x:Ticked" xml-node-type "attribute"
 
   index idx_prim as primary unique
      wb_no
      ws_no
      ro_no
      co_no
   .

define temp-table tt_xml_data no-undo
   serialize-name "ss:Data"

   field ns0         as character initial "http://www.w3.org/TR/REC-html40" serialize-name "xmlns" xml-node-type "attribute"

   field wb_no      as integer serialize-hidden
   field ws_no      as integer serialize-hidden
   field ro_no      as integer serialize-hidden
   field co_no      as integer serialize-hidden
   field cell_type  as character serialize-name "ss:Type" xml-node-type "attribute"
      /* String, Number, DateTime */
   field cell_value as character xml-node-type "TEXT"

   field inline_xml as logical serialize-hidden
 
   index idx_prim as primary unique
      wb_no
      ws_no
      ro_no
      co_no
   .

define temp-table tt_xml_data_inline no-undo
   field wb_no      as integer
   field ws_no      as integer
   field ro_no      as integer
   field co_no      as integer

   field hash_value as character
   field cell_value as character
 
   index idx_prim as primary unique
      wb_no
      ws_no
      ro_no
      co_no

   index idx_hash as unique
      hash_value
   .

define temp-table tt_xml_namedcell no-undo
   serialize-name "NamedCell"

   field wb_no      as integer serialize-hidden
   field ws_no      as integer serialize-hidden
   field ro_no      as integer serialize-hidden
   field co_no      as integer serialize-hidden
   field cell_name  as character serialize-name "ss:Name" xml-node-type "attribute"
   field refers_to  as character initial ? serialize-name "ss:RefersTo" xml-node-type "attribute"
   field hidden     as integer initial ? serialize-name "ss:Hidden" xml-node-type "attribute" 
 
   index idx_prim as primary unique
      wb_no
      ws_no
      ro_no
      co_no
   .

define temp-table tt_xml_styles no-undo
   serialize-name "Styles"

   field wb_no      as integer serialize-hidden

   index idx_prim as primary unique
      wb_no
   .

define temp-table tt_xml_style no-undo
   serialize-name "Style"

   field wb_no      as integer serialize-hidden
   field style_id   as character serialize-name "ss:ID" xml-node-type "attribute"
   field style_name as character initial ? serialize-name "ss:Name" xml-node-type "attribute"
   field num_id     as integer initial 0 serialize-hidden
   field int_id     as integer initial 0 serialize-hidden
   field font_id    as integer initial 0 serialize-hidden
   field align_id   as integer initial 0 serialize-hidden
   field prot_id    as integer initial 0 serialize-hidden
   field bord_id    as integer initial 0 serialize-hidden

   index idx_prim as primary unique
      wb_no
      style_id

   index idx_id as unique
      num_id
      int_id
      font_id
      align_id
      prot_id
      bord_id
   .

define temp-table tt_xml_protection no-undo
   serialize-name "Protection"

   field prot_id      as integer serialize-hidden
   field protected    as integer initial ? serialize-name "ss:Protected" xml-node-type "attribute"
   field hide_formula as integer initial ? serialize-name "x:HideFormula" xml-node-type "attribute"

   index idx_prim as primary unique
      prot_id

   index idx_prot as unique
      protected
      hide_formula
   .

define temp-table tt_xml_borders no-undo
   serialize-name "Borders"

   field bord_id       as integer serialize-hidden
   field border_bottom as character initial ? serialize-hidden
   field border_left   as character initial ? serialize-hidden
   field border_right  as character initial ? serialize-hidden
   field border_top    as character initial ? serialize-hidden

   index idx_prim as primary unique
      bord_id

   index idx_bord as unique
      border_bottom
      border_left
      border_right
      border_top
   .

define temp-table tt_xml_border no-undo
   serialize-name "Border"

   field bord_id      as integer serialize-hidden
   field position     as character initial ? serialize-name "ss:Position" xml-node-type "attribute"
      /* Left, Top, Right, Bottom, DiagonalLeft, and DiagonalRight */
   field border_color as character initial ? serialize-name "ss:Color" xml-node-type "attribute"
   field line_style   as character initial ? serialize-name "ss:LineStyle" xml-node-type "attribute"
      /* None, Continuous, Dash, Dot, DashDot, DashDotDot, SlantDashDot, and Double */
   field weight       as integer initial ? serialize-name "ss:Weight" xml-node-type "attribute"
      /* 0-Hairline 1-Thin 2-Medium 3-Thick */

   index idx_prim as primary unique
      bord_id
      position
   .

define temp-table tt_xml_alignment no-undo
   serialize-name "Alignment"

   field align_id         as integer serialize-hidden
   field align_vertical   as character initial ? serialize-name "ss:Vertical" xml-node-type "attribute"
   field align_horizontal as character initial ? serialize-name "ss:Horizontal" xml-node-type "attribute"

   index idx_prim as primary unique
      align_id

   index idx_align as unique
      align_vertical
      align_horizontal
   .

define temp-table tt_xml_font no-undo
   serialize-name "Font"

   field font_id        as integer serialize-hidden
   field font_name      as character initial ? serialize-name "ss:FontName" xml-node-type "attribute"
   field font_family    as character initial ? serialize-name "x:Family" xml-node-type "attribute"
   field font_size      as integer initial ? serialize-name "ss:Size" xml-node-type "attribute"
   field font_color     as character initial ? serialize-name "ss:Color" xml-node-type "attribute"
   field font_bold      as integer initial ? serialize-name "ss:Bold" xml-node-type "attribute"
   field font_italic    as integer initial ? serialize-name "ss:Italic" xml-node-type "attribute"
   field font_vertical  as character initial ? serialize-name "ss:VerticalAlign" xml-node-type "attribute" 
      /* None, Subscript, Superscript */

   index idx_prim as primary unique
      font_id

   index idx_font as unique
      font_name
      font_family
      font_size
      font_color
      font_bold
      font_italic
      font_vertical
   .

define temp-table tt_xml_interior no-undo
   serialize-name "Interior"

   field int_id      as integer serialize-hidden
   field int_color   as character initial ? serialize-name "ss:Color" xml-node-type "attribute"
   field int_pattern as character initial ? serialize-name "ss:Pattern" xml-node-type "attribute"

   index idx_prim as primary unique
      int_id

   index idx_int as unique
      int_color
      int_pattern
   .

define temp-table tt_xml_numberformat no-undo
   serialize-name "NumberFormat"

   field num_id     as integer serialize-hidden
   field num_format as character initial ? serialize-name "ss:Format" xml-node-type "attribute"

   index idx_prim as primary unique
      num_id

   index idx_num as unique
      num_format
   .

define dataset ds_xml 
   for tt_xml_workbook, tt_xml_documentproperties, 
                        tt_xml_officedocumentsettings, 
                        tt_xml_excelworkbook,
                        tt_xml_styles, tt_xml_style, tt_xml_alignment,
                                                     tt_xml_borders, tt_xml_border,
                                                     tt_xml_font,
                                                     tt_xml_interior, 
                                                     tt_xml_numberformat,
                                                     tt_xml_protection,
                        tt_xml_worksheet, tt_xml_names, tt_xml_namedrange,
                                          tt_xml_table, tt_xml_column, tt_xml_row, tt_xml_cell, tt_xml_data, 
                                                                                                tt_xml_namedcell,
                                          tt_xml_worksheetoptions, tt_xml_pagesetup, tt_xml_layout,
                                                                                     tt_xml_header,
                                                                                     tt_xml_footer,
                                                                                     tt_xml_pagemargins,
                                                                   tt_xml_print,
                                                                   tt_xml_panes, tt_xml_pane
   data-relation drdocumentproperties     for tt_xml_workbook, tt_xml_documentproperties relation-fields(wb_no, wb_no) nested
   data-relation drofficedocumentsettings for tt_xml_workbook, tt_xml_officedocumentsettings relation-fields(wb_no, wb_no) nested
   data-relation drexcelworkbook          for tt_xml_workbook, tt_xml_excelworkbook relation-fields(wb_no, wb_no) nested
   data-relation drstyles                 for tt_xml_workbook, tt_xml_styles relation-fields(wb_no, wb_no) nested
   data-relation drstyle                  for tt_xml_styles, tt_xml_style relation-fields(wb_no, wb_no) nested
   data-relation dralignment              for tt_xml_style, tt_xml_alignment relation-fields(align_id, align_id) nested
   data-relation drborders                for tt_xml_style, tt_xml_borders relation-fields(bord_id, bord_id) nested
   data-relation drborder                 for tt_xml_borders, tt_xml_border relation-fields(bord_id, bord_id) nested
   data-relation drfont                   for tt_xml_style, tt_xml_font relation-fields(font_id, font_id) nested
   data-relation drinterior               for tt_xml_style, tt_xml_interior relation-fields(int_id, int_id) nested
   data-relation drnumberformat           for tt_xml_style, tt_xml_numberformat relation-fields(num_id, num_id) nested
   data-relation drprotection             for tt_xml_style, tt_xml_protection relation-fields(prot_id, prot_id) nested
   data-relation drworksheet              for tt_xml_workbook, tt_xml_worksheet relation-fields(wb_no, wb_no) nested
   data-relation drnames                  for tt_xml_worksheet, tt_xml_names relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drnamedrange             for tt_xml_names, tt_xml_namedrange relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drtable                  for tt_xml_worksheet, tt_xml_table relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drcolumn                 for tt_xml_table, tt_xml_column relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drrow                    for tt_xml_table, tt_xml_row relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drcell                   for tt_xml_row, tt_xml_cell relation-fields(wb_no, wb_no, ws_no, ws_no, ro_no, ro_no) nested
   data-relation drdata                   for tt_xml_cell, tt_xml_data relation-fields(wb_no, wb_no, ws_no, ws_no, ro_no, ro_no, co_no, co_no) nested
   data-relation drnamedcell              for tt_xml_cell, tt_xml_namedcell relation-fields(wb_no, wb_no, ws_no, ws_no, ro_no, ro_no, co_no, co_no) nested
   data-relation drworksheetoptions       for tt_xml_worksheet, tt_xml_worksheetoptions relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drpagesetup              for tt_xml_worksheetoptions, tt_xml_pagesetup relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drlayout                 for tt_xml_pagesetup, tt_xml_layout relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drheader                 for tt_xml_pagesetup, tt_xml_header relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drfooter                 for tt_xml_pagesetup, tt_xml_footer relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drpagemargins            for tt_xml_pagesetup, tt_xml_pagemargins relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drprint                  for tt_xml_worksheetoptions, tt_xml_print relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drpanes                  for tt_xml_worksheetoptions, tt_xml_panes relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   data-relation drpane                   for tt_xml_panes, tt_xml_pane relation-fields(wb_no, wb_no, ws_no, ws_no) nested
   .

function createXMLWorkbook returns integer (input pauthor              as character,
                                            input plast_author         as character,
                                            input pcreated             as datetime-tz,
                                            input plast_saved          as datetime-tz,
                                            input pcompany             as character,
                                            input pversion             as character,
                                            input pwindow_height       as integer,
                                            input pwindow_width        as integer,
                                            input pwindow_topx         as integer,
                                            input pwindow_topy         as integer,
                                            input pprotect_structure   as logical,
                                            input pprotect_windows     as logical,
                                            input pfont_name           as character,
                                            input pfont_family         as character,
                                            input pfont_size           as integer,
                                            input pfont_color          as character):
   define variable wwb_no as integer initial ? no-undo.
   define buffer tt_xml_workbook               for tt_xml_workbook.
   define buffer tt_xml_documentproperties     for tt_xml_documentproperties.
   define buffer tt_xml_styles                 for tt_xml_styles.
   define buffer tt_xml_style                  for tt_xml_style.
   define buffer tt_xml_officedocumentsettings for tt_xml_officedocumentsettings.
   define buffer tt_xml_excelworkbook          for tt_xml_excelworkbook.
   define buffer tt_xml_alignment              for tt_xml_alignment.
   define buffer tt_xml_interior               for tt_xml_interior.
   define buffer tt_xml_numberformat           for tt_xml_numberformat.
   define buffer tt_xml_protection             for tt_xml_protection.
   define buffer tt_xml_font                   for tt_xml_font.
   define buffer tt_xml_borders                for tt_xml_borders.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      create tt_xml_workbook.
      assign tt_xml_workbook.wb_no = gwb_no
             gwb_no                = gwb_no + 1.
      validate tt_xml_workbook.

      create tt_xml_documentproperties.
      assign tt_xml_documentproperties.wb_no       = tt_xml_workbook.wb_no
             tt_xml_documentproperties.author      = pauthor
             tt_xml_documentproperties.last_author = plast_author
             tt_xml_documentproperties.created     = pcreated
             tt_xml_documentproperties.last_saved  = plast_saved
             tt_xml_documentproperties.company     = pcompany
             tt_xml_documentproperties.version     = pversion.
      validate tt_xml_documentproperties.

      create tt_xml_styles.
      assign tt_xml_styles.wb_no = tt_xml_workbook.wb_no.
      validate tt_xml_styles.

      create tt_xml_style.
      assign tt_xml_style.wb_no      = tt_xml_workbook.wb_no
             tt_xml_style.style_id   = "Default"
             tt_xml_style.style_name = "Normal".
      validate tt_xml_style.

      create tt_xml_alignment.
      assign tt_xml_alignment.align_id       = 0
             tt_xml_alignment.align_vertical = "Bottom".
      validate tt_xml_alignment.

      create tt_xml_interior.
      assign tt_xml_interior.int_id = 0.
      validate tt_xml_interior.

      create tt_xml_numberformat.
      assign tt_xml_numberformat.num_id = 0.
      validate tt_xml_numberformat.

      create tt_xml_protection.
      assign tt_xml_protection.prot_id = 0.
      validate tt_xml_protection.

      create tt_xml_font.
      assign tt_xml_font.font_id     = 0
             tt_xml_font.font_name   = pfont_name
             tt_xml_font.font_family = pfont_family
             tt_xml_font.font_size   = pfont_size
             tt_xml_font.font_color  = pfont_color.
      validate tt_xml_font.

      create tt_xml_borders.
      assign tt_xml_borders.bord_id = 0.
      validate tt_xml_borders.

      create tt_xml_officedocumentsettings.
      assign tt_xml_officedocumentsettings.wb_no = tt_xml_workbook.wb_no.
      validate tt_xml_officedocumentsettings.

      create tt_xml_excelworkbook.
      assign tt_xml_excelworkbook.wb_no             = tt_xml_workbook.wb_no
             tt_xml_excelworkbook.window_height     = pwindow_height
             tt_xml_excelworkbook.window_width      = pwindow_width
             tt_xml_excelworkbook.window_topx       = pwindow_topx
             tt_xml_excelworkbook.window_topy       = pwindow_topy
             tt_xml_excelworkbook.protect_structure = pprotect_structure
             tt_xml_excelworkbook.protect_windows   = pprotect_windows.
      validate tt_xml_excelworkbook.

      assign wwb_no = tt_xml_workbook.wb_no.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wwb_no.
end function.

function createXMLWorksheet returns integer (input pwb_no                 as integer,
                                             input pname                  as character,
                                             input pno_grid_lines         as logical,
                                             input pprotect_objects       as logical,
                                             input pprotect_scenarios     as logical,
                                             input pfit_to_page           as character,
                                             input pfit_width             as integer,
                                             input pfit_height            as integer,
                                             input pis_selected           as character,
                                             input pdefault_row_height    as integer,
                                             input pheader_margin         as decimal,
                                             input pfooter_margin         as decimal,
                                             input ppagemargin_bottom     as decimal,
                                             input ppagemargin_left       as decimal,
                                             input ppagemargin_right      as decimal,
                                             input ppagemargin_top        as decimal,
                                             input porientation           as character,
                                             input pprint_area            as character):
   define variable wws_no as integer initial ? no-undo.
   define buffer tt_xml_worksheet        for tt_xml_worksheet.
   define buffer tt_xml_worksheetoptions for tt_xml_worksheetoptions.
   define buffer tt_xml_names            for tt_xml_names.
   define buffer tt_xml_namedrange       for tt_xml_namedrange.
   define buffer tt_xml_table            for tt_xml_table.
   define buffer tt_xml_pagesetup        for tt_xml_pagesetup.
   define buffer tt_xml_header           for tt_xml_header.
   define buffer tt_xml_footer           for tt_xml_footer.
   define buffer tt_xml_pagemargins      for tt_xml_pagemargins.
   define buffer tt_xml_layout           for tt_xml_layout.
   define buffer tt_xml_print            for tt_xml_print.
   define buffer tt_xml_panes            for tt_xml_panes.
   define buffer tt_xml_pane             for tt_xml_pane.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      create tt_xml_worksheet.
      assign tt_xml_worksheet.name  = pname
             tt_xml_worksheet.wb_no = pwb_no
             tt_xml_worksheet.ws_no = gws_no
             gws_no                 = gws_no + 1.
      validate tt_xml_worksheet.

      create tt_xml_worksheetoptions.
      assign tt_xml_worksheetoptions.wb_no             = tt_xml_worksheet.wb_no
             tt_xml_worksheetoptions.ws_no             = tt_xml_worksheet.ws_no
             tt_xml_worksheetoptions.no_grid_lines     = pno_grid_lines
             tt_xml_worksheetoptions.protect_objects   = pprotect_objects
             tt_xml_worksheetoptions.protect_scenarios = pprotect_scenarios
             tt_xml_worksheetoptions.fit_to_page       = pfit_to_page
             tt_xml_worksheetoptions.is_selected       = pis_selected.
      validate tt_xml_worksheetoptions.

      create tt_xml_table.
      assign tt_xml_table.wb_no              = tt_xml_worksheet.wb_no
             tt_xml_table.ws_no              = tt_xml_worksheet.ws_no
             tt_xml_table.default_row_height = pdefault_row_height.
      validate tt_xml_table. 

      create tt_xml_names.
      assign tt_xml_names.wb_no = tt_xml_worksheet.wb_no
             tt_xml_names.ws_no = tt_xml_worksheet.ws_no.
      validate tt_xml_names. 

      if pprint_area <> ? then do:
         create tt_xml_namedrange.
         assign tt_xml_namedrange.wb_no      = tt_xml_worksheet.wb_no
                tt_xml_namedrange.ws_no      = tt_xml_worksheet.ws_no
                tt_xml_namedrange.range_name = "Print_Area"
                tt_xml_namedrange.refers_to  = pprint_area.
         validate tt_xml_namedrange. 
      end.

      create tt_xml_pagesetup.
      assign tt_xml_pagesetup.wb_no = tt_xml_worksheet.wb_no
             tt_xml_pagesetup.ws_no = tt_xml_worksheet.ws_no.
      validate tt_xml_pagesetup.

      create tt_xml_header.
      assign tt_xml_header.wb_no  = tt_xml_worksheet.wb_no
             tt_xml_header.ws_no  = tt_xml_worksheet.ws_no
             tt_xml_header.margin = pheader_margin.
      validate tt_xml_header.

      create tt_xml_footer.
      assign tt_xml_footer.wb_no  = tt_xml_worksheet.wb_no
             tt_xml_footer.ws_no  = tt_xml_worksheet.ws_no
             tt_xml_footer.margin = pfooter_margin.
      validate tt_xml_footer.

      create tt_xml_pagemargins.
      assign tt_xml_pagemargins.wb_no  = tt_xml_worksheet.wb_no
             tt_xml_pagemargins.ws_no  = tt_xml_worksheet.ws_no
             tt_xml_pagemargins.bottom = ppagemargin_bottom
             tt_xml_pagemargins.left   = ppagemargin_left
             tt_xml_pagemargins.right  = ppagemargin_right
             tt_xml_pagemargins.top    = ppagemargin_top.
      validate tt_xml_pagemargins.

      create tt_xml_layout.
      assign tt_xml_layout.wb_no       = tt_xml_worksheet.wb_no
             tt_xml_layout.ws_no       = tt_xml_worksheet.ws_no
             tt_xml_layout.orientation = porientation.
      validate tt_xml_layout.

      create tt_xml_print.
      assign tt_xml_print.wb_no                 = tt_xml_worksheet.wb_no
             tt_xml_print.ws_no                 = tt_xml_worksheet.ws_no
             tt_xml_print.fit_width             = pfit_width
             tt_xml_print.fit_height            = pfit_height
             tt_xml_print.valid_printer_info    = ""
             tt_xml_print.vertical_resolution   = 0
             tt_xml_print.horizontal_resolution = 0.
      validate tt_xml_print.

      create tt_xml_panes.
      assign tt_xml_panes.wb_no = tt_xml_worksheet.wb_no
             tt_xml_panes.ws_no = tt_xml_worksheet.ws_no.
      validate tt_xml_panes.

      assign wws_no = tt_xml_worksheet.ws_no.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wws_no.
end function.

function getXMLBordID returns integer (input pborder_bottom as character,
                                       input pborder_left   as character,
                                       input pborder_right  as character,
                                       input pborder_top    as character):
   define variable wbord_id    as integer initial ? no-undo.
   define variable wposition   as character no-undo.
   define variable wline_style as character no-undo.
   define variable wweight     as integer no-undo.
   define buffer tt_xml_borders for tt_xml_borders.
   define buffer tt_xml_border  for tt_xml_border.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_borders
         where tt_xml_borders.border_bottom = pborder_bottom
           and tt_xml_borders.border_left   = pborder_left
           and tt_xml_borders.border_right  = pborder_right
           and tt_xml_borders.border_top    = pborder_top
      exclusive-lock no-error.
      if not available tt_xml_borders then do:
         create tt_xml_borders.
         assign tt_xml_borders.bord_id       = gbord_id
                gbord_id                     = gbord_id + 1
                tt_xml_borders.border_bottom = pborder_bottom
                tt_xml_borders.border_left   = pborder_left
                tt_xml_borders.border_right  = pborder_right
                tt_xml_borders.border_top    = pborder_top.
         validate tt_xml_borders.

         if pborder_bottom <> ? then do:
            assign wposition   = "Bottom"
                   wline_style = if entry(1,pborder_bottom,"|") = "" then ? else entry(1,pborder_bottom,"|")
                   wweight     = if entry(2,pborder_bottom,"|") = "" then ? else integer(entry(2,pborder_bottom,"|")).

            create tt_xml_border.
            assign tt_xml_border.bord_id    = tt_xml_borders.bord_id
                   tt_xml_border.position   = wposition
                   tt_xml_border.line_style = wline_style
                   tt_xml_border.weight     = wweight.
            validate tt_xml_border.
         end.
         if pborder_left <> ? then do:
            assign wposition   = "Left"
                   wline_style = if entry(1,pborder_left,"|") = "" then ? else entry(1,pborder_left,"|")
                   wweight     = if entry(2,pborder_left,"|") = "" then ? else integer(entry(2,pborder_left,"|")).

            create tt_xml_border.
            assign tt_xml_border.bord_id    = tt_xml_borders.bord_id
                   tt_xml_border.position   = wposition
                   tt_xml_border.line_style = wline_style
                   tt_xml_border.weight     = wweight.
            validate tt_xml_border.
         end.
         if pborder_right <> ? then do:
            assign wposition   = "Right"
                   wline_style = if entry(1,pborder_right,"|") = "" then ? else entry(1,pborder_right,"|")
                   wweight     = if entry(2,pborder_right,"|") = "" then ? else integer(entry(2,pborder_right,"|")).

            create tt_xml_border.
            assign tt_xml_border.bord_id    = tt_xml_borders.bord_id
                   tt_xml_border.position   = wposition
                   tt_xml_border.line_style = wline_style
                   tt_xml_border.weight     = wweight.
            validate tt_xml_border.
         end.
         if pborder_top <> ? then do:
            assign wposition   = "Top"
                   wline_style = if entry(1,pborder_top,"|") = "" then ? else entry(1,pborder_top,"|")
                   wweight     = if entry(2,pborder_top,"|") = "" then ? else integer(entry(2,pborder_top,"|")).

            create tt_xml_border.
            assign tt_xml_border.bord_id    = tt_xml_borders.bord_id
                   tt_xml_border.position   = wposition
                   tt_xml_border.line_style = wline_style
                   tt_xml_border.weight     = wweight.
            validate tt_xml_border.
         end.
      end.

      assign wbord_id = tt_xml_borders.bord_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wbord_id.
end function.

function getXMLProtID returns integer (input pprotected    as integer,
                                       input phide_formula as integer):
   define variable wprot_id as integer initial ? no-undo.
   define buffer tt_xml_protection for tt_xml_protection.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_protection
         where tt_xml_protection.protected    = pprotected
           and tt_xml_protection.hide_formula = phide_formula
      exclusive-lock no-error.
      if not available tt_xml_protection then do:
         create tt_xml_protection.
         assign tt_xml_protection.prot_id      = gprot_id
                gprot_id                       = gprot_id + 1
                tt_xml_protection.protected    = pprotected
                tt_xml_protection.hide_formula = phide_formula.
         validate tt_xml_protection.
      end.

      assign wprot_id = tt_xml_protection.prot_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wprot_id.
end function.

function getXMLAlignID returns integer (input palign_vertical   as character,
                                        input palign_horizontal as character):
   define variable walign_id as integer initial ? no-undo.
   define buffer tt_xml_alignment for tt_xml_alignment.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_alignment
         where tt_xml_alignment.align_vertical   = palign_vertical
           and tt_xml_alignment.align_horizontal = palign_horizontal
      exclusive-lock no-error.
      if not available tt_xml_alignment then do:
         create tt_xml_alignment.
         assign tt_xml_alignment.align_id         = galign_id
                galign_id                         = galign_id + 1
                tt_xml_alignment.align_vertical   = palign_vertical
                tt_xml_alignment.align_horizontal = palign_horizontal.
         validate tt_xml_alignment.
      end.

      assign walign_id = tt_xml_alignment.align_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return walign_id.
end function.

function getXMLFontID returns integer (input pfont_name     as character,
                                       input pfont_family   as character,
                                       input pfont_size     as integer,
                                       input pfont_color    as character,
                                       input pfont_bold     as integer,
                                       input pfont_italic   as integer,
                                       input pfont_vertical as character):
   define variable wfont_id as integer initial ? no-undo.
   define buffer tt_xml_font for tt_xml_font.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_font
         where tt_xml_font.font_name     = pfont_name
           and tt_xml_font.font_family   = pfont_family
           and tt_xml_font.font_size     = pfont_size
           and tt_xml_font.font_color    = pfont_color
           and tt_xml_font.font_bold     = pfont_bold
           and tt_xml_font.font_italic   = pfont_italic
           and tt_xml_font.font_vertical = pfont_vertical
      exclusive-lock no-error.
      if not available tt_xml_font then do:
         create tt_xml_font.
         assign tt_xml_font.font_id       = gfont_id
                gfont_id                  = gfont_id + 1
                tt_xml_font.font_name     = pfont_name
                tt_xml_font.font_family   = pfont_family
                tt_xml_font.font_size     = pfont_size
                tt_xml_font.font_color    = pfont_color
                tt_xml_font.font_bold     = pfont_bold
                tt_xml_font.font_italic   = pfont_italic
                tt_xml_font.font_vertical = pfont_vertical.
         validate tt_xml_font.
      end.

      assign wfont_id = tt_xml_font.font_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wfont_id.
end function.

function getXMLIntID returns integer (input pint_color   as character,
                                      input pint_pattern as character):
   define variable wint_id as integer initial ? no-undo.
   define buffer tt_xml_interior for tt_xml_interior.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_interior
         where tt_xml_interior.int_color   = pint_color
           and tt_xml_interior.int_pattern = pint_pattern
      exclusive-lock no-error.
      if not available tt_xml_interior then do:
         create tt_xml_interior.
         assign tt_xml_interior.int_id      = gint_id
                gint_id                     = gint_id + 1
                tt_xml_interior.int_color   = pint_color
                tt_xml_interior.int_pattern = pint_pattern.
         validate tt_xml_interior.
      end.

      assign wint_id = tt_xml_interior.int_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wint_id.
end function.

function getXMLNumID returns integer (input pnum_format as character):
   define variable wnum_id as integer initial ? no-undo.
   define buffer tt_xml_numberformat for tt_xml_numberformat.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_numberformat
         where tt_xml_numberformat.num_format = pnum_format
      exclusive-lock no-error.
      if not available tt_xml_numberformat then do:
         create tt_xml_numberformat.
         assign tt_xml_numberformat.num_id     = gnum_id
                gnum_id                        = gnum_id + 1
                tt_xml_numberformat.num_format = pnum_format.
         validate tt_xml_numberformat.
      end.

      assign wnum_id = tt_xml_numberformat.num_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wnum_id.
end function.

function getXMLStyleID returns character (input pwb_no             as integer,
                                          input pnum_format        as character,
                                          input pint_color         as character,
                                          input pint_pattern       as character,
                                          input pfont_name         as character,
                                          input pfont_family       as character,
                                          input pfont_size         as integer,
                                          input pfont_color        as character,
                                          input pfont_bold         as integer,
                                          input pfont_italic       as integer,
                                          input pfont_vertical     as character,
                                          input palign_vertical    as character,
                                          input palign_horizontal  as character,
                                          input pborder_bottom     as character,
                                          input pborder_left       as character,
                                          input pborder_right      as character,
                                          input pborder_top        as character,
                                          input pprot_protected    as integer,
                                          input pprot_hideformula  as integer):
   define variable wnum_id   as integer no-undo.
   define variable wint_id   as integer no-undo.
   define variable wfont_id  as integer no-undo.
   define variable walign_id as integer no-undo.
   define variable wprot_id  as integer no-undo.
   define variable wbord_id  as integer no-undo.
   define variable wstyle_id as character initial ? no-undo.
   define buffer tt_xml_style for tt_xml_style.

   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      assign wnum_id   = getXMLNumID(pnum_format)
             wint_id   = getXMLIntID(pint_color, pint_pattern)
             wfont_id  = getXMLFontID(pfont_name, pfont_family, pfont_size, pfont_color, pfont_bold, pfont_italic, pfont_vertical)
             walign_id = getXMLAlignID(palign_vertical, palign_horizontal)
             wprot_id  = getXMLProtID(pprot_protected, pprot_hideformula)
             wbord_id  = getXMLBordID(pborder_bottom, pborder_left, pborder_right, pborder_top).

      if wnum_id = ? or wint_id = ? or wfont_id = ? or walign_id = ? or wprot_id = ? or wbord_id = ? then
         undo trans-blk, leave trans-blk.

      find first tt_xml_style
         where tt_xml_style.wb_no      = pwb_no
           and tt_xml_style.num_id     = wnum_id
           and tt_xml_style.int_id     = wint_id
           and tt_xml_style.font_id    = wfont_id
           and tt_xml_style.align_id   = walign_id
           and tt_xml_style.prot_id    = wprot_id
           and tt_xml_style.bord_id    = wbord_id
      exclusive-lock no-error.
      if not available tt_xml_style then do:
         create tt_xml_style.
         assign tt_xml_style.wb_no      = pwb_no
                tt_xml_style.style_id   = "s" + trim(string(gstyle_id,">>>>>>9"))
                gstyle_id               = gstyle_id + 1
                tt_xml_style.num_id     = wnum_id
                tt_xml_style.int_id     = wint_id
                tt_xml_style.font_id    = wfont_id
                tt_xml_style.align_id   = walign_id
                tt_xml_style.prot_id    = wprot_id
                tt_xml_style.bord_id    = wbord_id.
         validate tt_xml_style.
      end.

      assign wstyle_id = tt_xml_style.style_id.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wstyle_id.
end function.

function createXMLCell returns logical (input pwb_no            as integer,
                                        input pws_no            as integer,
                                        input pro_no            as integer,
                                        input pco_no            as integer,
                                        input pcolumn_width     as decimal,
                                        input pcell_type        as character,
                                        input pcell_value       as character,
                                        input pcell_formula     as character,
                                        input pauto_fit_height  as integer,
                                        input pauto_fit_width   as integer,
                                        input pnum_format       as character,
                                        input pint_color        as character,
                                        input pint_pattern      as character,
                                        input pfont_name        as character,
                                        input pfont_family      as character,
                                        input pfont_size        as integer,
                                        input pfont_color       as character,
                                        input pfont_bold        as integer,
                                        input pfont_italic      as integer,
                                        input pfont_vertical    as character,
                                        input palign_vertical   as character,
                                        input palign_horizontal as character,
                                        input pborder_bottom    as character,
                                        input pborder_left      as character,
                                        input pborder_right     as character,
                                        input pborder_top       as character,
                                        input pprot_protected   as integer,
                                        input pprot_hideformula as integer,
                                        input pinline_xml       as logical):
   define variable wtrans_ok as logical no-undo.
   define variable i         as integer no-undo.
   define variable j         as integer no-undo.
   define variable wstyle_id as character no-undo.
   define variable wtoday    as date no-undo.
   define variable wtime     as integer no-undo.
   define buffer tt_xml_table       for tt_xml_table.
   define buffer tt_xml_row         for tt_xml_row.
   define buffer tt_xml_column      for tt_xml_column.
   define buffer tt_xml_cell        for tt_xml_cell.
   define buffer tt_xml_data        for tt_xml_data.
   define buffer tt_xml_data_inline for tt_xml_data_inline.
   define buffer tt_xml_namedrange  for tt_xml_namedrange.
   define buffer tt_xml_namedcell   for tt_xml_namedcell.
   
   TRANS-BLK:
   repeat on error undo trans-blk, leave trans-blk
          on endkey undo trans-blk, leave trans-blk:

      find first tt_xml_table
         where tt_xml_table.wb_no = pwb_no
           and tt_xml_table.ws_no = pws_no
      exclusive-lock no-error.
      if not available tt_xml_table then 
         undo trans-blk, leave trans-blk.

      assign wstyle_id = getXMLStyleID(pwb_no,
                                       pnum_format, 
                                       pint_color, pint_pattern, 
                                       pfont_name, pfont_family, pfont_size, pfont_color, pfont_bold, pfont_italic, pfont_vertical,
                                       palign_vertical, palign_horizontal,
                                       pborder_bottom, pborder_left, pborder_right, pborder_top,
                                       pprot_protected, pprot_hideformula).

      find first tt_xml_column
         where tt_xml_column.wb_no = pwb_no
           and tt_xml_column.ws_no = pws_no
           and tt_xml_column.co_no = pco_no
      exclusive-lock no-error.
      if not available tt_xml_column then do:
         create tt_xml_column.
         assign tt_xml_column.wb_no = pwb_no
                tt_xml_column.ws_no = pws_no
                tt_xml_column.co_no = pco_no.
      end.

      find first tt_xml_row
         where tt_xml_row.wb_no = pwb_no
           and tt_xml_row.ws_no = pws_no
           and tt_xml_row.ro_no = pro_no
      exclusive-lock no-error.
      if not available tt_xml_row then do:
         create tt_xml_row.
         assign tt_xml_row.wb_no     = pwb_no
                tt_xml_row.ws_no     = pws_no
                tt_xml_row.ro_no     = pro_no.
      end.

      find first tt_xml_cell
         where tt_xml_cell.wb_no = pwb_no
           and tt_xml_cell.ws_no = pws_no
           and tt_xml_cell.ro_no = pro_no
           and tt_xml_cell.co_no = pco_no
      exclusive-lock no-error.
      if not available tt_xml_cell then do:
         create tt_xml_cell.
         assign tt_xml_cell.wb_no      = pwb_no
                tt_xml_cell.ws_no      = pws_no
                tt_xml_cell.ro_no      = pro_no
                tt_xml_cell.co_no      = pco_no
                tt_xml_cell.style_id   = wstyle_id.
      end.

      find first tt_xml_data
         where tt_xml_data.wb_no = pwb_no
           and tt_xml_data.ws_no = pws_no
           and tt_xml_data.ro_no = pro_no
           and tt_xml_data.co_no = pco_no
      exclusive-lock no-error.
      if not available tt_xml_data then do:
         create tt_xml_data.
         assign tt_xml_data.wb_no      = pwb_no
                tt_xml_data.ws_no      = pws_no
                tt_xml_data.ro_no      = pro_no
                tt_xml_data.co_no      = pco_no
                tt_xml_data.ns0        = ? when not pinline_xml
                tt_xml_data.inline_xml = pinline_xml.
      end.

      find first tt_xml_namedrange
         where tt_xml_namedrange.wb_no      = pwb_no
           and tt_xml_namedrange.ws_no      = pws_no
           and tt_xml_namedrange.range_name = "Print_Area"
      no-lock no-error.
      if available tt_xml_namedrange then do:
         find first tt_xml_namedcell
            where tt_xml_namedcell.wb_no = pwb_no
              and tt_xml_namedcell.ws_no = pws_no
              and tt_xml_namedcell.ro_no = pro_no
              and tt_xml_namedcell.co_no = pco_no
         exclusive-lock no-error.
         if not available tt_xml_namedcell then do:
            create tt_xml_namedcell.
            assign tt_xml_namedcell.wb_no    = pwb_no
                   tt_xml_namedcell.ws_no    = pws_no
                   tt_xml_namedcell.ro_no    = pro_no
                   tt_xml_namedcell.co_no    = pco_no.
         end.
         assign tt_xml_namedcell.cell_name = tt_xml_namedrange.range_name.
      end.

      assign tt_xml_column.column_width         = if tt_xml_column.column_width = ? then pcolumn_width 
                                                  else max(tt_xml_column.column_width,pcolumn_width)
             tt_xml_column.auto_fit_width       = pauto_fit_width
             tt_xml_row.auto_fit_height         = pauto_fit_height
             tt_xml_table.expanded_column_count = max(tt_xml_table.expanded_column_count,pco_no)
             tt_xml_table.expanded_row_count    = max(tt_xml_table.expanded_row_count,pro_no)
             tt_xml_data.cell_type              = pcell_type
             tt_xml_cell.style_id               = wstyle_id
             tt_xml_cell.formula                = pcell_formula.

      if tt_xml_data.inline_xml = true then do:
         find first tt_xml_data_inline
            where tt_xml_data_inline.wb_no = pwb_no
              and tt_xml_data_inline.ws_no = pws_no
              and tt_xml_data_inline.ro_no = pro_no
              and tt_xml_data_inline.co_no = pco_no
         exclusive-lock no-error.
         if not available tt_xml_data_inline then do:
            create tt_xml_data_inline.
            assign tt_xml_data_inline.wb_no = pwb_no
                   tt_xml_data_inline.ws_no = pws_no
                   tt_xml_data_inline.ro_no = pro_no
                   tt_xml_data_inline.co_no = pco_no.
         end.

         assign wtoday                        = today
                wtime                         = time
                tt_xml_data.cell_value        = "XMLHASH_" + string(year(wtoday),"9999") + string(month(wtoday),"99") + string(day(wtoday),"99") + string(wtime,"99999")
                                              + "_" + string(random(1,200000000),"999999999")
                tt_xml_data_inline.hash_value = tt_xml_data.cell_value
                tt_xml_data_inline.cell_value = pcell_value.
      end.
      else
         assign tt_xml_data.cell_value        = pcell_value.

      assign wtrans_ok = true.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
   return wtrans_ok.
end function.

function validFile returns logical (input f as character):
   if f = ? or f = "" then return false.
   else do:
      assign file-info:file-name = f.
      return not (file-info:file-type = ? or index(file-info:file-type,"F") <= 0 or index(file-info:file-type,"R") <= 0).
   end.
end function.

function getExcelVersion returns decimal:
   define variable vchExcel as com-handle no-undo.
   define variable wversion as decimal no-undo.

   create "Excel.Application" vchExcel no-error.
   if error-status:error then
      assign wversion = ?.
   else do:
      assign wversion = vchExcel:Version.
      release object vchExcel no-error.
      assign vchExcel = ?.
   end.
   return wversion.
end function.

function getExcelExtension returns character:
   define variable wver as decimal no-undo.
   assign wver = getExcelVersion().
   if wver = ? then return ?.
   else if wver >= {&XLS_EXCEL_2007} then return "xlsx".
   else return "xls".
end function.

function columnName returns character (input i as integer):
   define variable wp as character initial " ABCDEFGHIJKLMNOPQRSTUVWXYZ" no-undo.
   if i <= 0 or i >= 700 then return ?.
   assign i = i - 1.
   return trim(substring(wp,integer(truncate(i / 26,0)) + 1,1) 
             + substring(wp,(i mod 26) + 2,1)).
end function.

function columnNumber returns integer (input c as character):
   define variable i as integer no-undo.
   do i = 1 to 700:
      if columnName(i) = c then return i.
   end.
   return 0.
end function.

function createMerge returns logical (input pxls_no   as integer,
                                      input psheet_no as integer,
                                      input pmerge    as character):
   define buffer tt_excel_merge for tt_excel_merge.
   find first tt_excel_merge
      where tt_excel_merge.xls_no   = pxls_no
        and tt_excel_merge.sheet_no = psheet_no
        and tt_excel_merge.merge    = pmerge
   exclusive-lock no-error.
   if not available tt_excel_merge then do:
      create tt_excel_merge.
      assign tt_excel_merge.xls_no   = pxls_no
             tt_excel_merge.sheet_no = psheet_no
             tt_excel_merge.merge    = pmerge.
   end.
   return true.
end function.

function createCell returns logical (input pxls_no          as integer,
                                     input psheet_no        as integer,
                                     input prow_no          as integer,
                                     input pcolumn_no       as integer,
                                     input pcell_value      as character,
                                     input pcell_properties as character):
   define buffer tt_excel for tt_excel.
   find first tt_excel
      where tt_excel.xls_no    = pxls_no
        and tt_excel.sheet_no  = psheet_no
        and tt_excel.row_no    = prow_no
        and tt_excel.column_no = pcolumn_no
   exclusive-lock no-error.
   if not available tt_excel then do:
      create tt_excel.
      assign tt_excel.xls_no    = pxls_no
             tt_excel.sheet_no  = psheet_no
             tt_excel.row_no    = prow_no
             tt_excel.column_no = pcolumn_no.
             tt_excel.cell_name = columnName(tt_excel.column_no) + trim(string(tt_excel.row_no,">>>>>>9")).
   end.
   assign tt_excel.cell_value      = if pcell_value = ? then "" else pcell_value
          tt_excel.cell_properties = pcell_properties.
   return true.
end function.

function createFormula returns logical (input pxls_no          as integer,
                                        input psheet_no        as integer,
                                        input prow_no          as integer,
                                        input pcolumn_no       as integer,
                                        input pformula         as character,
                                        input pcell_properties as character):
   define buffer tt_excel for tt_excel.
   find first tt_excel
      where tt_excel.xls_no    = pxls_no
        and tt_excel.sheet_no  = psheet_no
        and tt_excel.row_no    = prow_no
        and tt_excel.column_no = pcolumn_no
   exclusive-lock no-error.
   if not available tt_excel then do:
      create tt_excel.
      assign tt_excel.xls_no    = pxls_no
             tt_excel.sheet_no  = psheet_no
             tt_excel.row_no    = prow_no
             tt_excel.column_no = pcolumn_no.
             tt_excel.cell_name = columnName(tt_excel.column_no) + trim(string(tt_excel.row_no,">>>>>>9")).
   end.
   assign tt_excel.cell_value      = if pformula = ? then "" else pformula
          tt_excel.cell_properties = pcell_properties
          tt_excel.is_formula      = true.
   return true.
end function.

function createSheet returns logical(pxls_no     as integer,
                                     psheet_no   as integer,
                                     psheet_name as character):
   define buffer tt_excel_sheet for tt_excel_sheet.
   find first tt_excel_sheet
      where tt_excel_sheet.xls_no   = pxls_no
        and tt_excel_sheet.sheet_no = psheet_no
   exclusive-lock no-error.
   if not available tt_excel_sheet then do:
      create tt_excel_sheet.
      assign tt_excel_sheet.xls_no     = pxls_no
             tt_excel_sheet.sheet_no   = psheet_no.
   end.
   assign tt_excel_sheet.sheet_name = psheet_name.
   return true.
end function.

function createSheetWithProperties returns logical(pxls_no           as integer,
                                                   psheet_no         as integer,
                                                   psheet_name       as character,
                                                   psheet_properties as character):
   define buffer tt_excel_sheet for tt_excel_sheet.
   find first tt_excel_sheet
      where tt_excel_sheet.xls_no   = pxls_no
        and tt_excel_sheet.sheet_no = psheet_no
   exclusive-lock no-error.
   if not available tt_excel_sheet then do:
      create tt_excel_sheet.
      assign tt_excel_sheet.xls_no     = pxls_no
             tt_excel_sheet.sheet_no   = psheet_no.
   end.
   assign tt_excel_sheet.sheet_name       = psheet_name
          tt_excel_sheet.sheet_properties = psheet_properties.
   return true.
end function.


function getProperty returns character (input pname as character,
                                        input plist as character):
   define variable i      as integer no-undo.
   define variable wvalue as character initial ? no-undo.

   ENTRY-LOOP:
   do i = 1 to num-entries(plist,"|"):
      if entry(1,entry(i,plist,"|"),"=") <> pname then next entry-loop.
      assign wvalue = entry(2,entry(i,plist,"|"),"=").
      leave entry-loop.
   end.  /* OF ENTRY-LOOP */
   return wvalue.
end function.

function setProperty returns character (input pname  as character,
                                        input pvalue as character,
                                        input plist  as character):
   define variable i    as integer no-undo.
   define variable fset as logical initial false no-undo.
   ENTRY-LOOP:
   do i = 1 to num-entries(plist,"|"):
      if entry(1,entry(i,plist,"|"),"=") <> pname then next entry-loop.
      assign entry(i,plist,"|") = pname + "=" + pvalue
             fset               = true.
      leave entry-loop.
   end.  /* OF ENTRY-LOOP */
   assign plist = (plist + "|" + pname + "=" + pvalue) when not fset.
   return plist.
end function.

function numSheets returns integer (input pxls_no as integer):
   define buffer tt_excel for tt_excel.
   find last tt_excel where tt_excel.xls_no = pxls_no no-lock no-error.
   return (if available tt_excel then tt_excel.sheet_no else 3).
end function.

function getXMLColumnWidth returns decimal (input pcell_value as character):
   return length(pcell_value) * 6.
end function.

function getColor returns integer (input pcolor as character):
   case pcolor:
      when "BLACK"       then return 1.
      when "GRAY"        then return 15.
      when "WHITE"       then return 2.
      when "RED"         then return 3.
      when "GREEN"       then return 4.
      when "LIGHTGREEN"  then return 35.
      when "BLUE"        then return 5.
      when "LIGHTBLUE"   then return 37.
      when "YELLOW"      then return 6.
      when "LIGHTYELLOW" then return 36.
      when "MAGENTA"     then return 7.
      when "CYAN"        then return 8.
      when "ORANGE"      then return 46.
      when "LIGHTORANGE" then return 40.
      otherwise return -4142.
   end case.
end function.

function getXMLColor returns character (input pcolor as character):
   case pcolor:
      when "BLACK"       then return "#000000".
      when "GRAY"        then return "#BFBFBF".
      when "WHITE"       then return "#FFFFFF".
      when "RED"         then return "#963634".
      when "LIGHTRED"    then return "#DA9694".
      when "GREEN"       then return "#76933C".
      when "LIGHTGREEN"  then return "#CCFFCC".
      when "BLUE"        then return "#538DD5".
      when "LIGHTBLUE"   then return "#8DB4E2".
      when "YELLOW"      then return "#FFFF00".
      when "LIGHTYELLOW" then return "#FFFF99".
      when "MAGENTA"     then return "#FF0066".
      when "CYAN"        then return "#66FFFF".
      when "ORANGE"      then return "#E26B0A".
      when "LIGHTORANGE" then return "#FABF8F".
      otherwise return ?.
   end case.
end function.

function getPaper returns integer (input ppaper as character):
   case ppaper:
      when "LETTER"    then return 1.
      when "TABLOID"   then return 3.
      when "LEDGER"    then return 4.
      when "LEGAL"     then return 5.
      when "STATEMENT" then return 6.
      when "EXECUTIVE" then return 7.
      when "A3"        then return 8.
      when "A4"        then return 9.
      when "A5"        then return 11.
      when "B4"        then return 12.
      when "B5"        then return 13.
      otherwise return 1.
   end case.
end function.

function createReference returns logical (input pref_hdl as com-handle):
   define variable ffound as logical initial no no-undo.
   define buffer tt_excel_handle for tt_excel_handle.

   find first tt_excel_handle
      where tt_excel_handle.ref_hdl = pref_hdl
   no-lock no-error.
   if not available tt_excel_handle then do:
      create tt_excel_handle.
      assign tt_excel_handle.ref_hdl = pref_hdl
             tt_excel_handle.ref_seq = gxls_ref_seq
             gxls_ref_seq            = gxls_ref_seq + 1
             ffound                  = true.
   end.
   return ffound.
end function.

function deleteReferences returns integer:
   define variable ncnt as integer no-undo.
   define buffer tt_excel_handle for tt_excel_handle.

   for each tt_excel_handle use-index idx_no 
   exclusive-lock:
      release object tt_excel_handle.ref_hdl no-error.
      delete tt_excel_handle.
      assign ncnt = ncnt + 1.
   end.
   return ncnt.
end function.

procedure outputXLS:
   define input parameter pfile_name as character no-undo.
   define input parameter pxls_no    like tt_excel.xls_no no-undo.
   define output parameter ptrans_ok as logical initial false no-undo.
   define output parameter perr_msg  as character initial "" no-undo.
   define variable vchExcel               as com-handle no-undo.
   define variable vchExcel2              as com-handle no-undo.
   define variable vchWorkBook            as com-handle no-undo.
   define variable vchWorkBooks           as com-handle no-undo.
   define variable vchSheet               as com-handle no-undo.
   define variable vchWorkSheets          as com-handle no-undo.
   define variable vchWorkSheet           as com-handle no-undo.
   define variable vchSheets              as com-handle no-undo.
   define variable vchRange               as com-handle no-undo.
   define variable vchFont                as com-handle no-undo.
   define variable vchBorder              as com-handle no-undo.
   define variable vchInterior            as com-handle no-undo.
   define variable vchCells               as com-handle no-undo.
   define variable vchEntireColumn        as com-handle no-undo.
   define variable vchPageSetup           as com-handle no-undo.
   define variable vchCenterHeaderPicture as com-handle no-undo.
   define variable wCenterHeaderPicture   as character no-undo.
   define variable vchLeftHeaderPicture   as com-handle no-undo.
   define variable wLeftHeaderPicture     as character no-undo.
   define variable vchRightHeaderPicture  as com-handle no-undo.
   define variable wRightHeaderPicture    as character no-undo.
   define variable i                      as integer no-undo.
   define buffer tt_excel       for tt_excel.
   define buffer tt_excel_sheet for tt_excel_sheet.
   define buffer tt_excel_merge for tt_excel_merge.

   TRANS-BLK:
   repeat on error undo trans-blk, retry trans-blk
          on endkey undo trans-blk, retry trans-blk:
      if retry then do:
         assign perr_msg = "Generic failure in outputXLS::TRANS-BLK".
         leave trans-blk.
      end.

      if not can-find(first tt_excel where tt_excel.xls_no = pxls_no) then do:
         assign perr_msg = "Excel records for file >" + trim(string(pxls_no,">>9")) + "< do not exist".
         undo trans-blk, leave trans-blk.
      end. 

      assign file-info:file-name = pfile_name.
      if file-info:file-type <> ? then
         os-delete value(file-info:full-pathname) no-error.

      /* PER MICROSOFT KB188546 THE EXCEL.EXE PROCESS SPAWNED BY THE INITIAL 'create "Excel.Application"' WILL BE USED BY 
         BOTH xls.i AND ANY SPREADSHEET THE USER OPENS DURING THE DURATION OF THE xls.i SPREADSHEET GENERATION.  THIS IS 
         CAUSING PROBLEMS WITH 'GHOST SPREADSHEETS' AND/OR MALFORMED SPREADSHEETS.  THE WORKAROUND IS TO CREATE A 
         'SACRIFICIAL ANODE' vchExcel2 EXCEL APPLICATION.  THIS SPAWNS TWO EXCEL.EXE PROCESSES WITH THE FIRST ONE BEING THE
         ONE THE USER WILL USE WHILE THE SECOND ONE BUILDS THE REPORT.
      */
      create "Excel.Application" vchExcel2 no-error.
      if not valid-handle(vchExcel2) then do:
         assign perr_msg = "Failed to open second Excel COM object.  Microsoft Excel does not exist on this system? (outputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.

      create "Excel.Application" vchExcel no-error.
      if not valid-handle(vchExcel) then do:
         assign perr_msg = "Failed to open Excel COM object.  Microsoft Excel does not exist on this system? (outputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.
      vchExcel:Visible            = false.
      vchExcel:DisplayAlerts      = false.

      vchWorkBooks = vchExcel:WorkBooks no-error.
      createReference(vchWorkBooks).

      vchWorkBook = vchWorkbooks:Add no-error.
      if not valid-handle(vchWorkBook) then do:
         assign perr_msg = "Failed to create Excel Workbook".
         undo trans-blk, leave trans-blk.
      end.
      createReference(vchWorkBook).

      vchWorkSheets = vchWorkBook:WorkSheets.
      createReference(vchWorkSheets).
      if vchWorkSheets:Count < numSheets(pxls_no) then
         do i = vchWorkSheets:Count to numSheets(pxls_no):
            vchWorkSheet = vchWorkSheets:Add().
            createReference(vchWorkSheet).
            vchWorkSheet:Activate.
         end.

      XLS-LOOP:
      for each tt_excel
         where tt_excel.xls_no = pxls_no
      no-lock
      break by tt_excel.xls_no
            by tt_excel.sheet_no
            by tt_excel.row_no
            by tt_excel.column_no
      on error undo xls-loop, retry xls-loop
      on endkey undo xls-loop, retry xls-loop:
         if retry then do:
            assign perr_msg = "Generic failure in outputXLS::XLS-LOOP".
            leave trans-blk.
         end.

         if first-of(tt_excel.sheet_no) then do:
            vchSheets = vchExcel:Sheets no-error.
            createReference(vchSheets).
            vchSheet  = vchExcel:Sheets:Item(tt_excel.sheet_no) no-error.
            if not valid-handle(vchSheet) then do:
               assign perr_msg = "Failed to find sheet number >" + trim(string(tt_excel.sheet_no,">>>9")) 
                               + "< for excel file number >" + trim(string(tt_excel.xls_no,">>>>>>9")) 
                               + "< (outputXLS::XLS-LOOP)".
               undo trans-blk, leave trans-blk.
            end.
            createReference(vchSheet).
            vchSheet:Select().
         end.

         vchRange = vchExcel:Range(tt_excel.cell_name) no-error.
         createReference(vchRange).

         if tt_excel.is_formula then 
            vchRange:Formula = tt_excel.cell_value.
         else
            vchRange:Value = tt_excel.cell_value.

         if getProperty("TEXT",tt_excel.cell_properties) = "TRUE" then
            vchRange:NumberFormat = "@".
         else if getProperty("FORMAT",tt_excel.cell_properties) <> ? then
            vchRange:NumberFormat = getProperty("FORMAT",tt_excel.cell_properties).
         if getProperty("BOLD",tt_excel.cell_properties) = "TRUE" then do:
            vchFont = vchRange:Font no-error.
            createReference(vchFont).
            vchFont:Bold = true.
         end.
         if getProperty("FONTSIZE",tt_excel.cell_properties) <> ? then do:
            vchFont = vchRange:Font no-error.
            createReference(vchFont).
            vchFont:Size = integer(getProperty("FONTSIZE",tt_excel.cell_properties)).
         end.
         if getProperty("UNDERLINE",tt_excel.cell_properties) = "TRUE" then do:
            vchBorder = vchRange:Borders({&XLS_BORDERBOTTOM}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEMEDIUM}.
         end.
         if getProperty("SUPERSCRIPT",tt_excel.cell_properties) = "TRUE" then do:
            vchFont = vchRange:Font no-error.
            createReference(vchFont).
            vchFont:Superscript = true.
         end.
         if getProperty("SUBSCRIPT",tt_excel.cell_properties) = "TRUE" then do:
            vchFont = vchRange:Font no-error.
            createReference(vchFont).
            vchFont:Subscript = true.
         end.
         if getProperty("CENTER",tt_excel.cell_properties) = "TRUE" then
            vchRange:HorizontalAlignment = {&XLS_ALIGNCENTER}.
         else if getProperty("RIGHTALIGN",tt_excel.cell_properties) = "TRUE" then
            vchRange:HorizontalAlignment = {&XLS_ALIGNRIGHT}.
         else if getProperty("LEFTALIGN",tt_excel.cell_properties) = "TRUE" then
            vchRange:HorizontalAlignment = {&XLS_ALIGNLEFT}.
         if getProperty("BGCOLOR",tt_excel.cell_properties) <> ? then do:
            vchInterior = vchRange:Interior no-error.
            createReference(vchInterior).
            vchInterior:ColorIndex = getColor(getProperty("BGCOLOR",tt_excel.cell_properties)).
         end.
         if getProperty("FGCOLOR",tt_excel.cell_properties) <> ? then do:
            vchFont = vchRange:Font no-error.
            createReference(vchFont).
            vchFont:ColorIndex = getColor(getProperty("FGCOLOR",tt_excel.cell_properties)).
         end.
         if getProperty("BOX",tt_excel.cell_properties) = "TRUE" then do:
            vchBorder = vchRange:Borders({&XLS_BORDERLEFT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEMEDIUM}.

            vchBorder = vchRange:Borders({&XLS_BORDERTOP}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEMEDIUM}.

            vchBorder = vchRange:Borders({&XLS_BORDERRIGHT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEMEDIUM}.

            vchBorder = vchRange:Borders({&XLS_BORDERBOTTOM}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEMEDIUM}.
         end.
         else if getProperty("LIGHTBOX",tt_excel.cell_properties) = "TRUE" then do:
            vchBorder = vchRange:Borders({&XLS_BORDERLEFT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINELIGHT}.

            vchBorder = vchRange:Borders({&XLS_BORDERTOP}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINELIGHT}.

            vchBorder = vchRange:Borders({&XLS_BORDERRIGHT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINELIGHT}.

            vchBorder = vchRange:Borders({&XLS_BORDERBOTTOM}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINELIGHT}.
         end.
         else if getProperty("HEAVYBOX",tt_excel.cell_properties) = "TRUE" then do:
            vchBorder = vchRange:Borders({&XLS_BORDERLEFT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEHEAVY}.

            vchBorder = vchRange:Borders({&XLS_BORDERTOP}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEHEAVY}.

            vchBorder = vchRange:Borders({&XLS_BORDERRIGHT}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEHEAVY}.

            vchBorder = vchRange:Borders({&XLS_BORDERBOTTOM}) no-error.
            createReference(vchBorder).
            vchBorder:LineStyle = {&XLS_LINECONTINUOUS}.
            vchBorder:Weight    = {&XLS_LINEHEAVY}.
         end.

         if last-of(tt_excel.sheet_no) then do:
            vchCells = vchSheet:Cells no-error.
            createReference(vchCells).

            vchRange = vchCells:Select().
            createReference(vchRange).

            vchEntireColumn = vchCells:EntireColumn no-error.
            createReference(vchEntireColumn).
            vchEntireColumn:AutoFit().

            MERGE-LOOP:
            for each tt_excel_merge
               where tt_excel_merge.xls_no   = tt_excel.xls_no
                 and tt_excel_merge.sheet_no = tt_excel.sheet_no
            no-lock
            on error undo merge-loop, retry merge-loop
            on endkey undo merge-loop, retry merge-loop:
               if retry then do:
                  assign perr_msg = "Generic failure in outputXLS::MERGE-LOOP".
                  leave trans-blk.
               end.
               vchRange = vchSheet:Range(tt_excel_merge.merge) no-error.
               createReference(vchRange).
               vchRange:Merge.
            end.  /* OF MERGE-LOOP */

            SHEET-LOOP:
            for first tt_excel_sheet
               where tt_excel_sheet.xls_no   = tt_excel.xls_no
                 and tt_excel_sheet.sheet_no = tt_excel.sheet_no
            no-lock
            on error undo sheet-loop, retry sheet-loop
            on endkey undo sheet-loop, retry sheet-loop:
               if retry then do:
                  assign perr_msg = "Generic failure in outputXLS::SHEET-LOOP".
                  leave trans-blk.
               end.
               assign vchSheet:Name = tt_excel_sheet.sheet_name.

               /* ORIENTATION/FITTOPAGE/PRINTAREA REQUIRE A VALID DEFAULT/ACTIVE PRINTER.  IF THERE ISN"T ONE --WHICH HAPPENS 
                  MORE OFTEN THAN NOT WITH VIRTUALIZED/CITRIX/RDP DESKTOPS-- THE CODE BELOW WILL FAIL WITH A 
                  'Failed to set Orientation' ERROR.  ADD A DEFAULT PRINTER (eg. Microsoft Office Document Writer) AND ALL WILL 
                  BE WELL IN THE WORLD OF EXCEL.

                  THE PAGESETUP OBJECT IS INCREDIBLY SLOW TO WRITE BUT MUCH FASTER TO READ SO READ THE PAGESETUP ATTRIBUTE TO
                  SEE IF IT NEEDS CHANGING BEFORE IT IS CHANGED.
               */
               if getProperty("LANDSCAPE",tt_excel_sheet.sheet_properties) = "TRUE" then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).
                  if vchPageSetup:Orientation <> {&XLS_LANDSCAPE} then
                     assign vchPageSetup:Orientation = {&XLS_LANDSCAPE}.
               end.
               else if getProperty("PORTRAIT",tt_excel_sheet.sheet_properties) = "TRUE" then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).
                  if vchPageSetup:Orientation <> {&XLS_PORTRAIT} then
                     assign vchPageSetup:Orientation = {&XLS_PORTRAIT}.
               end.
               if getProperty("FITTOPAGE",tt_excel_sheet.sheet_properties) <> ? then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).
                  assign vchPageSetup:Zoom           = false
                         vchPageSetup:FitToPagesWide = integer(getProperty("FITTOPAGE",tt_excel_sheet.sheet_properties))
                         vchPageSetup:FitToPagesTall = integer(getProperty("FITTOPAGE",tt_excel_sheet.sheet_properties)).
               end.
               if getProperty("PRINTAREA",tt_excel_sheet.sheet_properties) <> ? then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).
                  if vchPageSetup:PrintArea <> getProperty("PRINTAREA",tt_excel_sheet.sheet_properties) then
                     assign vchPageSetup:PrintArea = getProperty("PRINTAREA",tt_excel_sheet.sheet_properties).
               end.
               if getProperty("PAPER",tt_excel.cell_properties) <> ? then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).
                  if vchPageSetup:PaperSize <> getPaper(getProperty("PAPER",tt_excel.cell_properties)) then
                     assign vchPageSetup:PaperSize = getPaper(getProperty("PAPER",tt_excel.cell_properties)).
               end.

               assign wLeftHeaderPicture = getProperty("LEFTHEADERPICTURE",tt_excel_sheet.sheet_properties).
               if validFile(wLeftHeaderPicture) then do:
                  assign file-info:file-name = wLeftHeaderPicture
                         wLeftHeaderPicture  = file-info:full-pathname.
               end.
               else 
                  assign wLeftHeaderPicture = ?.

               if validFile(wLeftHeaderPicture) then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).

                  assign vchLeftHeaderPicture = vchPageSetup:LeftHeaderPicture no-error.
                  createReference(vchLeftHeaderPicture).

                  assign vchLeftHeaderPicture:FileName         = wLeftHeaderPicture
                         vchPageSetup:LeftHeader               = "&G"
                         vchPageSetup:ScaleWithDocHeaderFooter = true
                         vchPageSetup:AlignMarginsHeaderFooter = true.
               end.

               assign wCenterHeaderPicture = getProperty("CENTERHEADERPICTURE",tt_excel_sheet.sheet_properties).
               if validFile(wCenterHeaderPicture) then do:
                  assign file-info:file-name = wCenterHeaderPicture
                         wCenterHeaderPicture  = file-info:full-pathname.
               end.
               else 
                  assign wCenterHeaderPicture = ?.

               if validFile(wCenterHeaderPicture) then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).

                  assign vchCenterHeaderPicture = vchPageSetup:CenterHeaderPicture no-error.
                  createReference(vchCenterHeaderPicture).

                  assign vchCenterHeaderPicture:Filename       = wCenterHeaderPicture
                         vchPageSetup:CenterHeader             = "&G"
                         vchPageSetup:ScaleWithDocHeaderFooter = true
                         vchPageSetup:AlignMarginsHeaderFooter = true.
               end.

               assign wRightHeaderPicture = getProperty("RIGHTHEADERPICTURE",tt_excel_sheet.sheet_properties).
               if validFile(wRightHeaderPicture) then do:
                  assign file-info:file-name = wRightHeaderPicture
                         wRightHeaderPicture  = file-info:full-pathname.
               end.
               else 
                  assign wRightHeaderPicture = ?.

               if validFile(wRightHeaderPicture) then do:
                  assign vchPageSetup = vchSheet:PageSetup no-error.
                  createReference(vchPageSetup).

                  assign vchRightHeaderPicture = vchPageSetup:RightHeaderPicture no-error.
                  createReference(vchRightHeaderPicture).

                  assign vchRightHeaderPicture:Filename        = wRightHeaderPicture
                         vchPageSetup:RightHeader              = "&G"
                         vchPageSetup:ScaleWithDocHeaderFooter = true
                         vchPageSetup:AlignMarginsHeaderFooter = true.
               end.

               if getProperty("PASSWORD",tt_excel_sheet.sheet_properties) <> ? then do:
                  assign vchSheet:EnableSelection = 1.
                  vchSheet:Protect(getProperty("PASSWORD",tt_excel_sheet.sheet_properties),true,true,true).
               end.
            end.  /* OF SHEET-LOOP */

            vchRange = vchSheet:Range("A1").
            createReference(vchRange).
            vchRange:Select().
         end.
      end.  /* OF XLS-LOOP */

      vchWorkBook:SaveAs(pfile_name,,,,False,False,).
      vchWorkBook:Close().

      assign ptrans_ok = true.
      leave trans-blk.
   end.  /* OF TRANS-BLK */

   deleteReferences().

   assign vchSheet               = ?
          vchSheets              = ?
          vchWorkBook            = ?
          vchWorkBooks           = ?
          vchWorkSheet           = ?
          vchWorkSheets          = ?
          vchRange               = ?
          vchFont                = ?
          vchBorder              = ?
          vchInterior            = ?
          vchCells               = ?
          vchEntireColumn        = ?
          vchPageSetup           = ?
          vchCenterHeaderPicture = ?
          vchLeftHeaderPicture   = ?
          vchRightHeaderPicture  = ?.

   if valid-handle(vchExcel2) then do:
      vchExcel2:Quit().
      release object vchExcel2 no-error.
      assign vchExcel2 = ?.
   end.
   if valid-handle(vchExcel) then do:
      /* INVOKE THE Quit() METHOD *AFTER* ALL OTHER COM-HANDLE REFERENCES TO EXCEL HAVE BEEN RELEASED
         (PER PROGRESS KB18762)
       */
      vchExcel:Quit().
      release object vchExcel no-error.
      assign vchExcel = ?.
   end.
end procedure.

function getRow returns integer (input pcell as character):
   define variable i  as integer no-undo.
   do i = 1 to length(pcell):
      if _isInteger(substring(pcell,i)) then return integer(substring(pcell,i)). 
   end.
   return ?.
end function.

function getColumn returns character (input pcell as character):
   define variable i  as integer no-undo.
   do i = 1 to length(pcell):
      if _isInteger(substring(pcell,i)) then return replace(pcell,substring(pcell,i),""). 
   end.
   return ?.
end function.

procedure apply_merging:
   define output parameter ptrans_ok as logical initial false no-undo.
   define output parameter perr_msg  as character initial "" no-undo.
   define variable wcell    as character no-undo.
   define variable wco_from as integer no-undo.
   define variable wro_from as integer no-undo.
   define variable wco_to   as integer no-undo.
   define variable wro_to   as integer no-undo.
   define variable wro      as integer no-undo.
   define variable wco      as integer no-undo.
   define buffer tt_xml_row     for tt_xml_row.
   define buffer tt_xml_cell    for tt_xml_cell.
   define buffer bf_xml_cell    for tt_xml_cell.
   define buffer tt_xml_data    for tt_xml_data.
   define buffer tt_excel_merge for tt_excel_merge.
   define buffer tt_excel_sheet for tt_excel_sheet.

   TRANS-BLK:
   repeat on error undo trans-blk, retry trans-blk
          on endkey undo trans-blk, retry trans-blk:
      if retry then do:
         assign perr_msg = "General failure in apply_merging::TRANS-BLK".
         undo trans-blk, leave trans-blk.
      end.

      MERGE-LOOP:
      for each tt_excel_merge
      no-lock, first tt_excel_sheet
         where tt_excel_sheet.xls_no   = tt_excel_merge.xls_no
           and tt_excel_sheet.sheet_no = tt_excel_merge.sheet_no
      no-lock
      on error undo merge-loop, retry merge-loop
      on endkey undo merge-loop, retry merge-loop:
         if retry then do:
            assign perr_msg = "General failure in apply_merging::MERGE-LOOP".
            undo trans-blk, leave trans-blk.
         end.
         assign wcell    = entry(1,tt_excel_merge.merge,":")
                wco_from = columnNumber(getColumn(wcell))
                wro_from = getRow(wcell)
                wcell    = entry(2,tt_excel_merge.merge,":")
                wco_to   = columnNumber(getColumn(wcell))
                wro_to   = getRow(wcell).

         find first tt_xml_cell
            where tt_xml_cell.wb_no = tt_excel_sheet.wb_no
              and tt_xml_cell.ws_no = tt_excel_sheet.ws_no
              and tt_xml_cell.co_no = wco_from
              and tt_xml_cell.ro_no = wro_from
         exclusive-lock no-error.
         if not available tt_xml_cell then do:
            assign perr_msg = "Cell " + entry(1,tt_excel_merge.merge,":") + "not found in workbook/sheet " 
                            + trim(string(tt_excel_merge.xls_no,">>>>>>>>9")) + "/" + trim(string(tt_excel_merge.sheet_no,">>>>>>>>9"))  
                            + "[" + trim(string(tt_excel_sheet.wb_no,">>>>>>>>9")) + "." + trim(string(tt_excel_sheet.ws_no,">>>>>>>>9")) + "." 
                            + trim(string(wco_from,">>>>>>>>9")) + "." + trim(string(wro_from,">>>>>>>>9")) + "." 
                            + "] (apply_merging::MERGE-LOOP)".
            undo trans-blk, leave trans-blk.
         end.

         assign tt_xml_cell.merge_across = wco_to - wco_from when (wco_to <> wco_from) 
                tt_xml_cell.merge_down   = wro_to - wro_from when (wro_to <> wro_from).

         ROW-LOOP:
         for each tt_xml_row
            where tt_xml_row.wb_no  = tt_xml_cell.wb_no
              and tt_xml_row.ws_no  = tt_xml_cell.ws_no
              and tt_xml_row.ro_no  > wro_from
              and tt_xml_row.ro_no <= wro_to
         exclusive-lock
         on error undo row-loop, retry row-loop
         on endkey undo row-loop, retry row-loop:
            if retry then do:
               assign perr_msg = "General failure in apply_merging::ROW-LOOP".
               undo trans-blk, leave trans-blk.
            end.
            delete tt_xml_row.
         end.  /* OF ROW-LOOP */

         CELL-LOOP:
         for each bf_xml_cell
            where bf_xml_cell.wb_no  = tt_xml_cell.wb_no
              and bf_xml_cell.ws_no  = tt_xml_cell.ws_no
              and bf_xml_cell.ro_no >= wro_from
              and bf_xml_cell.ro_no <= wro_to
              and bf_xml_cell.co_no >= wco_from
              and bf_xml_cell.co_no <= wco_to
         exclusive-lock
         on error undo cell-loop, retry cell-loop
         on endkey undo cell-loop, retry cell-loop:
            if retry then do:
               assign perr_msg = "General failure in apply_merging::CELL-LOOP".
               undo trans-blk, leave trans-blk.
            end.
            if bf_xml_cell.ro_no = wro_from and bf_xml_cell.co_no = wco_from then next cell-loop.
            delete bf_xml_cell.
         end.  /* OF CELL-LOOP */
      end.  /* OF MERGE-LOOP */

      assign ptrans_ok = true.
      leave trans-blk.
   end.  /* OF TRANS-BLK */   
end procedure.

procedure backfill_cells:
   define output parameter ptrans_ok as logical initial false no-undo.
   define output parameter perr_msg  as character initial "" no-undo.
   define variable wco as integer no-undo.
   define variable wro as integer no-undo.
   define buffer tt_xml_table  for tt_xml_table.
   define buffer tt_xml_row    for tt_xml_row.
   define buffer tt_xml_column for tt_xml_column.
   define buffer tt_xml_cell   for tt_xml_cell.
   define buffer tt_xml_data   for tt_xml_data.

   TRANS-BLK:
   repeat on error undo trans-blk, retry trans-blk
          on endkey undo trans-blk, retry trans-blk:
      if retry then do:
         assign perr_msg = "General failure in backfill_cells::TRANS-BLK".
         undo trans-blk, leave trans-blk.
      end.

      TABLE-LOOP:
      for each tt_xml_table
      no-lock
      on error undo table-loop, retry table-loop
      on endkey undo table-loop, retry table-loop:
         if retry then do:
            assign perr_msg = "General failure in backfill_cells::TABLE-LOOP".
            undo trans-blk, leave trans-blk.
         end.

         COLUMN-LOOP:
         do wco = 1 to tt_xml_table.expanded_column_count
            on error undo column-loop, retry column-loop
            on endkey undo column-loop, retry column-loop:
            if retry then do:
               assign perr_msg = "General failure in backfill_cells::COLUMN-LOOP".
               undo trans-blk, leave trans-blk.
            end.

            find first tt_xml_column
               where tt_xml_column.wb_no = tt_xml_table.wb_no
                 and tt_xml_column.ws_no = tt_xml_table.ws_no
                 and tt_xml_column.co_no = wco
            exclusive-lock no-error.
            if not available tt_xml_column then do:
               create tt_xml_column.
               assign tt_xml_column.wb_no = tt_xml_table.wb_no
                      tt_xml_column.ws_no = tt_xml_table.ws_no
                      tt_xml_column.co_no = wco.
            end.
         end.  /* OF COLUMN-LOOP */

         ROW-LOOP:
         do wro = 1 to tt_xml_table.expanded_row_count
            on error undo row-loop, retry row-loop
            on endkey undo row-loop, retry row-loop:
            if retry then do:
               assign perr_msg = "General failure in backfill_cells::ROW-LOOP".
               undo trans-blk, leave trans-blk.
            end.
            
            find first tt_xml_row
               where tt_xml_row.wb_no = tt_xml_table.wb_no
                 and tt_xml_row.ws_no = tt_xml_table.ws_no
                 and tt_xml_row.ro_no = wro
            exclusive-lock no-error.
            if not available tt_xml_row then do:
               create tt_xml_row.
               assign tt_xml_row.wb_no = tt_xml_table.wb_no
                      tt_xml_row.ws_no = tt_xml_table.ws_no
                      tt_xml_row.ro_no = wro.
            end.
         end.  /* OF ROW-LOOP */

/* NOT REQUIRED WHEN COLUMNS, ROWS, CELLS HAVE INDEX ATTRIBUTE SET.  PERFORMANCE BOOST NOT HAVING TO BACKFILL EVERY CELL
   OF EVERY SHEET.

         RO-LOOP:
         do wro = 1 to tt_xml_table.expanded_row_count
            on error undo ro-loop, retry ro-loop
            on endkey undo ro-loop, retry ro-loop:
            if retry then do:
               assign perr_msg = "General failure in backfill_cells::RO-LOOP".
               undo trans-blk, leave trans-blk.
            end.

            CO-LOOP:
            do wco = 1 to tt_xml_table.expanded_column_count
               on error undo co-loop, retry co-loop
               on endkey undo co-loop, retry co-loop:
               if retry then do:
                  assign perr_msg = "General failure in backfill_cells::CO-LOOP".
                  undo trans-blk, leave trans-blk.
               end.

               find first tt_xml_cell
                  where tt_xml_cell.wb_no = tt_xml_table.wb_no
                    and tt_xml_cell.ws_no = tt_xml_table.ws_no
                    and tt_xml_cell.ro_no = wro
                    and tt_xml_cell.co_no = wco
               exclusive-lock no-error.
               if not available tt_xml_cell then do:
                  create tt_xml_cell.
                  assign tt_xml_cell.wb_no    = tt_xml_table.wb_no
                         tt_xml_cell.ws_no    = tt_xml_table.ws_no
                         tt_xml_cell.ro_no    = wro
                         tt_xml_cell.co_no    = wco
                         tt_xml_cell.style_id = "Default".
               end.
            end.  /* OF RO-LOOP */
         end.  /* OF CO-LOOP */
*/
      end.  /* OF TABLE-LOOP */ 

      assign ptrans_ok = true.
      leave trans-blk.
   end.  /* OF TRANS-BLK */   
end procedure.

function formatXMLNumber returns character (input pcell_value as character):
   return trim(replace(replace(pcell_value,"$",""),",","")).
end function.

function formatXMLDate returns character (input pcell_value as character):
   return pcell_value.
end function.

function clearXMLTables returns logical:
   define buffer tt_xml_workbook               for tt_xml_workbook.
   define buffer tt_xml_documentproperties     for tt_xml_documentproperties.
   define buffer tt_xml_officedocumentsettings for tt_xml_officedocumentsettings.
   define buffer tt_xml_excelworkbook          for tt_xml_excelworkbook.
   define buffer tt_xml_styles                 for tt_xml_styles.
   define buffer tt_xml_style                  for tt_xml_style.
   define buffer tt_xml_alignment              for tt_xml_alignment.
   define buffer tt_xml_borders                for tt_xml_borders.
   define buffer tt_xml_border                 for tt_xml_border.
   define buffer tt_xml_font                   for tt_xml_font.
   define buffer tt_xml_interior               for tt_xml_interior.
   define buffer tt_xml_numberformat           for tt_xml_numberformat.
   define buffer tt_xml_protection             for tt_xml_protection.
   define buffer tt_xml_worksheet              for tt_xml_worksheet.
   define buffer tt_xml_names                  for tt_xml_names.
   define buffer tt_xml_namedrange             for tt_xml_namedrange.
   define buffer tt_xml_table                  for tt_xml_table.
   define buffer tt_xml_column                 for tt_xml_column.
   define buffer tt_xml_row                    for tt_xml_row.
   define buffer tt_xml_cell                   for tt_xml_cell.
   define buffer tt_xml_data                   for tt_xml_data.
   define buffer tt_xml_namedcell              for tt_xml_namedcell.
   define buffer tt_xml_worksheetoptions       for tt_xml_worksheetoptions.
   define buffer tt_xml_pagesetup              for tt_xml_pagesetup.
   define buffer tt_xml_layout                 for tt_xml_layout.
   define buffer tt_xml_header                 for tt_xml_header.
   define buffer tt_xml_footer                 for tt_xml_footer.
   define buffer tt_xml_pagemargins            for tt_xml_pagemargins.
   define buffer tt_xml_print                  for tt_xml_print.
   define buffer tt_xml_panes                  for tt_xml_panes.
   define buffer tt_xml_pane                   for tt_xml_pane.

   empty temp-table tt_xml_workbook.
   empty temp-table tt_xml_documentproperties.
   empty temp-table tt_xml_officedocumentsettings.
   empty temp-table tt_xml_excelworkbook.
   empty temp-table tt_xml_styles.
   empty temp-table tt_xml_style.
   empty temp-table tt_xml_alignment.
   empty temp-table tt_xml_borders.
   empty temp-table tt_xml_border.
   empty temp-table tt_xml_font.
   empty temp-table tt_xml_interior.
   empty temp-table tt_xml_numberformat.
   empty temp-table tt_xml_protection.
   empty temp-table tt_xml_worksheet.
   empty temp-table tt_xml_names.
   empty temp-table tt_xml_namedrange.
   empty temp-table tt_xml_table.
   empty temp-table tt_xml_column.
   empty temp-table tt_xml_row.
   empty temp-table tt_xml_cell.
   empty temp-table tt_xml_data.
   empty temp-table tt_xml_namedcell.
   empty temp-table tt_xml_worksheetoptions.
   empty temp-table tt_xml_pagesetup.
   empty temp-table tt_xml_layout.
   empty temp-table tt_xml_header.
   empty temp-table tt_xml_footer.
   empty temp-table tt_xml_pagemargins.
   empty temp-table tt_xml_print.
   empty temp-table tt_xml_panes.
   empty temp-table tt_xml_pane.

   return true.
end function.

procedure outputXML:
   define input parameter pfile_name as character no-undo.
   define input parameter pxls_no    as integer no-undo.
   define output parameter ptrans_ok as logical initial false no-undo.
   define output parameter perr_msg  as character initial "" no-undo.
   define variable wtemp_file        as character no-undo.
   define variable wraw_line         as character no-undo.
   define variable i                 as integer no-undo.
   define variable wwb_no            as integer no-undo.
   define variable wws_no            as integer no-undo.
   define variable fok               as logical no-undo.
   define variable wsheet_name       as character no-undo.
   define variable wuser_name        as character no-undo.
   define variable worientation      as character no-undo.
   define variable wfont_bold        as integer no-undo.
   define variable wfont_italic      as integer no-undo.
   define variable wfont_vertical    as character no-undo.
   define variable wint_pattern      as character no-undo.
   define variable wcell_type        as character no-undo.
   define variable wtrans_ok         as logical no-undo.
   define variable werr_msg          as character no-undo.
   define variable wcell_value       as character no-undo.
   define variable wcell_formula     as character no-undo.
   define variable wborder_bottom    as character no-undo.
   define variable wborder_left      as character no-undo.
   define variable wborder_right     as character no-undo.
   define variable wborder_top       as character no-undo.
   define variable walign_horizontal as character no-undo.
   define variable wno_grid_lines    as logical no-undo.
   define variable wcolumn_width     as decimal no-undo.
   define variable widx              as integer no-undo.
   define variable whash             as character no-undo.
   define buffer tt_excel           for tt_excel.
   define buffer tt_excel_sheet     for tt_excel_sheet.
   define buffer tt_xml_data_inline for tt_xml_data_inline.

   TRANS-BLK:
   repeat on error undo trans-blk, retry trans-blk
          on endkey undo trans-blk, retry trans-blk:
      if retry then do:
         assign perr_msg = "General failure in outputXML::TRANS-BLK".
         undo trans-blk, leave trans-blk.
      end.

      assign wtemp_file = session:temp-directory 
                        + "xmlgen." 
                        + _getUserName()
                        + string(year(today),"9999")
                        + string(month(today),"99")
                        + string(day(today),"99")
                        + string(time,"99999")
                        + ".xml"
             i          = 0
             wuser_name = _getUserName().

      CELL-LOOP:
      for each tt_excel
         where tt_excel.xls_no = pxls_no
      no-lock
      break by tt_excel.xls_no
            by tt_excel.sheet_no
            by tt_excel.row_no
            by tt_excel.column_no
      on error undo cell-loop, retry cell-loop
      on endkey undo cell-loop, retry cell-loop:
         if retry then do:
            assign perr_msg = "Generic failure in outputXML::CELL-LOOP".
            leave trans-blk.
         end.

         if first-of(tt_excel.xls_no) then do:
            assign wwb_no = createXMLWorkbook(wuser_name,                 /* AUTHOR */
                                              "",                         /* LAST AUTHOR */
                                              now,                        /* CREATED */
                                              now,                        /* LAST SAVED */
                                              "",                         /* COMPANY */
                                              "14.00",                    /* EXCEL VERSION */
                                              10110,                      /* WINDOW HEIGHT */
                                              26700,                      /* WINDOW WIDTH */
                                              240,                        /* WINDOW TOPX */
                                              45,                         /* WINDOW TOPY */
                                              false,                      /* PROTECT STRUCTURE */
                                              false,                      /* PROTECT WINDOWS */
                                              {&XML_DEFAULT_FONT},        /* DEFAULT FONT */
                                              {&XML_DEFAULT_FONT_FAMILY}, /* DEFAULT FONT FAMILY */
                                              {&XML_DEFAULT_FONT_SIZE},   /* DEFAULT FONT SIZE */
                                              {&XML_DEFAULT_FONT_COLOR}). /* DEFAULT FONT COLOR */
            if wwb_no = ? then do:
               assign perr_msg = "Failed call to outputXML:createXMLWorkbook()".
               undo trans-blk, leave trans-blk.
            end.
         end.

         if first-of(tt_excel.sheet_no) then do:
            find first tt_excel_sheet
               where tt_excel_sheet.xls_no   = tt_excel.xls_no
                 and tt_excel_sheet.sheet_no = tt_excel.sheet_no
            exclusive-lock no-error.

            assign wsheet_name    = if available tt_excel_sheet then tt_excel_sheet.sheet_name
                                    else ("Sheet" + trim(string(tt_excel.sheet_no,">>>>>>>>9")))
                   worientation   = if getProperty("LANDSCAPE",tt_excel_sheet.sheet_properties) <> "TRUE" then "Portrait"
                                    else "Landscape"
                   wno_grid_lines = if getProperty("GRIDLINES",tt_excel_sheet.sheet_properties) = "FALSE" then true
                                    else false
                   wws_no         = createXMLWorksheet(wwb_no,                            /* WORKBOOK */
                                                       wsheet_name,                       /* SHEET NAME */
                                                       wno_grid_lines,                    /* NO GRID LINES */
                                                       false,                             /* PROTECT OBJECTS */
                                                       false,                             /* PROTECT SCENARIOS */
                                                       "",                                /* FIT TO PAGE */
                                                       0,                                 /* FIT WIDTH */
                                                       0,                                 /* FIT HEIGHT */
                                                       "",                                /* IS SELECTED */
                                                       {&XML_DEFAULT_ROW_HEIGHT},         /* DEFAULT ROW HEIGHT */
                                                       {&XML_DEFAULT_HEADER_MARGIN},      /* HEADER MARGIN */
                                                       {&XML_DEFAULT_FOOTER_MARGIN},      /* FOOTER MARGIN */
                                                       {&XML_DEFAULT_PAGE_MARGIN_BOTTOM}, /* PAGE MARGIN BOTTOM */
                                                       {&XML_DEFAULT_PAGE_MARGIN_LEFT},   /* PAGE MARGIN LEFT */
                                                       {&XML_DEFAULT_PAGE_MARGIN_RIGHT},  /* PAGE MARGIN RIGHT */
                                                       {&XML_DEFAULT_PAGE_MARGIN_TOP},    /* PAGE MARGIN TOP */
                                                       worientation,                      /* ORIENTATION */
                                                       ?).                                /* PRINT AREA */
            if wws_no = ? then do:
               assign perr_msg = "Failed call to outputXML:createXMLWorksheet()".
               undo trans-blk, leave trans-blk.
            end.

            /* LINK BETWEEN xls_no/sheet_no and wb_no/ws_no REQUIRED FOR MERGING CELLS.
               CREATE ONE IF IT DOESN'T ALREADY EXIST.
             */
            if not available tt_excel_sheet then do:
               create tt_excel_sheet.
               assign tt_excel_sheet.xls_no     = tt_excel.xls_no
                      tt_excel_sheet.sheet_no   = tt_excel.sheet_no
                      tt_excel_sheet.sheet_name = wsheet_name.
            end.
            assign tt_excel_sheet.wb_no = wwb_no
                   tt_excel_sheet.ws_no = wws_no.
         end.

         assign wcell_type        = if tt_excel.is_formula then "String" 
                                    else if getProperty("NUMBER",tt_excel.cell_properties) = "TRUE" then "Number"
                                    else if getProperty("DATETIME",tt_excel.cell_properties) = "TRUE" then "DateTime"
                                    else "String"
                wcell_value       = if tt_excel.is_formula then ?
                                    else if wcell_type = "Number" then formatXMLNumber(tt_excel.cell_value)
                                    else if wcell_type = "DateTime" then formatXMLDate(tt_excel.cell_value)
                                    else tt_excel.cell_value
                wcell_formula     = if tt_excel.is_formula then tt_excel.cell_value
                                    else ?
                wcolumn_width     = if getProperty("COLUMNWIDTH",tt_excel.cell_properties) = ? then getXMLColumnWidth(tt_excel.cell_value)
                                    else decimal(getProperty("COLUMNWIDTH",tt_excel.cell_properties))
                wfont_bold        = if getProperty("BOLD",tt_excel.cell_properties) = "TRUE" then 1 else 0
                wfont_italic      = if getProperty("ITALIC",tt_excel.cell_properties) = "TRUE" then 1 else 0
                wfont_vertical    = if getProperty("SUPERSCRIPT",tt_excel.cell_properties) = "TRUE" then "Superscript"
                                    else if getProperty("SUBSCRIPT",tt_excel.cell_properties) = "TRUE" then "Subscript"
                                    else ?
                wint_pattern      = if getProperty("BGCOLOR",tt_excel.cell_properties) = ? then ? else "Solid"
                walign_horizontal = if getProperty("CENTER",tt_excel.cell_properties) = "TRUE" then "Center"
                                    else if getProperty("RIGHTALIGN",tt_excel.cell_properties) = "TRUE" then "Right"
                                    else if getProperty("LEFTALIGN",tt_excel.cell_properties) = "TRUE" then "Left"
                                    else ?
                wborder_bottom    = if getProperty("UNDERLINE",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("LIGHTBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("BOX",tt_excel.cell_properties) = "TRUE" then "Continuous|2"
                                    else if getProperty("HEAVYBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|3"
                                    else ?
                wborder_left      = if getProperty("LIGHTBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("BOX",tt_excel.cell_properties) = "TRUE" then "Continuous|2"
                                    else if getProperty("HEAVYBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|3"
                                    else ?
                wborder_right     = if getProperty("LIGHTBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("BOX",tt_excel.cell_properties) = "TRUE" then "Continuous|2"
                                    else if getProperty("HEAVYBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|3"
                                    else ?
                wborder_top       = if getProperty("OVERLINE",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("LIGHTBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|1"
                                    else if getProperty("BOX",tt_excel.cell_properties) = "TRUE" then "Continuous|2"
                                    else if getProperty("HEAVYBOX",tt_excel.cell_properties) = "TRUE" then "Continuous|3"
                                    else ?
                fok               = createXMLCell(wwb_no,                                                       /* WORKBOOK */
                                                  wws_no,                                                       /* WORKSHEET */
                                                  tt_excel.row_no,                                              /* ROW */
                                                  tt_excel.column_no,                                           /* COLUMN */
                                                  wcolumn_width,                                                /* COLUMN WIDTH */
                                                  wcell_type,                                                   /* CELL TYPE */
                                                  wcell_value,                                                  /* CELL VALUE */
                                                  wcell_formula,                                                /* CELL FORMULA */
                                                  0,                                                            /* AUTO FIT HEIGHT */
                                                  1,                                                            /* AUTO FIT WIDTH */
                                                  getProperty("FORMAT",tt_excel.cell_properties),               /* NUM FORMAT */
                                                  getXMLColor(getProperty("BGCOLOR",tt_excel.cell_properties)), /* INTERIOR COLOR */
                                                  wint_pattern,                                                 /* INTERIOR PATTERN */
                                                  "Calibri",                                                    /* FONT NAME */
                                                  "Swiss",                                                      /* FONT FAMILY */
                                                  11,                                                           /* FONT SIZE */
                                                  getXMLColor(getProperty("FGCOLOR",tt_excel.cell_properties)), /* FONT COLOR */
                                                  wfont_bold,                                                   /* FONT BOLD */
                                                  wfont_italic,                                                 /* FONT ITALIC */
                                                  wfont_vertical,                                               /* FONT VERTICAL ALIGNMENT */
                                                  "Bottom",                                                     /* ALIGNMENT VERTICAL */
                                                  walign_horizontal,                                            /* ALIGNMENT HORIZONTAL */
                                                  wborder_bottom,                                               /* BORDER BOTTOM */
                                                  wborder_left,                                                 /* BORDER LEFT */
                                                  wborder_right,                                                /* BORDER RIGHT */
                                                  wborder_top,                                                  /* BORDER TOP */
                                                  ?,                                                            /* PROTECTION PROTECTED */
                                                  ?,                                                            /* PROTECTION HIDE FORMULA */
                                                  wcell_value matches "*</*").                                  /* INLINE XML */
         if not fok then do:
            assign perr_msg = "Failed call to outputXML:createXMLCell()".
            undo trans-blk, leave trans-blk.
         end.
      end.  /* OF CELL-LOOP */

      run backfill_cells(output wtrans_ok,
                         output werr_msg).
      if not wtrans_ok then do:
         assign perr_msg = werr_msg.
         undo trans-blk, leave trans-blk.
      end.

      run apply_merging(output wtrans_ok,
                        output werr_msg).
      if not wtrans_ok then do:
         assign perr_msg = werr_msg.
         undo trans-blk, leave trans-blk.
      end.

      dataset ds_xml:write-xml("FILE",wtemp_file,true).

      input stream sxmlin from value(wtemp_file).
      output stream sxmlout to value(pfile_name).      
      FILE-LOOP:
      repeat:
         import stream sxmlin unformatted wraw_line.
         assign i    = i + 1.

         if i = 2 then
            put stream sxmlout unformatted "<?mso-application progid='Excel.Sheet'?>" skip.
         if not wraw_line matches "*ds_xml*" then do:
            assign widx = index(wraw_line,"XMLHASH_").

            if widx <> 0 then do:
               assign whash = substring(wraw_line,widx,31).
               find first tt_xml_data_inline
                  where tt_xml_data_inline.hash_value = whash
               no-lock no-error.
               if available tt_xml_data_inline then
                  assign wraw_line = replace(wraw_line,whash,tt_xml_data_inline.cell_value).
            end.

            put stream sxmlout unformatted wraw_line skip.
         end.
      end.
      output stream sxmlout close.
      input stream sxmlin close.

      os-delete value(wtemp_file).

      clearXMLTables().

      assign ptrans_ok = true.
      leave trans-blk.
   end.  /* OF TRANS-BLK */
end procedure.

procedure inputXLS:
   define input parameter pfile_name as character no-undo.
   define input parameter pcolumns   as integer no-undo.
   define input parameter prows      as integer no-undo.
   define output parameter pxls_no   as integer no-undo.
   define output parameter ptrans_ok as logical initial false no-undo.
   define output parameter perr_msg  as character initial "" no-undo.
   define variable vchExcel     as com-handle no-undo.
   define variable vchWorkBook  as com-handle no-undo.
   define variable vchWorkBooks as com-handle no-undo.
   define variable vchSheet     as com-handle no-undo.
   define variable vchSheets    as com-handle no-undo.
   define variable vchUsedRange as com-handle no-undo.
   define variable vchRows      as com-handle no-undo.
   define variable vchColumns   as com-handle no-undo.
   define variable vchRange     as com-handle no-undo.
   define variable s        as integer no-undo. /* SHEET NUMBER */
   define variable r        as integer no-undo. /* ROW NUMBER */
   define variable c        as integer no-undo. /* COLUMN NUMBER */
   define variable nsheets  as integer no-undo.
   define variable nrows    as integer no-undo.
   define variable ncolumns as integer no-undo.
   define buffer tt_excel       for tt_excel.
   define buffer tt_excel_sheet for tt_excel_sheet.

   TRANS-BLK:
   repeat on error undo trans-blk, retry trans-blk
          on endkey undo trans-blk, retry trans-blk:
      if retry then do:
         assign perr_msg = "Generic failure in inputXLS::TRANS-BLK".
         leave trans-blk.
      end.

      if lookup(entry(num-entries(pfile_name,"."),pfile_name,"."),{&XLS_EXT_LIST}) = 0 then do:
         assign perr_msg = pfile_name + " is not an Excel spreadsheet (inputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.

      assign file-info:file-name = pfile_name.
      if file-info:file-type = ? then do:
         assign perr_msg = pfile_name + " does not exist (inputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.
      else if index(file-info:file-type,"F") <= 0 then do:
         assign perr_msg = pfile_name + " is not a file (inputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.
      else if index(file-info:file-type,"R") <= 0 then do:
         assign perr_msg = pfile_name + " is not readable (inputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.

      create "Excel.Application" vchExcel no-error.
      if not valid-handle(vchExcel) then do:
         assign perr_msg = "Failed to open Excel COM object.  Microsoft Excel does not exist on this system? (inputXLS::TRANS-BLK)".
         undo trans-blk, leave trans-blk.
      end.
      vchExcel:Visible = false.

      vchWorkBooks = vchExcel:WorkBooks no-error.
      createReference(vchWorkBooks).

      vchWorkBook = vchWorkBooks:Open(file-info:full-pathname) no-error.
      createReference(vchWorkBook).

      vchSheets = vchExcel:Sheets no-error.
      createReference(vchSheets).

      assign nsheets = vchSheets:Count
             gxls_no = gxls_no + 1.

      SHEET-LOOP:
      repeat s = 1 to nsheets
         on error undo sheet-loop, retry sheet-loop
         on endkey undo sheet-loop, retry sheet-loop:
         if retry then do:
            assign perr_msg = "Generic failure in inputXLS::SHEET-LOOP".
            undo trans-blk, leave trans-blk.
         end.

         vchSheet = vchExcel:Sheets:Item(s):Select no-error.
         createReference(vchSheet).

         vchSheet = vchExcel:Sheets:Item(s) no-error.
         createReference(vchSheet).

         vchUsedRange = vchSheet:UsedRange no-error.
         createReference(vchUsedRange).

         vchRows = vchUsedRange:Rows no-error.
         createReference(vchRows).

         vchColumns = vchUsedRange:Columns no-error.
         createReference(vchColumns).

         assign nrows    = if (prows = 0) then vchRows:Count else prows
                ncolumns = if (pcolumns = 0) then vchColumns:Count else pcolumns.

         find first tt_excel_sheet
            where tt_excel_sheet.xls_no   = gxls_no
              and tt_excel_sheet.sheet_no = s
         exclusive-lock no-error.
         if not available tt_excel_sheet then do:
            create tt_excel_sheet.
            assign tt_excel_sheet.xls_no     = gxls_no
                   tt_excel_sheet.sheet_no   = s
                   tt_excel_sheet.sheet_name = vchSheet:Name.
         end.

         ROW-LOOP:
         do r = 1 to nrows:
            COLUMN-LOOP:
            repeat c = 1 to ncolumns
               on error undo column-loop, retry column-loop
               on endkey undo column-loop, retry column-loop:
               if retry then do:
                  assign perr_msg = "Generic failure in inputXLS::COLUMN-LOOP".
                  undo trans-blk, leave trans-blk.
               end.

               create tt_excel.
               assign tt_excel.xls_no     = tt_excel_sheet.xls_no
                      tt_excel.sheet_no   = tt_excel_sheet.sheet_no
                      tt_excel.row_no     = r
                      tt_excel.column_no  = c.
                      tt_excel.cell_name  = columnName(tt_excel.column_no) 
                                          + trim(string(tt_excel.row_no,">>>>>>>9")).

               vchRange = vchSheet:Range(tt_excel.cell_name) no-error.
               createReference(vchRange).

               assign tt_excel.cell_value = vchRange:Value no-error.
               if error-status:error then do:
                  assign perr_msg = "Failed to get value at range >" + tt_excel.cell_name + "< (inputXLS::COLUMN-LOOP)".
                  undo trans-blk, leave trans-blk.
               end.
            end.  /* OF COLUMN-LOOP */
         end.  /* OF ROW-LOOP */
      end.  /* SHEET-LOOP */

      vchExcel:DisplayAlerts = false.

      assign ptrans_ok = true
             pxls_no   = gxls_no.
      leave trans-blk.
   end.  /* OF TRANS-BLK */

   deleteReferences().

   assign vchWorkBook  = ?
          vchWorkBooks = ?
          vchSheet     = ?
          vchSheets    = ?
          vchUsedRange = ?
          vchRows      = ?
          vchColumns   = ?
          vchRange     = ?.

   if valid-handle(vchExcel) then do:
      /* INVOKE THE Quit() METHOD *AFTER* ALL OTHER COM-HANDLE REFERENCES TO EXCEL HAVE BEEN RELEASED
         (PER PROGRESS KB18762)
       */
      vchExcel:Quit().
      release object vchExcel no-error.
      assign vchExcel = ?.
   end.
end procedure.
&ENDIF
