/***************************************************************************
    File          : calcUtils.i
    Purpose       : Procedures for use with Open Office Calc (spreadsheet) programming

    Syntax        : { appl/utils/oo-calcutils.i }

    Author(s)     : Patrick Hulst
    Created       : 2004-03-18

    Usage (REALLY basic idea):
        Start Calc:                     RUN open_calc.
        Choice: Open an existing book:  RUN open_book(filename).
                Create a new book:      RUN new_book.
        Column Headers:                 RUN write_col_hdr (column, "Label").
        Column Data:                    RUN write_cell_data (column, row, "Data").
        Save the Sheet:                 RUN save_book ("Directory", "Sheet Name").
        Close Calc:                     RUN close_calc. (optional!)
        Clean up the COM-HANDLES        RUN CleanUp.

***************************************************************************/

/*********** Define Global Variables ***********/
DEFINE VARIABLE chOpenOffice    AS COM-HANDLE   NO-UNDO.
DEFINE VARIABLE chWorkBook      AS COM-HANDLE   NO-UNDO.
DEFINE VARIABLE chDesktop       AS COM-HANDLE   NO-UNDO.
DEFINE VARIABLE chWorkSheet     AS COM-HANDLE   NO-UNDO.
DEFINE VARIABLE chCell          AS COM-HANDLE   NO-UNDO.

DEFINE VARIABLE iRow            AS INTEGER      NO-UNDO.
DEFINE VARIABLE iCol            AS INTEGER      NO-UNDO.
DEFINE VARIABLE chrLabel        AS CHARACTER    NO-UNDO.
DEFINE VARIABLE cc              AS RAW          NO-UNDO.

ASSIGN
    chOpenOffice = ?
    chWorkBook   = ?
    chDesktop    = ?
    chWorkSheet  = ?.


/****************************************************************************/
/***** Function Definitions ***********************************************/
/****************************************************************************/

/****************************************************************************/
FUNCTION col_letter RETURNS CHARACTER
    (INPUT ip_Col    AS INT):
/* Purpose: Returns the column letter for the column number              */
/****************************************************************************/

    DEF VAR res     AS CHAR NO-UNDO.
    DEF VAR l1      AS INT  NO-UNDO.
    DEF VAR l2      AS INT  NO-UNDO.

    /* Columns in OO start with 0; this function expects column 1 to be "A" */
    ASSIGN ip_Col = ip_Col + 1.

    /* Now get the column letter... */
    ASSIGN
        l2  = TRUNC((ip_Col - 1) / 26, 0).
        l1  = ip_Col - (26 * l2).
        res = CHR(64 + l2) + CHR(64 + l1).
        res = TRIM(res, CHR(64)).

    /* And return the value */
    RETURN res.
END FUNCTION. /* col_letter */

/****************************************************************************/
FUNCTION col_number RETURNS INTEGER
    ( INPUT ip_ColLetter AS CHARACTER ):
/* Purpose: returns the column number for the column letter passed int      */
/****************************************************************************/

    DEF VAR i           AS INT  NO-UNDO.
    DEF VAR intCurr     AS INT  NO-UNDO.
    DEF VAR intReturn   AS INT  NO-UNDO.

    /* upper case */
    ASSIGN ip_ColLetter = CAPS(ip_ColLetter).
    DO i = 1 TO LENGTH(ip_ColLetter) -  1 :
        ASSIGN
            intCurr   = ASC(SUBSTR(ip_ColLetter, i, 1)) - 64.
            intReturn = intReturn + (intCurr * 26).
    END.

    /* add the last letter. */
    ASSIGN 
        intReturn = intReturn + ASC(SUBSTR(ip_ColLetter, LENGTH(ip_ColLetter), 1)) - 64
        intReturn = intReturn - 1.      /* Subtract 1 -> 00 columns start at 0! */

    RETURN intReturn.
END FUNCTION. /* col_number */


/****************************************************************************/
/***** Procedure Definitions **********************************************/
/****************************************************************************/

/*******************************************************************/
PROCEDURE align_cell:
/* Purpose:     Align the cell specified.                          */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT          NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT          NO-UNDO.        /* Row Number */
    DEF INPUT PARAM ip_Horizontal   AS CHAR         NO-UNDO.        /* Values: BLOCK, CENTER, LEFT, REPEAT, RIGHT, STANDARD */
    DEF INPUT PARAM ip_Vertical     AS CHAR         NO-UNDO.        /* Values: BOTTOM, CENTER, STANDARD, TOP */
    DEF INPUT PARAM ip_Orientation  AS CHAR         NO-UNDO.        /* Values: BOTTOMTOP, STANDARD, STACKED, TOPBOTTOM */

    /* Get the current cell */
    ASSIGN chCell  = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).

    /* Set the various attributes (if applicable) */
    IF ip_Horizontal  > "" THEN chCell:HoriJustify = ip_Horizontal.
    IF ip_Vertical    > "" THEN chCell:VertJustify = ip_Vertical.
    IF ip_Orientation > "" THEN chCell:Orientation = ip_Orientation.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* align_cell */


/*******************************************************************/
PROCEDURE align_cell_range:
/* Purpose:     Align the row specified.                           */
/*******************************************************************/
    DEF INPUT PARAM ip_Range        AS CHAR         NO-UNDO.        /* eg B2:D4 */
    DEF INPUT PARAM ip_Horizontal   AS CHAR         NO-UNDO.        /* Values: BLOCK, CENTER, LEFT, REPEAT, RIGHT, STANDARD */
    DEF INPUT PARAM ip_Vertical     AS CHAR         NO-UNDO.        /* Values: BOTTOM, CENTER, STANDARD, TOP */
    DEF INPUT PARAM ip_Orientation  AS CHAR         NO-UNDO.        /* Values: BOTTOMTOP, STANDARD, STACKED, TOPBOTTOM */

    /* Get the cell range */
    ASSIGN chCell  = chWorkSheet:GetCellRangeByName(ip_Range).

    /* Set the various attributes (if applicable) */
    IF ip_Horizontal  > "" THEN chCell:HoriJustify = ip_Horizontal.
    IF ip_Vertical    > "" THEN chCell:VertJustify = ip_Vertical.
    IF ip_Orientation > "" THEN chCell:Orientation = ip_Orientation.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* align_cell_range */


/*******************************************************************/
PROCEDURE autofit_column:
/* Purpose:     Use the Calc OptimalWidth attribute to set widths  */
/*              for columns.                                       */
/* Parameters:  Pass in as Column#                                 */
/*******************************************************************/
    DEF INPUT PARAM ip_Col  AS INT  NO-UNDO.       /* column number */

    chWorkSheet:COLUMNS(ip_Col):OptimalWidth = TRUE.

END PROCEDURE. /* autofit_column */


/*******************************************************************/
PROCEDURE autofit_row:
/* Purpose:     Use the Calc OptimalHeight attribute to set height */
/*              for row                                            */
/* Parameters:  Pass in as Row#                                    */
/*******************************************************************/
    DEF INPUT PARAM ip_Row  AS INT  NO-UNDO.       /* row number */

    chWorkSheet:Rows(ip_Row):OptimalHeight = TRUE.

END PROCEDURE. /* autofit_column */

 
/****************************************************************************/
PROCEDURE CleanUp:
/* Purpose:     cleans up the com-handles                                   */
/****************************************************************************/

    /* Release all the com handles */
    IF chOpenOffice     <> ? THEN RELEASE OBJECT chOpenOffice.
    IF chWorkBook       <> ? THEN RELEASE OBJECT chWorkBook.
    IF chDesktop        <> ? THEN RELEASE OBJECT chDesktop.
    IF chWorkSheet      <> ? THEN RELEASE OBJECT chWorkSheet.

    ASSIGN
        chOpenOffice    = ?
        chWorkBook      = ?
        chDesktop       = ?
        chWorkSheet     = ?.

END PROCEDURE. /* CleanUp */


/*******************************************************************/
PROCEDURE close_Calc:
/* Purpose:     Close the Calc program.                            */
/*******************************************************************/

    /* Close the workbook */
    IF chWorkbook <> ? THEN
        chWorkbook:Close(TRUE).
    
    /* Close the program */
    IF chDesktop <> ? THEN
        chDesktop:TERMINATE().

END PROCEDURE. /* close_Calc */


/*******************************************************************/
PROCEDURE col_function:
/* Purpose:     Provides a utility to create a simple "column"     */
/*              function                                           */
/*              Valid functions are:                               */
/*              SUM: sums all numerical values                     */
/*              COUNT: total # of all values (including chars)     */
/*              COUNTNUMS: total number of all numerical cells     */
/*              AVERAGE: average of all numerical cells            */
/*              MAX: largest numerical value                       */
/*              MIN: smallest numerical value                      */
/*              PRODUCT: product of all numerical values           */
/*              STDEV: standard deviation                          */
/*              VAR: variance                                      */
/*              STDEVP: standard devt'n based on total population  */
/*              VARP: varianced based on total population          */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.        /* Row Number */
    DEF INPUT PARAM ip_StartRow     AS INT  NO-UNDO.
    DEF INPUT PARAM ip_EndRow       AS INT  NO-UNDO.
    DEF INPUT PARAM ip_Function     AS CHAR NO-UNDO.
    DEF VAR         chrFormula      AS CHAR NO-UNDO.

    /* Get the current cell */
    ASSIGN chCell  = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).

    /* And set the value */
    ASSIGN
        chrFormula = "=" +
                        ip_Function +
                        "(" +
                        col_letter(ip_Col) + STRING(ip_StartRow) +
                        ":" +
                        col_letter(ip_Col) + STRING(ip_EndRow) +
                        ")".

    chCell:SetFormula(chrFormula).

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* col_function */


/*******************************************************************/
PROCEDURE delete_col:
/* Purpose:     Deletes columns below the passed in col #          */
/*******************************************************************/
    DEF INPUT PARAM ip_Col      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_NumCols  AS INT  NO-UNDO.

    IF ip_NumCols < 1 THEN RETURN.

    /* delete columns */
    chWorkSheet:Columns:DeleteByIndex(ip_Col, ip_NumCols).

END PROCEDURE. /* delete_col */


/*******************************************************************/
PROCEDURE delete_row:
/* Purpose:     Deletes rows below the passed in row #             */
/*******************************************************************/
    DEF INPUT PARAM ip_Row      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_NumRows  AS INT  NO-UNDO.

    /* Make sure we're deleting at least one */
    IF ip_NumRows < 1 THEN RETURN.

    /* delete rows */
    chWorkSheet:Rows:DeleteByIndex(ip_Row, ip_NumRows).

END PROCEDURE. /* delete_row */


/*******************************************************************/
PROCEDURE freeze_panes:  
/* Purpose:     freezes panes with the specified number of columns */
/*              and rows. To freeze only horizontally, specify     */
/*              ip_Row as 0. To freeze only vertically, specify    */
/*              ip_Col as 0.                                       */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.        /* Row Number */

    chWorkBook:getCurrentController():freezeAtPosition(ip_Col, ip_Row).

END PROCEDURE. /* freeze_panes */


/****************************************************************************/
PROCEDURE format_cell:
/* Purpose:     Formats a cell with a specific format.                      */
/****************************************************************************/
    DEF INPUT PARAM ip_Col                  AS INT          NO-UNDO.
    DEF INPUT PARAM ip_Row                  AS INT          NO-UNDO.
    DEF INPUT PARAM ip_DataType             AS CHAR         NO-UNDO.
    DEF INPUT PARAM ip_Decimals             AS INT          NO-UNDO.

    DEF VAR         chNumberFormats         AS COM-HANDLE   NO-UNDO.
    DEF VAR         chLocalSettings         AS COM-HANDLE   NO-UNDO.
    DEF VAR         chrFormats              AS CHAR         NO-UNDO.
    DEF VAR         intNumberFormatId       AS DEC          NO-UNDO.
    DEF VAR         chrNumberFormatString   AS CHAR         NO-UNDO.


DEF VAR mpLocale    AS COM-HANDLE NO-UNDO.

    /* Get the current cell */
    chCell              = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).
    chNumberFormats     = chWorkBook:NumberFormats.
/*
    chrFormats          = chNumberFormats:queryKeys(16, mpLocale, FALSE).

    CASE ip_DataType:
        WHEN "Decimal" THEN chrNumberFormatString = "###,###,###,##0." + FILL("0", ip_Decimals).
        OTHERWISE           chrNumberFormatString = "".
    END CASE. /* ip_DataType */



    intNumberFormatId = chNumberFormats:queryKey(chrNumberFormatString, chLocalSettings, TRUE).
*/

    /* And get the value */
    ASSIGN
        chCell:NumberFormat = 8. /*intNumberFormatId. */


    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

    IF chNumberFormats <> ? THEN DO:
        RELEASE OBJECT chNumberFormats.
        ASSIGN chNumberFormats = ?.
    END.

    RETURN.

END PROCEDURE. /* format_cell */


/*******************************************************************/
PROCEDURE get_cell_data:
/* Purpose:     Gets the cell data (char) from a worksheet         */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.        /* Row Number */
    DEF OUTPUT PARAM op_CellData    AS CHAR NO-UNDO.

    /* Get the current cell */
    ASSIGN chCell  = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).

    /* And get the value */
    ASSIGN op_CellData = STRING(chCell:FORMULA).

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* get_cell_data */


/*******************************************************************/
PROCEDURE get_last_row:
/* Purpose:     Returns the last row as an integer.                */
/*******************************************************************/
    DEF INPUT PARAM ip_Col      AS INT  NO-UNDO.
    DEF OUTPUT PARAM op_LastRow AS INT  NO-UNDO.

    DEF VAR         chrRange    AS CHAR NO-UNDO.

    /* Set up the range */
    ASSIGN chrRange = col_letter(ip_Col) + "10:" + col_letter(ip_Col) + "14".

    /* Get the current row */
    ASSIGN chCell  = chWorkSheet:GetCellRangeByName(chrRange).
    op_LastRow = chCell:getRows:COUNT.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* get_last_row */


/*******************************************************************/
PROCEDURE insert_col:
/* Purpose:     Inserts columns below the passed in col #          */
/*******************************************************************/
    DEF INPUT PARAM ip_Col      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_NumCols  AS INT  NO-UNDO.

    /* Make sure we're deleting at least one */
    IF ip_NumCols < 1 THEN RETURN.

    /* Insert columns */
    chWorkSheet:Columns:InsertByIndex(ip_Col, ip_NumCols).

END PROCEDURE. /* insert_col */


/*******************************************************************/
PROCEDURE insert_row:  
/* Purpose:     Inserts rows below the passed in row #             */
/*******************************************************************/
    DEF INPUT PARAM ip_Row      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_NumRows  AS INT  NO-UNDO.

    IF ip_NumRows < 1 THEN RETURN.

    chWorksheet:rows:insertByIndex(ip_Row, ip_NumRows).

END PROCEDURE. /* insert_row */


/****************************************************************************/
PROCEDURE Maximize_Window:
/* Purpose:     Maximizes the OO screen.                                    */
/****************************************************************************/
    DEF VAR chFrame     AS COM-HANDLE   NO-UNDO.
    DEF VAR chWindow    AS COM-HANDLE   NO-UNDO.
    DEF VAR chRect      AS COM-HANDLE   NO-UNDO.

    chFrame  = chDesktop:getCurrentFrame().
    chWindow = chFrame:getContainerWindow().
    chRect   = chWindow:setPosSize(1,1,800,600,15).

    RETURN.

END PROCEDURE. /* Maximize_Window */


/****************************************************************************/
PROCEDURE Minimize_Window:
/* Purpose:     Minimizes the OO screen.                                    */
/****************************************************************************/
    DEF VAR chFrame     AS COM-HANDLE   NO-UNDO.
    DEF VAR chWindow    AS COM-HANDLE   NO-UNDO.
    DEF VAR chRect      AS COM-HANDLE   NO-UNDO.

    chFrame  = chDesktop:getCurrentFrame().
    chWindow = chFrame:getContainerWindow().
    chRect   = chWindow:setPosSize(1,1,1,1,15).

    RETURN.

END PROCEDURE. /* Minimize_Window */


/*******************************************************************/
PROCEDURE new_book:
/* Purpose:     Open a new "BOOK" in Calc and create a worksheet.  */
/*******************************************************************/

    chWorkBook  = chDesktop:loadComponentFromURL("private:factory/scalc", "_blank", 0, cc).
    chWorkSheet = chWorkBook:Sheets:getByIndex(0).

END PROCEDURE. /* new_book */


/*******************************************************************/
PROCEDURE open_book:
/* Purpose:     open a "BOOK" in Calc.                             */
/*******************************************************************/
    DEF INPUT PARAM ip_FileName    AS CHAR   NO-UNDO. /* spreadsheet name */

    IF chWorkbook = ? THEN LEAVE.

    ASSIGN
        ip_FileName = "file:///" + TRIM(ip_FileName)
        ip_FileName = REPLACE(ip_FileName, "\", "/").

    chWorkbook  = chDesktop:loadComponentFromURL(ip_FileName, "_blank", 0, cc).
    chWorksheet = chWorkbook:Worksheets:Item(1).

END PROCEDURE. /* open_book */


/*******************************************************************/
PROCEDURE open_calc:
/* Purpose:     Start OpenOffice & open a DDE conversation with    */
/*              the Calc System topic                              */
/*******************************************************************/

    /* Try to connect to existing instance of OpenOffice */
    CREATE "com.sun.star.ServiceManager" chOpenOffice CONNECT NO-ERROR.

    /* If some error happened then most likely there was no */
    /* instance of OO running so start a new one. */
    IF ERROR-STATUS:GET-MESSAGE(1) <> "" THEN
        CREATE "com.sun.star.ServiceManager" chOpenOffice.

    /* Start up the OO desktop now.  Everything fires from there. */
    chDesktop = chOpenOffice:createInstance("com.sun.star.frame.Desktop").
END PROCEDURE. /* open_calc */


/*******************************************************************/
PROCEDURE page_break:
/* Purpose:     Inserts a manual page break for reporting          */
/*******************************************************************/
    DEF INPUT PARAM ip_Row      AS INT  NO-UNDO.

    DEF VAR         chrRange    AS CHAR NO-UNDO.

    /* Set up the range */
    ASSIGN chrRange = "A" + TRIM(STRING(ip_Row)) + ":A" + TRIM(STRING(ip_Row)).

    /* Get the current row */
    ASSIGN chCell  = chWorkSheet:GetCellRangeByName(chrRange).
    chCell:getRows:IsStartOfNewPage = TRUE.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* page_break */


/*******************************************************************/
PROCEDURE row_function:
/* Purpose:     Provides a utility to create a simple "row"        */
/*              function                                           */
/*              SUM: sums all numerical values                     */
/*              COUNT: total # of all values (including chars)     */
/*              COUNTNUMS: total number of all numerical cells     */
/*              AVERAGE: average of all numerical cells            */
/*              MAX: largest numerical value                       */
/*              MIN: smallest numerical value                      */
/*              PRODUCT: product of all numerical values           */
/*              STDEV: standard deviation                          */
/*              VAR: variance                                      */
/*              STDEVP: standard devt'n based on total population  */
/*              VARP: varianced based on total population          */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.        /* Row Number */
    DEF INPUT PARAM ip_StartCol     AS INT  NO-UNDO.
    DEF INPUT PARAM ip_EndCol       AS INT  NO-UNDO.
    DEF INPUT PARAM ip_Function     AS CHAR NO-UNDO.
    DEF VAR         chrFormula      AS CHAR NO-UNDO.

    /* Get the current cell */
    ASSIGN chCell  = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).

    /* And set the value */
    ASSIGN
        chrFormula = "=" +
                        ip_Function +
                        "(" +
                        col_letter(ip_StartCol) + STRING(ip_Row) +
                        ":" +
                        col_letter(ip_EndCol) + STRING(ip_Row) +
                        ")".

    chCell:SetFormula(chrFormula).

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* row_function */


/*******************************************************************/
PROCEDURE save_book:
/* Purpose:     Save a "BOOK" in Calc                              */
/*******************************************************************/
    DEF INPUT PARAM ip_OutputPath   AS CHAR NO-UNDO.    /* directory path */
    DEF INPUT PARAM ip_FileName     AS CHAR NO-UNDO.    /* spreadsheet name */

    DEF VAR chrFileName             AS CHAR NO-UNDO.

    ASSIGN
        chrFileName = "file:///" + ip_OutputPath + ip_FileName + ".sxc"
        chrFileName = chrFileName + "\"
        chrFileName = REPLACE(chrFileName, "\\", "\")
        chrFileName = REPLACE(chrFileName, "/\", "\")
        chrFileName = REPLACE(chrFileName, "\", "/").

    chWorkBook:storeAsURL(chrFileName, cc).

END PROCEDURE. /* save_book */


/*******************************************************************/
PROCEDURE set_col_width:
/* Purpose:     Set the width of a column.                         */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.
    DEF INPUT PARAM ip_ColWidth     AS DEC  NO-UNDO.

    IF ip_Col > 0 THEN chWorksheet:Columns(ip_Col):Width = ip_ColWidth.

END PROCEDURE. /* set_col_width */


/*******************************************************************/
PROCEDURE set_font_style:
/* Purpose:     Set font style for a cell                          */
/* Parameters:  Row #                                              */
/*              Column #                                           */
/*              Font Name                                          */
/*              Size (points)                                      */
/*              Bold (TRUE/FALSE)                                  */
/*              Underline (See below)                              */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT  NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.        /* Row Number */
    DEF INPUT PARAM ip_Font         AS CHAR NO-UNDO.        /* Font Name */
    DEF INPUT PARAM ip_Size         AS INT  NO-UNDO.        /* Point Size */
    DEF INPUT PARAM ip_Bold         AS LOG  NO-UNDO.        /* Bold (weight = 150 for bold, 100 for normal) */
    DEF INPUT PARAM ip_Underline    AS INT  NO-UNDO.        /* NONE = 0, SINGLE =1, DOUBLE=2, DOTTED = 3
                                                               DONTKNOW=4, DASH=5, LONGDASH=6, DASHDOT=7,
                                                               DASHDOTDOT=8, SMALLWAVE=9, WAVE =10, DOUBLEWAVE=11,
                                                               BOLD=12, BOLDDOTTED=13, BOLDLONGDASH= 14,
                                                               BOLDDASHDOT=15, BOLDDASHDOTDOT=16, BOLDWAVE = 17  */

    /* Get the current cell */
    ASSIGN chCell  = chWorkSheet:GetCellByPosition(ip_Col,ip_Row).

    /* Set the various attributes (if applicable) */
    chCell:CharFontName  = IF ip_Font > ""      THEN ip_Font        ELSE chCell:CharFontName.
    chCell:CharHeight    = IF ip_Size > 0       THEN ip_Size        ELSE chCell:CharHeight.
    chCell:CharWeight    = IF ip_Bold           THEN 150            ELSE 100.
    chCell:CharUnderline = IF ip_Underline > 0  THEN ip_UnderLine   ELSE 0.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* set_font_style */


/*******************************************************************/
PROCEDURE set_footer:
/* Purpose:     Creates a footer for the document                  */
/*******************************************************************/
    DEF INPUT PARAM ip_Text     AS CHAR         NO-UNDO.

    DEF VAR chStyleFamilies     AS COM-HANDLE   NO-UNDO.
    DEF VAR chPageStyles        AS COM-HANDLE   NO-UNDO.
    DEF VAR chDefaultPage       AS COM-HANDLE   NO-UNDO.
    DEF VAR chFooterText        AS COM-HANDLE   NO-UNDO.
    DEF VAR chFooterContent     AS COM-HANDLE   NO-UNDO.

    /* first we need to get the default page style for this workbook/sheet */
    chStyleFamilies = chWorkBook:StyleFamilies.
    chPageStyles = chStyleFamilies:getByName("PageStyles").
    chDefaultPage = chPageStyles:getByName("Default").

    /* Turn on Footers */
    chDefaultPage:FooterIsOn = TRUE.

    /* Same Footer for both left (even) & right (odd) pages */
    chDefaultPage:FooterIsShared = TRUE.

    /* Set up the Footer */
    chFooterContent = chDefaultPage:RightPageFooterContent.
    chFooterText = chFooterContent:CenterText.
    chFooterText:STRING = ip_Text.
    chDefaultPage:RightPageFooterContent = chFooterContent.

    /* Clean up now */
    IF chStyleFamilies <> ? THEN DO:
        RELEASE OBJECT chStyleFamilies.
        ASSIGN chStyleFamilies = ?.
    END.

    IF chPageStyles <> ? THEN DO:
        RELEASE OBJECT chPageStyles.
        ASSIGN chPageStyles = ?.
    END.

    IF chDefaultPage <> ? THEN DO:
        RELEASE OBJECT chDefaultPage.
        ASSIGN chDefaultPage = ?.
    END.

    IF chFooterText <> ? THEN DO:
        RELEASE OBJECT chFooterText.
        ASSIGN chFooterText = ?.
    END.

    IF chFooterContent <> ? THEN DO:
        RELEASE OBJECT chFooterContent.
        ASSIGN chFooterContent = ?.
    END.

END PROCEDURE. /* set_footer */


/*******************************************************************/
PROCEDURE set_header:
/* Purpose:     Creates a header for the document                  */
/*******************************************************************/
    DEF INPUT PARAM ip_Text     AS CHAR         NO-UNDO.

    DEF VAR chStyleFamilies     AS COM-HANDLE   NO-UNDO.
    DEF VAR chPageStyles        AS COM-HANDLE   NO-UNDO.
    DEF VAR chDefaultPage       AS COM-HANDLE   NO-UNDO.
    DEF VAR chHeaderText        AS COM-HANDLE   NO-UNDO.
    DEF VAR chHeaderContent     AS COM-HANDLE   NO-UNDO.

    /* first we need to get the default page style for this workbook/sheet */
    chStyleFamilies = chWorkBook:StyleFamilies.
    chPageStyles = chStyleFamilies:getByName("PageStyles").
    chDefaultPage = chPageStyles:getByName("Default").

    /* Turn on headers */
    chDefaultPage:HeaderIsOn = TRUE.

    /* Same header for both left (even) & right (odd) pages */
    chDefaultPage:HeaderIsShared = TRUE.

    /* Set up the header */
    chHeaderContent = chDefaultPage:RightPageHeaderContent.
    chHeaderText = chHeaderContent:CenterText.
    chHeaderText:STRING = ip_Text.
    chDefaultPage:RightPageHeaderContent = chHeaderContent.

    /* Clean up now */
    IF chStyleFamilies <> ? THEN DO:
        RELEASE OBJECT chStyleFamilies.
        ASSIGN chStyleFamilies = ?.
    END.

    IF chPageStyles <> ? THEN DO:
        RELEASE OBJECT chPageStyles.
        ASSIGN chPageStyles = ?.
    END.

    IF chDefaultPage <> ? THEN DO:
        RELEASE OBJECT chDefaultPage.
        ASSIGN chDefaultPage = ?.
    END.

    IF chHeaderText <> ? THEN DO:
        RELEASE OBJECT chHeaderText.
        ASSIGN chHeaderText = ?.
    END.

    IF chHeaderContent <> ? THEN DO:
        RELEASE OBJECT chHeaderContent.
        ASSIGN chHeaderContent = ?.
    END.

END PROCEDURE. /* set_header */


/*******************************************************************/
PROCEDURE set_row_height:
/* Purpose:     Set the height of a row.                           */
/* Parameters:  ip_Row    row #                                    */
/*              ip_RowHeight in inches 1/4 = 0.25 passed in        */
/*******************************************************************/
    DEF INPUT PARAM ip_Row          AS INT  NO-UNDO.
    DEF INPUT PARAM ip_RowHeight    AS DEC  NO-UNDO.

    IF ip_Row > 0 THEN chWorksheet:Rows(ip_Row):HEIGHT = ip_RowHeight.

END PROCEDURE. /* set_row_height */


/*******************************************************************/
PROCEDURE show_col:
/* Purpose:     Shows or hides a column                            */
/*******************************************************************/
    DEF INPUT PARAM ip_Col      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_Show     AS LOG  NO-UNDO.        /* TRUE = visible, FALSE = hidden */

    DEF VAR         chrRange    AS CHAR NO-UNDO.

    /* Set up the range */
    ASSIGN chrRange = col_letter(ip_Col) + "1:" + col_letter(ip_Col) + "1".

    /* Get the current column */
    ASSIGN chCell  = chWorkSheet:GetCellRangeByName(chrRange).
    chCell:getColumns:IsVisible = ip_Show.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* show_col */


/*******************************************************************/
PROCEDURE show_row:
/* Purpose:     Shows or hides a row                               */
/*******************************************************************/
    DEF INPUT PARAM ip_Row      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_Show     AS LOG  NO-UNDO.        /* TRUE = visible, FALSE = hidden */

    DEF VAR         chrRange    AS CHAR NO-UNDO.

    /* Set up the range */
    ASSIGN chrRange = "A" + TRIM(STRING(ip_Row)) + ":A" + TRIM(STRING(ip_Row)).

    /* Get the current row */
    ASSIGN chCell  = chWorkSheet:GetCellRangeByName(chrRange).
    chCell:getRows:IsVisible = ip_Show.

    /* Clean up after ourselves */
    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* show_row */


/*******************************************************************/
PROCEDURE write_cell_data:
/* Purpose:     Write data into a worksheet.                       */
/*******************************************************************/
    DEF INPUT PARAM ip_Col          AS INT      NO-UNDO.        /* Column Number */
    DEF INPUT PARAM ip_Row          AS INT      NO-UNDO.        /* Row Number */
    DEF INPUT PARAM ip_Data         AS CHAR     NO-UNDO.
    DEF VAR         decTest         AS DEC      NO-UNDO.

    ASSIGN chCell = chWorkSheet:GetCellByPosition(iCol,iRow).

    ASSIGN decTest = DEC(ip_Data) NO-ERROR.

    IF ERROR-STATUS:ERROR THEN chCell:SetFormula(ip_Data).
    ELSE                       chCell:SetValue(DEC(ip_Data)).

    IF chCell <> ? THEN DO:
        RELEASE OBJECT chCell.
        ASSIGN chCell  = ?.
    END.

END PROCEDURE. /* write_cell_data */


/*******************************************************************/
PROCEDURE write_col_hdr:
/* Purpose:     Write row 1 column headers in a worksheet.         */
/*******************************************************************/
    DEF INPUT PARAM ip_Col      AS INT  NO-UNDO.
    DEF INPUT PARAM ip_Label    AS CHAR NO-UNDO.

    RUN write_cell_data (ip_Col, 0, ip_Label).
    RUN set_font_style
        (ip_Col,    /* Column ip_Col */
         0,         /* Row 1 */
         "",        /* Default font */
         10,        /* 12 points */
         TRUE,      /* Bold */
         1).        /* Single Underline */

END PROCEDURE. /* write_col_hdr */


