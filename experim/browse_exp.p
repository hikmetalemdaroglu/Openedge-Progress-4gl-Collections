/*------------------------------------------------------------------------

  File:         browse_exp.p

  Description: 

  Author:       Stefan Houtzager       

  Created:      20-04-2006

------------------------------------------------------------------------*/
/*  SESSION:DEBUG-ALERT = TRUE. */

DEFINE TEMP-TABLE tempBuffer NO-UNDO
FIELD Hdl AS HANDLE
FIELD Tbl AS CHARACTER.

DEFINE TEMP-TABLE tempCol NO-UNDO
  FIELD ColHandle  AS HANDLE
  FIELD CallHandle AS HANDLE
  FIELD ColWidth   AS INTEGER
  FIELD CalcAction AS CHARACTER
  FIELD ColLbl     AS CHARACTER
  FIELD DataType   AS CHARACTER
  FIELD AscDesc    AS CHARACTER
  FIELD AllowSort  AS LOGICAL.

DEFINE TEMP-TABLE tempFillin NO-UNDO
  FIELD FillinHandle AS HANDLE
  FIELD ColHandle    AS HANDLE.

DEFINE TEMP-TABLE tempIdx NO-UNDO
  FIELD TbleName AS CHARACTER 
  FIELD FldName  AS CHARACTER
  FIELD FldName2 AS CHARACTER
  FIELD FldName3 AS CHARACTER
  FIELD Prim     AS LOGICAL
  FIELD AscDesc  LIKE tempCol.AscDesc.

DEFINE TEMP-TABLE tempFilterElement NO-UNDO
  FIELD cField    AS CHARACTER 
  FIELD cElement  AS CHARACTER
  FIELD iEntryNum AS INTEGER
  FIELD lChartype AS LOGICAL
  FIELD lUsed     AS LOGICAL
  INDEX EntryNum iEntryNum.

DEFINE BUFFER buftempFilterElement FOR tempFilterElement.

DEFINE TEMP-TABLE tempQuerySlice NO-UNDO
  FIELD cBuffer AS CHARACTER
  FIELD cSlice  AS CHARACTER.

DEFINE TEMP-TABLE tempNewQuerySlice NO-UNDO
  LIKE tempQuerySlice.

DEFINE VARIABLE cText           AS CHARACTER NO-UNDO INITIAL "WeRtYuIoPlKjHgfdsWeNHbvjLLLhMlKjHtReWsCfTgHjKpLDrTgHjKlPlIuYtReWKoPlNhgfre".
DEFINE VARIABLE cBaseQuery      AS CHARACTER NO-UNDO.
DEFINE VARIABLE cPrivateParts   AS CHARACTER NO-UNDO.
DEFINE VARIABLE cFunc           AS CHARACTER NO-UNDO.
DEFINE VARIABLE cIndexInfo      AS CHARACTER NO-UNDO.
DEFINE VARIABLE cIdxFieldList   AS CHARACTER NO-UNDO.
DEFINE VARIABLE OCXFile         AS CHARACTER NO-UNDO.
DEFINE VARIABLE hEntry          AS HANDLE    NO-UNDO.
DEFINE VARIABLE hStartingWidget AS HANDLE    NO-UNDO.
DEFINE VARIABLE hLastTabItem    AS HANDLE    NO-UNDO.
DEFINE VARIABLE hPrevFillin     AS HANDLE    NO-UNDO.
DEFINE VARIABLE hCaller         AS HANDLE    NO-UNDO.
DEFINE VARIABLE hCancel         AS HANDLE    NO-UNDO.
DEFINE VARIABLE hOk             AS HANDLE    NO-UNDO.
DEFINE VARIABLE hBrowse         AS HANDLE    NO-UNDO.
DEFINE VARIABLE hColumn         AS HANDLE    NO-UNDO.
DEFINE VARIABLE hPrevCol        AS HANDLE    NO-UNDO.
DEFINE VARIABLE hQuery          AS HANDLE    NO-UNDO.
DEFINE VARIABLE hBuffer         AS HANDLE    NO-UNDO.
DEFINE VARIABLE hWin            AS HANDLE    NO-UNDO.
DEFINE VARIABLE hFrame          AS HANDLE    NO-UNDO.
DEFINE VARIABLE hFill           AS HANDLE    NO-UNDO.
DEFINE VARIABLE hSuper          AS HANDLE    NO-UNDO.
DEFINE VARIABLE CtrlFrame       AS HANDLE    NO-UNDO.
DEFINE VARIABLE hCall           AS HANDLE    NO-UNDO. 
DEFINE VARIABLE hCurCol         AS HANDLE    NO-UNDO.
DEFINE VARIABLE chCtrlFrame     AS COMPONENT-HANDLE NO-UNDO. 
DEFINE VARIABLE i               AS INTEGER   NO-UNDO.
DEFINE VARIABLE j               AS INTEGER   NO-UNDO. 
DEFINE VARIABLE iBuffers        AS INTEGER   NO-UNDO. 
DEFINE VARIABLE iPrevWidth      AS INTEGER   NO-UNDO.
DEFINE VARIABLE hIcon           AS INTEGER   NO-UNDO.
DEFINE VARIABLE ret             AS INTEGER   NO-UNDO.
DEFINE VARIABLE iIdx            AS INTEGER   NO-UNDO.
DEFINE VARIABLE iNumPars        AS INTEGER   NO-UNDO.
DEFINE VARIABLE iXcoordinate    AS INTEGER   NO-UNDO.
DEFINE VARIABLE iPosition       AS INTEGER INITIAL 1  NO-UNDO.
DEFINE VARIABLE iTitleBarHeight AS INTEGER INITIAL 35 NO-UNDO.
DEFINE VARIABLE UIB_S           AS LOGICAL    NO-UNDO.
DEFINE VARIABLE lIndexedRepos   AS LOGICAL    NO-UNDO.
DEFINE VARIABLE lDummy          AS LOGICAL    NO-UNDO.
/* choose values for custom appearance */
&SCOPED-DEFINE BORDERWIDTH 5  
&SCOPED-DEFINE BORDERHEIGHT 10 
&SCOPED-DEFINE NUMROWS 20  
&SCOPED-DEFINE ROWHEIGHT 13 
&SCOPED-DEFINE FONT 4 
&SCOPED-DEFINE WINMESSAGE 20 

PROCEDURE GetDC EXTERNAL "user32.dll" :
  DEFINE INPUT  PARAMETER hWnd AS LONG.
  DEFINE RETURN PARAMETER hdc  AS LONG.
END PROCEDURE.

PROCEDURE DeleteDC EXTERNAL "gdi32.dll" :
  DEFINE INPUT  PARAMETER hdc          AS LONG.
  DEFINE RETURN PARAMETER ReturnValue  AS LONG.
END PROCEDURE.

PROCEDURE LoadImageA EXTERNAL "user32.dll":
  DEFINE INPUT  PARAMETER hInst       AS LONG.
  DEFINE INPUT  PARAMETER lpsz        AS CHARACTER.
  DEFINE INPUT  PARAMETER un1         AS LONG.
  DEFINE INPUT  PARAMETER n1          AS LONG.
  DEFINE INPUT  PARAMETER n2          AS LONG.
  DEFINE INPUT  PARAMETER un2         AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

PROCEDURE CreateCompatibleDC EXTERNAL "gdi32.dll" :
  DEFINE INPUT  PARAMETER hdc         AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

PROCEDURE SelectObject EXTERNAL "gdi32.dll" :
  DEFINE INPUT  PARAMETER hdc         AS LONG.
  DEFINE INPUT  PARAMETER hObject     AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

PROCEDURE DeleteObject EXTERNAL "gdi32.dll" :
  DEFINE INPUT  PARAMETER hObject     AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

PROCEDURE BitBlt EXTERNAL "gdi32.dll" :
  DEFINE INPUT  PARAMETER hDestDC     AS LONG.
  DEFINE INPUT  PARAMETER x           AS LONG.
  DEFINE INPUT  PARAMETER y           AS LONG.
  DEFINE INPUT  PARAMETER nWidth      AS LONG.
  DEFINE INPUT  PARAMETER nHeight     AS LONG.
  DEFINE INPUT  PARAMETER hSrcDC      AS LONG.
  DEFINE INPUT  PARAMETER xSrc        AS LONG.
  DEFINE INPUT  PARAMETER ySrc        AS LONG.
  DEFINE INPUT  PARAMETER dwRop       AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

PROCEDURE ReleaseDC EXTERNAL "user32.dll" :
  DEFINE INPUT  PARAMETER hWnd        AS LONG.
  DEFINE INPUT  PARAMETER hdc         AS LONG.
  DEFINE RETURN PARAMETER ReturnValue AS LONG.
END PROCEDURE.

FUNCTION getString RETURNS CHARACTER PRIVATE (INPUT pihCurrentCol AS HANDLE) FORWARD. 
FUNCTION openQuery RETURNS LOGICAL PRIVATE () FORWARD. 

ON 'ENTRY':U ANYWHERE DO: 
  /* ensure repaint after leaving the browse */
  DEFINE VARIABLE hWigetLeave AS HANDLE NO-UNDO. 

  hWigetLeave = LAST-EVENT:WIDGET-LEAVE.
  IF VALID-HANDLE(hWigetLeave) 
      AND hWigetLeave:TYPE = 'BROWSE':U 
  THEN DO: 
    hCurCol = hPrevCol. 
    RUN CtrlFrame.Msgblst32.Message ({&WINMESSAGE},1522,0,input-output i). 
  END.
END.

/* Assign your buttons, menu-items etc from where you run this procedure as private-data 
  <function-name>|<dialog/window>|<widgetname1=fieldname1,widgetname2=fieldname2,etc>  
   function-name to retrieve the func-record 
   "dialog" to create a dialog-box
   edit the third parameter for a selection-dialog to assign cell-values in the browse to fillin-screenvalues in the starting window */
ASSIGN cPrivateParts       = SELF:PRIVATE-DATA 
       hStartingWidget     = SELF
       cFunc               = ENTRY(1,cPrivateParts,'|':U).

FIND Func WHERE Func.Func = cFunc NO-LOCK NO-ERROR.
FIND Qry OF Func NO-LOCK NO-ERROR.

ASSIGN lIndexedRepos = INDEX(Qry.Qry,'INDEXED-REPOSITION':U) NE 0
       cBaseQuery    = "FOR ":U + Qry.Qry.

IF SEARCH("calcsupers\":U + cFunc + "super.r") NE ? OR SEARCH("calcsupers\":U + cFunc + "super.p") NE ? THEN DO:
  RUN VALUE("calcsupers\":U + cFunc + "super.p") PERSISTENT SET hSuper.
  THIS-PROCEDURE:ADD-SUPER-PROCEDURE(hSuper).
END.

CREATE QUERY hQuery.       

DO iBuffers = 1 TO NUM-ENTRIES(Qry.TblList):
  CREATE BUFFER hBuffer FOR TABLE ENTRY(1,ENTRY(2,ENTRY(iBuffers,Qry.TblList),'.':U),' ':U).
  hQuery:ADD-BUFFER(hBuffer).
  CREATE tempBuffer.
  ASSIGN tempBuffer.Hdl = hBuffer
         tempBuffer.Tbl = ENTRY(1,ENTRY(2,ENTRY(iBuffers,Qry.TblList),'.':U),' ':U)

         iIdx           = 1.
  /* chop the query into slices for manipulation of the querystring in FUNCTION openQuery */
  CREATE tempQuerySlice.
  ASSIGN tempQuerySlice.cSlice  = SUBSTRING(Qry.Qry,
                                            iPosition,
                                            IF INDEX(Qry.Qry,',':U,iPosition + 1) = 0 THEN LENGTH(Qry.Qry) - iPosition + 1
                                            ELSE INDEX(Qry.Qry,',':U,IF iPosition NE 1 THEN iPosition + 1 ELSE iPosition) - iPosition)
         iPosition              = INDEX(Qry.Qry,',':U,iPosition + 1)
         tempQuerySlice.cBuffer = tempBuffer.Tbl.

  /* fill temptable with index-info to be able to indicate non-indexed sortcolumns */
  DO WHILE hBuffer:INDEX-INFORMATION(iIdx) NE ?:    
    CREATE tempIdx.
    ASSIGN tempIdx.TbleName = hBuffer:NAME 
           tempIdx.FldName  = ENTRY(5, hBuffer:INDEX-INFORMATION(iIdx))
           tempIdx.FldName2 = IF NUM-ENTRIES(hBuffer:INDEX-INFORMATION(iIdx)) >= 7 THEN ENTRY(7, hBuffer:INDEX-INFORMATION(iIdx)) ELSE '':U
           tempIdx.FldName3 = IF NUM-ENTRIES(hBuffer:INDEX-INFORMATION(iIdx)) >= 9 THEN ENTRY(9, hBuffer:INDEX-INFORMATION(iIdx)) ELSE '':U
           tempIdx.Prim     = ENTRY(3, hBuffer:INDEX-INFORMATION(iIdx)) = '1':U
           tempIdx.AscDesc  = IF tempIdx.Prim = TRUE AND ENTRY(6, hBuffer:INDEX-INFORMATION(iIdx)) = '0':U THEN 'asc':U ELSE 
                              IF tempIdx.Prim = TRUE AND ENTRY(6, hBuffer:INDEX-INFORMATION(iIdx)) = '1':U THEN 'desc':U ELSE '':U
           iIdx             = iIdx + 1. 
  END.
END.

IF NUM-ENTRIES(cPrivateParts,'|':U) = 1 
    OR (NUM-ENTRIES(cPrivateParts,'|':U) > 1 AND NOT ENTRY(2, cPrivateParts, '|':U) = 'dialog':U) THEN DO:
   CREATE WINDOW hWin ASSIGN
         HIDDEN             = YES
         TITLE              = func.wintitle
         RESIZE             = NO
         SCROLL-BARS        = no
         STATUS-AREA        = no
         KEEP-FRAME-Z-ORDER = yes
         THREE-D            = yes
         FONT               = {&FONT}
         MESSAGE-AREA       = no
         SENSITIVE          = yes
      TRIGGERS:
         ON WINDOW-CLOSE PERSISTENT RUN disable_UI IN THIS-PROCEDURE.
      END TRIGGERS. 
   
   CREATE FRAME hFrame ASSIGN
         PARENT             = hWin
         HIDDEN             = YES
         THREE-D            = YES 
         FONT               = {&FONT}
         SENSITIVE          = YES 
         NAME               = 'DEFAULT-FRAME':U.
END.
ELSE DO:
  CREATE DIALOG-BOX hFrame ASSIGN
      HIDDEN             = YES
      TITLE              = func.wintitle
      THREE-D            = YES 
      FONT               = {&FONT}
      SENSITIVE          = YES 
   TRIGGERS:
      ON WINDOW-CLOSE, ENDKEY, END-ERROR PERSISTENT RUN disable_UI IN THIS-PROCEDURE.
   END TRIGGERS. 
  
  CREATE BUTTON hCancel ASSIGN
      FRAME         = hFrame
      WIDTH-PIXELS  = 75
      HEIGHT-PIXELS = 25
      LABEL         = "Annuleer"
      AUTO-END-KEY  = YES
      SENSITIVE     = YES
    TRIGGERS:
       ON CHOOSE PERSISTENT RUN cancelButton IN THIS-PROCEDURE. 
    END TRIGGERS. 

  CREATE BUTTON hOk ASSIGN
      FRAME          = hFrame
      WIDTH-PIXELS   = 75
      HEIGHT-PIXELS  = 25
      LABEL          = "OK"
      AUTO-GO        = YES
      DEFAULT        = TRUE
      SENSITIVE      = YES
    TRIGGERS:
       ON CHOOSE PERSISTENT RUN okButton IN THIS-PROCEDURE. 
    END TRIGGERS.       

  ASSIGN hFrame:CANCEL-BUTTON  = hCancel
         hFrame:DEFAULT-BUTTON = hOk.
END.

CREATE BROWSE hBrowse ASSIGN 
        X                 = {&BORDERWIDTH}
        Y                 = {&BORDERHEIGHT}
        WIDTH-PIXELS      = 40
        ROW-HEIGHT-PIXELS = {&ROWHEIGHT}
        QUERY             = hQuery
        FRAME             = hFrame 
        READ-ONLY         = TRUE
        SEPARATORS        = TRUE
        FIT-LAST-COLUMN   = TRUE 
        COLUMN-RESIZABLE  = TRUE
        SENSITIVE         = TRUE
        FONT              = {&FONT}
        ROW-MARKERS       = FALSE
  TRIGGERS:
    ON END-RESIZE ANYWHERE PERSISTENT RUN endResize   IN THIS-PROCEDURE. 
    ON SCROLL-NOTIFY  PERSISTENT RUN endResize   IN THIS-PROCEDURE. 
    ON START-SEARCH   PERSISTENT RUN startSearch IN THIS-PROCEDURE.
    ON ROW-DISPLAY    PERSISTENT RUN rowDisplay  IN THIS-PROCEDURE. 
    ON DEFAULT-ACTION PERSISTENT RUN okButton    IN THIS-PROCEDURE. 
  END TRIGGERS.

iXcoordinate = hBrowse:X.

FOR EACH bcol OF func NO-LOCK:
  FIND FIRST tempIdx WHERE 
              tempIdx.TbleName = BCol.TBLE  
          AND tempIdx.FldName  = BCol.NAME NO-ERROR.
  CREATE tempCol. 
  ASSIGN hColumn               = IF BCol.DATA-TYPE NE ? THEN hBrowse:ADD-LIKE-COLUMN(BCol.TBLE + '.':U + BCol.NAME)
                                 ELSE hBrowse:ADD-CALC-COLUMN("char":U,BCol.DEF-FORMAT,"":U,BCol.DEF-LABEL,BCol.SEQ)
         hColumn:LABEL         = BCol.LBL
         hColumn:LABEL-FGCOLOR = IF AVAIL tempIdx THEN ? ELSE 9
         hColumn:LABEL         = " ":U + IF ((IF BCol.DATA-TYPE NE ? THEN (BCol.DATA-TYPE = "INTEGER":U OR BCol.DATA-TYPE = "DECIMAL":U) ELSE FALSE) OR 
                                             (FONT-TABLE:GET-TEXT-WIDTH-PIXELS(hColumn:LABEL) + 16 > FONT-TABLE:GET-TEXT-WIDTH-PIXELS(STRING(cText,BCol.FRMAT)))) 
                                     THEN hColumn:LABEL + "      ":U  
                                     ELSE hColumn:LABEL
         hColumn:WIDTH-PIXELS  = MAX(FONT-TABLE:GET-TEXT-WIDTH-PIXELS(hColumn:LABEL,hBrowse:FONT),FONT-TABLE:GET-TEXT-WIDTH-PIXELS(IF BCol.DATA-TYPE = ? OR BCol.DATA-TYPE = "CHARACTER":U THEN STRING(cText,BCol.FRMAT) ELSE BCol.FRMAT,hBrowse:FONT))
         hCurCol               = IF NOT VALID-HANDLE(hCurCol) AND AVAIL tempIdx AND tempIdx.Prim = TRUE THEN hColumn ELSE hCurCol
         hBrowse:WIDTH-PIXELS  = hBrowse:WIDTH-PIXELS + hColumn:WIDTH-PIXELS
         tempCol.ColHandle     = hColumn
         tempCol.ColWidth      = hColumn:WIDTH-PIXELS
         tempCol.CalcAction    = IF BCol.DATA-TYPE NE ? THEN ? ELSE BCol.DISP-NAME
         tempCol.ColLbl        = BCol.LBL
         tempCol.DataType      = BCol.DATA-TYPE
         tempCol.AllowSort     = BCol.DATA-TYPE NE ?.
  ASSIGN tempCol.AscDesc       = tempIdx.ascdesc WHEN AVAIL tempIdx.
 
  /* create a callobject for each calc-field */
  IF tempCol.CalcAction NE ? THEN DO:
    CREATE CALL hCall.
    ASSIGN hCall:IN-HANDLE      = hSuper
           hCall:CALL-TYPE      = FUNCTION-CALL-TYPE 
           hCall:CALL-NAME      = ENTRY(1,BCol.DISP-NAME) 
           hCall:NUM-PARAMETERS = hQuery:NUM-BUFFERS
           tempCol.CallHandle   = hCall. 
  END.

  /* for each browsecolumn create a filterfield above it when sorts on the column are allowed */
  IF tempCol.AllowSort
  THEN DO:
    CREATE FILL-IN hFill ASSIGN 
           X                 = iXcoordinate
           Y                 = 1
           WIDTH-PIXELS      = hColumn:WIDTH-PIXELS 
           HEIGHT-PIXELS     = 20
           FRAME             = hFrame 
           SENSITIVE         = TRUE
           NAME              = BCol.TBLE + '|':U + BCol.NAME
           DATA-TYPE         = /* IF BCol.DATA-TYPE = 'DECIMAL':u THEN 'INTEGER' ELSE IF BCol.DATA-TYPE = 'DATE':u THEN 'CHARACTER':U ELSE BCol.DATA-TYPE */ 'CHARACTER':U
                                   /* Convert decimals to integer as zero's and other signs would distort the query-prepare, date to character for the same reason */
           FORMAT            = IF BCol.DATA-TYPE NE 'LOGICAL':u THEN REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(BCol.FRMAT,'9':u,/* IF BCol.DATA-TYPE = 'INTEGER':U OR BCol.DATA-TYPE = 'DECIMAL':U THEN '>':u ELSE */ 'X':U),'!':U,'X':U),'%':U,'':U),'.':U,'':U),',':U,'':U),'/':U,'':U),'-':U,'':U),'>':U,'X':U)
                               ELSE 'X(3)':U
                                   /* zero's and other signs in the screen-value would distort the query-prepare */ 
           TOOLTIP           = "Press enter to filter"
           PRIVATE-DATA      = BCol.DATA-TYPE
    TRIGGERS:
      ON RETURN PERSISTENT RUN startSearch IN THIS-PROCEDURE.
    END TRIGGERS.

    CREATE tempFillin.
    ASSIGN tempFillin.FillinHandle = hFill
           tempFillin.ColHandle    = tempCol.ColHandle.
  END.
  iXcoordinate = iXcoordinate + hColumn:WIDTH-PIXELS + 4.
END.

hQuery:QUERY-PREPARE(cBaseQuery). 
hQuery:QUERY-OPEN.

ASSIGN hBrowse:Y                  = IF VALID-HANDLE(hFill) THEN hBrowse:Y + hFill:HEIGHT-PIXELS ELSE hBrowse:Y
       hBrowse:HEIGHT-PIXELS      = ((hBrowse:ROW-HEIGHT-PIXELS + 4) * {&NUMROWS}) + 18
       hFrame:WIDTH-PIXELS        = hBrowse:WIDTH-PIXELS  + IF VALID-HANDLE(hWin) THEN (2 * {&BORDERWIDTH}) ELSE (4 * {&BORDERWIDTH})
       hFrame:HEIGHT-PIXELS       = hBrowse:Y + hBrowse:HEIGHT-PIXELS  + (2 * {&BORDERHEIGHT}) + IF NOT VALID-HANDLE(hWin) THEN (hCancel:HEIGHT-PIXELS + iTitleBarHeight) ELSE 0.
IF VALID-HANDLE(hWin) 
THEN ASSIGN
       hWin:HEIGHT-PIXELS         = hFrame:HEIGHT-PIXELS
       hWin:WIDTH-PIXELS          = hFrame:WIDTH-PIXELS
       hWin:MAX-HEIGHT-PIXELS     = hWin:HEIGHT-PIXELS 
       hWin:MAX-WIDTH-PIXELS      = hWin:WIDTH-PIXELS
       hWin:VIRTUAL-HEIGHT-PIXELS = hWin:HEIGHT-PIXELS
       hWin:VIRTUAL-WIDTH-PIXELS  = hWin:WIDTH-PIXELS
       hWin:VISIBLE               = TRUE.
ELSE ASSIGN 
      hCancel:X       = hFrame:WIDTH-PIXELS - (3 * {&BORDERWIDTH}) - hCancel:WIDTH-PIXELS
      hCancel:Y       = hFrame:HEIGHT-PIXELS - {&BORDERHEIGHT} - hCancel:HEIGHT-PIXELS - iTitleBarHeight
      hOk:X           = hCancel:X - hCancel:WIDTH-PIXELS - 5
      hOk:Y           = hCancel:Y
      hCancel:VISIBLE = YES
      hOk:VISIBLE     = YES.

/* define taborder */ 

ASSIGN hEntry = hFrame:FIRST-CHILD.

FOR EACH tempFillin BREAK BY tempFillin.FillinHandle:
    IF FIRST(tempFillin.FillinHandle) THEN 
      ASSIGN hEntry:FIRST-TAB-ITEM = tempFillin.FillinHandle
             hEntry                = tempFillin.FillinHandle. 
    ELSE tempFillin.FillinHandle:MOVE-AFTER-TAB-ITEM(hPrevFillin).
    hPrevFillin = tempFillin.FillinHandle.
END.

IF NOT VALID-HANDLE(hPrevFillin) THEN
    ASSIGN hEntry:FIRST-TAB-ITEM = hBrowse
           hLastTabItem          = hBrowse.
ELSE hLastTabItem = hPrevFillin.

IF VALID-HANDLE(hLastTabItem) THEN lDummy = hBrowse:MOVE-AFTER-TAB-ITEM(hLastTabItem).
IF VALID-HANDLE(hOk)          THEN lDummy = hOk:MOVE-AFTER-TAB-ITEM(hBrowse).
IF VALID-HANDLE(hCancel)      THEN lDummy = hCancel:MOVE-AFTER-TAB-ITEM(hOk). 
ASSIGN hCaller        = SOURCE-PROCEDURE
       hFrame:VISIBLE = TRUE.

/* force row-display trigger to populate calcfields */
APPLY "END":U TO hBrowse.
APPLY "HOME":U TO hBrowse.

CREATE CONTROL-FRAME CtrlFrame ASSIGN
         FRAME           = hFrame
         ROW             = 1
         COLUMN          = 1
         HEIGHT          = 1
         WIDTH           = 1
         NAME = "CtrlFrame":U.

OCXFile = SEARCH( "browse-exp.wrx":U ).
IF OCXFile = ? THEN
       OCXFile = SEARCH(SUBSTRING(THIS-PROCEDURE:FILE-NAME, 1,
                     R-INDEX(THIS-PROCEDURE:FILE-NAME, ".":U), "CHARACTER":U) + "wrx":U).
IF OCXFile <> ? THEN
  ASSIGN
    chCtrlFrame = CtrlFrame:COM-HANDLE
    UIB_S = chCtrlFrame:LoadControls( OCXFile, "CtrlFrame":U)
  .
 /*------------------------------------------------------------------------------
   Purpose:     listen for WM_ERASEBKGND messages
 ------------------------------------------------------------------------------*/
   chCtrlFrame:Msgblst32:MsgList(0)    = {&WINMESSAGE}.    /* = WM_ERASEBKGND */ 
   chCtrlFrame:Msgblst32:MsgPassage(0) = -1.    /* = let PSC handle the message first */
   chCtrlFrame:Msgblst32:hWndTarget    = hFrame:HWND. 

RUN CtrlFrame.Msgblst32.Message ({&WINMESSAGE},1522,0,input-output i).

IF NOT VALID-HANDLE(hWin) THEN DO:
    RUN setEntry.
    WAIT-FOR GO OF hFrame.
END.
     
PROCEDURE cancelButton: 
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/ 
  RUN DISABLE_UI.
END PROCEDURE.

PROCEDURE CtrlFrame.Msgblst32.Message. 
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters:  Required for OCX.
    MsgVal
    wParam
    lParam
    lplRetVal
  Notes:       On purpose not defined private  
------------------------------------------------------------------------------*/

DEFINE INPUT        PARAMETER p-MsgVal    AS INTEGER NO-UNDO.
DEFINE INPUT        PARAMETER p-wParam    AS INTEGER NO-UNDO.
DEFINE INPUT        PARAMETER p-lParam    AS INTEGER NO-UNDO.
DEFINE INPUT-OUTPUT PARAMETER p-lplRetVal AS INTEGER NO-UNDO.

DEFINE VARIABLE hdc                AS INTEGER NO-UNDO.
DEFINE VARIABLE hdc_bitmap         AS INTEGER NO-UNDO.
DEFINE VARIABLE last_bitmap_handle AS INTEGER NO-UNDO.

IF NOT VALID-HANDLE(hBrowse) THEN RETURN.

IF hBrowse:VISIBLE AND NOT ENTRY(1,PROGRAM-NAME(2),' ':U) = 'endResize':U  THEN PROCESS EVENTS. /* Necessary on at least some machines to make the (re)paint work  */

CASE p-MsgVal :
   WHEN {&WINMESSAGE} /* = WM_ERASEBKGND */ THEN   
     FOR EACH tempCol: 
       /* Do a 'next' in case no bitmap has to be painted on the columnheader. This is the case when:
          - sorting on the column is allowed but the column is not the current one while there is a valid current-column
          - sorting on the column is allowed but the column is not the one on the primairy index while there is no valid current-column 
            (a asc/desc bitmap has to be painted if the query is just opened on the 'default' sorting column, which is the first field in the primary index)
          - 
       */  
       IF (
            (
                (
                     VALID-HANDLE(hBrowse:CURRENT-COLUMN) = TRUE
                     AND tempCol.ColHandle NE hBrowse:CURRENT-COLUMN 
                ) 

                OR 

                (
                     VALID-HANDLE(hBrowse:CURRENT-COLUMN) = FALSE 
                     AND tempCol.ColHandle NE hCurCol
                )
            ) AND tempCol.AllowSort = TRUE
          )
          
         OR tempCol.ColHandle:WIDTH-PIXELS < 16 THEN NEXT.
       
       RUN GetDC (hBrowse:HWND, 
                  OUTPUT hdc).

       IF hdc <> 0 THEN DO:
         RUN LoadImageA (0, /* hdc */
                         IF tempCol.AllowSort = FALSE THEN SEARCH('nosorts.bmp':U) ELSE IF tempCol.AscDesc = "asc":u THEN SEARCH('up.bmp') ELSE SEARCH('down.bmp':U), 
                         0, /* IMAGE_BITMAP */
                         14, /* imagewidth */
                         14, /* imageheigth */
                         16 + 4096, /* LR_LOADFROMFILE + LR_LOADMAP3DCOLORS */ 
                         OUTPUT hIcon).      
         /* create other device context for bitmap */
         RUN CreateCompatibleDC(hdc, OUTPUT hdc_bitmap).
         /* use that bitmap */
         RUN SelectObject(hdc_bitmap, hIcon, OUTPUT last_bitmap_handle).  
         /* paint to window device context, 13369376 is copy from original bitmap */
         RUN BitBlt (hdc, tempCol.ColHandle:x + tempCol.ColHandle:WIDTH-PIXELS - 14, 2, 12, 12, hdc_bitmap, 0, 0, 13369376, OUTPUT ret).
          /* clear the bitmap object */
         RUN SelectObject(hdc_bitmap, last_bitmap_handle, OUTPUT ret).
         /* delete unused device context */
         RUN DeleteDC(hdc_bitmap, OUTPUT ret).

         RUN ReleaseDC (hBrowse:HWND, 
                        hdc, 
                        OUTPUT ret).
         RUN DeleteObject(hIcon, OUTPUT ret). 
       END. /* if hdc<>0 */
     END. /* FOR EACH tempCol */
  END CASE.
END PROCEDURE.

PROCEDURE DISABLE_UI: 
/*------------------------------------------------------------------------------
  Purpose:     remove all objects created in this-procedure from memory 
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/ 
  DELETE OBJECT hQuery.
  FOR EACH tempBuffer:
    DELETE OBJECT tempBuffer.Hdl.
  END.
  DELETE WIDGET hBrowse.
  IF VALID-HANDLE(hSuper) THEN DELETE PROCEDURE hSuper.
  FOR EACH tempCol:
    IF VALID-HANDLE(tempCol.CallHandle) THEN DELETE OBJECT tempCol.CallHandle.
  END.
  FOR EACH tempFillin:
    DELETE WIDGET tempFillin.FillinHandle.
  END.
  IF NOT VALID-HANDLE(hWin) THEN APPLY 'go':U TO hFrame.
  hFrame:VISIBLE = FALSE.
  DELETE WIDGET hFrame.
  IF VALID-HANDLE(hWin) THEN DELETE WIDGET hWin.
  IF THIS-PROCEDURE:PERSISTENT THEN DELETE PROCEDURE THIS-PROCEDURE. 
END PROCEDURE.

PROCEDURE endResize PRIVATE:
/*------------------------------------------------------------------------------
  Purpose:    Change column-label(s) of resized colums & resize associated
              filter-fillins 
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/
  DEFINE VARIABLE hCurrentCol AS HANDLE    NO-UNDO.

  FOR EACH tempCol: 
     ASSIGN hCurrentCol         = tempCol.ColHandle.
                                                                                                         
     IF hCurrentCol:WIDTH-PIXELS NE tempCol.ColWidth THEN
     DO: 
        ASSIGN iPrevWidth       = tempCol.ColWidth 
               tempCol.ColWidth = tempCol.ColHandle:WIDTH-PIXELS NO-ERROR.

        IF (FONT-TABLE:GET-TEXT-WIDTH-PIXELS(hCurrentCol:LABEL)  + 16  > hCurrentCol:WIDTH-PIXELS) OR 
           (iPrevWidth < hCurrentCol:WIDTH-PIXELS AND hCurrentCol:LABEL NE " " + tempCol.ColLbl) THEN
               hCurrentCol:LABEL = " " + (IF (FONT-TABLE:GET-TEXT-WIDTH-PIXELS(tempCol.ColLbl) + 16 <= hCurrentCol:WIDTH-PIXELS) THEN tempCol.ColLbl
                                          ELSE IF hCurrentCol:WIDTH-PIXELS <= 16 THEN "":U
                                          ELSE getString(hCurrentCol)
                                         ) 
                                       + IF tempCol.DataType = "integer":U THEN FILL(" ":U,6) ELSE "":u. 
     END.
     FIND FIRST tempFillin 
        WHERE tempFillin.ColHandle = hCurrentCol
      NO-ERROR.
     
     IF AVAIL tempFillin THEN DO:
         IF hCurrentCol:X + hCurrentCol:WIDTH-PIXELS <= hBrowse:WIDTH-PIXELS THEN
            ASSIGN tempFillin.FillinHandle:WIDTH-PIXELS = IF hCurrentCol:X + hCurrentCol:WIDTH-PIXELS >= hBrowse:X THEN hCurrentCol:WIDTH-PIXELS 
                                                          ELSE 1
                   tempFillin.FillinHandle:X            = IF hCurrentCol:X >= hBrowse:X THEN hCurrentCol:X + hBrowse:X ELSE hBrowse:X
                   tempFillin.FillinHandle:VISIBLE      = TRUE NO-ERROR.
         ELSE IF hCurrentCol:X <= hBrowse:X + hBrowse:WIDTH-PIXELS THEN 
            ASSIGN tempFillin.FillinHandle:WIDTH-PIXELS = hBrowse:WIDTH-PIXELS - hCurrentCol:X
                   tempFillin.FillinHandle:X            = hCurrentCol:X
                   tempFillin.FillinHandle:VISIBLE      = TRUE NO-ERROR.
         ELSE ASSIGN tempFillin.FillinHandle:X       = hBrowse:X
                     tempFillin.FillinHandle:VISIBLE = FALSE.
     END.
     ERROR-STATUS:ERROR = FALSE.
  END.
  PROCESS EVENTS.
  RUN CtrlFrame.Msgblst32.Message ({&WINMESSAGE},1522,0,input-output i).
END PROCEDURE.

PROCEDURE okButton PRIVATE:
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/
  DEFINE VARIABLE iTel         AS INTEGER   NO-UNDO.
  DEFINE VARIABLE hWidget      AS HANDLE    NO-UNDO.
  DEFINE VARIABLE cWidgetName  AS CHARACTER NO-UNDO.
  DEFINE VARIABLE cFieldName   AS CHARACTER NO-UNDO.

  FIND FIRST tempBuffer. 

  IF NUM-ENTRIES(cPrivateParts,'|':U) > 2 THEN
  DO iTel = 1 TO NUM-ENTRIES(ENTRY(3,cPrivateParts,'|':U),',':U):
      ASSIGN hWidget     = hStartingWidget:FRAME:FIRST-CHILD:FIRST-CHILD
             cWidgetName = ENTRY(1,ENTRY(iTel,ENTRY(3,cPrivateParts,'|':U),',':U),'=':U)
             cFieldName  = ENTRY(2,ENTRY(iTel,ENTRY(3,cPrivateParts,'|':U),',':U),'=':U).
      DO WHILE VALID-HANDLE(hWidget):
          IF hWidget:NAME = cWidgetName 
          THEN DO:
              hWidget:SCREEN-VALUE = STRING(tempBuffer.Hdl:BUFFER-FIELD(cFieldName):BUFFER-VALUE()).
              LEAVE.
          END.
          hWidget = hWidget:NEXT-SIBLING.
      END.
  END.
  RUN DISABLE_UI.
END PROCEDURE.

PROCEDURE rowDisplay PRIVATE:
/*------------------------------------------------------------------------------
  Purpose:     Assign screen-value of calc-fields
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/ 
  FOR EACH tempCol WHERE tempCol.CalcAction NE ?:
    IF VALID-HANDLE(tempCol.ColHandle) AND VALID-HANDLE(tempCol.CallHandle) THEN 
    DO:  
      DO iNumPars = 1 TO tempCol.CallHandle:NUM-PARAMETERS: 
        ASSIGN hBuffer = hQuery:GET-BUFFER-HANDLE(iNumPars).
        tempCol.CallHandle:SET-PARAMETER(iNumPars,"HANDLE":U,"INPUT":U, hBuffer).
      END. 
      tempCol.CallHandle:INVOKE().
      tempCol.ColHandle:SCREEN-VALUE = tempCol.CallHandle:RETURN-VALUE. 
    END.
  END. 
END PROCEDURE.

PROCEDURE setEntry:
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/
  APPLY "entry":U TO hEntry.
END PROCEDURE.

PROCEDURE startSearch PRIVATE:
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/ 
  openQuery(). 
  RUN CtrlFrame.Msgblst32.Message ({&WINMESSAGE},1522,0,input-output i). 
END PROCEDURE.

FUNCTION getString RETURNS CHARACTER PRIVATE (INPUT pihCurrentCol AS HANDLE):
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/
  DEFINE VARIABLE iSubtract     AS INTEGER   NO-UNDO.
  DEFINE VARIABLE cReturnString AS CHARACTER NO-UNDO.

  FIND tempCol WHERE tempCol.ColHandle = pihCurrentCol NO-ERROR.
  IF AVAIL tempCol THEN 
  DO WHILE TRUE:
    ASSIGN iSubtract     = iSubtract + 1
           cReturnString = IF LENGTH(tempCol.ColLbl) - iSubtract > 0 THEN
                            SUBSTRING(tempCol.ColLbl,
                                      1,
                                      LENGTH(tempCol.ColLbl) - iSubtract) 
                           ELSE "":U 
           cReturnString = IF iSubtract > 3 AND LENGTH(cReturnString) = 1 THEN cReturnString + 
                              (IF FONT-TABLE:GET-TEXT-WIDTH-PIXELS(" " + cReturnString + "..":U) < pihCurrentCol:WIDTH-PIXELS - 16 THEN "..":U
                               ELSE IF FONT-TABLE:GET-TEXT-WIDTH-PIXELS(" " + cReturnString + ".":U) < pihCurrentCol:WIDTH-PIXELS - 16 THEN ".":U
                               ELSE "":U)
                           ELSE IF LENGTH(cReturnString) = 1 THEN cReturnString + 
                            (IF FONT-TABLE:GET-TEXT-WIDTH-PIXELS(" " + cReturnString + ".":U) < pihCurrentCol:WIDTH-PIXELS - 16 THEN ".":U 
                             ELSE "":U)
                           ELSE cReturnString + (IF cReturnString NE "" THEN "...":U ELSE ""). 
                                       
    IF FONT-TABLE:GET-TEXT-WIDTH-PIXELS(cReturnString) < pihCurrentCol:WIDTH-PIXELS - 16 OR pihCurrentCol:WIDTH-PIXELS <= 16 THEN LEAVE.
  END.
  RETURN cReturnString.
END FUNCTION.

FUNCTION openQuery RETURNS LOGICAL PRIVATE:
/*------------------------------------------------------------------------------
  Purpose:     
  Parameters: 
  Notes:       
------------------------------------------------------------------------------*/
  DEFINE VARIABLE hTemp        AS HANDLE    NO-UNDO. 
  DEFINE VARIABLE cFilter      AS CHARACTER NO-UNDO.
  DEFINE VARIABLE cIndFlds     AS CHARACTER NO-UNDO.
  DEFINE VARIABLE cFilterQuery AS CHARACTER NO-UNDO.
  DEFINE VARIABLE iTeller      AS INTEGER   NO-UNDO.  

  FIND tempCol WHERE tempCol.ColHandle = IF VALID-HANDLE(hBrowse:CURRENT-COLUMN) THEN hBrowse:CURRENT-COLUMN ELSE hPrevCol NO-ERROR.
  
  IF VALID-HANDLE(hBrowse:CURRENT-COLUMN) AND AVAIL tempCol AND tempCol.AllowSort = TRUE THEN 
    ASSIGN tempCol.AscDesc = IF tempCol.AscDesc = "ASC":u AND hPrevCol =  hBrowse:CURRENT-COLUMN THEN "DESC":U ELSE "ASC":U
           hPrevCol        = hBrowse:CURRENT-COLUMN 
           hTemp           = hPrevCol:BUFFER-FIELD.

  EMPTY TEMP-TABLE tempNewQuerySlice.

  FOR EACH tempQuerySlice:
     ASSIGN cFilter   = "":U
            cIndFlds  = "":U.
     EMPTY TEMP-TABLE tempFilterElement.
     FOR EACH tempFillin: /* build initial filterstring, finally to be used in building the new query-string  */
        IF ENTRY(1,tempFillin.FillinHandle:NAME,'|':U) = tempQuerySlice.cBuffer 
             AND tempFillin.FillinHandle:SCREEN-VALUE NE "":U 
        THEN DO:
            CREATE tempFilterElement.
            ASSIGN tempFilterElement.cField    = ENTRY(2,tempFillin.FillinHandle:NAME,'|':U)
                   tempFilterElement.lChartype = tempFillin.FillinHandle:DATA-TYPE = "CHARACTER":u 
                   tempFilterElement.cElement  = (IF tempFillin.FillinHandle:PRIVATE-DATA NE "CHARACTER":u THEN "STRING(":u ELSE "":u) 
                                                   + REPLACE(tempFillin.FillinHandle:NAME,'|':U,'.':U) 
                                                   + (IF tempFillin.FillinHandle:PRIVATE-DATA NE "CHARACTER":u THEN ")":u ELSE "":u) 
                                                   + " BEGINS ":U + "~'" + tempFillin.FillinHandle:SCREEN-VALUE + "~' ":U.
        END.
     END.

     FIND FIRST tempFilterElement 
         WHERE tempFilterElement.lChartype = TRUE 
         NO-ERROR.
     IF AVAIL tempFilterElement THEN /* assign number (iEntryNum) in filterstring concerning index-use, with no respect to eventual whereclause
                                        (could be done later) in tempQuerySlice.cSlice */
     FOR EACH tempIdx WHERE tempIdx.TbleName = tempQuerySlice.cBuffer:
        FIND FIRST tempFilterElement 
            WHERE tempFilterElement.cField    = tempIdx.FldName 
              AND tempFilterElement.lChartype = TRUE 
              AND tempFilterElement.lUsed     = FALSE
            NO-ERROR. 

        IF AVAIL tempFilterElement 
        THEN DO: 
            FIND LAST buftempFilterElement. /* last iEntryNum to count from */
            
            ASSIGN tempFilterElement.iEntryNum = buftempFilterElement.iEntryNum + 1
                   tempFilterElement.lUsed     = YES. 
            IF tempIdx.FldName2 NE '':U 
            THEN DO:
               FIND FIRST tempFilterElement 
                   WHERE tempFilterElement.cField    = tempIdx.FldName2 
                     AND tempFilterElement.lChartype = TRUE 
                     AND tempFilterElement.lUsed     = FALSE
                   NO-ERROR. 
               IF AVAIL tempFilterElement
               THEN DO: 
                  ASSIGN tempFilterElement.iEntryNum = buftempFilterElement.iEntryNum + 2
                         tempFilterElement.lUsed     = YES.  
                  IF tempIdx.FldName3 NE '':U 
                  THEN DO:
                     FIND FIRST tempFilterElement 
                         WHERE tempFilterElement.cField    = tempIdx.FldName3 
                           AND tempFilterElement.lChartype = TRUE
                           AND tempFilterElement.lUsed     = FALSE
                         NO-ERROR. 
                     IF AVAIL tempFilterElement
                     THEN ASSIGN tempFilterElement.iEntryNum = buftempFilterElement.iEntryNum + 3
                                 tempFilterElement.lUsed     = YES.  
                  END.
               END.
           END.
        END.
     END.
     FIND LAST buftempFilterElement NO-ERROR.

     /* now build the filterstring, first assign the left elements a iEntryNum  */
     FOR EACH tempFilterElement: 
         IF tempFilterElement.iEntryNum = 0
         THEN ASSIGN iTeller                     = iTeller + 1
                     tempFilterElement.iEntryNum = buftempFilterElement.iEntryNum + iTeller.
         ASSIGN cFilter = IF tempFilterElement.iEntryNum = 1 THEN tempFilterElement.cElement 
                          ELSE cFilter + ' AND ':U + tempFilterElement.cElement.
     END.

     CREATE tempNewQuerySlice.
     ASSIGN tempNewQuerySlice.cBuffer = tempQuerySlice.cBuffer
            tempNewQuerySlice.cSlice  = IF cFilter = '':U THEN tempQuerySlice.cSlice
                                        ELSE IF INDEX(tempQuerySlice.cSlice," WHERE ":U) = 0 THEN 
                                            IF INDEX(tempQuerySlice.cSlice," OUTER-JOIN ":U) = 0 THEN REPLACE(tempQuerySlice.cSlice," NO-LOCK":U, " WHERE ":U + cFilter + " NO-LOCK":U)
                                            ELSE REPLACE(tempQuerySlice.cSlice," OUTER-JOIN ":U, " WHERE ":U + cFilter + " OUTER-JOIN ":U)
                                        ELSE IF INDEX(tempQuerySlice.cSlice," OUTER-JOIN ":U) = 0 THEN REPLACE(tempQuerySlice.cSlice," NO-LOCK":U, " AND ":U + cFilter + " NO-LOCK":U) 
                                        ELSE REPLACE(tempQuerySlice.cSlice," OUTER-JOIN ":U, " AND ":U + cFilter + " OUTER-JOIN ":U).
  END.

  FOR EACH tempNewQuerySlice:
     cFilterQuery = cFilterQuery + tempNewQuerySlice.cSlice.
  END.

  cFilterQuery = "FOR ":U + REPLACE(cFilterQuery,'INDEXED-REPOSITION':U,'':U) + (IF AVAIL tempCol AND tempCol.AllowSort = TRUE THEN (" BY ":U + hPrevCol:BUFFER-FIELD:NAME +
                        IF tempCol.AscDesc = "DESC":u THEN " ":U + tempCol.AscDesc ELSE "":u) ELSE "":U) + IF lIndexedRepos THEN ' INDEXED-REPOSITION':U ELSE '':U.
/*   MESSAGE 'cFilterQuery ' cFilterQuery. */
  hQuery:QUERY-PREPARE(cFilterQuery).
  hQuery:QUERY-OPEN.
  RETURN FALSE.
END FUNCTION.
