/* hikmetalemdaroglu@gmail.com */
FUNCTION FDATE_TO_DATETIME RETURN DATETIME(INPUT P-DATE AS DATE):  
  DEF VAR DT   AS DATETIME NO-UNDO. 
  DEF VAR VNOW AS CHAR NO-UNDO.
  VNOW = STRING(SUBSTRING(STRING(NOW),12,12)).
  ASSIGN dt = DATETIME(STRING(P-DATE,"99-99-9999") + " " + VNOW).
  RETURN DT.
END FUNCTION.
// DEF VAR AA AS DATETIME.
// DEF VAR VDATE AS DATE.
// VDATE = 05/25/2018.
// AA = FDATE_TO_DATETIME(VDATE).
// MESSAGE AA VIEW-AS ALERT-BOX INFO BUTTONS OK.


