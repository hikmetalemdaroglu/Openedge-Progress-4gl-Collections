/* hikmetalemdaroglu@gmail.com */
FUNCTION FDAYNAME return char (input p-date as date):
  DEFINE VARIABLE daynum  AS INTEGER   NO-UNDO.
  DEFINE VARIABLE dayname AS CHARACTER NO-UNDO.
  DEFINE VARIABLE daylist AS CHARACTER NO-UNDO FORMAT "x(9)"
    INITIAL "Pazar,Pazartesi,Sali,Carsamba,Persembe,Cuma,Cumartesi".
  ASSIGN 
   daynum  = WEEKDAY(P-DATE)
   dayname = ENTRY(daynum, daylist). 
  RETURN STRING(DAYNUM) +  "|" + DAYNAME.
END FUNCTION.

// DEF VAR DAY AS CHAR. 
// DAY = FDAYNAME(TODAY). 
// MESSAGE DAY VIEW-AS ALERT-BOX INFO BUTTONS OK.
