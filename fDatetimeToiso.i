/* hikmetalemdaroglu@gmail.com */
FUNCTION FISO_DATE RETURN CHAR (INPUT PR-TIME AS CHAR):   
   DEF VAR VTIME AS DATETIME NO-UNDO.
   ASSIGN VTIME = DATETIME(PR-TIME).
   ASSIGN PR-TIME = REPLACE(SUBSTRING(ISO-DATE(VTIME),1,23),"/", "_").
   ASSIGN PR-TIME = REPLACE(PR-TIME,":","").     
   RETURN PR-TIME.
END FUNCTION.

// DEF VAR XX AS CHAR.
// XX = FISO_DATE(STRING(NOW)).