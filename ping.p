/* ----------------------------------------------------------------- *
   -w timeout out.  
   -n trynum.
         
   DEF VAR VMSG AS CHAR NO-UNDO.
   
   Call Procedure 
   
		RUN PING.P (INPUT "hürriyet.con" , OUTPUT VMSG).
		MESSAGE VMSG VIEW-AS ALERT-BOX INFO BUTTONS OK.   
 * ---------------------------------------------------------------- */

DEF INPUT  PARAM PR-PING_ADR AS CHAR NO-UNDO.
DEF OUTPUT PARAM PR-MSG      AS CHAR NO-UNDO.

DEFINE VARIABLE cCommandLine  AS CHARACTER NO-UNDO.
DEFINE VARIABLE cImportedLine AS CHARACTER NO-UNDO.

ASSIGN cCommandLine = "c:\windows\system32\ping" + " " + PR-PING_ADR + " " + "-n 4 -w 10"
       PR-MSG = "900|" + PR-PING_ADR + " NO CONNECT".
INPUT THROUGH value(cCommandLine).
REPEAT:
  IMPORT UNFORMATTED cImportedLine.
  IF INDEX (cImportedLine, "Reply") > 0 OR 
     INDEX (cImportedLine, "Cevab") > 0 THEN DO:

     PR-MSG = "100|" + PR-PING_ADR + " CONNECTION OK".
     LEAVE.
  END.
END.
INPUT CLOSE.
