/*------------------------------------------------------------------------

  File:          special/_delxftr.p 
  
  Description:   Remove XFTR from template. Used on write-event of the XFTR               
 
  Parameters:    

  Author:        Stefan Houtzager
  
  Created:       01/01/2005
  
  Remarks:          

  Revisions:
---------------------------------------------------------------------------*/ 
DEFINE INPUT        PARAMETER trg-recid AS INTEGER   NO-UNDO.
DEFINE INPUT-OUTPUT PARAMETER trg-code  AS CHARACTER NO-UNDO.

DEFINE VARIABLE cResult AS CHARACTER NO-UNDO.

/* If the file is a template then don't mark the XFTR for removal */
RUN adeuib/_uibinfo.p (trg-recid,?,"TEMPLATE":U, OUTPUT cResult). 

IF cResult = "YES" THEN RETURN.

/* mark */
trg-code = "/* verwijder XFTR */":U.
RETURN.
