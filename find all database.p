DEFINE VARIABLE i AS INTEGER NO-UNDO.
 
DO i = 1 TO NUM-DBS:
    MESSAGE
        LDBNAME(i) SKIP /* Database logical name */
        PDBNAME(i) SKIP /* Database physical name */
        DBPARAM(i) SKIP /* Database connection parameters */
        DBTYPE(i) SKIP  /* Database type */
        DBVERSION(i) SKIP /* Database version */
        DBRESTRICTION(i)SKIP /* Database unsupported features */
        SDBNAME(i) SKIP /* Database or schema holder logical name */
        DBCOLLATION(i) SKIP /* Database collating sequence */
        DBCODEPAGE(i) /* Database code page */
        VIEW-AS ALERT-BOX INFO BUTTONS OK.
END.