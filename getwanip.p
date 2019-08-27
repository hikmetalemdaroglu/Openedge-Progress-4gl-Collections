/*  GET WAN IP FROM ipify.org     
    for xml or json output see ipify.org
    hikmetalemdaroglu@gmail.com          
*/

DEF OUTPUT PARAM VOUT_STR AS CHAR NO-UNDO. 
DEF VAR VURL     AS CHAR NO-UNDO.

ASSIGN VURL = "http://api.ipify.org".

RUN PR-GET_WANIP (INPUT VURL, OUTPUT VOUT_STR).    

PROCEDURE PR-GET_WANIP:
    DEF INPUT  PARAM PR-URL AS CHAR NO-UNDO.
    DEF OUTPUT PARAM PR-STR AS CHAR NO-UNDO.
    DEF VAR hobj AS com-handle no-undo.

    CREATE "MSXML2.ServerXMLHTTP" hobj NO-ERROR.
    NO-RETURN-VALUE HOBJ:setTimeouts(5000,5000,15000,15000) NO-ERROR.
    NO-RETURN-VALUE HOBJ:open("GET", PR-URL, "FALSE") NO-ERROR .
    NO-RETURN-VALUE HOBJ:setRequestHeader("Content-Type", "application/x-www-form-urlencoded") NO-ERROR.
    NO-RETURN-VALUE HOBJ:SEND(PR-URL) NO-ERROR.
    PR-STR = hOBJ:responseText NO-ERROR.
    RELEASE OBJECT HOBJ.
END PROCEDURE.

