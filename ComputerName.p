USING System.* .
USING System.Net.* .

DEFINE OUTPUT PARAM pr-strMachineName AS CHAR NO-UNDO.
DEFINE OUTPUT PARAM PR-LOCALIP AS CHAR FORMAT "X(12)" NO-UNDO EXTENT 5.

DEFINE VARIABLE ipHost  AS IPHostEntry NO-UNDO .
DEFINE VARIABLE ipAddr  AS IPAddress   NO-UNDO EXTENT .
DEF VAR IX AS INT.

ASSIGN pr-strMachineName = Dns:GetHostName().
ASSIGN ipHost = Dns:GetHostByName(pr-strMachineName).
ASSIGN ipAddr = ipHost:AddressList.
REPEAT IX = 1 TO 5:
   PR-LOCALIP[IX] = ipAddr[IX]:ToString() NO-ERROR.
END.
