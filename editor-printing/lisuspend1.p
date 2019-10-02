/* ----------------------------------------------------------------------------------- *
   Program : LISUSPEND1.P       Path : LI                                    
   Aciklama: ISTENEN OLCU BIRIMI VE SURE KADAR SISTEMI BEKLETIR.
   UYARLAMA:1.0         ILK YAZIM    : 22/09/2005 YZN : ha. 
   UYARLAMA:2.0         SON Uyarlama :   /  /     YZN :                
   -------------------------------------------------------------------------------------
   Telif Hakký © 2002-2004 Fokus Yazýlým Tek.San.Tic.Ltd.
   Tüm Haklarý Saklýdýr. Bu döküman kopyalanamaz.
   
   Copyright 2002-2004 © Fokus Yazýlým Tek.San.Tic.Ltd.
   All rights reserved worldwide.  This is an unpublished work. 
   --------------------------------------------------------------------------------------
       Y   = YEAR
       M   = MONTH
       D   = DAY
       H   = HAUR
       MN  = MINUTE
       S   = SECOND
   -------------------------------------------------------------------------------------- */  
DEF INPUT PARAM PR-UNIT   AS CHAR NO-UNDO.
DEF INPUT PARAM PR-VALUE  AS INT  NO-UNDO.

DEF VAR D-CURRTIME AS INT                       NO-UNDO.
DEF VAR D-NEXTTIME AS INT                       NO-UNDO.
DEF VAR D-PERIOD   AS DECI FORMAT ">>>,>>>,>>9" NO-UNDO.

IF PR-UNIT = "Y"  THEN D-PERIOD = DECIMAL(PR-VALUE * 365 * 24 * 60 * 60).

IF PR-UNIT = "M"  THEN D-PERIOD = DECIMAL(PR-VALUE * 30  * 24 * 60 * 60).
IF PR-UNIT = "D"  THEN D-PERIOD = DECIMAL(PR-VALUE * 24  * 60 * 60).
IF PR-UNIT = "H"  THEN D-PERIOD = DECIMAL(PR-VALUE * 60  * 60).
IF PR-UNIT = "MN" THEN D-PERIOD = DECIMAL(PR-VALUE * 60).
IF PR-UNIT = "S"  THEN D-PERIOD = PR-VALUE.

ASSIGN D-CURRTIME = TIME.
ASSIGN D-NEXTTIME = D-CURRTIME + D-PERIOD.


REPEAT:
    IF TIME >= D-NEXTTIME THEN LEAVE.
END.
