/* ---------------------------------------------------------------- *
 *                      ligunbul1.p                                 * 
 *                                                                  *
 *           Herhangi bir ayin en son gununu bulur.                 *
 *                                                                  *
 *             (c)... Hikmet ALEMDAROGLU 2000                       *
 * GIRILEN TARIHTEN SONRA GELEN ILK AYIN BIRINCI GUNUNU             *
 * BULUP BIR EKSILTEREK AYIN SON GUNUNU BULUR                       * 
 * 
 * LAST UPDATE : 21/06/2016 
   
  DEF VAR VIN_DATE  AS DATE NO-UNDO.
  DEF VAR VOUT_DATE AS DATE NO-UNDO.
  
  ASSIGN VIN_DATE = TODAY.
  RUN ORTAK/LASTDAYOFMONTH2.P (INPUT VIN_DATE, OUTPUT VOUT_DATE).
  
  MESSAGE VOUT_DATE VIEW-AS ALERT-BOX INFO BUTTONS OK.
 
 * --------------------------------------------------------------- */
   define input  param pr-indate  as date format "99/99/9999" no-undo.
   define output param pr-outdate as date format "99/99/9999" no-undo.
   
   def var d-day   as int format "99" no-undo.
   def var d-month as int format "99" no-undo.
   def var d-year  as int format "99" no-undo.

   d-day   = 1.
   d-month = month(pr-indate).
   d-year  = year(pr-indate).
   
   if d-month = 12 then assign d-year  = d-year + 1
                               d-month = 1.
   else d-month = d-month + 1.
   
   pr-outdate = date(d-month,d-day,d-year).
   pr-outdate = pr-outdate - 1.   
