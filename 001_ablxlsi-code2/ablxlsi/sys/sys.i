/* sys/sys.i SYSTEM RELATED FUNCTIONS AND DECLARATIONS
             KURT GUNDERSON 01/06/2009
   PURPOSE:

   MODIFICATIONS:
      01/06/2009 KAG Original.
      10/28/2009 KAG Added _cando() function.
      08/04/2010 KAG BUGFIX: _getUserName from os-getenv("USRENAME") failing
                     on appserver at client site; environmental variable
                     doesn't always exist.
      10/06/2010 KAG Added getFullPath()
      12/31/2010 KAG Added _getUptime() and _getDatabaseName()
      02/16/2011 KAG Scoped _user to internal procedure.
      05/04/2011 KAG Added _center()
      09/15/2011 KAG Added YEARLY LOGIC to _getDate().
      10/31/2012 KAG *nix compatibility
      12/17/2012 KAG Added _getComputerName().
      01/10/2013 KAG Added _validFile(), _validDirectory, _getDirectory().
      06/05/2013 KAG Added MSS logic to _numUserConnections().

   NOTES:
      PROCEDURES/FUNCTIONS IN THIS LIBRARY SHOULD NEVER REFERENCE DATABASE
      FIELDS/TABLES BUT CAN REFERENCE SYSTEM/VIRTUAL TABLES.

      _numUserConnections() RETURNS THE NUMBER OF CONNECTIONS/SESSIONS THE
      USER HAS OPEN.  FOR MSS SQL SERVER DATABASES, THIS IS ONLY TRUE OF
      GUI CLIENTS (NOT BOTH GUI AND CHUI CLIENTS) AND IS ONLY TRUE WHEN THE
      CONNECTION TO THE MSS DATABASE IS NOT LOGGED INTO VIA A SERVICE ACCOUNT.

*/
/******************************************************************************/
&IF DEFINED(SYS_I) = 0 &THEN
&GLOBAL-DEFINE SYS_I
define stream ssysi.

function _numUserConnections returns integer (input u as character):
   define variable i as integer initial 0 no-undo.
&IF DBTYPE("REACCT":U) = "MSS" &THEN

   define variable wqry_str as character no-undo.
   define variable wproc_id as integer no-undo.

   assign wqry_str = "select count(distinct sys.sysprocesses.hostprocess) from sys.sysprocesses "
                   + "   where sys.sysprocesses.program_name = N'rx.exe' "
                   + "     and sys.sysprocesses.nt_username  = N'" + u + "' "
          wproc_id = 0.
   run stored-proc reacct.send-sql-statement wproc_id = proc-handle no-error (wqry_str).
   if not error-status:error then
      for each reacct.proc-text-buffer where proc-handle = wproc_id:
         assign i = integer(reacct.proc-text-buffer.proc-text).
      end.
   if wproc_id > 0 then
      close stored-proc reacct.send-sql-statement where proc-handle = wproc_id.

&ELSEIF DBTYPE("REACCT":U) = "ORACLE" &THEN



&ELSE

   define buffer bf_connect for _connect.
   for each bf_connect where bf_connect._connect-name = u no-lock: assign i = i + 1. end.

&ENDIF   
   return i.
end function.

function _getFileSeparator returns character:
   return if opsys begins "WIN" then chr(92) 
          else chr(47).
end function.

function _validFile returns logical (input f as character):
   if f = ? then return false.
   else do:
      assign file-info:file-name = f.
      return not (file-info:file-type = ? or index(file-info:file-type,"F") <= 0 or index(file-info:file-type,"R") <= 0).
   end.
end function.

function _validDirectory returns logical (input f as character):
   if f = ? then return false.
   else do:
      assign file-info:file-name = f.
      return not (file-info:file-type = ? or index(file-info:file-type,"D") <= 0 or index(file-info:file-type,"R") <= 0 or index(file-info:file-type,"W") <= 0).
   end.
end function.

/* RETURNS THE DIRECTORY PORTION OF A RELATIVE OR ABSOLUTE PATHED FILE 
 */
function _getDirectory returns character (input f as character):
   define variable i      as integer no-undo.
   define variable wdir   as character initial "" no-undo.
   define variable xdl    as character initial "" no-undo.
   define variable wdl    as character initial "" no-undo.

   assign wdl = _getFileSeparator().

   if num-entries(f,wdl) < 2 then
      assign wdir = ".".
   else
      do i = 1 to num-entries(f,wdl) - 1:
         assign wdir = wdir + xdl + entry(i,f,wdl)
                xdl  = wdl.
      end.
   return wdir.
end function.

function _center returns character (input s as character, input l as integer):
   assign l = (l - length(s)) / 2
          s = fill(" ",l) + s.
   return s.
end function.

function _cando returns logical (input t as character,input s as character):
   return (lookup("*",s) <> 0 or lookup(t,s) <> 0) and not lookup("!" + t,s) <> 0.
end function.

function _getEOM returns date (input d as date):
   return (date(month(d),28,year(d)) + 4) - day(date(month(d),28,year(d)) + 4).
end function.

function _getBOM returns date (input d as date):
   return date(month(d),1,year(d)).
end function.

function _getDate returns date (input pd as date, 
                                input pt as character):
   case pt:
      when "TODAY" then
         assign pd = pd.
      when "YESTERDAY" or
      when "LASTDAY" then
         assign pd = pd - 1.
      when "TOMORROW" or 
      when "NEXTDAY" then
         assign pd = pd + 1.
      /* FIRST OF THE MONTH LOGIC */
      when "FIRSTDAYOFMONTH" or
      when "BOM" then
         assign pd = _getBOM(pd).
      when "NEXTFIRSTDAYOFMONTH" then
         assign pd = _getEOM(pd) + 1.
      when "LASTFIRSTDAYOFMONTH" then
         assign pd = _getBOM(_getBOM(pd) - 1).
      /* END OF THE MONTH LOGIC */
      when "LASTDAYOFMONTH" or
      when "EOM" then
         assign pd = _getEOM(pd).
      when "NEXTLASTDAYOFMONTH" then
         assign pd = _getEOM(_getEOM(pd) + 1).
      when "LASTLASTDAYOFMONTH" then
         assign pd = _getBOM(pd) - 1.
      /* WEEKDAY LOGIC */
      when "FIRSTDAYOFWEEK" then
         assign pd = pd - (weekday(pd) - 1).
      when "LASTDAYOFWEEK" then
         assign pd = pd + (7 - weekday(pd)).
      when "NEXTWEEKDAY" then
         repeat:
            assign pd = pd + 1.
            if weekday(pd) > 1 and weekday(pd) < 7 then leave.
         end.
      when "LASTWEEKDAY" then
         repeat:
            assign pd = pd - 1.
            if weekday(pd) > 1 and weekday(pd) < 7 then leave.
         end.
      /* QUARTER LOGIC */
      when "LASTFIRSTDAYOFQUARTER" then
         assign pd = _getDate(_getDate(pd,"LASTLASTDAYOFQUARTER"),"FIRSTDAYOFQUARTER").
      when "LASTLASTDAYOFQUARTER" then
         assign pd = _getDate(pd,"FIRSTDAYOFQUARTER") - 1.
      when "FIRSTDAYOFQUARTER" then
         assign pd = _getBOM(date(month(pd) - ((month(pd) - 1) mod 3),1,year(pd))).
      when "LASTDAYOFQUARTER" then
         assign pd = _getEOM(date(month(pd) - ((month(pd) - 1) mod 3) + 2,1,year(pd))).
      when "NEXTFIRSTDAYOFQUARTER" then
         assign pd = _getDate(pd,"LASTDAYOFQUARTER") + 1.
      when "NEXTLASTDAYOFQUARTER" then
         assign pd = _getDate(_getDate(pd,"NEXTFIRSTDAYOFQUARTER"),"LASTDAYOFQUARTER").
      /* YEARLY LOGIC */
      when "FIRSTDAYOFYEAR" then
         assign pd = date(1,1,year(pd)).
      when "LASTDAYOFYEAR" then
         assign pd = date(12,31,year(pd)).
      /* ELSE */
      otherwise
         assign pd = pd.
   end case.
   return pd.
end function.

function _getDateStamp returns character (input pd as date, 
                                         input pt as character):
   define variable wd as date no-undo.
   assign wd = _getDate(pd,pt).
   return string(year(wd),"9999") + string(month(wd),"99") + string(day(wd),"99").
end function.

function _getPeriodStamp returns character (input pd as date, 
                                            input pt as character):
   define variable wd as date no-undo.
   assign wd = _getDate(pd,pt).
   return string(year(wd),"9999") + string(month(wd),"99").
end function.

function _getUserName returns character:
   return userid(ldbname("DICTDB")).
end function.

function _getDatabaseName returns character:
   return ldbname("DICTDB").
end function.

function _getComputerName returns character:
   define variable wname as character no-undo.

   assign wname = if opsys begins "WIN" then os-getenv("COMPUTERNAME")
                  else os-getenv("HOSTNAME").
   if wname = ? and opsys = "UNIX" then do:
      input stream ssysi through "hostname -f" no-echo.
      import stream ssysi unformatted wname.
      input stream ssysi close.
   end.

   return wname.
end function.

function _getUptime returns integer:
   define buffer bf_actother for _actother.
   find first bf_actother no-lock.
   return bf_actother._other-uptime.
end function.

function _getHomeDirectory returns character:
   define variable woutput_dir as character no-undo.
   define variable hShell      as com-handle no-undo.
   define variable hFolder     as com-handle no-undo.

   if opsys begins "WIN" then do: /* WINDOWS */
      create 'WScript.Shell' hShell no-error.
      if valid-handle(hShell) then do:
         assign hFolder = hShell:SpecialFolders.
         file-info:file-name = hFolder:Item("MyDocuments") no-error.
      end.

      if file-info:full-pathname = ? then
         assign file-info:file-name = os-getenv("USERPROFILE") + "\Documents".
      if file-info:full-pathname = ? then
         assign file-info:file-name = os-getenv("USERPROFILE") + "\My Documents".
      if file-info:full-pathname = ? then do:
         load "SOFTWARE" base-key "HKEY_CURRENT_USER".
         use "SOFTWARE".
         get-key-value section "MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\SHELL FOLDERS"
            key "PERSONAL"
            value woutput_dir.
         unload "SOFTWARE".
         file-info:file-name = woutput_dir.
      end.

      if file-info:full-pathname = ? then
         assign file-info:file-name = ".".

      if valid-handle(hFolder) then release object hFolder no-error.
      if valid-handle(hShell) then release object hShell no-error.

      assign hShell  = ?
             hFolder = ?.
   end.
   else do:                    /* UNIX */
      assign file-info:file-name = os-getenv("HOME").

      if file-info:full-pathname = ? then
         assign file-info:file-name = ".".
   end.

   return file-info:full-pathname + _getFileSeparator().
end function.

function _getPID returns integer:
   define buffer bf_myconnection for _myconnection.
   for first bf_myconnection no-lock: end.
   return bf_myconnection._myconn-pid.
end function.

function _isInteger returns logical (input c as character):
   define variable x as integer no-undo.
   assign x = integer(c) no-error.
   return (x <> ?) and (error-status:error = no).
end function.

function _isDate returns logical (input c as character):
   define variable x as date no-undo.
   assign x = date(c) no-error.
   return (x <> ?) and (error-status:error = no).
end function.

function _isDecimal returns logical (input c as character):
   define variable x as decimal no-undo.
   assign x = decimal(c) no-error.
   return (x <> ?) and (error-status:error = no).
end function.

function _printRunTime returns character (input dstart as date,
                                          input tstart as integer,
                                          input dend   as date,
                                          input tend   as integer):
   define variable wrt    as character initial "" no-undo.
   define variable wstart as decimal no-undo.
   define variable wend   as decimal no-undo.

   if (dstart > dend) or (dstart = dend and tstart > tend) then return "error".
   assign wstart = decimal(dstart) + (tstart / 86400)
          wend   = decimal(dend) + (tend / 86400)
          wend   = round((wend - wstart) * 86400,0)
          wrt    = wrt + trim(string(truncate(wend / 86400,0),">>>>>>>>9")) + "d"
          wend   = wend mod 86400
          wrt    = wrt + trim(string(truncate(wend / 3600,0),">9")) + "h"
          wend   = wend mod 3600
          wrt    = wrt + trim(string(truncate(wend / 60,0),">9")) + "m"
          wend   = wend mod 60
          wrt    = wrt + trim(string(truncate(wend,0),">9")) + "s".
   return wrt.
end function.

function getFullPath returns char (input inf as character).
   define buffer bf_user for _user.
   define variable wfull as character initial "" no-undo.
   define variable wexp  as character no-undo.

   if opsys begins "WIN" then do:
      assign inf  = replace(inf,chr(47),chr(92))
             wexp = "*" + chr(92)  + "*".
      if inf matches wexp or inf matches "*printer*" then
         assign wfull = inf.
      else do:
         find bf_user where bf_user._userid = userid(ldbname(1)) no-lock no-error.
         if available bf_user then do:
            if replace(bf_user._user-misc,chr(47),chr(92)) matches "*~\*" then
               assign wfull = bf_user._user-misc + chr(92) + inf.
         end.
   
         assign wfull = _getHomeDirectory() + inf when wfull = "".
      end.
   end.
   else do:
      assign inf  = replace(inf,chr(92),chr(47))
             wexp = "*" + chr(47)  + "*".
      if inf matches wexp or inf matches "*printer*" then
         assign wfull = inf.
      else
         assign wfull = _getHomeDirectory() + inf.
   end.
   return wfull.
end function.

procedure ShellExecuteA external "shell32":U:
   define input parameter hwnd         as long.
   define input parameter lpOperation  as character.
   define input parameter lpFile       as character.
   define input parameter lpParameters as character.
   define input parameter lpDirectory  as character.
   define input parameter nShowCmd     as long.
   define return parameter hInstance as long.
end procedure.
&ENDIF
