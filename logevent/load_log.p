/* load_log.p
 *
 * You need a database table named log_event with these fields:
 *
 *   lge_date      as date "99/99/99"
 *   lge_time      as integer ">>,>>9"
 *   lge_type      as character "x(8)"
 *   lge_id        as integer ">>>,>>9"
 *   lge_user_name as character "x(8)"
 *   lge_msgnum    as integer ">>>9"
 *   lge_msgtxt    as character "x(60)"
 *
 * and indexes:
 *
 *   event_time, primary on lge_date, lge_time
 *   event_msg on lge_msgnum
 *
 */

define buffer lge for log_event.

define variable i         as integer   no-undo.
define variable lg        as character no-undo format "x(30)".
define variable ln        as character no-undo.
define variable curr_date as date      no-undo.
define variable curr_time as integer   no-undo.

update lg label "Log File" with side-labels.

input from value( lg ).

pause 0 before-hide.

form curr_date i with frame a.

repeat:

  ln = "".
  import unformatted ln.
  if ln = "" then next.

  /* lines that start with "    " are date changes...
   *
   */

  if substring( ln, 1, 4 ) = "    " then
    do:
      curr_date =
        date(
          lookup(
            substring( ln, 21, 3 ),
            "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec"
          ),
          integer( substring( ln, 25, 2 )),
          integer( substring( ln, 37, 4 ))
        ).
      display curr_date with frame a.
      next.
    end.

  /* message lines have a ":" in the 19th position
   *
   */

  if ( curr_date = ? ) then
    next.
   else
    do transaction:

      create log_event.

      if ( substring( ln, 19, 1 ) = ":" ) then
        assign
          log_event.lge_date = curr_date
          log_event.lge_time =
            integer( substring( ln, 1, 2 )) * 3600 +
            integer( substring( ln, 4, 2 )) * 60 +
            integer( substring( ln, 7, 2 ))
          log_event.lge_type =
            trim( substring( ln, 10, 6 ))
          log_event.lge_id   = 
            integer( substring( ln, 16, 3 ))
          log_event.lge_msgnum =
            integer(
              substring(
                ln,
                r-index( ln, "(" ) + 1,
                length( ln ) - r-index( ln, "(" ) - 1 ))
          log_event.lge_msgtxt =
            substring( ln, 21 )
          no-error.

      /* if an error occurs then store the whole line as the message
       * text and mark the message as wierd
       */

      if (( error-status:error = true ) or ( substring( ln, 19, 1 ) <> ":" )) then
        assign
          log_event.lge_date   = curr_date
          log_event.lge_time   = curr_time
          log_event.lge_type   = "Unknown"
          log_event.lge_id     = ?
          /**
          log_event.lge_msgnum = -1
           **/
          log_event.lge_msgtxt = ln.
       else
        curr_time = log_event.lge_time.

      /* attempt to fix up certain "unknown" messages
       *
       */

      if log_event.lge_msgnum = -1 then
        do:

          if lookup( substring( ln, 10, 9 ), "marked af,rfutil -c,bi file t,switched ,this is a,can't swi,backup ai,database " ) > 0 then
            assign
              log_event.lge_type = "RFUTIL"
              log_event.lge_time =
                integer( substring( ln, 1, 2 )) * 3600 +
                integer( substring( ln, 4, 2 )) * 60 +
                integer( substring( ln, 7, 2 ))
              log_event.lge_msgnum =
                integer(
                  substring(
                    ln,
                    r-index( ln, "(" ) + 1,
                    length( ln ) - r-index( ln, "(" ) - 1 ))
              log_event.lge_msgtxt =
                substring( ln, 10 ).

        end.
  
      /* I want the UNIX id not the Progress id thus
       * the use of msg 452 rather than 708
       */

      if log_event.lge_msgnum = 452 then
        do:
          log_event.lge_user_name =
            substring(
              ln,
              index( ln, "Login by " ) + 9,
              index( ln, " on " ) - index( ln, "Login by " ) - 9
            ).
        end.
       else if log_event.lge_type = "Usr" then
        do:
          find last lge where lge.lge_id = log_event.lge_id and lge.lge_msgnum = 452 no-lock no-error.
          if available lge then log_event.lge_user_name = lge.lge_user_name.
        end.

    end.

  i = i + 1.
  if i modulo 100 = 0 then display i with frame a.

end.

return.
