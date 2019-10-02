/*******************************************************************************
 *******************************************************************************
 **                                                                           **
 **                                                                           **
 **  Copyright 2003-2006 Tom Bascom, Greenfield Technologies                  **
 **  http://www.greenfieldtech.com                                            **
 **                                                                           **
 **  dbStat.p is free software; you can redistribute it and/or modify it      **
 **  under the terms of the GNU General Public License (GPL) as published     **
 **  by the Free Software Foundation; either version 2 of the License, or     **
 **  at your option) any later version.                                       **
 **                                                                           **
 **  dbStat.p is distributed in the hope that it will be useful, but WITHOUT  **
 **  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or    **
 **  FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License     **
 **  for more details.                                                        **
 **                                                                           **
 **                                                                           **
 **  See http://www.fsf.org for more information about the GPL.               **
 **                                                                           **
 **                                                                           **
 *******************************************************************************
 *******************************************************************************
 *
 * dbstat.p
 *
 * sample and gather statistics about record count, size and standard deviation
 *
 * to run:

	$ export DB=sports2000			# the database being analyzed
	$ export LOGDIR=/tmp			# where to write the log file
	$ export SHRTNM=s2k			# a short name for the logs
	$ export DBANAL=s2k.dbanalys.out	# optional previously obtained db analysis output
	$ export TBLIST=customer		# optional comma delimited list of tables to report on

	$ $DLC/bin/_progres -b $DB -p dbstat.p -param "$LOGDIR|$SHRTNM|$DBANAL|$TBLIST" $* >> $LOGDIR/$SHRTNM.err 2>&1 &

 *
 */

&global-define SAMPLE    500								/* an integer sample size		*/
&global-define MAXLOOK 25000								/* the maximum number of attempts	*/

&global-define GWIDTH     40								/* ASCII-art graph width		*/
&global-define GHEIGHT    20								/* ASCII-art graph height		*/
 
define variable i       as integer no-undo.						/* number of successful samples		*/
define variable j       as integer no-undo.						/* number of sampling attempts		*/
define variable k       as integer no-undo.						/* sum of sampled values		*/
define variable v       as integer no-undo extent {&SAMPLE}.				/* sampled values used to calc std dev	*/
define variable s       as integer no-undo.						/* record length			*/
define variable x       as integer no-undo.						/* intermediate value used in calc	*/
define variable y       as decimal no-undo.						/* intermediate value used in calc	*/
define variable z       as decimal no-undo format ">>,>>>,>>>,>>9".			/* total number of sample attempts	*/
define variable stime   as integer no-undo.						/* start time				*/
define variable max-rid as integer no-undo.						/* maximum recid for this storage area	*/
define variable min-rec as integer no-undo label "Min"     format ">>,>>9".		/* minimum sampled record size		*/
define variable max-rec as integer no-undo label "Max"     format ">>,>>9".		/* maximum sampled record size		*/
define variable avg-rec as integer no-undo label "Mean"    format ">>,>>9".		/* average sampled record size		*/
define variable std-dev as integer no-undo label "Std Dev" format ">>,>>9".		/* standard deviation of samples	*/
define variable rec-cnt as integer no-undo label "Records" format ">>>>>>>>>9".		/* estimated number of records in table	*/
define variable xrec    as integer no-undo label "xRecs".				/* expected number of records		*/
define variable xavg    as integer no-undo label "xAvg".				/* expected average record size		*/

define variable g       as integer   no-undo extent {&GWIDTH} format ">>9".
define variable graph   as character no-undo label "Graph" format "x({&GWIDTH})".

define variable bh as handle no-undo.							/* a handle for a dynamic buffer	*/
define variable qh as handle no-undo.							/* a handle for a dynamic query		*/

define variable db_analys  as character no-undo.					/* name of dbanalys file		*/

define variable log_dir    as character no-undo.
define variable log_id     as character no-undo.
define variable log_name   as character no-undo.
define variable log_file   as character no-undo.
define variable tbl_list   as character no-undo.

assign
  log_dir = "."
  log_id  = "dbstat." + ldbname( 1 )
.

if num-entries( session:parameter, "|" ) >= 1 then log_dir   = entry( 1, session:parameter, "|" ).
if num-entries( session:parameter, "|" ) >= 2 then log_id    = entry( 2, session:parameter, "|" ).
if num-entries( session:parameter, "|" ) >= 3 then db_analys = entry( 3, session:parameter, "|" ).
if num-entries( session:parameter, "|" ) >= 4 then tbl_list  = entry( 4, session:parameter, "|" ).

if log_dir = "" then log_dir = ".".
if log_id  = "" then log_id  = "dbstat." + ldbname( 1 ).

log_name = log_id + "." + string( month( today ), "99" ) + "." + string( day( today ), "99" ).
log_file = log_dir + "/" + log_name.

stime = time.										/* when did we start running?		*/

create query qh.

output to value( log_file ) unbuffered.

for each _file no-lock where not _file._hidden:						/* loop through the tables in DICTDB	*/

  if tbl_list <> "" and lookup( _file._file-name, tbl_list ) = 0 then next.		/* that's *really* ugly!		*/

  find _storageobject no-lock								/* we need to find the *current*	*/
    where _storageobject._db-recid = _file._db-recid					/* storage area -- not the *initial*	*/
      and _storageobject._object-type = 1						/* storage area that holds the table	*/
      and _storageobject._object-number = _file._file-num.				/* ( _file._ianum = initial area)	*/

  find _area no-lock where _area._area-number = _storageobject._area-number.
  find _areastatus no-lock where _areastatus._areastatus-areanum = _storageobject._area-number.

  /* narrow the sample space down by noting that there cannot be a recid
   * past the area's hiwater mark.  it would be nice if we could establish
   * a hwm on a per table basis but, sadly, that isn't possible :-(
   */

  max-rid = ( _areastatus-hiwater * exp( 2, _area-recbits )) + exp( 2, _area-recbits ) - 1.

  create buffer bh for table _file._file-name.						/* we need a buffer for the table	*/

  qh:set-buffers( bh ).									/* set up a dynamic query...		*/
  qh:query-prepare( "for each " + _file._file-name + " no-lock" ).
  qh:query-open().

  assign										/* initialize stats and counters	*/
    i = 0
    j = 0
    k = 0
    s = 0										/* reset the successful sample counter	*/
    v = 0
    min-rec = 999999
    max-rec = 0
  .

  do while s < {&SAMPLE}:								/* make sure that there are at least	*/
											/* {&SAMPLE} records in the table	*/
    qh:get-next().
    if qh:query-off-end then
      leave.
     else
      s = s + 1.									/* count a record			*/

  end.

  if s = 0 then next.									/* skip table if there are no records	*/

  display
    _file._file-name label "Table Name" format "x(15)"
   with
    width 132
  .

  qh:query-close().									/* set up the dynamic query again...	*/
  qh:set-buffers( bh ).									/* we will use this if we're in "grab	*/
  qh:query-prepare( "for each " + _file._file-name + " no-lock" ).			/* them all" rather than "sample" mode.	*/
  qh:query-open().

  stat_loop: do while i < min( s, {&SAMPLE} ):						/* don't try to find more records than	*/
											/* exist...				*/
    assign
      j = j + 1										/* how many times have we looked?	*/
      z = z + 1										/* (in total across all tables...)	*/
    .

    /* if j > ( {&MAXLOOK} * {&SAMPLE} ) then leave stat_loop.	*/			/* don't try forever			*/
    if j > ( {&MAXLOOK} ) then leave stat_loop.						/* don't try forever			*/

    if s >= {&SAMPLE} then								/* if we expect more than {&SAMPLE}	*/
      do:										/* records then randomly probe for one	*/
											/* by recid...				*/
        qh:query-close().
        qh:set-buffers( bh ).

        qh:query-prepare(
          "for each " + _file._file-name +						/* look for records in _file-name	*/
          " no-lock where recid( " +							/* no-lock, where the RECID is a random	*/
          _file._file-name + " ) = " +							/* recid between 1 and the recid of the	*/
          string( random( 1, max-rid ))							/* area hi-water mark			*/
        ).

        qh:query-open().

      end.										/* otherwise just get them all ;-)	*/

    qh:get-next().

    if not qh:query-off-end then							/* did we find a record?		*/
      assign
        i = i + 1									/* count it!				*/
        x = bh:record-length + 2							/* dbanalysis is 2 bytes longer...	*/
        min-rec = min( x, min-rec )							/* is it the smallest so far?		*/
        max-rec = max( x, max-rec )							/* is it the biggest so far?		*/
        k = k + x									/* add it to the summation		*/
        v[i] = x									/* track the individual values		*/
      .

  end.

  qh:query-close().									/* clean up...				*/
  delete object bh.

  if min-rec = 999999 then next.							/* too sparsely populated -- give up	*/


  assign
    avg-rec = k / i									/* compute the mean record size		*/
    g = 0
    y = max( 1, (( max-rec - min-rec ) + 1 ) / ( {&GWIDTH} - 1 ))
    x = 0
  .

  do s = 1 to i:									/* calculate the sum of the squares of	*/
    x = x + exp(( avg-rec - v[s] ), 2 ).						/* the differences from the average	*/
    g[integer(( v[s] - min-rec ) / y ) + 1] = g[integer(( v[s] - min-rec ) / y ) + 1] + 1.
  end.

  std-dev = exp(( x / ( i - 1 )), 0.5 ).						/* the standard deviation		*/

  rec-cnt = integer(( if i < {&SAMPLE} then i else (( max-rid / j ) * i ))).		/* estimate the record count based on 	*/
											/* density of sample...			*/
  display
    rec-cnt
    min-rec
    max-rec
    avg-rec
    std-dev
  .

  if db_analys > "" then								/* compare to known dbanalys output	*/
    do:

      assign										/* we might not find anything (the grep	*/
        xrec = -1									/* command below won't, for instance,	*/
        xavg = -1									/* return lines that are split due to	*/
      .											/* long table names).			*/

      input through value( 'grep -i "PUB.' + _file._file-name + ' " ' + db_analys ).	/* the trailing " " after the file name	*/
      import ^ xrec ^ ^ ^ xavg no-error.						/* prevents substring problems!		*/
      input close.

    end.

  x = 0.
  do i = 1 to {&GWIDTH}:								/* find the max value for the y-axis	*/
    x = max( x, g[i] ).									/* perhaps the variable ought to be "y"	*/
  end.

  do j = {&GHEIGHT} to 1 by -1:								/* create the chart line by line...	*/

    if rec-cnt >= {&SAMPLE} then							/* if there is anything to chart ;-)	*/
      do:
        display (( j ) * ( x / {&GHEIGHT} )) @ g[1] format ">,>>>,>>9" with no-label.	/* label the y-axis			*/
        graph = "".

        do i = 1 to min( {&GWIDTH}, (( max-rec - min-rec ) + 1 )):			/* if columns are less than 1 unit	*/
											/* apart then the width is max - min	*/
          graph =
           graph +
           ( if ( g[i] > (( j - 1 ) * ( x / {&GHEIGHT} ))) then				/* if we found samples in this range	*/
               "*"									/* mark the spot!			*/
              else ( if ( g[i] = 0 ) and ( j = 1 ) then					/* if it is the baseline we might want	*/
               " "									/* to output "_"			*/
              else									/* the bar isn't this high but we need	*/
               " " )									/* to align the next column		*/
           ).

        end.
        display graph.									/* spit out the line			*/

      end.

    if j = ( {&GHEIGHT} - 1 ) then							/* display expected values obtained	*/
      display										/* from dbanalys			*/
        xrec @ rec-cnt
        xavg @ avg-rec
      .
    if j = ( {&GHEIGHT} - 2 ) then							/* display variance			*/
      display
        exp( exp(( xrec - rec-cnt ) / xrec, 2 ), 0.5 ) * 100 format ">,>>9.99%" @ rec-cnt
        exp( exp(( xavg - avg-rec ) / xavg, 2 ), 0.5 ) * 100 format ">,>>9.99%" @ avg-rec
      .

    if rec-cnt >= {&SAMPLE} or ( j >= ( {&GHEIGHT} - 2 )) then down 1.

  end.

  if rec-cnt >= {&SAMPLE} then
    do:

      /* show how wide each column is
       */

      graph = fill( "-", min( {&GWIDTH}, (( max-rec - min-rec ) + 1 ))).
      display graph y.
      down 1.

      /* show min & max record sizes
       */

      graph = string( min-rec ).
      graph = graph + fill( " ", min( {&GWIDTH}, (( max-rec - min-rec ) + 1 )) - length( graph ) - 4 ) + string( max-rec, ">>>9" ).
      display graph.
      down 1.

      /* show standard deviations & mean
       */

      graph = fill( "_", min( {&GWIDTH}, (( max-rec - min-rec ) + 1 ))).

      do k = 1 to 9:
        x = avg-rec + ( std-dev * k ).
        if x <= max-rec then substr( graph, integer(( x - min-rec ) / y ) + 1, 1 ) = string( k ).
        x = avg-rec - ( std-dev * k ).
        if x >= min-rec then substr( graph, integer(( x - min-rec ) / y ) + 1, 1 ) = string( k ).
      end.

      /* insert the mean last so that it overwrites std-dev if values are closely packed
       */

      substr( graph, integer(( avg-rec - min-rec ) / y ) + 1, 1 ) = "^".

      display graph.
      down 1.

      /* show suggested RPB break points
       */

      graph = fill( "-", min( {&GWIDTH}, (( max-rec - min-rec ) + 1 ))).

      /* standard v9 RPB break points -- for OE10 this should look at the area create & toss limits
       *
       */

      if       min-rec <=   27 and max-rec >=   27 then substr( graph, integer((   27 - min-rec ) / y ) + 1, 1 ) = string( "8" ).
       else if min-rec <=   57 and max-rec >=   57 then substr( graph, integer((   57 - min-rec ) / y ) + 1, 1 ) = string( "7" ).
       else if min-rec <=  116 and max-rec >=  116 then substr( graph, integer((  116 - min-rec ) / y ) + 1, 1 ) = string( "6" ).
       else if min-rec <=  234 and max-rec >=  234 then substr( graph, integer((  234 - min-rec ) / y ) + 1, 1 ) = string( "5" ).
       else if min-rec <=  473 and max-rec >=  473 then substr( graph, integer((  473 - min-rec ) / y ) + 1, 1 ) = string( "4" ).
       else if min-rec <=  946 and max-rec >=  946 then substr( graph, integer((  946 - min-rec ) / y ) + 1, 1 ) = string( "3" ).
       else if min-rec <= 1893 and max-rec >= 1893 then substr( graph, integer(( 1893 - min-rec ) / y ) + 1, 1 ) = string( "2" ).
       else if min-rec <= 3786 and max-rec >= 3786 then substr( graph, integer(( 3786 - min-rec ) / y ) + 1, 1 ) = string( "1" ).
       else if min-rec <= 7572 and max-rec >= 7572 then substr( graph, integer(( 7572 - min-rec ) / y ) + 1, 1 ) = string( "0" ).

      display graph.
      down 1.

      down 1.										/* skip an extra  line			*/

    end.

end.

delete object qh.									/* finish cleaning up			*/

display {&SAMPLE} z string( time - stime, "hh:mm:ss" ).					/* how long did it take?		*/

output close.

if os-getenv( "MAILLOG" ) > "" then os-command value( "mailx -s " + log_name + " " + os-getenv( "MAILLOG" ) + " < " + log_file ).

return.
