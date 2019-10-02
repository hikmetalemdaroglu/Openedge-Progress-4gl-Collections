/*******************************************************************************
 *******************************************************************************
 **                                                                           **
 **                                                                           **
 **  Copyright 2003-2006 Tom Bascom, Greenfield Technologies                  **
 **  http://www.greenfieldtech.com                                            **
 **                                                                           **
 **  chkarea.p is free software; you can redistribute it and/or modify it     **
 **  under the terms of the GNU General Public License (GPL) as published     **
 **  by the Free Software Foundation; either version 2 of the License, or     **
 **  at your option) any later version.                                       **
 **                                                                           **
 **  chkarea.p is distributed in the hope that it will be useful, but WITHOUT **
 **  ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or    **
 **  FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License     **
 **  for more details.                                                        **
 **                                                                           **
 **  See TERMS.TXT for more information regarding the Terms and Conditions    **
 **  of use and alternative licensing options for this software.              **
 **                                                                           **
 **  A copy of the GPL is in GPL.TXT which was provided with this package.    **
 **                                                                           **
 **  See http://www.fsf.org for more information about the GPL.               **
 **                                                                           **
 **                                                                           **
 *******************************************************************************
 *******************************************************************************
 *
 * chkarea.p
 *
 * With V9 areas, promon no longer shows you what space is used.
 * This is a replacement.
 *
 * 09/20/01 mikem
 * 10/23/02 tom		reformatted
 * 08/03/04 tom		added hi & lo, warning tt
 * 07/14/06 tom		added Var Size, RPB, CSZ, RMChn, MaxB and changed to %Allocated
 * 07/26/06 tom		output control
 * 10/17/06 tom		added area_data TT & moved exception report to the top
 * 10/17/06 tom		supress lo threshold warning for areas with less than 1024 blocks
 * 10/17/06 tom		changed back to reporting 1k blocks
 *
 */

define variable blksz      as integer   no-undo.
define variable bisz       as integer   no-undo.

define variable total-free as decimal   no-undo.
define variable used       as decimal   format ">>>9" no-undo.

define variable gt-free    as decimal no-undo format ">,>>>,>>>,>>>,>>9".
define variable gt-used    as decimal no-undo format ">,>>>,>>>,>>>,>>9".

define variable hi_lim     as integer no-undo format ">>9%".
define variable lo_lim     as integer no-undo format ">>9%".

define variable log_dir    as character no-undo.
define variable log_id     as character no-undo.
define variable log_name   as character no-undo.
define variable log_file   as character no-undo.

hi_lim = 70.		/* use 50 or 60 for initial planning -- 70 or 80 for production	monitoring */
lo_lim = 20.

define temp-table capacity_warn
  field areaname     as character format "x(25)"         label "Area Name"
  field percent_used as decimal   format ">>>,>>9%"      label "%Used"
  field suggest      as integer   format ">,>>>,>>>,>>9" label "Suggest"
  field projected    as decimal   format ">>>,>>9%"      label "%Projected"
  field xnote        as character format "x(40)"         label "Note"
.

define temp-table area_data
  field areanum      as integer   format ">>>"           label "Num"
  field areaname     as character format "x(25)"         label "Area Name"
  field blks-alloc   as decimal   format ">,>>>,>>>,>>9" label "Allocated"
  field vsize        as decimal   format ">,>>>,>>>,>>9" label "Var Size"
  field totblks      as decimal   format ">,>>>,>>>,>>9" label "Total"
  field blks-used    as decimal   format ">,>>>,>>>,>>9" label "KB Used"
  field pct-alloc    as decimal   format ">>>,>>9%"      label "%Allocated"
  field maxb         as decimal   format ">,>>>,>>>,>>9" label "Max Blocks (M)"
  field pct-maxb     as decimal   format ">>9.99999%"    label "%Max Blocks"
  field rpb          as integer   format      ">>>9"     label "RPB"
  field clstr-sz     as integer   format      ">>>9"     label "CSZ"
  field rmchn-sz     as integer   format ">,>>>,>>9"     label "RMChn"
.


function area-info returns logical (
  input  area-id    as integer,
  output blks_alloc as decimal,
  output blks_var   as decimal,
  output blks_free  as decimal,
  output blks_max   as decimal,
  output rpb        as integer,
  output csz        as integer,
  output rmchn      as integer
  ).

  find _AreaStatus no-lock where _AreaStatus-AreaNum = area-id.

  if (_AreaStatus-Freenum = ? ) then
    blks_free = _AreaStatus-TotBlocks - _AreaStatus-Hiwater.
   else
    blks_free = _AreaStatus-TotBlocks - _AreaStatus-HiWater + _AreaStatus-FreeNum.

  find _Area no-lock where _Area._Area-Num = _AreaStatus._AreaStatus-AreaNum.

  blks_max = exp( 2, 31 - _Area._Area-RecBits ).

  if area-id = 3 then
    blks_free = blks_free * bisz.
   else
    blks_free = blks_free * blksz.

  assign
    blks_alloc = 0
    blks_var   = 0
    rpb        = exp( 2, _Area._Area-recbits )
&IF DECIMAL(SUBSTRING(PROVERSION,1,INDEX(PROVERSION,".") + 1)) >= 10.0
&THEN
    csz        = _Area._Area-clustersize
&ELSE
    csz        = 0
&ENDIF
    rmchn      = _AreaStatus-RMNum
  .

  for each _AreaExtent no-lock where _AreaExtent._Area-Number = area-id:
    if _Extent-type >= 4 and _Extent-type <= 7 then
      blks_var = max( _Extent-size, _AreaStatus-TotBlocks - blks_alloc ).
     else
      blks_alloc = blks_alloc + integer( _Extent-Size / 1 /* ( _Area-BlockSize / 1024 ) */ ).
  end.

  return true.

end.


assign
  log_dir = "."
  log_id  = "chkarea." + ldbname( 1 )
.

if num-entries( session:parameter, "|" ) >= 1 then log_dir = entry( 1, session:parameter, "|" ).
if num-entries( session:parameter, "|" ) >= 2 then log_id  = entry( 2, session:parameter, "|" ).

if log_dir = "" then log_dir = ".".
if log_id  = "" then log_id  = "chkarea." + ldbname( 1 ).

log_name = log_id + "." + string( month( today ), "99" ) + "." + string( day( today ), "99" ).
log_file = log_dir + "/" + log_name.

output to value ( log_file ).

find first _DbStatus no-lock.
find first _AreaStatus no-lock.

put unformatted today " " string( time, "hh:mm:ss" ) " Area Status Check for DB: " _AreaStatus-Lastextent " " (( if integer( dbversion(1)) >= 10 then "OpenEdge " else "Progress " ) + dbversion(1) ) to 132 skip.
put skip(1).

blksz = ( _DbStatus._DbStatus-DBBlkSize / 1024 ).
bisz  = ( _DbStatus._DbStatus-BIBlkSize / 1024 ).

for each _AreaStatus no-lock where _AreaStatus-AreaNum >= 3 and ( not _areaStatus-Areaname matches "*After Image Area*" ):

  create area_data.
  assign
    area_data.areanum  = _AreaStatus-Areanum
    area_data.areaname = _AreaStatus-Areaname
  .

  if areanum = 3 then
    area_data.totblks  = _AreaStatus-Totblocks * bisz.
   else
    area_data.totblks  = _AreaStatus-Totblocks * blksz.

  area-info( _AreaStatus-AreaNum, blks-alloc, vsize, total-free, maxb, rpb, clstr-sz, rmchn-sz ).

  assign
    blks-used = ( totblks - total-free )
    gt-used   = gt-used + blks-used
    gt-free   = gt-free + total-free
  .

  used = ( blks-used / totblks ) * 100.

  if blks-alloc = 0 then
    pct-alloc = used.
   else
    pct-alloc = (( blks-used ) / blks-alloc ) * 100.

  pct-maxb = 100 * (( blks-used / blksz ) / maxb ).
  maxb = maxb / ( 1024 * 1024 ).

  if  ( _AreaStatus-Areanum >= 6 ) and ( pct-alloc >= hi_lim ) or
      ( _AreaStatus-Areanum  = 3 ) and ( pct-alloc >= 90 ) then
    do:
      create capacity_warn.
      assign
        capacity_warn.areaname     = _AreaStatus-AreaName
        capacity_warn.percent_used = pct-alloc
        suggest   = ((( blks-used ) * 1 ) * 2 )
        projected = (( suggest -  ( blks-used )) / suggest ) * 100
        suggest   = ( truncate( suggest / 16, 0 ) + ( if ( suggest modulo 16 = 0 ) then 0 else 1 )) * 16
        xnote = "hi"
      .
    end.
   else if ( _AreaStatus-Areanum > 6 ) and (( pct-alloc <= lo_lim ) and ( _AreaStatus-totblocks > 1024 )) then
    do:
      create capacity_warn.
      assign
        capacity_warn.areaname     = _AreaStatus-AreaName
        capacity_warn.percent_used = pct-alloc
        suggest   = ((( blks-used ) * 1 ) * 2 )
        projected = (( suggest - ( blks-used )) / suggest ) * 100
        suggest   = ( truncate( suggest / 16, 0 ) + ( if ( suggest modulo 16 = 0 ) then 0 else 1 )) * 16
        xnote = "lo"
      .
    end.

end.

find first capacity_warn no-lock no-error.
if available capacity_warn then
  put "Storage Areas violating capacity thresholds (" lo_lim "/" hi_lim "):" skip.
 else
  put "All Storage Areas are within capacity thresholds (" lo_lim "/" hi_lim ")." skip.

for each capacity_warn no-lock:
  display capacity_warn except xnote with column 10 width 132.
end.

put skip(1).
put "Suggested sizing is in 1K blocks rounded UP to a multiple of 16 * db_block_size as used in a structure file." skip.

for each area_data no-lock:
  display area_data with width 255.
end.

put skip(1).
put "   Total Blocks: " gt-free + gt-used format ">,>>>,>>>,>>>,>>>" to 43 skip.
put "    Blocks Used: " gt-used to 43 skip.
put "    Blocks Free: " gt-free to 43 skip.
put "   Percent Used: " ( gt-used / ( gt-free + gt-used )) * 100 format ">>>,>>9%" to 44 skip.
put skip(1).
put "  DB Block Size: " _DbStatus._DbStatus-DBBlkSize to 43 skip.
put "  AI Block Size: " _DbStatus._DbStatus-AIBlkSize to 43 skip.
put "  BI Block Size: " _DbStatus._DbStatus-BIBlkSize to 43 skip.
put "BI Cluster Size: " _DbStatus._DbStatus-BIClSize  to 43 skip.

output close.

if os-getenv( "MAILLOG" ) > "" then os-command value( "mailx -s " + log_name + " " + os-getenv( "MAILLOG" ) + " < " + log_file ).

return.
