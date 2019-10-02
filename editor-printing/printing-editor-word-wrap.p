define variable editor-widget as widget-handle no-undo.
define variable line-length as integer initial 40 no-undo.
define variable line-quantity as integer initial 6 no-undo.
define variable iPosNow as integer no-undo.
define variable iPosPrev as integer initial 1 no-undo.
define variable line-number as integer no-undo.

define frame f-editor-widget
with 1 down no-box overlay
no-labels
at col 1 row 1
scrollable size 320 by 18.

create editor editor-widget assign
frame = frame f-editor-widget:handle row=1 column=1
inner-chars = line-length
inner-lines = line-quantity + 1
max-chars = (line-length * 18).

assign editor-widget:screen-value =
/* The following long line may be wrapped by the email system. Everything
between the double quotes is one line. */
"I have a editor-box field with 6 lines. The option 'word-wrap' is on. But
now I want to print this editor field line by line. The problem now is that
I can not detect where in the editor-box a skip is made because of the
word-wrap option.".

lineLoop:
do line-number = 2 to line-quantity + 1:
assign iPosNow = editor-widget:convert-to-offset(line-number,1).
if iPosNow eq 0 then assign iPosNow = line-quantity * line-length.
editor-widget:set-selection(iPosPrev,iPosNow).
if length(editor-widget:selection-text) eq 0 then leave lineLoop.
display editor-widget:selection-text format "x(40)"
with down frame f-display.
down with frame f-display.
assign iPosPrev = iPosNow.
end. /* line-number = 1 to inner-lines */
