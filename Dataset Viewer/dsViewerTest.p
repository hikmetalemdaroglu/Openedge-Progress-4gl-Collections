/* singleproc.p */                               
define variable hChangeDataSet as handle no-undo.
/* Temp-table definitions */

DEFINE TEMP-TABLE ttOrder 
    FIELD OrderNum AS INTEGER FORMAT "zzzzzzzzz9"
    FIELD OrderDate AS DATE FORMAT "99/99/99"
    FIELD ShipDate AS DATE FORMAT "99/99/99" 
    FIELD PromiseDate AS DATE FORMAT "99/99/99"
    FIELD OrderTotal AS DECIMAL FORMAT "->,>>>,>>9.99"
    INDEX OrderNum IS UNIQUE PRIMARY OrderNum.

DEFINE TEMP-TABLE ttOrderLine BEFORE-TABLE ttOrderLineBefore
     FIELD OrderNum AS INTEGER FORMAT "zzzzzzzzz9"
     FIELD LineNum AS INTEGER FORMAT ">>9"
     FIELD ItemNum AS INTEGER FORMAT "zzzzzzzzz9"
     FIELD Price AS DECIMAL FORMAT "->,>>>,>>9.99"
     FIELD Qty AS INTEGER FORMAT "->>>>9"
     FIELD Discount AS INTEGER FORMAT ">>9%"
     FIELD ExtendedPrice AS DECIMAL FORMAT "->>>,>>9.99"
     INDEX OrderNum_LineNum IS UNIQUE PRIMARY OrderNum LineNum.

/* ProDataSet definition */

DEFINE DATASET dsOrderOrderLine FOR ttOrder, ttOrderLine
    DATA-RELATION drOrderOrderLine FOR ttOrder, ttOrderLine
    RELATION-FIELDS (OrderNum, OrderNum).

/* Data-Source Definitions */

DEFINE QUERY qOrder FOR Order.
DEFINE DATA-SOURCE srcOrder FOR QUERY qORder.
DEFINE DATA-SOURCE srcOrderLine FOR OrderLine.

/* Attach Data Sources */

BUFFER ttOrder:ATTACH-DATA-SOURCE(DATA-SOURCE srcOrder:HANDLE,?,?).
BUFFER ttOrderLine:ATTACH-DATA-SOURCE(DATA-SOURCE srcOrderLine:HANDLE,?,?).

/* Prepare Query */

QUERY qOrder:QUERY-PREPARE("FOR EACH Order WHERE shipdate <> ?").

/* Populate the ProDataSet */

DATASET dsOrderOrderLine:FILL().
TEMP-TABLE ttOrderLine:TRACKING-CHANGES = TRUE.
FOR EACH ttOrderLine where ttOrderLine.OrderNum = 1:
  assign ttOrderLine.Qty = 20.
END.

FOR EACH ttOrderLine where ttOrderLine.OrderNum = 2:
  delete ttOrderLine.
END.

create ttOrderLine.
ttOrderLine.OrderNUm = 20.
ttOrderLine.LineNUm = 9.
create ttOrderLine.
ttOrderLine.OrderNUm = 20.
ttOrderLine.LineNUm = 10.
display "OK".
TEMP-TABLE ttOrderLine:TRACKING-CHANGES = FALSE.

CREATE DATASET hChangeDataSet.
hChangeDataSet:CREATE-LIKE(DATASET dsOrderOrderLine:HANDLE,"cds").
hChangeDataSet:GET-CHANGES(DATASET dsOrderOrderLine:HANDLE).

{dsViewer.i}
run displayDataset(input DATASET dsOrderOrderLine).

wait-for "esc" of this-procedure.
