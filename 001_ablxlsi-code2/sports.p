define variable wxls_no     as integer initial 1 no-undo.
define variable wsheet_no   as integer initial 0 no-undo.
define variable wsheet_name as character no-undo.
define variable wtrans_ok   as logical no-undo.
define variable werr_msg    as character no-undo.
define variable wrow_no     as integer no-undo.
{xls.i}

CUSTOMER-LOOP:
for each customer
   where customer.custnum <= 100
no-lock:
   assign wsheet_no   = wsheet_no + 1
          wsheet_name = "Customer " + trim(string(customer.custnum,">>>>>9"))
          wrow_no     = 7.

   createSheetWithProperties(wxls_no,wsheet_no,wsheet_name,"LANDSCAPE=TRUE|FITTOPAGE=1").
   createCell(wxls_no,wsheet_no,1,1,"Customer Name:","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,1,2,customer.name,"TEXT=TRUE").
   createCell(wxls_no,wsheet_no,2,1,"City:","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,2,2,customer.city,"TEXT=TRUE").
   createCell(wxls_no,wsheet_no,3,1,"State:","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,3,2,customer.state,"TEXT=TRUE").
   createCell(wxls_no,wsheet_no,4,1,"Country:","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,4,2,customer.country,"TEXT=TRUE").
   createCell(wxls_no,wsheet_no,5,1,"Balance:","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,5,2,string(customer.balance,"->>>,>>>,>>>,>>9.99"),"NUMBER=TRUE|FORMAT=$#,##0.00_);($#,##0.00)").
   
   createCell(wxls_no,wsheet_no,6,1,"Order","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,2,"Ordered","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,3,"Shipped","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,4,"Promised","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,5,"Carrier","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,6,"PO","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,7,"Terms","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   createCell(wxls_no,wsheet_no,6,8,"Status","TEXT=TRUE|BOLD=TRUE|BGCOLOR=GRAY").
   
   ORDER-LOOP:
   for each order
      where order.custnum = customer.custnum
   no-lock:
      createCell(wxls_no,wsheet_no,wrow_no,1,trim(string(order.ordernum,">>>>>>>>9")),"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,2,string(order.orderdate,"99/99/9999"),"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,3,string(order.shipdate,"99/99/9999"),"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,4,string(order.promisedate,"99/99/9999"),"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,5,order.carrier,"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,6,order.po,"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,7,order.terms,"TEXT=TRUE").
      createCell(wxls_no,wsheet_no,wrow_no,8,order.orderstatus,"TEXT=TRUE").
      assign wrow_no = wrow_no + 1.
   end.  /* OF ORDER-LOOP */
end.  /* OF CUSTOMER-LOOP */

run outputXML(input "sports.xml",
              input wxls_no,
              output wtrans_ok,
              output werr_msg).

message "XML creation ran" string(wtrans_ok,"successfully/with errors") skip werr_msg
   view-as alert-box information.
