  /* reading each line from editor */
  
  define var LvNotes as char view-as EDITOR SIZE 40 BY 10
  SCROLLBAR-VERTICAL.
  define var i as int.
  define frame xx
    LvNotes.
    
  do :
     LvNotes:INSERT-string("dsasdasd asdasdda asdasdasd 
  asdasd asdasd asdasd asdasd asdasd asdasd asdasd asdasd").
     assign LvNotes.
  end.  
 
  do with frame x-det :
    do  i = 1 to LvNotes:num-lines  :
       LvNotes:set-selection(LvNotes:convert-to-offset(i,1),
                             LvNotes:convert-to-offset(i + 1,1) - 1).

        message "Line" i  ":" LvNotes:SELECTION-TEXT view-as alert-box.


      
    end.
  end.  
