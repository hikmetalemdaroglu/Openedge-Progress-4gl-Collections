&if defined(baseIncluded) = 0 &then 

  function getObjectHandle returns handle
    (input pcFileName as character):
    define variable hObject as handle no-undo.
    hObject = session:first-procedure.
    do while valid-handle(hObject) and hObject:filename ne pcFileName:
      hObject = hObject:next-sibling.
    end.
    return hObject.
  end function.
  
  function launch returns logical
    (input pcFileName as character,
     input phStack    as handle):
    define variable hObject as handle no-undo.
    run value(pcFileName) persistent set hObject.
    if valid-handle(phStack) then phStack:add-super-procedure(hObject).
  end function.

  &global-define baseIncluded

&endif

