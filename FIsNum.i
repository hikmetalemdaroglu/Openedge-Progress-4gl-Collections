function fisNum returns logical ( input p-str as character ):
  define variable n as decimal no-undo.
  assign n = decimal( p-str ) no-error.
  return (error-status:num-messages = 0).
end.

MESSAGE 
  fisNum( "123" ) SKIP
  fisNum( "xyz" )
VIEW-AS ALERT-BOX INFO BUTTONS OK.

