DEF INPUT PARAM PR-FILENAME AS CHAR NO-UNDO.
DEF VAR PR-RUN AS CHAR NO-UNDO.

IF PR-FILENAME <> "" THEN ASSIGN PR-FILENAME = "C:\TEMP\" + PR-FILENAME.
IF SEARCH (PR-FILENAME) = ? THEN ASSIGN PR-FILENAME = "".

IF SEARCH("c:\prg\notepad++\NOTEPAD++.EXE") = ? THEN PR-RUN = "NOTEPAD.EXE".
ELSE PR-RUN = "c:\prg\notepad++\NOTEPAD++.EXE" + " " +
              "-nosession" + " " +
              "-noPlugins" + " " +
              "-notabbar".
IF PR-FILENAME <> "" THEN ASSIGN PR-RUN = PR-RUN + " " + PR-FILENAME.
OS-COMMAND NO-WAIT VALUE(PR-RUN).
