Attribute VB_Name = "mMain"
' Tok komunikacije - opis komandi:

' ---> USER*
' <--- *
' ---> *
' <--- User(varijabla)

' ---> EXEC*
' <--- *

' ---> DELE filename.ext*
' <--- *

' ---> WIPE*
' <--- *

' ---> QUIT*
' <--- *

' ---> GIVE filename.ext*
' <--- *
' ---> *
' <--- FileSize(varijabla)OK*
' ---> *
' <--- blok1(512 bajta)
' ---> *
' <--- blok2(512 bajta)
' ---> *
' ...
' <--- blokN(<=512 bajta)

' ---> TAKE filename.ext:size*
' <--- *
' ---> *
' <--- *
' ---> blok1(512 bajta)
' <--- *
' ---> blok2(512 bajta)
' <--- *
' ...
' ---> blokN(<=512 bajta)
' <--- *


