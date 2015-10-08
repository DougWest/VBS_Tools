Option Explicit

Dim iHour, apm

apm = " AM"

For iHour = 1 to 12
  wscript.echo iHour & ":00" & apm
  wscript.echo iHour & ":15" & apm
  wscript.echo iHour & ":30" & apm
  wscript.echo iHour & ":45" & apm
Next

apm = " PM"
For iHour = 1 to 12
  wscript.echo iHour & ":00" & apm
  wscript.echo iHour & ":15" & apm
  wscript.echo iHour & ":30" & apm
  wscript.echo iHour & ":45" & apm
Next

