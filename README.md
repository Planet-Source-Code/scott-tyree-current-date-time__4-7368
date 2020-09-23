<div align="center">

## Current Date & Time


</div>

### Description

This script will get the system format date/time and display the results in an easy to read format. All elements of the date/time have been broken out into individual customizable pieces so it's very easy to adapt to your needs. Unlike my usual code, this has been highly commented!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott Tyree](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-tyree.md)
**Level**          |Beginner
**User Rating**    |2.7 (19 globes from 7 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-tyree-current-date-time__4-7368/archive/master.zip)

### API Declarations

Please feel free to change the code however you see fit.


### Source Code

```
<%
' Define variable names
Dim vMonth
Dim vDay
Dim vYear
Dim vTime
Dim vHour
Dim vMinute
Dim v0Min
Dim vDate
Dim vWeekday
Dim vFinal
Dim vAMPM
Dim vExt
Dim vSec
' Changes the numeric month to text
If Month(Now())=1 Then
vMonth="Jan"
ElseIf Month(Now())=2 Then
vMonth="Feb"
ElseIf Month(Now())=3 Then
vMonth="Mar"
ElseIf Month(Now())=4 Then
vMonth="Apr"
ElseIf Month(Now())=5 Then
vMonth="May"
ElseIf Month(Now())=6 Then
vMonth="Jun"
ElseIf Month(Now())=7 Then
vMonth="Jul"
ElseIf Month(Now())=8 Then
vMonth="Aug"
ElseIf Month(Now())=9 Then
vMonth="Sep"
ElseIf Month(Now())=10 Then
vMonth="Oct"
ElseIf Month(Now())=11 Then
vMonth="Nov"
ElseIf Month(Now())=12 Then
vMonth="Dec"
End If
' Adds the leading 0 to the minutes if less than 10
If Minute(Now())<=09 Then
vMin="0" & Minute(Now())
ElseIf Minute(Now())>09 Then
vMin= Minute(Now())
End If
' Adds the leading 0 to the seconds if less than 10
If Second(Now())<=09 Then
vSec="0" & Second(Now())
ElseIf Second(Now())>09 Then
vSec= Second(Now())
End If
' Sets AM/PM based on hour
If Hour(Now())<12 Then
vAMPM="am"
ElseIf Hour(Now())>11 Then
vAMPM="pm"
ElseIf Hour(Now())=24 Then
vAMPM="am"
End If
' Makes the 24 server clock show like a 12 hour clock
If Hour(Now())=00 Then
vHour="1"
ElseIf Hour(Now())<=12 Then
vHour=Hour(Now())
ElseIf Hour(Now())=13 Then
vHour="1"
ElseIf Hour(Now())=14 Then
vHour="2"
ElseIf Hour(Now())=15 Then
vHour="3"
ElseIf Hour(Now())=16 Then
vHour="4"
ElseIf Hour(Now())=17 Then
vHour="5"
ElseIf Hour(Now())=18 Then
vHour="6"
ElseIf Hour(Now())=19 Then
vHour="7"
ElseIf Hour(Now())=20 Then
vHour="8"
ElseIf Hour(Now())=21 Then
vHour="9"
ElseIf Hour(Now())=22 Then
vHour="10"
ElseIf Hour(Now())=23 Then
vHour="11"
ElseIf Hour(Now())=24 Then
vHour="12"
End If
' Changes the numeric weekday to text
If Weekday(Now())=1 Then
vWeekday="Sun"
ElseIf Weekday(Now())=2 Then
vWeekday="Mon"
ElseIf Weekday(Now())=3 Then
vWeekday="Tue"
ElseIf Weekday(Now())=4 Then
vWeekday="Wed"
ElseIf Weekday(Now())=5 Then
vWeekday="Thu"
ElseIf Weekday(Now())=6 Then
vWeekday="Fri"
ElseIf Weekday(Now())=7 Then
vWeekday="Sat"
End If
' Pulls the day of the week
vDay=Day(Now())
' Adds the st, nd, rd, th extention to the date
If vDay=1 Or vDay=21 Or vDay=31 Then
vExt="st"
ElseIf vDay=2 Or vDay=22 Then
vExt="nd"
ElseIf vDay=3 Or vDay=23 Then
vExt="rd"
ElseIf vDay=4 Or vDay=5 Or vDay=6 Or vDay=7 Or vDay=8 Or vDay=9 Or vDay=10 Or vDay=11 Or vDay=12 Or vDay=13 Or vDay=14 Or vDay=15 Or vDay=16 Or vDay=17 Or vDay=18 Or vDay=19 Or vDay=20 Or vDay=24 Or vDay=25 Or vDay=26 Or vDay=27 Or vDay=28 Or vDay=29 Or vDay=30 Then
vExt="th"
End If
' Pulls the seconds
vSecond = vSec
' Pulls the current year
vYear = Year(Now())
' Combines all the related elements to create the time module
vTime = vHour & ":" & vMin & ":" & vSecond
' Combines all the related elements to create the day/date module
vDate = vWeekday & ", " & vMonth & " " & vDay & "" & vExt & " " & vYear
' Brings it all together
vFinal = vTime & "" & vAMPM & " - " & vDate
%>
The current time And Date Is: <%= vFinal %>
```

