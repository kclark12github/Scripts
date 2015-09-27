'TimeStamp.vbs
'	Visual Basic Script Used to Format Date and Time Suitable for use in File Names...
'   Copyright © 2015, Ken Clark
'*********************************************************************************************************************************
'
'   Modification History:
'   Date:       Developer:		Description:
'   02/15/10	Ken Clark		Created;
'=================================================================================================================================
Private Function TimeStamp()
    myDate = Now()
    yyyy = Year(myDate)
    MM = Fill(Month(myDate))    
    dd = Fill(Day(myDate))
    hh = Fill(Hour(myDate))
    m = Fill(Minute(myDate))
    ss = Fill(Second(myDate))
    myDateFormat= yyyy & MM & dd & "." & hh & m & ss
End Function
Private Function Fill(num)
    If(Len(num)=1) Then Fill = "0" & num Else Fill = num
End Function
If Not IsNull(WScript.StdOut) Then objStdOut.WriteLine "TimeStamp: " & TimeStamp()
If Not IsNull(WScript.StdOut) Then WScript.StdOut.Close
WScript.Quit