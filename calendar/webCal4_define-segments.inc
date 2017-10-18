<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 10/05/1999

' this include file calculates the variables needed
' to create an HTML table displaying the proper
' number of time segments each day

intFactor = 60/intErval

' these define the range of time to display
' the first number is the 24-hour time of day

intRange1 = intHourStart * intFactor
intRange2 = intHourEnd * intFactor - 1

' this function takes a time and converts it to the
' proper number of table segments based on the
' specified interval

intSegments = intFactor

function segments(strTime)
	intMin = Minute(strTime)
	intAdd = intErval/2
	for z = 0 to 60/intErval - 1
		if intMin < intAdd then
			intSegments = z
			exit for
		end if
		intAdd = intAdd + intErval
	next

' events set to end at 0000 hours actually
' end at midnight, so set to the artificial
' 2400 hours
	
	if Hour(strTime) = 0 then
		intHour = 24
	else
		intHour = Hour(strTime)
	end if
	
	segments = intSegments + (intHour * intFactor)
end function

' ZERO-BASED count of time segements/day
' we actually go through 25 hours to accomodate
' events that end in the 12:00 AM hour

intTotal = (1500/intErval) - 1
intHeight = 24/intFactor - 1

%>