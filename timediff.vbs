timediff()
Sub timediff()
clockin = inputbox("clockin","clockin")
clockout = inputbox("clockout","clockout")

If datepart("h", clockout) = datepart("h", clockin) then
	HourDif = 00
elseIF datepart("h",clockout) = 00 then
	HourDif = 24 - datepart("h",clockin)
elseif datepart("h",clockout) < datepart("h", clockin) then
	Hourdif = datepart("h", clockout) - datepart("h",clockin) - 1
else
	HourDif = datepart("h", clockout) - datepart("h", clockin)

End if

'1
If datepart("n",clockin) > datepart("n",clockout) and HourDif = 1  then
	HourDif = 0 
	MinDif = 60 - datepart("n",clockin) + datepart("n", clockout)
elseif datepart("h",clockout) > datepart("h", clockin) and datepart("n", clockout) > datepart("n",clockin) then
	Mindif = datepart("n",clockout) - datepart("n",clockin)
	Hourdif = datepart("h",clockout) - datepart("h",clockin)
'2
elseif datepart("h",clockout) > datepart("h", clockin) and datepart("n", clockin) > datepart("n",clockout) then
	MinDif = 60 - datepart("n",clockin) + datepart("n",clockout)
	HourDif = datepart("h", clockout) - datepart("h", clockin) - 1
elseif datepart("n", clockin) = datepart("n", clockout) then
	Mindif = 00
elseIf datepart("n", clockout) = 00  then
	MinDif = 60 - datepart("n",clockin)
elseIf datepart("n", clockin) > datepart("n", clockout) then
	Hourdif = HourDif - 1
	MinDif = datepart("n",clockin) - datepart("n", clockout)
Else
	MinDif = datepart("n", clockout) - datepart("n", clockin)
End if

msgbox HourDif & ":" & MinDif
End Sub
timediff()
timediff()
