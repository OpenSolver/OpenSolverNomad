on setValue(targetRange, newValue)
	tell application id "com.microsoft.Excel" to set value of range targetRange to newValue
end setValue

tell application id "com.microsoft.Excel"
	try
		--set value of cell "B3" to (run VB macro "getrootdrivename")
		--return value of range "B3:C3"
	on error error_message number error_number
		if error_number = -10000 then
			display dialog "Excel is busy"
		else
			display dialog the error_message
		end if
	end try
end tell

on getLogFilePath()
	tell application id "com.microsoft.Excel"
		return (run VB macro "NOMAD_GetLogFilePath")
	end tell
end getLogFilePath

on getOptionData()
	tell application id "com.microsoft.Excel"
		return (run VB macro "NOMAD_GetOptionData")
	end tell
end getOptionData

on getVariableData()
	tell application id "com.microsoft.Excel"
		return (run VB macro "NOMAD_GetVariableData")
	end tell
end getVariableData

on updateVars(newVars, bestSolution, feasibility)
	tell application id "com.microsoft.Excel"
		return run VB macro "NOMAD_UpdateVar" arg1 newVars arg2 bestSolution arg3 feasibility
	end tell
end updateVars

--my setValue("B7:D7", {4, 5})
--my getLogFilePath()
my updateVars({{1}, {2}, {3}, {4}}, 2, true)