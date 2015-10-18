on getLogFilePath()
	tell application id "com.microsoft.Excel"
		return (run VB macro "OpenSolver.NOMAD_GetLogFilePath")
	end tell
end getLogFilePath

on getNumConstraints()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_GetNumConstraints")
  end tell
end getNumConstraints

on getNumVariables()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_GetNumVariables")
  end tell
end getNumVariables

on getVariableData()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_GetVariableData")
  end tell
end getVariableData

on getOptionData()
	tell application id "com.microsoft.Excel"
		return (run VB macro "OpenSolver.NOMAD_GetOptionData")
	end tell
end getOptionData

on updateVars(newVars, bestSolution, feasibility)
	tell application id "com.microsoft.Excel"
		return (run VB macro "OpenSolver.NOMAD_UpdateVar" arg1 newVars arg2 bestSolution arg3 feasibility)
	end tell
end updateVars

on recalculateValues()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_RecalculateValues")
  end tell
end recalculateValues

on getConstraintValues()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_GetValues")
  end tell
end getConstraintValues

on loadResult(retVal)
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_LoadResult" arg1 retVal)
  end tell
end loadResult

on getConfirmedAbort()
  tell application id "com.microsoft.Excel"
    return (run VB macro "OpenSolver.NOMAD_GetConfirmedAbort")
  end tell
end getConfirmedAbort
