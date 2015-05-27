// ExcelCallbacks.h
// Functions for NOMAD library to interface with Excel
// Implemented per-platform
// All functions should be C-style to avoid mangling issues on OS X:
// http://www.drdobbs.com/cpp/problems-when-linking-objective-c-and-c/240166238

#ifndef SRC_EXCELCALLBACKS_H_
#define SRC_EXCELCALLBACKS_H_

#include <string>

namespace OPENSOLVER {

// Names of macros in Excel that we will need to call
const char GET_LOG_FILE_PATH_NAME[] =   "OpenSolver.NOMAD_GetLogFilePath";
const char GET_NUM_CONSTRAINTS_NAME[] = "OpenSolver.NOMAD_GetNumConstraints";
const char GET_NUM_VARIABLES_NAME[] =   "OpenSolver.NOMAD_GetNumVariables";
const char GET_VARIABLE_DATA_NAME[] =   "OpenSolver.NOMAD_GetVariableData";
const char GET_OPTION_DATA_NAME[] =     "OpenSolver.NOMAD_GetOptionData";
const char SHOW_CANCEL_DIALOG_NAME[] =  "OpenSolver.NOMAD_ShowCancelDialog";
const char UPDATE_VAR_NAME[] =          "OpenSolver.NOMAD_UpdateVar";
const char RECALCULATE_VALUES_NAME[] =  "OpenSolver.NOMAD_RecalculateValues";
const char GET_VALUES_NAME[] =          "OpenSolver.NOMAD_GetValues";

// Exception indicating a call into Excel failed
struct FailedCallException : std::exception {
  std::string s;
  explicit FailedCallException(std::string failedFunctionName) {
    s = failedFunctionName + " failed";
  }
  ~FailedCallException() throw() {}
  const char* what() const throw() { return s.c_str(); }
};

// We need C-style names on OS X since we are implementing in Obj-C++
#ifdef __APPLE__
extern "C" {
#endif

/**
 * Gets the path to the log file from Excel
 *
 * @param logPath String to store the full path to the log file
 */
void GetLogFilePath(std::string* logPath);

/**
 * Gets number of constraints and objectives from Excel.
 *
 * @param numCons Set to the number of constraints (inc. objectives)
 * @param numObjs Set to the number of objectives
 */
void GetNumConstraints(int* numCons, int* numObjs);

/**
 * Gets number of variables from Excel.
 *
 * @param numVars Set to the number of variables in the model
 */
void GetNumVariables(int* numVars);

/**
 * Gets information about the variables from Excel
 *
 * @param numVars The number of variables in the model
 * @param lowerBounds Array to store the lower bound for each variable
 * @param upperBounds Array to store the upper bound for each variable
 * @param startingX Array to store the starting value of each variable
 * @param varTypes Array to store the type of each variable (see VarType)
 */
void GetVariableData(int numVars, double* lowerBounds, double* upperBounds,
                     double* startingX, int* varTypes);

/**
 * Gets solver parameters for NOMAD from Excel
 *
 * @param paramStrings Pointer to array to store parameter strings from Excel
 * @param numOptions Set to the number of parameter strings
 */
void GetOptionData(std::string **paramStrings, int* numOptions);

/**
 * Sets new values of variables in Excel
 *
 * @param newValues Array of new variable values to set
 * @param numVars The number of variables in the model
 * @param bestSolution Pointer to current best solution (NULL if no solution)
 * @param feasibility True if current best solution is feasible
 */
void UpdateVars(double* newVars, int numVars, const double* bestSolution,
                bool feasibility);

/**
 * Forces a recalculate in Excel
 */
void RecalculateValues();

/**
 * Gets the new values for each constraint/objective from Excel
 *
 * @param numCons The number of constraints in the model
 * @param newCons Array to store the new values of each constraint cell
 */
void GetConstraintValues(int numCons, double* newCons);

/**
 * Conduct an evaluation iteration in Excel
 *
 * Sets the values of variables in Excel, recalculates, and then reads out the
 * updated values of the constraint cells.
 * @param newValues Array of new variable values to set
 * @param numVars The number of variables in the model
 * @param numCons The number of constraints in the model
 * @param bestSolution Pointer to current best solution (NULL if no solution)
 * @param feasibility True if current best solution is feasible
 * @param newCons Array to store the new values of each constraint cell
 */
void EvaluateX(double* newVars, int numVars, int numCons,
               const double* bestSolution, bool feasibility, double* newCons);

#ifdef __APPLE__
/**
 * Load the NOMAD result into Excel (OS X-only)
 *
 * @param retVal The NOMAD return code
 */
void LoadResult(int retVal);
#endif
  
#ifdef __APPLE__
}  // extern "C"
#endif

}  // namespace OPENSOLVER

#endif  // SRC_EXCELCALLBACKS_H_
