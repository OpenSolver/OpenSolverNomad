// ExcelCallbacks.h
// Functions for NOMAD library to interface with Excel
// Implemented per-platform

#ifndef EXCELCALLBACKS_H_
#define EXCELCALLBACKS_H_

#include <string>

namespace OPENSOLVER {

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
 * @return The number of variables in the model
 */
int GetNumVariables(void);

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
 * @return The number of parameter strings
 */
int GetOptionData(string **paramStrings);

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

}  // namespace OPENSOLVER

#endif  // EXCELCALLBACKS_H_
