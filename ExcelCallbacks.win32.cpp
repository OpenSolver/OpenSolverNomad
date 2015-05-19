// ExcelCallbacks.win32.cpp
// Implementation of ExcelCallbacks.h for Windows

#include <atlbase.h>
#include <stdio.h>

#include <xlcall.h>
#include <framewrk.h>

#include <limits>
#include <string>

namespace OPENSOLVER {

void GetNumConstraints(int* numCons, int* numObjs) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetNumConstraints"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != 2) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_GetNumConstraints failed";
  }

  *numCons = static_cast<int>(xResult.val.array.lparray[0].val.num);
  *numObjs = static_cast<int>(xResult.val.array.lparray[1].val.num);

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);
  return;
}

int GetNumVariables(void) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetNumVariables"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeNum) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_GetNumVariables failed";
  }

  int result = static_cast<int>(xResult.val.num);

  Excel12f(xlFree, nullptr, 1, &xResult);
  return result;
}

void GetVariableData(int numVars, double* lowerBounds, double* upperBounds,
                     double* startingX, int* varTypes) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetVariableData"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != 4 * numVars) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_GetVariableData failed";
  }

  // Get the lower and upper bounds for each of the variables
  for (int i = 0; i < numVars; i++) {
    lowerBounds[i] = xResult.val.array.lparray[2 * i].val.num;
    upperBounds[i] = xResult.val.array.lparray[2 * i + 1].val.num;
  }

  // Get start point
  for (int i = 0; i < numVars; i++) {
    startingX[i] = xResult.val.array.lparray[2 * numVars + i].val.num;
  }

  // Get the variable types (real, integer, or binary)
  for (int i = 0; i < numVars; i++) {
    double rawType = xResult.val.array.lparray[3 * numVars + i].val.num;
    varTypes[i] = static_cast<int>(rawType);
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);
  return;
}

int GetOptionData(std::string **paramStrings) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetOptionData"));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_GetOptionData failed";
  }

  std::wstring ws;
  int n;
  int m = xResult.val.array.rows;
  *paramStrings = new std::string[m];

  for (int i = 0; i < m; ++i) {
    // Get the string value out of the result
    n = static_cast<int>(xResult.val.array.lparray[2 * i + 1].val.num);
    ws = std::wstring(xResult.val.array.lparray[2 * i].val.str);
    (*paramStrings)[i] = std::string(ws.begin(), ws.end()).substr(1, n);
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);
  return m;
}

void EvaluateX(double* newVars, int numVars, int numCons,
               const double* bestSolution, bool feasibility, double *newCons) {
  XLOPER12 xOpAbort, xOpConfirm;

  // Check for escape key press
  // http://msdn.microsoft.com/en-us/library/office/bb687825%28v=office.15%29.aspx
  Excel12f(xlAbort, &xOpAbort, 0);
  if (xOpAbort.val.xbool) {
    Excel12f(xlFree, nullptr, 2, &xOpAbort);

    int ret = Excel12f(xlUDF, &xOpConfirm, 1,
                       TempStr12(L"OpenSolver.NOMAD_ShowCancelDialog"));
    if (ret == xlretAbort || ret == xlretUncalced ||
        xOpConfirm.xltype != xltypeBool) {
      Excel12f(xlFree, nullptr, 1, &xOpConfirm);
      throw "NOMAD_ShowCancelDialog failed";
    }

    if (xOpConfirm.val.xbool) {
      Excel12f(xlFree, nullptr, 1, &xOpConfirm);
      throw "Abort";
    } else {
        // Clear the escape key press so we can resume
      Excel12f(xlAbort, nullptr, 1, TempBool12(false));
    }

    Excel12f(xlFree, nullptr, 1, &xOpConfirm);
  }

  static XLOPER12 xResult;

  // In this implementation, the upper limit is the largest
  // single column array (equals 2^20, or 1048576, rows in Excel 2007).
  if (numVars < 1 || numVars > 1048576) {
    return;
  }

  // Create an array of XLOPER12 values.
  XLOPER12 *xOpArray = static_cast<XLOPER12*>(
      malloc(static_cast<size_t>(numVars) * sizeof(XLOPER12)));

  // Create and initialize an xltypeMulti array
  // that represents a one-column array.
  XLOPER12 xOpMulti;
  xOpMulti.xltype = xltypeMulti|xlbitDLLFree;
  xOpMulti.val.array.lparray = xOpArray;
  xOpMulti.val.array.columns = 1;
  xOpMulti.val.array.rows = static_cast<RW>(numVars);

  // Initialize and populate the array of XLOPER12 values.
  for (int i = 0; i < numVars; i++) {
    xOpArray[i].xltype = xltypeNum;
    xOpArray[i].val.num = *(newVars+i);
  }

  // Create XLOPER12 objects for passing in solution and feasibility

  // Pass solution in as Double, or vbNothing if no solution
  XLOPER12 xOpSol;
  if (bestSolution == nullptr) {
    xOpSol.xltype = xltypeMissing|xlbitXLFree;
  } else {
    xOpSol.xltype = xltypeNum|xlbitXLFree;
    xOpSol.val.num = *bestSolution;
  }

  int ret;

  // Update variables
  ret = Excel12f(xlUDF, &xResult, 4, TempStr12(L"OpenSolver.NOMAD_UpdateVar"),
                 &xOpMulti, &xOpSol, TempBool12(!feasibility));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_UpdateVar failed";
  }

  // Recalculate values
  ret = Excel12f(xlUDF, nullptr, 1,
                 TempStr12(L"OpenSolver.NOMAD_RecalculateValues"));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_RecalculateValues failed";
  }

  // Get constraint values
  ret = Excel12f(xlUDF, &xResult, 1, TempStr12(L"OpenSolver.NOMAD_GetValues"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != numCons) {
    Excel12f(xlFree, nullptr, 1, &xResult);
    throw "NOMAD_GetValues failed";
  }

  for (int i = 0; i < numCons; i++) {
    // Check for error passed back from VBA and set to C++ NaN.
    // We need to catch errors separately as they are otherwise interpreted
    // as having value zero.
    if (xResult.val.array.lparray[i].xltype != xltypeNum) {
      newCons[i] = std::numeric_limits<double>::quiet_NaN();
    } else {
      newCons[i] = xResult.val.array.lparray[i].val.num;
    }
  }

  // Free memory allocated by Excel
  Excel12f(xlFree, nullptr, 1, &xResult);
  return;
}

}  // namespace OPENSOLVER
