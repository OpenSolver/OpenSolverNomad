// ExcelCallbacks.win32.cpp
// Implementation of ExcelCallbacks.h for Windows

#include "ExcelCallbacks.h"

#include <atlbase.h>
#include <stdio.h>

#include <xlcall.h>
#include <framewrk.h>

#include <cstdlib>
#include <limits>
#include <string>

namespace OPENSOLVER {

const size_t WCHARBUF = 100;

// Fills an XCHAR from char[] if the XCHAR is empty
void ConvertToXcharIfNeeded(XCHAR* destination, const char* source) {
  if (wcslen(destination) == 0) {
    auto destLength = mbstowcs(destination, source, WCHARBUF);
    if (destLength == -1) {
      std::string error = std::string("Error while converting") + source;
      throw std::exception(error.c_str());
    }
  }
}

// Check the return code from an Excel12 call for errors
bool CheckReturnCodeOkay(int ret) {
  return ret != xlretAbort && ret != xlretUncalced;
}

// Checks that the variant returned from Excel is an array of the given size
bool CheckIsArray(const XLOPER12& xResult, int expectedSize) {
  return xResult.xltype == xltypeMulti &&
         xResult.val.array.rows * xResult.val.array.columns == expectedSize;
}

// Converts XCHAR to std::string
std::string XcharToString(const XCHAR* s) {
  std::wstring ws(s);
  return std::string(ws.begin(), ws.end());
}

// Converts a string returned from Excel to std::string
// All strings need to be passed as a variant array with string then its length
std::string GetStringFromExcel(const XLOPER12* stringData) {
  // Get the length of the string
  if (stringData[1].xltype != xltypeNum) {
    throw std::exception("String length supplied by Excel not integer");
  }
  int n = static_cast<int>(stringData[1].val.num);

  if (stringData[0].xltype != xltypeStr) {
    throw std::exception("String data supplied was not a string type");
  }
  // Excel puts garbage in the first char
  return XcharToString(stringData[0].val.str).substr(1, n);
}

// Creates a FailedCallException using the provided XCHAR message
FailedCallException MakeFailedCallException(const XCHAR* message) {
  return FailedCallException(XcharToString(message));
}

// Checks for an escape keypress in Excel and reacts appropriately
void CheckForEscapeKeypress() {
  // Reference link:
  // http://msdn.microsoft.com/en-us/library/office/bb687825%28v=office.15%29.aspx
  static XCHAR ShowCancelDialogName[WCHARBUF];
  ConvertToXcharIfNeeded(ShowCancelDialogName, SHOW_CANCEL_DIALOG_NAME);

  static XLOPER12 xOpAbort;
  Excel12f(xlAbort, &xOpAbort, 0);
  BOOL escapePressed = xOpAbort.val.xbool;
  Excel12f(xlFree, nullptr, 1, &xOpAbort);

  if (escapePressed) {
    static XLOPER12 xOpConfirm;
    bool successful = false;
    BOOL abort = false;
    int ret = Excel12f(xlUDF, &xOpConfirm, 1, TempStr12(ShowCancelDialogName));
    if (CheckReturnCodeOkay(ret) && xOpConfirm.xltype == xltypeBool) {
      successful = true;
      abort = xOpConfirm.val.xbool;
    }
    Excel12f(xlFree, nullptr, 1, &xOpConfirm);

    if (!successful) {
      throw MakeFailedCallException(ShowCancelDialogName);
    }

    if (abort) {
      throw std::exception("Aborting through user action");
    }

    // Clear the escape key press so we can resume
    Excel12f(xlAbort, nullptr, 1, TempBool12(false));
  }
}

// Interface implementations

void GetLogFilePath(std::string* logPath) {
  static XCHAR GetLogFilePathName[WCHARBUF];
  ConvertToXcharIfNeeded(GetLogFilePathName, GET_LOG_FILE_PATH_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetLogFilePathName));
  if (CheckReturnCodeOkay(ret) && CheckIsArray(xResult, 2)) {
    successful = true;
    *logPath = GetStringFromExcel(xResult.val.array.lparray);
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetLogFilePathName);
  }
}

void GetNumConstraints(int* numCons, int* numObjs) {
  static XCHAR GetNumConstraintsName[WCHARBUF];
  ConvertToXcharIfNeeded(GetNumConstraintsName, GET_NUM_CONSTRAINTS_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetNumConstraintsName));
  if (CheckReturnCodeOkay(ret) && CheckIsArray(xResult, 2)) {
    successful = true;
    *numCons = static_cast<int>(xResult.val.array.lparray[0].val.num);
    *numObjs = static_cast<int>(xResult.val.array.lparray[1].val.num);
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetNumConstraintsName);
  }
}

void GetNumVariables(int* numVars) {
  static XCHAR GetNumVariablesName[WCHARBUF];
  ConvertToXcharIfNeeded(GetNumVariablesName, GET_NUM_VARIABLES_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetNumVariablesName));
  if (CheckReturnCodeOkay(ret) && xResult.xltype == xltypeNum) {
    successful = true;
    *numVars = static_cast<int>(xResult.val.num);
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetNumVariablesName);
  }
}

void GetVariableData(int numVars, double* lowerBounds, double* upperBounds,
                     double* startingX, int* varTypes) {
  static XCHAR GetVariableDataName[WCHARBUF];
  ConvertToXcharIfNeeded(GetVariableDataName, GET_VARIABLE_DATA_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetVariableDataName));
  if (CheckReturnCodeOkay(ret) && CheckIsArray(xResult, 4 * numVars)) {
    successful = true;

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
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetVariableDataName);
  }
}

void GetOptionData(std::string** paramStrings, int* numOptions) {
  static XCHAR GetOptionDataName[WCHARBUF];
  ConvertToXcharIfNeeded(GetOptionDataName, GET_OPTION_DATA_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetOptionDataName));
  if (CheckReturnCodeOkay(ret) && xResult.xltype == xltypeMulti) {
    successful = true;

    *numOptions = xResult.val.array.rows;
    *paramStrings = new std::string[*numOptions];

    for (int i = 0; i < *numOptions; ++i) {
      XLOPER12* stringData = xResult.val.array.lparray + 2 * i;
      (*paramStrings)[i] = GetStringFromExcel(stringData);
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetOptionDataName);
  }
}

void UpdateVars(double* newVars, int numVars, const double* bestSolution,
                bool feasibility) {
  // Set up the variant array of new variable values.
  XLOPER12 *xOpArray = new XLOPER12[numVars];
  for (int i = 0; i < numVars; i++) {
    xOpArray[i].xltype = xltypeNum;
    xOpArray[i].val.num = newVars[i];
  }

  // Create container for the one-column variant array of variable values.
  XLOPER12 xOpMulti;
  xOpMulti.xltype = xltypeMulti|xlbitDLLFree;
  xOpMulti.val.array.lparray = xOpArray;
  xOpMulti.val.array.columns = 1;
  xOpMulti.val.array.rows = static_cast<RW>(numVars);

  // Pass solution in as Double, or vbNothing if no solution

  XLOPER12 xOpSol;
  if (bestSolution == nullptr) {
    xOpSol.xltype = xltypeMissing|xlbitXLFree;
  } else {
    xOpSol.xltype = xltypeNum|xlbitXLFree;
    xOpSol.val.num = *bestSolution;
  }

  // Do update
  static XCHAR UpdateVarName[WCHARBUF];
  ConvertToXcharIfNeeded(UpdateVarName, UPDATE_VAR_NAME);

  bool successful = false;
  int ret = Excel12f(xlUDF, nullptr, 4, TempStr12(UpdateVarName),
                      &xOpMulti, &xOpSol, TempBool12(!feasibility));
  if (CheckReturnCodeOkay(ret)) {
    successful = true;
  }
  delete[] xOpArray;

  if (!successful) {
    throw MakeFailedCallException(UpdateVarName);
  }
}

void RecalculateValues() {
  static XCHAR RecalculateValuesName[WCHARBUF];
  ConvertToXcharIfNeeded(RecalculateValuesName, RECALCULATE_VALUES_NAME);

  int ret = Excel12f(xlUDF, nullptr, 1, TempStr12(RecalculateValuesName));
  if (!CheckReturnCodeOkay(ret)) {
    throw MakeFailedCallException(RecalculateValuesName);
  }
}

void GetConstraintValues(int numCons, double* newCons) {
  static XCHAR GetValuesName[WCHARBUF];
  ConvertToXcharIfNeeded(GetValuesName, GET_VALUES_NAME);

  static XLOPER12 xResult;
  bool successful = false;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetValuesName));
  if (CheckReturnCodeOkay(ret) && CheckIsArray(xResult, numCons)) {
    successful = true;
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
  }

  // Free memory allocated by Excel
  Excel12f(xlFree, nullptr, 1, &xResult);

  if (!successful) {
    throw MakeFailedCallException(GetValuesName);
  }
}

void EvaluateX(double* newVars, int numVars, int numCons,
               const double* bestSolution, bool feasibility, double *newCons) {
  CheckForEscapeKeypress();
  UpdateVars(newVars, numVars, bestSolution, feasibility);
  RecalculateValues();
  GetConstraintValues(numCons, newCons);
}

}  // namespace OPENSOLVER
