// ExcelCallbacks.win32.cpp
// Implementation of ExcelCallbacks.hpp for Windows

#include "ExcelCallbacks.hpp"

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
    size_t destLength;
    mbstowcs_s(&destLength, destination, WCHARBUF, source, WCHARBUF);
    if (destLength == -1) {
      std::string error = std::string("Error while converting") + source;
      throw std::exception(error.c_str());
    }
  }
}

// Check the return code from an Excel12 call for errors in the Excel API.
bool CheckReturnCodeOkay(int ret) {
  return ret == xlretSuccess;
}

// Check the return variant from an Excel12 call for errors in the VBA.
// A returned integer with value -1 indicates an error inside the VBA.
bool CheckReturnVariantOkay(const XLOPER12& result) {
  return !(result.xltype == xltypeNum && result.val.num == -1);
}

EXCEL_RC CheckReturn(int ret, const XLOPER12& result) {
  if (!CheckReturnCodeOkay(ret)) {
    return EXCEL_API_ERROR;
  } else if (!CheckReturnVariantOkay(result)) {
    return EXCEL_VBA_ERROR;
  } else {
    return SUCCESS;
  }
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
EXCEL_RC GetStringFromExcel(const XLOPER12* stringData, std::string* outstr) {
  // Get the length of the string
  if (stringData[1].xltype != xltypeNum) {
    return EXCEL_INVALID_RETURN;
  }
  int n = static_cast<int>(stringData[1].val.num);

  if (stringData[0].xltype != xltypeStr) {
    return EXCEL_INVALID_RETURN;
  }
  // Excel puts garbage in the first char
  *outstr = XcharToString(stringData[0].val.str).substr(1, n);
  return SUCCESS;
}

EXCEL_RC ShowCancelDialog() {
  static XCHAR ShowCancelDialogName[WCHARBUF];
  ConvertToXcharIfNeeded(ShowCancelDialogName, SHOW_CANCEL_DIALOG_NAME);
  static XLOPER12 xOpResult;

  // Show confirm escape dialog
  int ret = Excel12f(xlUDF, &xOpResult, 1, TempStr12(ShowCancelDialogName));
  EXCEL_RC rc = CheckReturn(ret, xOpResult);
  if (rc == SUCCESS) {
    if (xOpResult.xltype != xltypeNum || xOpResult.val.num != 0) {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free memory
  Excel12f(xlFree, nullptr, 1, &xOpResult);

  return AddLocationIfError(rc, SHOW_CANCEL_DIALOG_NUM);
}

// Interface implementations

EXCEL_RC CheckForEscapeKeypress(bool fullCheck) {
  // On a full check, we need to check whether Excel has a pending esc press.
  if (fullCheck) {
    // Reference link:
    // http://msdn.microsoft.com/en-us/library/office/bb687825%28v=office.15%29.aspx
    static XLOPER12 xOpAbort;
    Excel12f(xlAbort, &xOpAbort, 0);
    BOOL escapePressed = xOpAbort.val.xbool;
    Excel12f(xlFree, nullptr, 1, &xOpAbort);

    if (escapePressed) {
      // Show dialog to confirm escape keypress
      EXCEL_RC showDialogRc = ShowCancelDialog();
      if (showDialogRc != SUCCESS) {
        return showDialogRc;
      }
      // Clear the escape key press so we can resume
      Excel12f(xlAbort, nullptr, 1, TempBool12(false));
    }
  }

  // Now check if with Excel if an abort has been requested
  static XCHAR GetConfirmedAbort[WCHARBUF];
  ConvertToXcharIfNeeded(GetConfirmedAbort, GET_CONFIRMED_ABORT_NAME);

  static XLOPER12 xOpConfirm;
  int ret = Excel12f(xlUDF, &xOpConfirm, 1, TempStr12(GetConfirmedAbort));

  EXCEL_RC rc = CheckReturn(ret, xOpConfirm);
  if (rc == SUCCESS) {
    if (xOpConfirm.xltype == xltypeBool) {
      if (xOpConfirm.val.xbool) {
        rc = ESC_ABORT;
      }
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xOpConfirm);

  return AddLocationIfError(rc, CHECK_ESC_PRESS_NUM);
}

EXCEL_RC GetLogFilePath(std::string* logPath) {
  static XCHAR GetLogFilePathName[WCHARBUF];
  ConvertToXcharIfNeeded(GetLogFilePathName, GET_LOG_FILE_PATH_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetLogFilePathName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (CheckIsArray(xResult, 2)) {
      rc = GetStringFromExcel(xResult.val.array.lparray, logPath);
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_LOG_FILE_PATH_NUM);
}

EXCEL_RC GetNumConstraints(int* numCons, int* numObjs) {
  static XCHAR GetNumConstraintsName[WCHARBUF];
  ConvertToXcharIfNeeded(GetNumConstraintsName, GET_NUM_CONSTRAINTS_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetNumConstraintsName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (CheckIsArray(xResult, 2)) {
      *numCons = static_cast<int>(xResult.val.array.lparray[0].val.num);
      *numObjs = static_cast<int>(xResult.val.array.lparray[1].val.num);
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_NUM_CONSTRAINTS_NUM);
}

EXCEL_RC GetNumVariables(int* numVars) {
  static XCHAR GetNumVariablesName[WCHARBUF];
  ConvertToXcharIfNeeded(GetNumVariablesName, GET_NUM_VARIABLES_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetNumVariablesName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (xResult.xltype == xltypeNum) {
      *numVars = static_cast<int>(xResult.val.num);
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free up Excel-allocated array
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_NUM_VARIABLES_NUM);
}

EXCEL_RC GetVariableData(int numVars, double* lowerBounds, double* upperBounds,
                     double* startingX, int* varTypes) {
  static XCHAR GetVariableDataName[WCHARBUF];
  ConvertToXcharIfNeeded(GetVariableDataName, GET_VARIABLE_DATA_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetVariableDataName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (CheckIsArray(xResult, 4 * numVars)) {
      for (int i = 0; i < numVars; i++) {
        lowerBounds[i] = xResult.val.array.lparray[0 * numVars + i].val.num;
        upperBounds[i] = xResult.val.array.lparray[1 * numVars + i].val.num;
        startingX[i] =   xResult.val.array.lparray[2 * numVars + i].val.num;
        double rawType = xResult.val.array.lparray[3 * numVars + i].val.num;
        varTypes[i] = static_cast<int>(rawType);
      }
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_VARIABLE_DATA_NUM);
}

EXCEL_RC GetOptionData(std::string** paramStrings, int* numOptions) {
  static XCHAR GetOptionDataName[WCHARBUF];
  ConvertToXcharIfNeeded(GetOptionDataName, GET_OPTION_DATA_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetOptionDataName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (xResult.xltype == xltypeMulti) {
      *numOptions = xResult.val.array.rows;
      *paramStrings = new std::string[*numOptions];

      for (int i = 0; i < *numOptions; ++i) {
        XLOPER12* stringData = xResult.val.array.lparray + 2 * i;
        rc = GetStringFromExcel(stringData, *paramStrings + i);
        if (rc != SUCCESS) {
          break;
        }
      }
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_OPTION_DATA_NUM);
}

EXCEL_RC GetUseWarmstart(bool* useWarmstart) {
  static XCHAR GetUseWarmstartName[WCHARBUF];
  ConvertToXcharIfNeeded(GetUseWarmstartName, GET_USE_WARMSTART_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetUseWarmstartName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (xResult.xltype == xltypeBool) {
      *useWarmstart = (xResult.val.xbool != 0);
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_USE_WARMSTART);
}

EXCEL_RC UpdateVars(double* newVars, int numVars, const double* bestSolution,
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

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 4, TempStr12(UpdateVarName),
                      &xOpMulti, &xOpSol, TempBool12(!feasibility));
  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (xResult.xltype != xltypeNum || xResult.val.num != 0) {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  delete[] xOpArray;
  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, UPDATE_VARS_NUM);
}

EXCEL_RC RecalculateValues() {
  static XCHAR RecalculateValuesName[WCHARBUF];
  ConvertToXcharIfNeeded(RecalculateValuesName, RECALCULATE_VALUES_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(RecalculateValuesName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (xResult.xltype != xltypeNum || xResult.val.num != 0) {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, RECALCULATE_VALUES_NUM);
}

EXCEL_RC GetConstraintValues(int numCons, double* newCons) {
  static XCHAR GetValuesName[WCHARBUF];
  ConvertToXcharIfNeeded(GetValuesName, GET_VALUES_NAME);

  static XLOPER12 xResult;
  int ret = Excel12f(xlUDF, &xResult, 1, TempStr12(GetValuesName));

  EXCEL_RC rc = CheckReturn(ret, xResult);
  if (rc == SUCCESS) {
    if (CheckIsArray(xResult, numCons)) {
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
    } else {
      rc = EXCEL_INVALID_RETURN;
    }
  }

  // Free memory allocated by Excel
  Excel12f(xlFree, nullptr, 1, &xResult);

  return AddLocationIfError(rc, GET_CONSTRAINT_VALUES_NUM);
}

}  // namespace OPENSOLVER
