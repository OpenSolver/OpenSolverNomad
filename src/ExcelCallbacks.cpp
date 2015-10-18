// ExcelCallbacks.cpp
// Implementations of ExcelCallbacks.h common to all platforms

#include "ExcelCallbacks.hpp"

#include <string>

namespace OPENSOLVER {

EXCEL_RC AddLocationIfError(EXCEL_RC rc, int location) {
  if (rc != SUCCESS && rc < LOCATION_OFFSET) {
    rc += location * LOCATION_OFFSET;
  }
  return rc;
}

EXCEL_RC GetErrorCode(EXCEL_RC rc) {
  return rc % LOCATION_OFFSET;
}

EXCEL_RC GetLocation(EXCEL_RC rc) {
  return rc / LOCATION_OFFSET;
}

std::string GetExcelCallbackErrorMessage(EXCEL_RC rc) {
  EXCEL_RC err = GetErrorCode(rc);
  std::string messageErr;
  switch (err) {
    case ESC_ABORT:
      messageErr = "Aborted due to user cancellation.";
      break;
    case EXCEL_API_ERROR:
      messageErr = "Error contacting Excel.";
      break;
    case EXCEL_VBA_ERROR:
      messageErr = "An error occured inside Excel while running.";
      break;
    case EXCEL_INVALID_RETURN:
      messageErr = "Excel returned an invalid value.";
      break;
    default:
      messageErr = "Unknown error.";
      break;
  }
  EXCEL_RC location = GetLocation(rc);
  std::string messageLocation;
  switch (location) {
    case SHOW_CANCEL_DIALOG_NUM:
      messageLocation = "ShowCancelDialog";
      break;
    case CHECK_ESC_PRESS_NUM:
      messageLocation = "CheckEscapeKeypress";
      break;
    case GET_LOG_FILE_PATH_NUM:
      messageLocation = "GetLogFilePath";
      break;
    case GET_NUM_CONSTRAINTS_NUM:
      messageLocation = "GetNumConstraints";
      break;
    case GET_NUM_VARIABLES_NUM:
      messageLocation = "GetNumVariables";
      break;
    case GET_VARIABLE_DATA_NUM:
      messageLocation = "GetVariableData";
      break;
    case GET_OPTION_DATA_NUM:
      messageLocation = "GetOptionData";
      break;
    case UPDATE_VARS_NUM:
      messageLocation = "UpdateVars";
      break;
    case RECALCULATE_VALUES_NUM:
      messageLocation = "RecalculateValues";
      break;
    case GET_CONSTRAINT_VALUES_NUM:
      messageLocation = "GetConstraintValues";
      break;
    default:
      messageLocation = "unknown";
      break;
  }

  return messageErr + " Location: " + messageLocation + ".";
}

void ValidateReturnCode(EXCEL_RC rc) {
  if (rc != SUCCESS) {
    throw std::runtime_error(GetExcelCallbackErrorMessage(rc));
  }
}

EXCEL_RC EvaluateX(double* newVars, int numVars, int numCons,
                   const double* bestSolution, bool feasibility,
                   double* newCons) {
  EXCEL_RC rc;

  rc = CheckForEscapeKeypress(true);
  if (rc != SUCCESS) {
    goto ErrorHandler;
  }

  rc = UpdateVars(newVars, numVars, bestSolution, feasibility);
  if (rc != SUCCESS) {
    goto ErrorHandler;
  }

  rc = RecalculateValues();
  if (rc != SUCCESS) {
    goto ErrorHandler;
  }

  rc = GetConstraintValues(numCons, newCons);
  if (rc != SUCCESS) {
    goto ErrorHandler;
  }

  return SUCCESS;

ErrorHandler:
  // Confirm whether the error is the result of an escape keypress
  if (GetErrorCode(CheckForEscapeKeypress(false)) == ESC_ABORT) {
    return ESC_ABORT;
  } else {
    return rc;
  }
}

}  // namespace OPENSOLVER
