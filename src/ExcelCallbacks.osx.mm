// ExcelCallbacks.osx.mm
// Implementation of ExcelCallbacks.hpp for OS X

#include "ExcelCallbacks.hpp"

#import <Carbon/Carbon.h>
#import <Foundation/Foundation.h>

#include <string>

namespace OPENSOLVER {

NSAppleScript* GetCompiledScript() {
  static NSAppleScript* cachedScript;
  if (cachedScript == nil) {
    NSDictionary *errorDict;
    
    // Get path to folder containing executable
    NSString *execPath = NSProcessInfo.processInfo.arguments[0];
    NSURL *execUrl = [NSURL fileURLWithPath:execPath].URLByDeletingLastPathComponent;
    
    // Load compiled script from same folder as executable
    NSURL *scriptUrl = [[NSURL alloc]
                        initWithString:@"ExcelCallbacks.osx.scpt" relativeToURL:execUrl];
    cachedScript = [[NSAppleScript alloc] initWithContentsOfURL:scriptUrl error:&errorDict];
  }
  return cachedScript;
}

NSAppleEventDescriptor* RunScriptFunction(NSString* functionName, NSAppleEventDescriptor *params) {
  // Build Apple event to invoke user-defined handler in script
  // See http://appscript.sourceforge.net/nsapplescript.html
  NSAppleEventDescriptor *event;
  event = [NSAppleEventDescriptor appleEventWithEventClass:kASAppleScriptSuite
                                                   eventID:kASSubroutineEvent
                                          targetDescriptor:NSAppleEventDescriptor.nullDescriptor
                                                  returnID:kAutoGenerateReturnID
                                             transactionID:kAnyTransactionID];
  if (params != nil) {
    [event setDescriptor:params forKeyword:keyDirectObject];
  }
  [event setDescriptor:[NSAppleEventDescriptor descriptorWithString:functionName]
            forKeyword:keyASSubroutineName];
  
  NSAppleScript *script = GetCompiledScript();
  NSDictionary *errorDict;
  NSAppleEventDescriptor *result = [script executeAppleEvent:event error:&errorDict];
  // return nil if there was an error
  return (errorDict == nil) ? result : nil;
}

EXCEL_RC CheckReturn(NSAppleEventDescriptor* result) {
  if (result == nil) {
    return EXCEL_API_ERROR;
  } else if (result.descriptorType == 'long' && result.int32Value == -1) {
    return EXCEL_VBA_ERROR;
  }
  return SUCCESS;
}

bool CheckListDescriptor(NSAppleEventDescriptor* result, int expectedNumItems) {
  return result.descriptorType == 'list' && result.numberOfItems == expectedNumItems;
}

NSAppleEventDescriptor* GetVectorEntry(NSAppleEventDescriptor* vector, NSInteger i) {
  return [vector descriptorAtIndex:i];
}

NSAppleEventDescriptor* GetMatrixEntry(NSAppleEventDescriptor* matrix, NSInteger i, NSInteger j) {
  return GetVectorEntry(GetVectorEntry(matrix, i), j);
}

EXCEL_RC ConvertDescriptorToDouble(NSAppleEventDescriptor* result, double* outdoub) {
  if (result.descriptorType == 'doub' && result.data.length == sizeof(double)) {
    memcpy(outdoub, result.data.bytes, sizeof(double));
    return SUCCESS;
  } else {
    return EXCEL_INVALID_RETURN;
  }
}

EXCEL_RC ConvertDescriptorToInt(NSAppleEventDescriptor* result, int* outint) {
  if (result.descriptorType == 'long') {
    *outint = result.int32Value;
    return SUCCESS;
  } else {
    return EXCEL_INVALID_RETURN;
  }
}

bool ConvertDescriptorToBool(NSAppleEventDescriptor* result) {
  return result.int32Value;
}

bool CheckBoolDescriptor(NSAppleEventDescriptor* result) {
  return result.descriptorType == 'fals' || result.descriptorType == 'true';
}

EXCEL_RC ConvertDescriptorToString(NSAppleEventDescriptor* result, std::string* outstr) {
  NSAppleEventDescriptor* stringDesc = GetVectorEntry(result, 1);
  if (stringDesc.descriptorType != 'utxt') {
    return EXCEL_INVALID_RETURN;
  }

  *outstr = stringDesc.stringValue.UTF8String;

  int stringLength;
  EXCEL_RC rc = ConvertDescriptorToInt(GetVectorEntry(result, 2), &stringLength);
  if (rc != SUCCESS) {
    return rc;
  }

  if ((*outstr).length() != stringLength) {
    return EXCEL_INVALID_RETURN;
  }

  return SUCCESS;
}

extern "C" {

EXCEL_RC CheckForEscapeKeypress(bool /* fullCheck */) {
  @autoreleasepool {
    NSAppleEventDescriptor *result = RunScriptFunction(@"getConfirmedAbort", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      if (CheckBoolDescriptor(result)) {
        if (ConvertDescriptorToBool(result)) {
          rc = ESC_ABORT;
        }
      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, CHECK_ESC_PRESS_NUM);
  }
}

EXCEL_RC GetLogFilePath(std::string* logPath) {
  @autoreleasepool {
    NSAppleEventDescriptor *result = RunScriptFunction(@"getLogFilePath", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      result = GetVectorEntry(result, 1);
      if (CheckListDescriptor(result, 2)) {
        rc = ConvertDescriptorToString(result, logPath);
      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, GET_LOG_FILE_PATH_NUM);
  }
}

EXCEL_RC GetNumConstraints(int* numCons, int* numObjs) {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"getNumConstraints", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      result = GetVectorEntry(result, 1);
      if (CheckListDescriptor(result, 2)) {
        rc = ConvertDescriptorToInt(GetVectorEntry(result, 1), numCons);
        if (rc != SUCCESS) goto ExitFunction;

        rc = ConvertDescriptorToInt(GetVectorEntry(result, 2), numObjs);
        if (rc != SUCCESS) goto ExitFunction;

      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }

ExitFunction:
    return AddLocationIfError(rc, GET_NUM_CONSTRAINTS_NUM);
  }
}

EXCEL_RC GetNumVariables(int* numVars) {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"getNumVariables", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      rc = ConvertDescriptorToInt(result, numVars);
    }
    return AddLocationIfError(rc, GET_NUM_VARIABLES_NUM);
  }
}

EXCEL_RC GetVariableData(int numVars, double* lowerBounds, double* upperBounds, double* startingX,
                         int* varTypes) {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"getVariableData", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      if (CheckListDescriptor(result, 4 * numVars)) {
        for (int i = 0; i < numVars; ++i) {
          rc = ConvertDescriptorToDouble(GetVectorEntry(result, 0 * numVars + i + 1),
                                         lowerBounds + i);
          if (rc != SUCCESS) break;

          rc = ConvertDescriptorToDouble(GetVectorEntry(result, 1 * numVars + i + 1),
                                         upperBounds + i);
          if (rc != SUCCESS) break;

          rc = ConvertDescriptorToDouble(GetVectorEntry(result, 2 * numVars + i + 1),
                                         startingX + i);
          if (rc != SUCCESS) break;

          rc = ConvertDescriptorToInt(GetVectorEntry(result, 3 * numVars + i + 1),
                                      varTypes + i);
          if (rc != SUCCESS) break;
        }
      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, GET_VARIABLE_DATA_NUM);
  }
}

EXCEL_RC GetOptionData(std::string** paramStrings, int* numOptions) {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"getOptionData", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      if (result.descriptorType == 'list') {
        *numOptions = (int)result.numberOfItems;
        *paramStrings = new std::string[*numOptions];
        for (int i = 0; i < *numOptions; ++i) {
          rc = ConvertDescriptorToString(GetVectorEntry(result, i + 1), *paramStrings + i);
          if (rc != SUCCESS) break;
        }
      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, GET_OPTION_DATA_NUM);
  }
}

EXCEL_RC UpdateVars(double* newVars, int numVars, const double* bestSolution,
                bool feasibility) {
  @autoreleasepool {

    // Build array of new variables
    NSAppleEventDescriptor *newVarsContainer = [NSAppleEventDescriptor listDescriptor];
    for (int i = 0; i < numVars; ++i) {
      NSAppleEventDescriptor *value =
      [NSAppleEventDescriptor descriptorWithDescriptorType:'doub'
                                                     bytes:newVars + i
                                                    length:sizeof(double)];
      NSAppleEventDescriptor *innerArray = [NSAppleEventDescriptor listDescriptor];
      [innerArray insertDescriptor:value atIndex:1];
      [newVarsContainer insertDescriptor:innerArray atIndex:(i + 1)];
    }
  
    // Params
    NSAppleEventDescriptor *params = [NSAppleEventDescriptor listDescriptor];
    [params insertDescriptor:newVarsContainer atIndex:1];
    NSAppleEventDescriptor *bestSolutionContainer;
    if (bestSolution == nullptr) {
      bestSolutionContainer = [NSAppleEventDescriptor nullDescriptor];
    } else {
      bestSolutionContainer = [NSAppleEventDescriptor descriptorWithDescriptorType:'doub'
                                                                             bytes:bestSolution
                                                                            length:sizeof(double)];
    }
    [params insertDescriptor:bestSolutionContainer atIndex:2];
    [params insertDescriptor:[NSAppleEventDescriptor descriptorWithBoolean:!feasibility]
                     atIndex:3];
  
    NSAppleEventDescriptor* result = RunScriptFunction(@"updateVars", params);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      int retval;
      rc = ConvertDescriptorToInt(result, &retval);
      if (rc != SUCCESS || retval != 0) {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, UPDATE_VARS_NUM);
  }
}

EXCEL_RC RecalculateValues() {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"recalculateValues", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      int retval;
      rc = ConvertDescriptorToInt(result, &retval);
      if (rc != SUCCESS || retval != 0) {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, RECALCULATE_VALUES_NUM);
  }
}

EXCEL_RC GetConstraintValues(int numCons, double* newCons) {
  @autoreleasepool {
    NSAppleEventDescriptor* result = RunScriptFunction(@"getConstraintValues", nil);
    EXCEL_RC rc = CheckReturn(result);
    if (rc == SUCCESS) {
      if (CheckListDescriptor(result, numCons)) {
        // Copy in result values
        for (int i = 0; i < numCons; ++i) {
          rc = ConvertDescriptorToDouble(GetMatrixEntry(result, i + 1, 1), newCons + i);
          if (rc != SUCCESS) break;
        }
      } else if (result.descriptorType == 'list') {
        // There aren't the full number of constraints, indicating some of the numbers are errors.
        // Set all result values to NaN in response.
        for (int i = 0; i < numCons; ++i) {
          newCons[i] = std::numeric_limits<double>::quiet_NaN();
        }
      } else {
        rc = EXCEL_INVALID_RETURN;
      }
    }
    return AddLocationIfError(rc, GET_CONSTRAINT_VALUES_NUM);
  }
}

void LoadResult(int retVal) {
  @autoreleasepool {
    NSAppleEventDescriptor *params = [NSAppleEventDescriptor listDescriptor];
    [params insertDescriptor:[NSAppleEventDescriptor descriptorWithInt32:retVal] atIndex:1];
    RunScriptFunction(@"loadResult", params);
  }
}

}  // extern "C"

}  // namespace OPENSOLVER
