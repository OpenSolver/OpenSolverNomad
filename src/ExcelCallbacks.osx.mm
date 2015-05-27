//  ExcelCallbacks.osx.mm

#include "ExcelCallbacks.h"

#import <Carbon/Carbon.h>
#import <Foundation/Foundation.h>

#include <string>

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
  event = [NSAppleEventDescriptor appleEventWithEventClass: kASAppleScriptSuite
                                                   eventID: kASSubroutineEvent
                                          targetDescriptor: NSAppleEventDescriptor.nullDescriptor
                                                  returnID: kAutoGenerateReturnID
                                             transactionID: kAnyTransactionID];
  if (params != nil) {
    [event setDescriptor: params forKeyword: keyDirectObject];
  }
  [event setDescriptor: [NSAppleEventDescriptor descriptorWithString:functionName]
            forKeyword: keyASSubroutineName];
  
  NSAppleScript *script = GetCompiledScript();
  NSDictionary *errorDict;
  NSAppleEventDescriptor *result = [script executeAppleEvent:event error:&errorDict];
  // TODO error checking on error dict
  return result;
}

NSAppleEventDescriptor* GetVectorEntry(NSAppleEventDescriptor* vector, NSInteger i) {
  return [vector descriptorAtIndex:i];
}

NSAppleEventDescriptor* GetMatrixEntry(NSAppleEventDescriptor* matrix, NSInteger i, NSInteger j) {
  return GetVectorEntry(GetVectorEntry(matrix, i), j);
}

double ConvertNSDataToDouble(NSData* data) {
  double d;
  assert([data length] == sizeof(d));
  memcpy(&d, [data bytes], sizeof(d));
  return d;
}

double ConvertDescriptorToDouble(NSAppleEventDescriptor* result) {
  //TODO check type of descriptor
  return ConvertNSDataToDouble([result data]);
}

int ConvertDescriptorToInt(NSAppleEventDescriptor* result) {
  //TODO check type of descriptor
  return [result int32Value];
}

std::string ConvertDescriptorToString(NSAppleEventDescriptor* result) {
  // TODO error checking
  assert([result numberOfItems] == 2);
  // TODO type check
  NSString *stringValue = [GetVectorEntry(result, 1) stringValue];
  int stringLength = ConvertDescriptorToInt(GetVectorEntry(result, 2));

  if (stringValue.length != stringLength) {
    // Path returned didn't have the correct length
    // TODO throw exception
    stringValue = @"";
  }
  return std::string([stringValue UTF8String]);
}

extern "C" {

void GetLogFilePath(std::string* logPath) {
  NSAppleEventDescriptor *result = RunScriptFunction(@"getLogFilePath", nil);
  *logPath = ConvertDescriptorToString(GetVectorEntry(result, 1));
}

void GetNumConstraints(int* numCons, int* numObjs) {
  NSAppleEventDescriptor* result = RunScriptFunction(@"getNumConstraints", nil);
  *numCons = ConvertDescriptorToInt(GetVectorEntry(result, 1));
  *numObjs = ConvertDescriptorToInt(GetVectorEntry(result, 2));
}

void GetNumVariables(int* numVars) {
  NSAppleEventDescriptor* result = RunScriptFunction(@"getNumVariables", nil);
  // TODO typecheck
  *numVars = [result int32Value];
}

void GetVariableData(int numVars, double* lowerBounds, double* upperBounds, double* startingX,
                                int* varTypes) {
  NSAppleEventDescriptor* result = RunScriptFunction(@"getVariableData", nil);
  // TODO error handle
  assert([result numberOfItems] == 4 * numVars);
  for (int i = 0; i < numVars; ++i) {
    lowerBounds[i] = ConvertDescriptorToDouble(GetVectorEntry(result, 2 * i + 1));
    upperBounds[i] = ConvertDescriptorToDouble(GetVectorEntry(result, 2 * i + 2));
    startingX[i]   = ConvertDescriptorToDouble(GetVectorEntry(result, 2 * numVars + i + 1));
    varTypes[i]    = ConvertDescriptorToInt   (GetVectorEntry(result, 3 * numVars + i + 1));
  }
}

void GetOptionData(std::string** paramStrings, int* numOptions) {
  NSAppleEventDescriptor* result = RunScriptFunction(@"getOptionData", nil);
  *numOptions = (int)[result numberOfItems];
  *paramStrings = new std::string[*numOptions];
  for (int i = 0; i < *numOptions; ++i) {
    (*paramStrings)[i] = ConvertDescriptorToString(GetVectorEntry(result, i + 1));
  }
}

void UpdateVars(double* newVars, int numVars, const double* bestSolution,
                bool feasibility) {
  
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
  
  RunScriptFunction(@"updateVars", params);
}

void RecalculateValues() {
  RunScriptFunction(@"recalculateValues", nil);
}

void GetConstraintValues(int numCons, double* newCons) {
  NSAppleEventDescriptor* result = RunScriptFunction(@"getConstraintValues", nil);

  if ([result numberOfItems] == numCons) {
    for (int i = 0; i < numCons; ++i) {
      newCons[i] = ConvertDescriptorToDouble(GetMatrixEntry(result, i + 1, 1));
    }
  } else {
    // There aren't the full number of constraints
    // Set all results to NaN
    for (int i = 0; i < numCons; ++i) {
      newCons[i] = std::numeric_limits<double>::quiet_NaN();
    }
  }

}

void EvaluateX(double* newVars, int numVars, int numCons, const double* bestSolution,
                          bool feasibility, double* newCons) {
  // TODO see if we can detect an escape keypress?
  UpdateVars(newVars, numVars, bestSolution, feasibility);
  RecalculateValues();
  GetConstraintValues(numCons, newCons);
}

void LoadResult(int retVal) {
  NSAppleEventDescriptor *params = [NSAppleEventDescriptor listDescriptor];
  [params insertDescriptor:[NSAppleEventDescriptor descriptorWithInt32:retVal] atIndex:1];
  RunScriptFunction(@"loadResult", params);
}

}  // extern "C"
