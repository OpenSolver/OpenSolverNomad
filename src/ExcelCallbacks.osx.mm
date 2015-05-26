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

NSAppleEventDescriptor* GetArrayEntry(NSAppleEventDescriptor* array, NSInteger i, NSInteger j) {
  return [[array descriptorAtIndex:i] descriptorAtIndex:j];
}

extern "C" {

void GetLogFilePath(std::string* logPath) {
  NSAppleEventDescriptor *result = RunScriptFunction(@"getLogFilePath", nil);
  NSString *logFilePath = [GetArrayEntry(result, 1, 1) stringValue];
  NSUInteger pathLength = [GetArrayEntry(result, 1, 2) int32Value];
  
  if (logFilePath.length != pathLength) {
    // Path returned didn't have the correct length
    // TODO throw exception
    logFilePath = @"";
  }
  *logPath = std::string([logFilePath UTF8String]);
}

void GetNumConstraints(int* numCons, int* numObjs) {
  
}

void GetNumVariables(int* numVars) {
  
}

void GetVariableData(int numVars, double* lowerBounds, double* upperBounds, double* startingX,
                                int* varTypes) {
  
}

void GetOptionData(std::string** paramStrings, int* numOptions) {
  
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
  
}

void GetConstraintValues(int numCons, double* newCons) {
  
}

void EvaluateX(double* newVars, int numVars, int numCons, const double* bestSolution,
                          bool feasibility, double* newCons) {
  
}

}  // extern "C"
