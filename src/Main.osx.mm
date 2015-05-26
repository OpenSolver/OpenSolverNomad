// Main.osx.mm
// Functions for the final executable on OS X

#import <Foundation/Foundation.h>

#include "ExcelCallbacks.h"

#include <string>

int main(int argc, const char * argv[]) {
  @autoreleasepool {
    std::string logFilePath;
    OPENSOLVER::GetLogFilePath(&logFilePath);
    NSLog(@"%@", [NSString stringWithUTF8String:logFilePath.c_str()]);

    double newVars[4] = { 2.5, 3.2, 4.0, 5.999 };
    double bestSolution = 4.0;
    OPENSOLVER::UpdateVars(newVars, 4, &bestSolution, true);

    NSLog(@"Hello, World!");
  }
  return 0;
}
