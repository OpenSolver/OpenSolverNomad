// Main.osx.mm
// Functions for the final executable on OS X

#import <Foundation/Foundation.h>

// Foundation Framework defines 'check' macro that interferes with 'check' functions in NOMAD
#undef check

#include "ExcelCallbacks.h"
#include "NomadInterface.h"

#include <string>

int main(int argc, const char * argv[]) {
  @autoreleasepool {
    if (argc == 2) {
      const char* arg = argv[argc - 1];
      if (strcmp(arg, "-v") == 0) {
        printf(OPENSOLVER::VERSION);
        return EXIT_SUCCESS;
      } else if (strcmp(arg, "-nv") == 0) {
        printf("%s", NOMAD::VERSION.c_str());
        return EXIT_SUCCESS;
      }
    } else if (argc > 2) {
      //error
    }
    
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
