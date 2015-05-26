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
    
//    std::string logFilePath;
//    OPENSOLVER::GetLogFilePath(&logFilePath);
//    NSLog(@"%@", [NSString stringWithUTF8String:logFilePath.c_str()]);
//
//    double newVars[4] = { 2.5, 3.2, 4.0, 5.999 };
//    double bestSolution = 4.0;
//    OPENSOLVER::UpdateVars(newVars, 4, &bestSolution, true);
//
//    double newCons[10];
//    int numCons = 5;
//    OPENSOLVER::GetConstraintValues(numCons, newCons);
//
//    int numObjs;
//    OPENSOLVER::GetNumConstraints(&numCons, &numObjs);
//
//    int numVars;
//    OPENSOLVER::GetNumVariables(&numVars);
//
//    double* lowerBounds = new double[100];
//    double* upperBounds = new double[100];
//    double* startingX = new double[100];
//    int* varTypes = new int[100];
//    OPENSOLVER::GetVariableData(numVars, lowerBounds, upperBounds, startingX, varTypes);
//
//    std::string* params;
//    int numStrings;
//    OPENSOLVER::GetOptionData(&params, &numStrings);
//    NSLog(@"Hello, World!");
    return OPENSOLVER::RunNomad();
  }
}
