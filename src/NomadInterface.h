// NomadInterface.h
// Runs NOMAD solver on OpenSolver problem

#ifndef SRC_NOMADINTERFACE_H_
#define SRC_NOMADINTERFACE_H_

#include "nomad.hpp"

namespace OPENSOLVER {

const char VERSION[] = "1.1.2";
const int LOG_FILE_FAILED = -12;

// Should match the definition in VariableType enum in OpenSolverConsts module
enum VarType {
  CONTINUOUS = 0,
  INTEGER = 1,
  BINARY = 2
};

/**
 * Runs entire NOMAD process
 */
int RunNomad();

#ifdef __APPLE__
/**
 * Runs entire NOMAD process and loads result into Excel
 */
int RunNomadAndLoadResult();
#endif

/**
 * Converts an OpenSolver variable type to the corresponding NOMAD var type.
 */
NOMAD::bb_input_type VarTypeToNomad(int varType);

}  // namespace OPENSOLVER

#endif  // SRC_NOMADINTERFACE_H_
