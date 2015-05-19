// NomadInterface.h
// Runs NOMAD solver on OpenSolver problem

#ifndef NOMADINTERFACE_H_
#define NOMADINTERFACE_H_

#include "nomad.hpp"

namespace OPENSOLVER {

const char DLL_VERSION[] = "1.1.0";

// Should match the definition in VariableType enum in OpenSolverConsts module
enum VarType {
  CONTINUOUS = 0,
  INTEGER = 1,
  BINARY = 2
};

/**
 * Converts an OpenSolver variable type to the corresponding NOMAD var type.
 */
NOMAD::bb_input_type VarTypeToNomad(int varType);

}  // namespace OPENSOLVER

#endif  // NOMADINTERFACE_H_
