// NomadInterface.h
// Runs NOMAD solver on OpenSolver problem

#ifndef SRC_NOMADINTERFACE_H_
#define SRC_NOMADINTERFACE_H_

#include "nomad.hpp"

namespace OPENSOLVER {

const char VERSION[] = "1.3.1";

// Should match the definition of VariableType enum in OpenSolverConsts module
enum VarType {
  CONTINUOUS = 0,
  INTEGER = 1,
  BINARY = 2
};

// Should match the definition of NomadResult enum in SolverNomad module
enum NomadResult {
  LOG_FILE_ERROR = -12,
  USER_CANCELLED = -3,
  OPTIMAL = 0,
  ERROR_OCCURED = 1,
  SOLVE_STOPPED_ITER = 2,
  SOLVE_STOPPED_TIME = 3,
  INFEASIBLE = 4,
  SOLVE_STOPPED_ITER_INF = 10,
  SOLVE_STOPPED_TIME_INF = 11
};

/**
 * Runs entire NOMAD process
 */
NomadResult RunNomad();

#ifdef __APPLE__
/**
 * Runs entire NOMAD process and loads result into Excel
 */
NomadResult RunNomadAndLoadResult();
#endif

/**
 * Converts an OpenSolver variable type to the corresponding NOMAD var type.
 */
NOMAD::bb_input_type VarTypeToNomad(int varType);

}  // namespace OPENSOLVER

#endif  // SRC_NOMADINTERFACE_H_
