// NomadInterface.cpp

#include "NomadInterface.h"
#include "ExcelCallbacks.h"

#include <atlbase.h>
#include <stdio.h>
#include <windows.h>

#include <string>
#include <vector>

namespace OPENSOLVER {

NOMAD::bb_input_type VarTypeToNomad(int varType) {
  switch (varType) {
    case CONTINUOUS:
      return NOMAD::CONTINUOUS;
    case INTEGER:
      return NOMAD::INTEGER;
    case BINARY:
      return NOMAD::BINARY;
    default:
      throw "Unknown variable type";
  }
}

}  // namespace OPENSOLVER

NOMAD::Mads *mads;

//*--------------------------------------*/
//*            custom evaluator          */
//*--------------------------------------*/
class Excel_Evaluator : public NOMAD::Evaluator {
 private:
  int      _n;
  int      _m;
  double * _px;
  double * _fx;
  NOMAD::Mads* _mads;

 public:
  Excel_Evaluator(const NOMAD::Parameters &p, int n, int m) : Evaluator(p),
        _n(n),
        _m(m),
        _px(new double[_n]),
        _fx(new double[_m]) {}

  ~Excel_Evaluator(void) { delete [] _px; delete [] _fx; _mads = nullptr; }

  // eval_x:
  bool eval_x(NOMAD::Eval_Point& x,
              const NOMAD::Double& h_max,
              bool& count_eval) const override;
};

// eval_x:
bool Excel_Evaluator::eval_x(NOMAD::Eval_Point& x,
                             const NOMAD::Double& /*h_max*/,
                             bool& count_eval) const {
  for (int i = 0; i < _n; ++i) {
      _px[i] = x[i].value();
  }

  // Get current solution for status updating
  bool feasibility = true;
  const NOMAD::Eval_Point *bestPoint = mads->get_best_feasible();
  double* bestSol = nullptr;
  if (bestPoint == nullptr) {
    bestPoint = mads->get_best_infeasible();
    feasibility = false;
  }

  if (bestPoint != nullptr) {
    double bestValue = bestPoint->get_f().value();
    bestSol = &bestValue;
  }

  try {
    OPENSOLVER::EvaluateX(_px, _n, _m, bestSol, feasibility, _fx);
  } catch (exception&) {
     mads->force_quit(0);
     return false;
  }
  for (int i = 0; i < _m; ++i) {
      x.set_bb_output(i, _fx[i]);
  }
  count_eval = true;
  return true;
}

extern "C" BSTR _stdcall NomadVersion() {
  return CComBSTR(NOMAD::VERSION.c_str());
}

extern "C" BSTR _stdcall NomadDLLVersion() {
  return CComBSTR(OPENSOLVER::DLL_VERSION);
}

// This function must be called directly within VBA using
// retCode = NomadMain(SolveRelaxation).
// If Application.Run is used, the Excel12f calls will fail in 64-bit Office.
// TODO: try to remove this unused bool, seems to crash Excel if we take it out
extern "C" int _stdcall NomadMain(bool) {
  std::string logFilePath;
  try {
    // Get a temp path to write parameters etc to
    OPENSOLVER::GetLogFilePath(&logFilePath);
  } catch (exception&) {
    return OPENSOLVER::LOG_FILE_FAILED;
  }

  ofstream logFile(logFilePath, ios::out);
  NOMAD::Display out(logFile);
  out.precision(NOMAD::DISPLAY_PRECISION_STD);

  try {
    NOMAD::begin(0, nullptr);

    // Variable information
    int numVars;
    OPENSOLVER::GetNumVariables(&numVars);

    if (numVars < 1) {
      throw std::exception("No variables returned");
    }

    double * const lowerBounds =   new double[numVars];
    double * const upperBounds =   new double[numVars];
    double * const startingPoint = new double[numVars];
    int * const varTypes =         new int[numVars];

    OPENSOLVER::GetVariableData(numVars, lowerBounds, upperBounds,
                                startingPoint, varTypes);
    for (int i = 0; i < numVars; ++i) {
      if (upperBounds[i] >= 1e10) {
        upperBounds[i] = NOMAD::INF;
      }
    }

    NOMAD::Point x0(numVars);
    NOMAD::Point ub(numVars);
    NOMAD::Point lb(numVars);
    vector<NOMAD::bb_input_type> bbit(numVars);
    for (int i = 0; i < numVars; i++) {
      ub[i] = upperBounds[i];
      lb[i] = lowerBounds[i];
      x0[i] = startingPoint[i];
      bbit[i] = OPENSOLVER::VarTypeToNomad(varTypes[i]);
    }

    delete[] lowerBounds;
    delete[] upperBounds;
    delete[] startingPoint;
    delete[] varTypes;

    // Constraint/Objective info
    int numCons;
    int numObjs;
    OPENSOLVER::GetNumConstraints(&numCons, &numObjs);

    vector<NOMAD::bb_output_type> bbot(numCons);
    for (int i = 0; i < numObjs; i++) {
      bbot[i] = NOMAD::OBJ;
    }
    for (int i = numObjs; i < numCons; ++i) {
      bbot[i] = NOMAD::EB;
    }

    // User options
    string *paramStrings;
    int numStrings;
    OPENSOLVER::GetOptionData(&paramStrings, &numStrings);

    NOMAD::Parameter_Entries entries;
    NOMAD::Parameter_Entry *pe;
    string err;
    bool invalid = false;
    for (int i = 0; i < numStrings; ++i) {
      pe = new NOMAD::Parameter_Entry(*(paramStrings + i));
      if (pe->is_ok()) {
        entries.insert(pe);  // pe will be deleted by ~Parameter_Entries()
      } else {
        if ((pe->get_name() != "" && pe->get_nb_values() == 0) ||
            pe->get_name() == "STATS_FILE") {
          err = "invalid parameter: " + pe->get_name();
          invalid = true;
        }
        delete pe;
        if (invalid) {
          throw std::exception(err.c_str());
        }
      }
    }
    delete[] paramStrings;

    // Set all parameters
    NOMAD::Parameters p(out);
    p.set_DIMENSION(numVars);
    p.set_X0(x0);
    p.set_UPPER_BOUND(ub);
    p.set_LOWER_BOUND(lb);
    p.set_BB_INPUT_TYPE(bbit);
    p.set_BB_OUTPUT_TYPE(bbot);
    p.set_DISPLAY_STATS("bbe ( sol ) obj");
    p.read(entries);

    p.check();

    // Display parameters:
    out << p << endl;

    // Run NOMAD
    Excel_Evaluator ev(p, numVars, numCons);
    mads = new NOMAD::Mads (p, &ev);
    NOMAD::stop_type stopflag = mads->run();
    NOMAD::Slave::stop_slaves(out);
    NOMAD::end();

    bool feasibility = true;
    // Obtain Solution
    const NOMAD::Eval_Point *bestSol = mads->get_best_feasible();
    if (bestSol == nullptr) {
      bestSol = mads->get_best_infeasible();
      // Manually mark infeasibility (there isn't an infeasible flag)
      feasibility = false;
    }
    if (bestSol != nullptr) {
      double * const finalVars = new double[numVars];
      for (int i = 0; i < numVars; ++i) {
        finalVars[i] = (*bestSol)[i].value();
      }
      const double bestPoint = bestSol->get_f().value();
      OPENSOLVER::UpdateVars(finalVars, numVars, &bestPoint, feasibility);
      delete[] finalVars;
    }

    // Check if it reached the bounds of time and iterations
    int retval = 0;
    if (mads->get_stats().get_real_time() == p.get_max_time()) {
      retval = 3;
    } else if (mads->get_stats().get_bb_eval() == p.get_max_bb_eval()) {
      retval = 2;
    }

    // Free Memory
    delete mads;

    out << endl << endl << "NOMAD Solve Return Value: " << retval << endl;
    logFile.close();

    // Return values
    if (stopflag == NOMAD::CTRL_C) {
      retval = -3;
    } else if ((retval != 0) & (!feasibility)) {
      retval = 4;
    } else if (!feasibility) {
      retval = 10;
    }

    return retval;
  }
  catch (exception& e) {
    NOMAD::Slave::stop_slaves(out);
    NOMAD::end();
    out << e.what() << endl;
    logFile.close();
    return EXIT_FAILURE;
  }
}

