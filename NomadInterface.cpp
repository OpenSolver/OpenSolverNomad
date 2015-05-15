// NomadInterface.cpp

#include "NomadInterface.h"
#include "ExcelCallbacks.h"

#include <atlbase.h>
#include <stdio.h>
#include <windows.h>

#include <vector>

namespace OPENSOLVER {

NOMAD::bb_input_type VarTypeToNomad(int varType) {
  switch (varType) {
    case OPENSOLVER::CONTINUOUS:
      return NOMAD::CONTINUOUS;
    case OPENSOLVER::INTEGER:
      return NOMAD::INTEGER;
    case OPENSOLVER::BINARY:
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
  // ctor:
  Excel_Evaluator(const NOMAD::Parameters &p, int n, int m)
      : NOMAD::Evaluator(p),
        _n(n),
        _m(m),
        _px(new double[_n]),
        _fx(new double[_m]) {}

  // dtor:
  ~Excel_Evaluator(void) { delete [] _px; delete [] _fx; _mads = NULL; }

  // eval_x:
  bool eval_x(NOMAD::Eval_Point& x,
              const NOMAD::Double& h_max,
              bool& count_eval) const;
};

// eval_x:
bool Excel_Evaluator::eval_x(NOMAD::Eval_Point& x,
                             const NOMAD::Double& h_max,
                             bool& count_eval) const {
  for (int i = 0; i < _n; ++i) {
      _px[i] = x[i].value();
  }

  // Get current solution for status updating
  bool feasibility = true;
  const NOMAD::Eval_Point *bestPoint = mads->get_best_feasible();
  double* bestSol = NULL;
  if (bestPoint == NULL) {
    bestPoint = mads->get_best_infeasible();
    feasibility = false;
  }

  if (bestPoint != NULL) {
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
int _stdcall NomadMain(bool SolveRelaxation) {
  // Get a temp path to write parameters etc to
  DWORD dwRetVal = 0;
  UINT uRetVal   = 0;
  TCHAR lpTempPathBuffer[MAX_PATH];
  TCHAR szTempFileName[MAX_PATH];
  dwRetVal = GetTempPath(MAX_PATH, lpTempPathBuffer);

  // Generates a temporary file name.
  uRetVal = GetTempFileName(lpTempPathBuffer, TEXT("log"), 1,
                            szTempFileName);

  // display:
  ofstream myfile;
  myfile.open(szTempFileName, ios::out);

  /*===Need to try this- Added to work with Andres Sommerhoff's==============
  =====changes to getTempFolder which gives the user the option==============
  =====of changing their temp file through environment variables=============

  //check whether there is a temp path specified by the user in 
  //environment variables 
  char * EnvTempPath;
  EnvTempPath=getenv("OpenSolverTempPath");
  if (EnvTempPath!=NULL) {
      myfile.close();
      string strPath;
      strPath.append(EnvTempPath);
      strPath.append("\\Nom1.tmp");
      myfile.open(strPath, ios::out);
  }
  ===================================================================*/

  NOMAD::Display out(myfile);
  out.precision(NOMAD::DISPLAY_PRECISION_STD);

  try {
    // NOMAD initializations:
    NOMAD::begin(0, NULL);

    int n = OPENSOLVER::GetNumVariables();

    // If no variables are retrieved from Excel (due to an error or
    // otherwise), we cannot proceed.
    if (n < 1) {
      throw "No variables returned";
    }

    double * const lowerBounds = new double[n];
    double * const upperBounds = new double[n];
    double * const startingPoint = new double[n];
    int * const varTypes = new int[n];

    OPENSOLVER::GetVariableData(n, lowerBounds, upperBounds, startingPoint,
                                varTypes);

    // Initialise m(number of Constraints) and n(number of objectives)
    int m = 0;
    int nobj = 1;
    OPENSOLVER::GetNumConstraints(&m, &nobj);

    // parameters creation:
    // --------------------
    NOMAD::Parameters p(out);

    // Dimension:
    p.set_DIMENSION(n);

    // Definition of input types:
    vector<NOMAD::bb_input_type> bbit(n);
    for (int i = 0; i < n; i++) {
      bbit[i] = OPENSOLVER::VarTypeToNomad(
          static_cast<int>(OPENSOLVER::VarType(varTypes[i])));
    }

    p.set_BB_INPUT_TYPE(bbit);

    // Set upper and lower bounds and starting position
    NOMAD::Point x0(n);
    NOMAD::Point ub(n);
    NOMAD::Point lb(n);
    for (int i = 0; i < n; i++) {
      ub[i] = upperBounds[i];
      lb[i] = lowerBounds[i];
      x0[i] = startingPoint[i];
    }
    p.set_X0(x0);
    p.set_UPPER_BOUND(ub);
    p.set_LOWER_BOUND(lb);

    // definition of output types:
    vector<NOMAD::bb_output_type> bbot(m);
    for (int i = 0; i < nobj; i++) {
      bbot[i] = NOMAD::OBJ;
    }
    for (int i = nobj; i < m; ++i) {
      bbot[i] = NOMAD::EB;
    }
    p.set_BB_OUTPUT_TYPE(bbot);

    p.set_DISPLAY_STATS("bbe ( sol ) obj");

    // set user options
    OPENSOLVER::GetOptionData(&p);

    // parameters check:
    p.check();

    // display parameters:
    out << p << endl;

    // Nomad vars
    NOMAD::stop_type stopflag;

    // ========Running Nomad with Single Objective=========================
    // custom evaluator creation:
    Excel_Evaluator ev(p, n, m);
    // algorithm creation and execution:
    mads = new NOMAD::Mads (p, &ev);
    stopflag = mads->run();

    // End nomad run
    NOMAD::Slave::stop_slaves(out);
    NOMAD::end();

    bool feasibility = true;
    // Obtain Solution
    const NOMAD::Eval_Point *bestSol = mads->get_best_feasible();
    if (bestSol == NULL) {
      bestSol = mads->get_best_infeasible();
      // Manually mark infeasibility (there isn't an infeasible flag)
      feasibility = false;
    }
    if (bestSol != NULL) {
      double * const fx = new double[m];
      double * const px = new double[n];
      for (int i = 0; i < n; ++i) {
        px[i] = (*bestSol)[i].value();
      }
      const double bestPoint = bestSol->get_f().value();
      OPENSOLVER::EvaluateX(px, n, m, &bestPoint, feasibility, fx);
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
    myfile.close();

    // Return values
    if (stopflag == NOMAD::CTRL_C) {
      return -3;
    } else if ((retval != 0) & (!feasibility)) {
      retval = 4;
      return retval;
    } else if (!feasibility) {
      retval = 10;
      return retval;
    } else if (retval != 0) {
      return retval;
    } else {
      return EXIT_SUCCESS;
    }
  }
  catch (exception& e) {
    NOMAD::Slave::stop_slaves(out);
    NOMAD::end();
    out << e.what() << endl;
    myfile.close();
    return EXIT_FAILURE;
  }

  return EXIT_SUCCESS;
}

