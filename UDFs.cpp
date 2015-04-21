// UDFs.c

#include "nomad.hpp"

#include <atlbase.h>
#include <stdio.h>
#include <windows.h>

#include <xlcall.h>
#include <framewrk.h>

#include <limits>
#include <string>
#include <vector>

const char DLL_VERSION[] = "1.1.0";

void GetNumConstraints(int* numCons, int* nObj);
int GetNumVariables(void);
void EvaluateX(double *newVars, int size, double *newCons, int numCons);
void GetVariableData(double *LowerBounds, double *UpperBounds, double *X0,
                     int *type, int numVars);
void GetOptionData(NOMAD::Parameters *p);

//*--------------------------------------*/
//*            custom evaluator          */
//*--------------------------------------*/
class Excel_Evaluator : public NOMAD::Evaluator {
 private:
  int      _n;
  int      _m;
  double * _px;
  double * _fx;

 public:
  // ctor:
  Excel_Evaluator(const NOMAD::Parameters &p, int n, int m)
      : NOMAD::Evaluator(p),
        _n(n),
        _m(m),
        _px(new double[_n]),
        _fx(new double[_m]) {}

  // dtor:
  ~Excel_Evaluator(void) { delete [] _px; delete [] _fx; }

  // eval_x:
  bool eval_x(NOMAD::Eval_Point & x, const NOMAD::Double & h_max,
              bool & cnt_eval) const;
};

// eval_x:
bool Excel_Evaluator::eval_x(NOMAD::Eval_Point & x,
                             const NOMAD::Double & h_max,
                             bool & cnt_eval) const {
  for (int i = 0; i < _n; ++i) {
      _px[i] = x[i].value();
  }
  EvaluateX(_px , _n, _fx, _m);
  for (int i = 0; i < _m; ++i) {
      x.set_bb_output(i, _fx[i]);
  }
  cnt_eval = true;
  return true;
}

/*====================================================================================
Nomad multi objective class- could work for bi objectives but need to add support into
OpenSolver
======================================================================================
//Nomad MultiObj Evaluator Class
class XllMulti_Evaluator : public NOMAD::Multi_Obj_Evaluator {
private:
    Xll_Evaluator *mEval;
public:
    //Constructor
    XllMulti_Evaluator(const NOMAD::Parameters & p , int n , int m) : NOMAD::Multi_Obj_Evaluator(p)
    {
        mEval = new Xll_Evaluator(p,n,m);
    }
    //Deconstructor
    ~XllMulti_Evaluator(void)
    {
        delete mEval;
    }
    //Function + Constraint Information
    bool eval_x(NOMAD::Eval_Point &x, const NOMAD::Double &h_max, bool &count_eval)
    {
        return mEval->eval_x(x,h_max,count_eval);
    }        
};
========================================================================================*/

extern "C" BSTR _stdcall NomadVersion() {
  return CComBSTR(NOMAD::VERSION.c_str());
}

extern "C" BSTR _stdcall NomadDLLVersion() {
  return CComBSTR(DLL_VERSION);
}

NOMAD::Mads *mads;

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

    int n = GetNumVariables();

    // If no variables are retrieved from Excel (due to an error or
    // otherwise), we cannot proceed.
    if (n < 1) {
      throw "No variables returned";
    }

    double * const LowerBounds = new double[n];
    double * const UpperBounds = new double[n];
    double * const startingPoint = new double[n];
    int * const varType = new int[n];
    bool * const setBinaryBounds = new bool[n];

    GetVariableData(LowerBounds, UpperBounds, startingPoint, varType, n);

    // Initialise m(number of Constraints) and n(number of objectives)
    int m = 0;
    int nobj = 1;
    GetNumConstraints(&m, &nobj);

    // parameters creation:
    // --------------------
    NOMAD::Parameters p(out);

    // Dimension:
    p.set_DIMENSION(n);

    // Definition of input types:
    vector<NOMAD::bb_input_type> bbit(n);
    for (int i = 0; i < n; ++i) {
      if (!SolveRelaxation) {
        switch (varType[i]) {
          case 1:
            bbit[i] = NOMAD::CONTINUOUS;
            break;
          case 2:
            bbit[i] = NOMAD::INTEGER;
            break;
          case 3:
            bbit[i] = NOMAD::BINARY;
            break;
        }
        setBinaryBounds[i] = false;
      } else {
        // If solving a relaxation make all variables continuous
        bbit[i] = NOMAD::CONTINUOUS;
        switch (varType[i]) {
          case 1:
          case 2:
            setBinaryBounds[i] = false;
            break;
          case 3:
            setBinaryBounds[i] = true;
            break;
        }
      }
    }
    p.set_BB_INPUT_TYPE(bbit);

    // Set upper and lower bounds and starting position
    NOMAD::Point x0(n);
    NOMAD::Point ub(n);
    NOMAD::Point lb(n);
    for (int i = 0; i < n; i++) {
      if (!setBinaryBounds[i]) {
        ub[i] = UpperBounds[i];
        lb[i] = LowerBounds[i];
        x0[i] = startingPoint[i];
      } else {
        // If solving relaxation make bounds between 0 and 1
        ub[i] = 1;
        lb[i] = 0;
        x0[i] = 0;
      }
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

    // p.set_DISPLAY_DEGREE ( FULL_DISPLAY );

    p.set_DISPLAY_STATS("bbe ( sol ) obj");

    // set user options
    GetOptionData(&p);

    // parameters check:
    p.check();

    // display parameters:
    out << p << endl;

    // Nomad vars
    NOMAD::stop_type stopflag;

    /*=======================================================================
    Running Nomad for Multi Objective (bi-objective) - no support for this in
    OpenSolver yet
    =========================================================================
    //p.set_MULTI_OVERALL_BB_EVAL ((int)OptionData[0]); //could be set for multi obj

    //Evaluator Vars
    Xll_Evaluator *mSEval = NULL;
    XllMulti_Evaluator *mBEval = NULL;

    //Create evaluator and run mads based on number of objectives
    try
    {     
        if(nobj > 1) {
            mBEval = new XllMulti_Evaluator(p,n,m); //Bi-Objective Evaluator
            mads = new NOMAD::Mads(p, mBEval); //Run NOMAD  
            stopflag = mads->multi_run();
        }
        else {
            mSEval = new Xll_Evaluator(p,n,m); //Single Objective Evaluator
            mads = new NOMAD::Mads(p, mSEval); //Run NOMAD 
            stopflag = mads->run();
        }
    }
    catch(exception &e)
    {
        out<<"NOMAD Run Error:\n\n"<<e.what();
    }
    */

    // ========Running Nomad with Single Objective=========================
    // custom evaluator creation:
    Excel_Evaluator ev(p, n, m);
    // algorithm creation and execution:
    mads = new NOMAD::Mads (p, &ev);
    stopflag = mads->run();

    // ========End of Nomad run, Clean up and get values back==============
    // algorithm display:
    // out << mads << endl;

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
      EvaluateX(px, n, fx, m);
    }

    // Check if it reached the bounds of time and iterations
    int retval = 0;
    if (mads->get_stats().get_real_time() == p.get_max_time()) {
      retval = 3;
    } else if (mads->get_stats().get_bb_eval() == p.get_max_bb_eval()) {
      retval = 2;
    }

    // Free Memory
    // if(mSEval) delete mSEval; mSEval = NULL; //for multi-obj
    // if(mBEval) delete mBEval; mBEval = NULL; //for multi-obj
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

/*=========================================================================================
  OpenSolver VBA Function calls to evaluate model and spreadsheet for NOMAD
==========================================================================================*/

// Calls excel to get the number of constraints
// outputs = number of constraints (inc. objectives) and number of objectives
void GetNumConstraints(int* numCons, int* nObj) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetNumConstraints"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != 2) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_GetNumConstraints failed";
  }

  *numCons = static_cast<int>(xResult.val.array.lparray[0].val.num);
  *nObj = static_cast<int>(xResult.val.array.lparray[1].val.num);

  // Free up Excel-allocated array
  Excel12f(xlFree, 0, 1, &xResult);
  return;
}

// Calls excel to get the number of variables
// outputs = number of variables
int GetNumVariables(void) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetNumVariables"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeNum) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_GetNumVariables failed";
  }

  int result = static_cast<int>(xResult.val.num);

  Excel12f(xlFree, 0, 1, &xResult);
  return result;
}

// Calls excel to evaluate each new point of X
// inputs:  newVars = the new values to put into the sheet
//          size    = number of variables
//          newCons = the values of the constraints evaluated at the new point
//          numCons = number of constraints
void EvaluateX(double *newVars, int size, double *newCons, int numCons) {
  XLOPER12 xOpAbort, xOpConfirm;

  // Check for escape key press
  // http://msdn.microsoft.com/en-us/library/office/bb687825%28v=office.15%29.aspx
  Excel12f(xlAbort, &xOpAbort, 0);
  if (xOpAbort.val.xbool) {
    Excel12f(xlFree, 0, 2, &xOpAbort);

    int ret = Excel12f(xlUDF, &xOpConfirm, 1,
                       TempStr12(L"OpenSolver.NOMAD_ShowCancelDialog"));
    if (ret == xlretAbort || ret == xlretUncalced ||
        xOpConfirm.xltype != xltypeBool) {
      Excel12f(xlFree, 0, 1, &xOpConfirm);
      throw "NOMAD_ShowCancelDialog failed";
    }

    if (xOpConfirm.val.xbool) {
      Excel12f(xlFree, 0, 1, &xOpConfirm);
      mads->force_quit(0);
      return;
    } else {
        // Clear the escape key press so we can resume
      Excel12f(xlAbort, 0, 1, TempBool12(false));
    }

    Excel12f(xlFree, 0, 1, &xOpConfirm);
  }

  static XLOPER12 xResult;

  // In this implementation, the upper limit is the largest
  // single column array (equals 2^20, or 1048576, rows in Excel 2007).
  if (size < 1 || size > 1048576) {
    return;
  }

  // Create an array of XLOPER12 values.
  XLOPER12 *xOpArray = static_cast<XLOPER12*>(malloc((size_t)size *
                                              sizeof(XLOPER12)));

  // Create and initialize an xltypeMulti array
  // that represents a one-column array.
  XLOPER12 xOpMulti;
  xOpMulti.xltype = xltypeMulti|xlbitDLLFree;
  xOpMulti.val.array.lparray = xOpArray;
  xOpMulti.val.array.columns = 1;
  xOpMulti.val.array.rows = (RW) size;

  // Initialize and populate the array of XLOPER12 values.
  for (int i = 0; i < size; i++) {
    xOpArray[i].xltype = xltypeNum;
    xOpArray[i].val.num = *(newVars+i);
  }

  // Get current solution for status updating
  bool feasibility = true;
  const NOMAD::Eval_Point *bestSol = mads->get_best_feasible();
  if (bestSol == NULL) {
    bestSol = mads->get_best_infeasible();
    feasibility = false;
  }

  // Create XLOPER12 objects for passing in solution and feasibility

  // Pass solution in as Double, or vbNothing if no solution
  XLOPER12 xOpSol;
  if (bestSol == NULL) {
    xOpSol.xltype = xltypeMissing|xlbitXLFree;
  } else {
    xOpSol.xltype = xltypeNum|xlbitXLFree;
    xOpSol.val.num = bestSol->get_f().value();
  }

  int ret;

  // Update variables
  ret = Excel12f(xlUDF, &xResult, 4, TempStr12(L"OpenSolver.NOMAD_UpdateVar"),
                 &xOpMulti, &xOpSol, TempBool12(!feasibility));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_UpdateVar failed";
  }

  // Recalculate values
  ret = Excel12f(xlUDF, 0, 1,
                 TempStr12(L"OpenSolver.NOMAD_RecalculateValues"));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_RecalculateValues failed";
  }

  // Get constraint values
  ret = Excel12f(xlUDF, &xResult, 1, TempStr12(L"OpenSolver.NOMAD_GetValues"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != numCons) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_GetValues failed";
  }

  for (int i = 0; i < numCons; i++) {
    // Check for error passed back from VBA and set to C++ NaN.
    // We need to catch errors separately as they are otherwise interpreted
    // as having value zero.
    if (xResult.val.array.lparray[i].xltype != xltypeNum) {
      *(newCons + i) = std::numeric_limits<double>::quiet_NaN();
    } else {
      *(newCons + i) = xResult.val.array.lparray[i].val.num;
    }
  }

  // Free memory allocated by Excel
  Excel12f(xlFree, 0, 1, &xResult);
  return;
}

// Gets the variable data (bounds, starting points and variable types)
// inputs:  LowerBounds = Lower bounds of each variable from Excel
//          UpperBounds = Upper bounds of each variable from Excel
//          X0          = Starting point for each variable (must be within
//                        bounds, this is enforced by excel)
//          type        = Type of variable (continuous, integer, binary)
//          numVars     = Number of variables
void GetVariableData(double *LowerBounds, double *UpperBounds, double *X0,
                     int *type, int numVars) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetVariableData"));
  if (ret == xlretAbort || ret == xlretUncalced ||
      xResult.xltype != xltypeMulti ||
      xResult.val.array.rows * xResult.val.array.columns != 4 * numVars) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_GetVariableData failed";
  }

  // Get the lower and upper bounds for each of the variables
  for (int i = 0; i < numVars; i++) {
    *(LowerBounds + i)= xResult.val.array.lparray[2 * i].val.num;
    *(UpperBounds + i)= xResult.val.array.lparray[2 * i + 1].val.num;
    if (*(UpperBounds+i) >= 1e10) {
      *(UpperBounds+i) = NOMAD::INF;
    }
  }

  // Get start point
  for (int i = 0; i < numVars; i++) {
    *(X0 + i) = xResult.val.array.lparray[2 * numVars + i].val.num;
  }

  // Get the variable types (real, integer, or binary)
  for (int i = 0; i < numVars; i++) {
    double vartype = xResult.val.array.lparray[3 * numVars + i].val.num;
    *(type + i) = static_cast<int>(vartype);
  }

  // Free Excel-allocated memory
  Excel12f(xlFree, 0, 1, &xResult);
  return;
}

// Save the users options for tolerance and time limits etc.
// inputs:  OptionData[0]=max iterations
//          OptionData[1]=max time
//          OptionData[2]=tolerance-epsilon
void GetOptionData(NOMAD::Parameters *p) {
  static XLOPER12 xResult;

  int ret = Excel12f(xlUDF, &xResult, 1,
                     TempStr12(L"OpenSolver.NOMAD_GetOptionData"));
  if (ret == xlretAbort || ret == xlretUncalced) {
    Excel12f(xlFree, 0, 1, &xResult);
    throw "NOMAD_GetOptionData failed";
  }

  NOMAD::Parameter_Entries entries;
  NOMAD::Parameter_Entry *pe;
  std::string s;
  std::string err;
  wstring ws;
  int n;
  int m = xResult.val.array.rows;

  for (int i = 0; i < m; ++i) {
    // Get the string value out of the result
    n = static_cast<int>(xResult.val.array.lparray[2 * i + 1].val.num);
    ws = wstring(xResult.val.array.lparray[2 * i].val.str);
    s = string(ws.begin(), ws.end()).substr(1, n);

    // Add the parameter to the entries
    pe = new NOMAD::Parameter_Entry(s);
    if (pe->is_ok()) {
      entries.insert(pe);  // pe will be deleted by ~Parameter_Entries()
    } else {
      if ((pe->get_name() != "" && pe->get_nb_values() == 0) ||
          pe->get_name() == "STATS_FILE") {
        err = "invalid parameter: " + pe->get_name();
        delete pe;
        throw err;
      }
      delete pe;
    }
  }

  // Read all the new entries into p
  p->read(entries);

  // Free Excel-allocated memory
  Excel12f(xlFree, 0, 1, &xResult);
  return;
}
