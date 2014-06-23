// UDFs.c

//needs to be defined to stop a warning on compiling
#define _ITERATOR_DEBUG_LEVEL 0

#include "nomad.hpp"

#include <windows.h>

#include <xlcall.h>

#include <limits>

#include <stdio.h>

#include <string>

using namespace std;

void GetNumConstraints(int* numCons, int* nObj);
int GetNumVariables(void);
void EvaluateX(double *newVars, double size, double *newCons, double numCons);
void GetVariableData(double *LowerBounds, double *UpperBounds, double *X0, int *type, double numVars);
void getOptionData(double *OptionData);

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
	Excel_Evaluator ( const NOMAD::Parameters & p , int n , int m )
	  : NOMAD::Evaluator ( p               ) ,
		_n        ( n               ) ,
		_m        ( m               ) ,
		_px       ( new double [_n] ) ,
		_fx       ( new double [_m] )   {}

	// dtor:
	~Excel_Evaluator ( void ) { delete [] _px; delete [] _fx; }

	// eval_x:
	bool eval_x ( NOMAD::Eval_Point   & x        ,
		const NOMAD::Double & h_max    ,
		bool         & cnt_eval   ) const;
};

// eval_x:
bool Excel_Evaluator::eval_x ( NOMAD::Eval_Point   & x        ,
			    const NOMAD::Double & h_max    ,
			    bool         & cnt_eval   ) const {
	int i;
	for ( i = 0 ; i < _n ; ++i )
		_px[i] = x[i].value();
	EvaluateX(_px , _n,_fx,_m);
	for ( i = 0 ; i < _m ; ++i )
		x.set_bb_output ( i , _fx[i] );
	cnt_eval=true;
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

long _stdcall NomadMain (bool SolveRelaxation)
{
	//get a temp path to write parameters etc to
	DWORD dwRetVal = 0;
	UINT uRetVal   = 0;
	TCHAR lpTempPathBuffer[MAX_PATH];
	TCHAR szTempFileName[MAX_PATH]; 
	dwRetVal = GetTempPath(MAX_PATH,lpTempPathBuffer); 

	//  Generates a temporary file name. 
    uRetVal = GetTempFileName(lpTempPathBuffer, TEXT("log"), 1, szTempFileName); 

	// display:
	ofstream myfile;
	myfile.open(szTempFileName, ios::out);

	/*===========Need to try this- Added to work with Andres Sommerhoff's================
	=============changes to getTempFolder which gives the user the option==============
	=============of changing their temp file through environment variables=============

	//check whether there is a temp path specified by the user in environment variables 
	char * EnvTempPath;
	EnvTempPath=getenv("OpenSolverTempPath");
	if (EnvTempPath!=NULL) {
		myfile.close();
		string strPath;
		strPath.append(EnvTempPath);
		strPath.append("\\Nom1.tmp");
		myfile.open(strPath, ios::out);
	}
	==================================================================================*/

	NOMAD::Display out(myfile);
	out.precision ( NOMAD::DISPLAY_PRECISION_STD );

	try {
		// NOMAD initializations:
		NOMAD::begin ( 0 , NULL );
		
		int n = GetNumVariables();

		// If no variables are retrieved from Excel (due to an error or otherwise), we cannot proceed.
		if (n < 1) {
			return (long) EXIT_FAILURE;
		}

		double * const LowerBounds = new double[(int) n];
		double * const UpperBounds = new double[(int) n];
		double * const startingPoint = new double[(int) n];
		int * const varType = new int [(int) n];
		bool * const binaryVar = new bool [(int) n];

		GetVariableData(LowerBounds,UpperBounds,startingPoint,varType,n);

		//initialise m(number of Constraints) and n(number of objectives)
		int m=0;
		int nobj=1;
		GetNumConstraints(&m,&nobj);
		
		// parameters creation:
		// --------------------
		NOMAD::Parameters p ( out );

		// dimension:
		p.set_DIMENSION ( n );

		//definition of input types:
		vector<NOMAD::bb_input_type> bbit (n);
		for ( int i = 0 ; i < n ; ++i ) {
			if (!SolveRelaxation) {
				switch(varType[i])
				{
					case 1:
						bbit[i] = NOMAD::CONTINUOUS; break;
					case 2:
						bbit[i] = NOMAD::INTEGER; break;
					case 3:
						bbit[i] = NOMAD::BINARY; break;
				}
				binaryVar[i]=false;
			}
			else {
				//if solving a relaxation make all variables continuous
				bbit[i] = NOMAD::CONTINUOUS;
				switch (varType[i])
				{
					case 1:
					case 2:
						binaryVar[i]=false; break;
					case 3:
						binaryVar[i]=true; break;
				}	
			}
		}
		p.set_BB_INPUT_TYPE ( bbit );

		//Setting upper and lower bounds and starting position
		NOMAD::Point x0 (n);
		NOMAD::Point ub (n);                    
		NOMAD::Point lb (n);                    
		for (int i=0;i<n;i++) {
			if (binaryVar[i]==false) {
				ub[i]=UpperBounds[i];
				lb[i]=LowerBounds[i];
				x0[i]=startingPoint[i];
			}
			//if solve relaxation and binary variable make bounds between 0-1
			else {
				ub[i]=1;
				lb[i]=0;
				x0[i]=0;
			}
		}
		p.set_X0 (x0);
		p.set_UPPER_BOUND ( ub );
		p.set_LOWER_BOUND ( lb );

		// definition of output types:
		vector<NOMAD::bb_output_type> bbot (m);
		for(int i=0;i<nobj;i++)
            bbot[i] = NOMAD::OBJ;
		for ( int i = nobj ; i < m ; ++i )
			bbot[i] = NOMAD::EB;
		p.set_BB_OUTPUT_TYPE ( bbot );

		// p.set_DISPLAY_DEGREE ( FULL_DISPLAY );

		//getting parameter data
		double * const OptionData = new double[(int) 3];
		getOptionData(OptionData);

		p.set_DISPLAY_STATS ( "bbe ( sol ) obj" );

		//==========set user options============================================
		p.set_MAX_BB_EVAL ((int)OptionData[0]);
		p.set_MAX_TIME ((int)OptionData[1]);
		p.set_EPSILON(OptionData[2]);
		//p.set_MIN_MESH_SIZE(OptionData[2]);
		
		// parameters check:
		p.check();

		// display parameters:
		out << p << endl;

		//Nomad vars
		NOMAD::Mads *mads;
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

		//=========Running Nomad with Single Objective=============================
		// custom evaluator creation:
		Excel_Evaluator ev ( p , n , m );
		// algorithm creation and execution:
		mads = new NOMAD::Mads ( p , &ev  );
		stopflag = mads->run();

		//=========End of Nomad run, Clean up and get values back==================
		// algorithm display:
		//out << mads << endl;
		
		//end nomad run
		NOMAD::Slave::stop_slaves ( out );
		NOMAD::end();

		bool feasibility = true;
		//Obtain Solution
		const NOMAD::Eval_Point *bestSol = mads->get_best_feasible();
		if(bestSol == NULL) {
			bestSol = mads->get_best_infeasible();  
			//manually set as infeasible (no infeasible stop flag)
			feasibility=false;
		}
		if (bestSol!=NULL) {
			double * const fx = new double[(int) m];
			double * const px = new double[(int) n];
			for ( int i = 0 ; i < n ; ++i ) {
				px[i] = (*bestSol)[i].value();
			}
			EvaluateX(px , n,fx,m);
		}

		//check if it reached the bounds of time and iterations
		long retval=0;
		if (mads->get_stats().get_real_time()==OptionData[1])
		{
			retval = 3;
		}
		else if (mads->get_stats().get_bb_eval()==OptionData[0]){
			retval = 2;
		}
		
		//Free Memory
		//if(mSEval) delete mSEval; mSEval = NULL; //for multi-obj
		//if(mBEval) delete mBEval; mBEval = NULL; //for multi-obj
		delete mads;

		out<< endl << endl << "NOMAD Solve Return Value: " << retval << endl;
		myfile.close();

		//return values
		if ((retval !=0) & (!feasibility)) {
			retval=4;
			return retval;
		}
		else if (!feasibility) {
			retval=10;
			return retval;
		}
		else if (retval !=0) {
			return retval;
		}
		else {
			return (long) EXIT_SUCCESS;
		}
	}
	catch ( ... ) {
		NOMAD::Slave::stop_slaves ( out );
		NOMAD::end();
		myfile.close();
		return (long) EXIT_FAILURE;
	}

	return (long) EXIT_SUCCESS;
}

/*=========================================================================================
  OpenSolver VBA Function calls to evaluate model and spreadsheet for NOMAD
==========================================================================================*/

//Calls excel to get the number of constraints
//outputs=number of constraints including number of objectives and the number of objectives
void GetNumConstraints(int* numCons, int* nObj)
{
	static XLOPER12 xResult;
	XLOPER12 funcName;

	funcName.val.str=L"\034OpenSolver.getNumConstraints";
	funcName.xltype=xltypeStr;

	Excel12(xlUDF,&xResult,1,&funcName);
	*numCons=(int)xResult.val.array.lparray[0].val.num;
	*nObj=(int)xResult.val.array.lparray[1].val.num;

	return;
}

//Calls excel to get the number of variables
//outputs=number of variables
int GetNumVariables(void)
{
	XLOPER12 xRes;
    static HANDLE hOld = 0;
    short iRet;

    if (Excel12(xlGetInst, &xRes, 0) != xlretSuccess)
        iRet = -1;
    else
    {
    HANDLE hNew;

    hNew = (HANDLE)xRes.val.w;
    if (hNew != hOld)
            iRet = 0;
    else
            iRet = 1;
    hOld = hNew;
    }


	static XLOPER12 xResult;
	XLOPER12 funcName;
	funcName.val.str=L"\034OpenSolver.getNumVariables";
	funcName.xltype=xltypeStr;

	Excel12(xlUDF,&xResult,1,&funcName);

	return (int) xResult.val.num;
}

//Calls excel to evaluate each new point of X
//inputs:	newVars=the new values to put into the sheet
//			size=number of variables
//			newCons=the values of the constraints evaluated by excel at the new point(these our outputs)
//			numCons=number of constraints
void EvaluateX(double *newVars, double size, double *newCons, double numCons)
{
	static XLOPER12 xResult;
	XLOPER12 funcName, funcName1, funcName2;

// In this implementation, the upper limit is the largest
// single column array (equals 2^20, or 1048576, rows in Excel 2007).
    if(size < 1 || size > 1048576)
        return;

// Create an array of XLOPER12 values.
    XLOPER12 *xOpArray = (XLOPER12 *)malloc((size_t)size * sizeof(XLOPER12));

// Create and initialize an xltypeMulti array
// that represents a one-column array.
    XLOPER12 xOpMulti;
    xOpMulti.xltype = xltypeMulti|xlbitDLLFree;
    xOpMulti.val.array.lparray = xOpArray;
    xOpMulti.val.array.columns = 1;
    xOpMulti.val.array.rows = (RW) size;

// Initialize and populate the array of XLOPER12 values.
    for(unsigned short i = 0; i < size; i++)
    {
        xOpArray[i].xltype = xltypeNum;
        xOpArray[i].val.num = *(newVars+i);
    }

	funcName.xltype=xltypeStr;
	funcName.val.str=L"\024OpenSolver.updateVar";
	funcName1.xltype=xltypeStr;
	funcName1.val.str=L"\024OpenSolver.getValues";
	funcName2.xltype=xltypeStr;
	funcName2.val.str=L"\034OpenSolver.RecalculateValues";

	Excel12(xlUDF, 0, 2, &funcName, &xOpMulti);
	Excel12(xlUDF, 0,1,&funcName2);
	Excel12(xlUDF,&xResult,1,&funcName1);
	Excel12(xlFree,0,1,&xOpMulti);
	
	for (unsigned short i=0;i<numCons;i++) {
		// Check for error passed back from VBA and set to C++ NaN.
		// We need to catch errors separately as they are otherwise interpreted as having value zero.
		if (xResult.val.array.lparray[i].xltype != xltypeNum) {
		    *(newCons+i) = std::numeric_limits<double>::quiet_NaN();
		} else {
			*(newCons+i)=xResult.val.array.lparray[i].val.num;
		}
	}
	return;
}

//Calls to excel that get the variable data such as bounds, starting points and variable types
//inputs:	LowerBounds=returned lower bounds of each variable from excel
//			UpperBounds=returned upper bounds of each variable from excel
//			X0=returned starting point for solve (must be within bounds, this is checked by excel)
//			type=returned type of variable(continuous,integer,binary)
//			numVars=number of variables
void GetVariableData(double *LowerBounds, double *UpperBounds, double *X0, int *type, double numVars)
{
	static XLOPER12 xResult;

	XLOPER12 funcName;
	funcName.val.str=L"\035OpenSolver.getVariableData";
	funcName.xltype=xltypeStr;
	
	Excel12(xlUDF,&xResult,1,&funcName);
	
	//get the lower and upper bounds for each of the variables
	for (int i=0;i<numVars;i++) {
		*(LowerBounds+i)=xResult.val.array.lparray[2*i].val.num;
		*(UpperBounds+i)=xResult.val.array.lparray[2*i+1].val.num;
	}

	//get start point
	for ( int i=0;i<numVars;i++)
		*(X0+i)=xResult.val.array.lparray[2*(int)numVars+i].val.num;

	//get the variable types (real,integer,binary)
	for ( int i=0;i<numVars;i++)
		*(type+i)=(int)xResult.val.array.lparray[3*(int)numVars+i].val.num;
	return;
}

//Save the users options for tolerance and time limits etc.
//inputs:	OptionData[0]=max iterations
//			OptionData[1]=max time
//			OptionData[2]=tolerance-epsilon
void getOptionData(double *OptionData)
{
	static XLOPER12 xResult;
	XLOPER12 funcName;
	funcName.val.str=L"\033OpenSolver.getOptionData";
	funcName.xltype=xltypeStr;
	
	Excel12(xlUDF,&xResult,1,&funcName);

	for ( int i=0;i<3;i++)
		*(OptionData+i)=xResult.val.array.lparray[i].val.num;

	return;
}
