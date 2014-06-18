OpenSolverNomadDll
-----------------------------

This is the C++ NOMAD interface that links OpenSolver and NOMAD together. This project builds both the 32 and 64 bit dlls that are used by OpenSolver to call NOMAD.

This Dll is released under the GNU GPL License - See "OpenSolver GNU GPL License.txt" in the OpenSolver folder.

The output of this solution is: OpenSolverNomadDll.dll and OpenSolverDll64.dll


UDFs.cpp
- the OpenSolver specific code that sits between NOMAD and Excel's VBA

libnomad.lib
- pre-compiled NOMAD code provided by Jonathon Currie (of OPTI Matlab add-in)

src
- the CPP and .h NOMAD files

Excel2010XLLSDK
- Microsoft XLL SDK (software development kit)
- not allowed to be redistributed
- This can be downloaded from microsoft.com
- It is "MICROSOFT EXCEL 2010 SOFTWARE DEVELOPMENT KIT (SDK)"

NOMAD source
-The NOMAD source files. OpenSolverNomadDll needs the header files in here to run.


Building the dll
-----------------------------
This dll requires the Microsoft XLL SDK (software development kit) to be downloaded and installed to run. These files will need to be put into the file Excel2010XLLSDK in the OpenSolverNomadDll folder. This can be downloaded from the microsoft website (http://www.microsoft.com/en-us/download/details.aspx?id=20199). You will need to copy the INCLUDE, LIB and SRC folders from your downloaded file into this Excel2010XLLSDK folder. The main files that are needed are the libraries frmwrk32.lib and xlcall.lib from the LIB folder as well as the header files in the INCLUDE folder.


32 bit OpenSolverNomadDll.dll
-----------------------------
This should now be ready to build the 32 bit dll (to run on 32 bit Microsoft Office). Once the project has been built the OpenSolverNomadDll.dll file will be found in the Release folder from the main page of the solution.You will now need to copy this file (OpenSolverNomadDll.dll) into the main OpenSolver folder (where OpenSolver.xlam is) so that OpenSolver is using your version of the dll.


64 bit OpenSolverNomadDll.dll
-----------------------------
The 64 bit dll requires a bit more work to build. You will need to copy the file xlcall.cpp from the SRC folder of Excel2010XLLSDK into the main solution page. Then in the C++ project files right click on Soucre Files->Add->Existing Item... and choose the xlcall.cpp file. Under the configuration manager make sure that you are building your project with the x64 compiler not the win32 one. If this is not an option click on configuration manager... Then under platform click <add..> and choose x64. This project should now be ready to build. When this is built OpenSolverNomad.dll will be made in the folder x64->Release from the main project folder. You will then need to rename this folder OpenSolverNomadDll64.dll before copying it into the OpenSolver directory (where OpenSolver.xlam is) so that OpenSolver can use it.

Properties
-----------------------------
If you are having problems building your project right click on your project in the solution manager and go to properties. 
Under C/C++->Additional Include Directories you should have the directories of the NOAMD source code as well as the include files for Excel2010XLLSDK folder.
Under Linker->General->Additional Library Directories should be the directories of the correct versions (32 or 64 bit) Nomad Library libnomad.lib as well as xlcall32.lib and also the directory of frmwrk32.lib.
Under Linker->Input->Additional Dependencies should be xlcall32.lib;frmwrk32.lib;libnomad.lib
