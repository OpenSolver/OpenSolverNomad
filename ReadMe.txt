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
This dll requires the Microsoft XLL SDK (software development kit) to be downloaded and installed to run. These files will need to be put into the file Excel2010XLLSDK in the OpenSolverNomadDll folder. This can be downloaded from the microsoft website (http://www.microsoft.com/en-us/download/details.aspx?id=20199). You will need to copy the INCLUDE, LIB and SRC folders from your downloaded file into this Excel2010XLLSDK folder.

The 32 and 64 bit DLLs can be built by selecting either 'Win32' or 'x64' build configurations as appropriate. The resulting DLL is renamed to OpenSolverDll64.dll (only if 64 bit), and moved to the "OpenSolver Release" folder (which should reside in the same directory as the "OpenSolverNomadDll" folder).

Properties
-----------------------------
If you are having problems building your project right click on your project in the solution manager and go to properties. 
Under C/C++->Additional Include Directories you should have the directories of the NOAMD source code as well as the include files for Excel2010XLLSDK folder.
Under Linker->General->Additional Library Directories should be the directories of the correct versions (32 or 64 bit) Nomad Library libnomad.lib as well as xlcall32.lib.
Under Linker->Input->Additional Dependencies should be xlcall32.lib;libnomad.lib
