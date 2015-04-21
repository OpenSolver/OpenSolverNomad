# OpenSolver Nomad DLL

This is the C++ NOMAD interface that links OpenSolver and NOMAD together. This project builds both the 32 and 64 bit dlls that are used by OpenSolver to call NOMAD.

The output of this solution is the file `OpenSolverNomadDll.dll`.

##### Inputs:

`UDFs.cpp`
- the OpenSolver specific code that sits between NOMAD and Excel's VBA

`libnomad.lib`
- compiled NOMAD code produced by the [OpenSolverNomadLib](https://github.com/OpenSolver/OpenSolverNomadLib) repository.

`Excel2010XLLSDK`
- Microsoft XLL SDK (software development kit). This is not allowed to be redistributed, but can be downloaded from [microsoft.com](http://www.microsoft.com/en-nz/download/details.aspx?id=20199). After downloading, copy all of the folders into the `Excel2010XLLSDK` folder. You then need to build the `frmwrk32.lib` library, which is contained in the `SAMPLES/FRAMEWRK` folder. You need to make the following changes to the `Makefile` first:
  1. Change the line
      
        ```
        FRAMEWORK_BINARY   = "frmwrk32.lib"
        ```
		
     to
        
        ```make
        !if "$(TYPE)" == "DEBUG"
        FRAMEWORK_BINARY   = "frmwrk32d.lib"
        !else
        FRAMEWORK_BINARY   = "frmwrk32.lib"
        !endif
        ```
	  
  2. Change the `CPPFLAGS` lines to use `/MTd` and `/MT` for the `DEBUG` and `RELEASE` configurations respectively:
        
        ```make
        !if "$(TYPE)" == "DEBUG"
        CPPFLAGS        =/Od /W3 /WX /EHsc /Zi /MTd /Fd"$(FRAMEWORK_PDB)" /Fo"$(FRAMEWORKBUILDDIR)\\"
        !else
        CPPFLAGS        =/W3 /WX /EHsc /MT /Fo"$(FRAMEWORKBUILDDIR)\\"
        !endif
        ```

You also need access to the NOMAD header files. The project assumes that this repository and [OpenSolverNomadLib](https://github.com/OpenSolver/OpenSolverNomadLib) are in the same root directory, and is configured to search for the header files under this assumption. If you change the relative location of the repositories the project will need to be updated to reflect this.

## Building the DLL

The 32 and 64 bit DLLs can be built by selecting either `Win32` or `x64` build configurations as appropriate. The resulting DLL is copied into the appropriate solver folder in the [OpenSolver](https://github.com/OpenSolver/OpenSolver) repository (which should reside in the same directory as this repository).

## Properties

If you are having problems building your project right click on your project in the solution manager and go to properties. 
Under `C/C++`->`Additional Include Directories` you should have the directories of the NOMAD source code as well as the `INCLUDE files` for `Excel2010XLLSDK` folder.

Under `Linker`->`General`->`Additional Library Directories` should be the directories of the correct versions (32 or 64 bit) of the Nomad Library `libnomad.lib` (`libnomad-debug` for debug builds), `xlcall32.lib` and `frmwrk32.lib` (or `frmwrk32d.lib` for debug).
These files should also be included under `Linker`->`Input`->`Additional Dependencies`.

## OpenSolverNomadDll License

The `OpenSolverNomad.dll` files support the use of NOMAD in OpenSolver, and are licensed under the GNU GPL License for use by all OpenSolver users. 

Please see `NOMAD LICENSE.txt` and http://www.gerad.ca/nomad/Project/Home.html for details of the NOMAD license.

© Copyright by Andrew Mason and Matthew Milner, 2013
