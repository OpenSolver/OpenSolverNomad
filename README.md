# OpenSolver Nomad DLL

This is the C++ NOMAD interface that links OpenSolver and NOMAD together. This project builds both the 32 and 64 bit dlls that are used by OpenSolver to call NOMAD.

The output of this solution is: `OpenSolverNomadDll.dll` and `OpenSolverDll64.dll`

##### Inputs:

`UDFs.cpp`
- the OpenSolver specific code that sits between NOMAD and Excel's VBA

`libnomad.lib`
- compiled NOMAD code produced by the [OpenSolverNomadLib](https://github.com/OpenSolver/OpenSolverNomadLib) repository.

`Excel2010XLLSDK`
- Microsoft XLL SDK (software development kit). This is not allowed to be redistributed, but can be downloaded from [microsoft.com](http://www.microsoft.com/en-nz/download/details.aspx?id=20199). After downloading, copy the `INCLUDE`, `LIB` and `SRC` folders into this folder.

You also need access to the NOMAD header files. The project assumes that this repository and [OpenSolverNomadLib](https://github.com/OpenSolver/OpenSolverNomadLib) are in the same root directory, and is configured to search for the header files under this assumption. If you change the relative location of the repositories the project will need to be updated to reflect this.

## Building the DLL

The 32 and 64 bit DLLs can be built by selecting either `Win32` or `x64` build configurations as appropriate. The resulting DLL is renamed to `OpenSolverDll64.dll` (only if 64 bit) and copied into the [OpenSolver](https://github.com/OpenSolver/OpenSolver) repository (which should reside in the same directory as this repository).

## Properties

If you are having problems building your project right click on your project in the solution manager and go to properties. 
Under `C/C++`->`Additional Include Directories` you should have the directories of the NOMAD source code as well as the `INCLUDE files` for `Excel2010XLLSDK` folder.

Under `Linker`->`General`->`Additional Library Directories` should be the directories of the correct versions (32 or 64 bit) Nomad Library `libnomad.lib` (`libnomad-debug` for debug builds) as well as the correct version (32 or 64 bit) of `xlcall32.lib`.
These files should also be under `Linker`->`Input`->`Additional Dependencies`.

## OpenSolverNomadDll License

The `OpenSolverNomadDll` and `OpenSolverNomadDll64` files support the use of NOMAD in OpenSolver, and are licensed under the GNU GPL License for use by all OpenSolver users. 

Please see `NOMAD LICENSE.txt` and http://www.gerad.ca/nomad/Project/Home.html for details of the NOMAD license.

© Copyright by Andrew Mason and Matthew Milner, 2013
