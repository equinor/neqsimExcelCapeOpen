# neqsim.NET
[NeqSim (Non-Equilibrium Simulator)](https://equinor.github.io/neqsimhome/) is a tool for estimation of fluid behaviour for oil and gas production.

neqsim.NET is the Visual Studio project using the NeqSim library for development of a Excel interface and a Cape-Open interface. The Cape Open interface makes it possible to use NeqSim in Cape Open compatible simulation tools such as HYSYS, PROII and Unisim.

Folders:
NeqSimExcel - is the Excel interface for NeqSim. A published NeqSim Excel tool is available under releases. You will need to install Visual Studio Tools for Office (VSTO) to work with this project in Visual Studio.
NeqSimNET - is where the NeqSim dll (created with IKVM) is used - and where calculations are done
NeqSimCapeOpen - is where the implementation of the Cape Open thermo interface (both 1.0 and 1.1)
CapeOpenUnitOperations - is the implementation of Cape Open unit operations using NeqSim
NeqSimRegistration - is used for registration of NeqSim Cape Open on the local computer (so it is recogniced by the Cape Open tool)
NeqSimX.X.X - is the installer for NeqSim.
