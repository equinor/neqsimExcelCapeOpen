# NeqSim - Excel and Cape Open
[NeqSim](https://equinor.github.io/neqsimhome/) is a tool for estimation of fluid behaviour for oil and gas production.

NeqSim Excel and Cape Open is the Visual Studio project using the NeqSim library for development of an Excel interface and a Cape-Open interface. The Cape Open interface makes it possible to use NeqSim in Cape Open compatible simulation tools such as HYSYS, PROII and Unisim.

## Installation
An installation file for NeqSim (Excel and Cape Open) is available under [releases](https://github.com/equinor/neqsimExcelCapeOpen/releases).

## Documentation
NeqSim Excel [documentation](https://github.com/equinor/neqsim.NET/wiki/Getting-started-with-NeqSim-in-Excel).

## Discussion forum

Questions related to neqsim can be posted in the [github discussion pages](https://github.com/equinor/neqsim/discussions).

## Folders

* NeqSimExcel - is the Excel interface for NeqSim. A published NeqSim Excel tool is available under releases. You will need to install Visual Studio Tools for Office (VSTO) to work with this project in Visual Studio.
* NeqSimNET - is where the NeqSim dll (created with IKVM) is used - and where calculations are done
* NeqSimCapeOpen - is where the implementation of the Cape Open thermo interface (both 1.0 and 1.1)
* CapeOpenUnitOperations - is the implementation of Cape Open unit operations using NeqSim
* NeqSimRegistration - is used for registration of NeqSim Cape Open on the local computer (so it is recogniced by the Cape Open tool)
* NeqSimX.X.X - is the installer for NeqSim.


## Authors and contact persons

Even Solbraa (esolbraa@gmail.com),  Marlene Louise Lund

## Licence

NeqSim is distributed under the [Apache-2.0](https://github.com/equinor/neqsim/blob/master/LICENSE) licence.

## Acknowledgments

A number of master and PhD students at NTNU have contributed to development of NeqSim. We greatly acknowledge their contributions.
