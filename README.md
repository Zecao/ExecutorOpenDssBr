# ExecutorOpenDssBr
This project is a C# [OpenDSS](http://smartgrid.epri.com/SimulationTool.aspx) customization to batch calculate the power flow of brazilian eletricity distribution feeders. The feeders \*.dss files - see [FeederExample](https://github.com/Zecao/ExecutorOpenDssBr/tree/master/FeederExample) directory - are very similar to the ones created by the GeoPerdas.EXE software from the brazilian energy regulatory agency [ANEEL](http://aneel.gov.br/), but in this case, the .dss files were created using the project [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS), a simple **GeoPerdas** database reader also created by me. 

Objectives of this project:
1. Alternative to using the official software of ANEEL **GeoPerdas.EXE**, with the following benefits:
- faster execution time compared with **GeoPerdas.EXE**. 
- easier to detect and correct errors in the .dss files. BTW, I always try to converge 100% of feeders before run the official GeoPerdas.EXE.

2. Simulation of parameters and their effects on all feeders. Some examples of simulations:
- analisys of capacitor banks.
- modelling of reclosers and fuses.
- analisys of changing the power factor of loads. For example, simulation with metered power factor.
- considering of zero sequence impedances on the lines.
- analisys of other load models.

3. Report generation.
Currently, I developed the following reports, yet in a crude model, i.e. a .txt file is generated and later treated in a spreadsheet.
- voltage level in the primary buses of distribution transformers.
- disconnected loads.
- loads with critical voltage levels.
- tap of voltage regulator banks.

**Usage**: 
There are 2 directories in this project:

1. **ExecutorOpenDssBr**: the Visual Studio C# project files. You'll have to only set up the local path of the 2 resources directories:
1."*Caminho dos arquivos dos alimentadores \*.dss*": the feeder \*.dss files path.
2."*Caminho dos recursos permanentes*": the "global" resources files.

The project also uses 2 dll files: **EEPlus.dll** which allows Excel files to be read in the C# and **Auxiliares.dll** (the co author Daniel Rocha routines) already included in the project.    

2. **FeederExample**: This directory contains a 13.8 kV feeder *.dss files, created by the data base exporter [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS). The files must be placed in a subdirectory with the same feeder name. 
There are also a subdirectory with some .xlsx and .txt files with external resources to run the *.dss files, such as the "linecode" and the "load profiles" files.  I've included thise files in a separated subdirectory, as they are the same for all feeders.

Most of the code are in portuguese, but I'm making efforts (in every release) to translate some code to english.

Ezequiel C. Pereira
