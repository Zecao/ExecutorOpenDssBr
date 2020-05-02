# ExecutorOpenDss
This project is a C# [OpenDSS](http://smartgrid.epri.com/SimulationTool.aspx) customization to batch calculate the power flow of brazilian eletric distribution feeders. The project main objective is calculate energy and power losses. The feeder \*.dss files (see FeederExample directory) are very similar to the ones created by the brazilian regulatory agency [ANEEL](http://aneel.gov.br/) program (GeoPerdas.exe). In fact, those files are created from the same database (**GeoPerdas**), using the project [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS)

There are 2 directories in this project:
1. **ExecutorOpenDssBr**: the C# project. As it is a complete Visual Studio project, there wont be any dificult to compile it.
You'll have to only set up the local path of 2 resources directories:
1."*Caminho dos arquivos dos alimentadores \*.dss*", thats is the FeederExample path.
2."*Caminho dos recursos permanentes*", thats is the "global" resources files.

The project also uses 2 dll files: EEPlus.dll (that allows Excel files to be read in the C#) and Auxiliares.dll (the co author Daniel Rocha routines) already included in the project.    

2. **FeederExample**: a 13.8kV feeder *.dss files. This is the directory where one the fedeer .DSS files (created by the DB export [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS) must be, one sub-directory per feeder. 
There are also a "global" subdirectory named "recursos" with some .xlsx and .txt files with external information necessary to the "ExecutorOpenDssBr" run *.dss files, as the "linecode" and the "load profiles" files. As I use those files for all feeders, I've included them in a separated sub-directory.

Some code optimization can be done in this project, so any help is welcome. The code comments are in portuguese (sorry for non-portuguese natives).

Ezequiel C. Pereira
