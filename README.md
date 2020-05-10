# ExecutorOpenDss
This project is a C# [OpenDSS](http://smartgrid.epri.com/SimulationTool.aspx) customization to batch calculate the power flow of brazilian eletricity distribution feeders. The project main objective is calculate energy losses. The feeder \*.dss files (see FeederExample directory) are very similar to the ones created by the GeoPerdas.exe from the brazilian energy regulatory agency [ANEEL](http://aneel.gov.br/), but in this case, the files are created using the project [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS) from the database (**GeoPerdas**). 

There are 2 directories in this project:
1. **ExecutorOpenDssBr**: the C# project. As it is a complete Visual Studio project, there won't be any dificult to compile it.
You'll have to only set up the local path of 2 resources directories:
1."*Caminho dos arquivos dos alimentadores \*.dss*", thats is the FeederExample path.
2."*Caminho dos recursos permanentes*", thats is the "global" resources files.

The project also uses 2 dll files: EEPlus.dll (that allows Excel files to be read in the C#) and Auxiliares.dll (the co author Daniel Rocha routines) already included in the project.    

2. **FeederExample**: This directory contains a 13.8 kV feeder *.dss files, created by the DB export [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS). The files must be in a sub-directory with the same name from the feeder. 
There are also a resource subdirectory with some .xlsx and .txt files with external information to the "ExecutorOpenDssBr" run the *.dss files, such as the "linecode" and the "load profiles" files. As I use those files for all feeders, I've included them in a separated sub-directory.

Some code optimization can be done, so any help is welcome. The code comments are in portuguese, so I apologize to the non-portuguese speaking people.

Ezequiel C. Pereira
