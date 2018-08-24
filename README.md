# ExecutorOpenDss
This project is a C# OpenDSS (http://smartgrid.epri.com/SimulationTool.aspx) customization to batch calculate the power flow of brazilian eletric distribution feeders. The project main objective is calculate energy and power losses. The feeder *.dss files (see feeder example directory) are very similar to the ones created by the brazilian regulatory agency (http://aneel.gov.br/) program (GeoPerdas.exe).

There are 3 directories in this project:
1. ExecutorOpenDssBr: the C# project. 
It uses 2 dll files EEPlus.dll (some Excel stuff) and Auxiliares.dll (the co author Daniel Rocha routines) already included in the project.    

2. FeederExample: a 13.8kV feeder example. 

3. FME: A FME (check https://www.safe.com) project able to generate *.dss files, when connected to GeoPerdas SQLServer database (it mighty be useful for people who work for the any brazilian eletric utilies).
In this directory there are another 3 files used by the FME to generate the fedeers.dss files:
Alim_gdis_OpenDSS.txt -> just a txt.file with the feeders names to be generated.
recursos\curvasDeCargaAccess.xlsx -> external data from curves to consuption/demand transformation.
recursos\geracaoDistribuidaMT_Diaria_01_2016.xlsx -> external data from generation units.

