# ExecutorOpenDss
This project is a C# OpenDSS customization to batch calculate the power flow of brazilian eletric distribution feeders. The project main objective is calculate energy and power losses. The feeder *.dss files (see feeder example directory) are very similar to the ones created by GeoPerdas.exe (the brazilian regulatory agency program).

There are 3 directories in this project:
1. ExecutorOpenDssBr: the C# project. 
2. FeederExample: a 13.8kV feeder example. 
3. FME: A FME (check https://www.safe.com) project able to generate *.dss files, when connected to GeoPerdas SQLServer database (it mighty be useful for people who work for the eletric utilies).
