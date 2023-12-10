# ExecutorOpenDssBr
This project is a C# [OpenDSS](http://smartgrid.epri.com/SimulationTool.aspx) customization to batch calculate the power flow of Brazilian electricity distribution feeders. The feeders \*.dss files - see [FeederExample](https://github.com/Zecao/ExecutorOpenDssBr/tree/master/FeederExample) directory - are very similar to the ones created by the GeoPerdas.EXE software from the Brazilian energy regulatory agency [ANEEL](http://aneel.gov.br/). But in this case, the .dss files were created using the project [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS), a simple **GeoPerdas** database reader also created by me. 

Objectives of this project:
1. Alternative to using the official software of ANEEL, **GeoPerdas.EXE**, with the following benefits:
- faster execution time compared with **GeoPerdas.EXE**. 
- easier to detect and correct errors in the .dss files. BTW, I always try to converge 100% of feeders before running the official GeoPerdas.EXE.

2. Simulation of parameters and their effects on all feeders. Some examples:
- inclusion of capacitor banks.
- modeling of reclosers and fuses (SAIDI, SAIFI calculation).
- analysis of changing the power factor of loads. For example, simulation with metered power factor.
- zero sequence impedances on the lines.
- analysis of other load models.

3. Report generation.
Currently, the following reports were developed (yet in a crude model, i.e. a .txt file is generated and later treated in a spreadsheet).
- voltage level in the primary buses of distribution transformers.
- disconnected loads.
- loads with critical voltage levels.
- tap of voltage regulator banks.

**Usage**: 
1. The directory **FeederExample** contains 2 subdirectories: 1. **ABDD201** with a 13.8 kV feeder *.dss files created by the database exporter [ExportadorGeoPerdasDSS](https://github.com/Zecao/ExportadorGeoPerdasDSS) and; 2. the subdirectory **Recursos** which contains some resources files such as the *linecode* and the *load profiles* files (I've included these files in a separated subdirectory, as they are usually the same for all feeders).

2. You must configure the GUI TextBox "*Caminho dos recursos permanentes*" with the "*Recursos*" subdirectory and the GUI TextBox "*Caminho dos arquivos dos alimentadores \*.dss*" with the root subdirectory.

3. The list of feeders to be executed must be in the file *lstAlimentadores.m*.

**External projects**
In addition to **OpenDSSengine.dll** the project uses the following dlls from other projects (all of them included):
- OpenDSS Extensions from Unicamp (dss_capi, dss_sharp, libklusolvex, libwinpthread dlls)
- **EEPlus.dll**: which allows Excel files to be read in the C#;
- Dlls from [QuickGraph 3.6](https://archive.codeplex.com/?p=quickgraph). 

**Language** 
Most of the code is in Portuguese, but I'm making an effort (in every new release) to translate some code into English.

Ezequiel C. Pereira
