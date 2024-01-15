using ExecutorOpenDSS.AuxClasses;
using ExecutorOpenDSS.MainClasses;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExecutorOpenDSS
{
    public class GeneralParameters
    {
        // nome alim atual
        private string _nomeAlimAtual;

        //
        public TipoDiasMes _objTipoDeDiasDoMes;

        // txt files for results and reports
        private readonly string _nomeArqCurvasDeCarga = "CurvasDeCarga";
        private readonly string _arquivoResPerdasHorario = "perdasAlimHorario.txt";
        private readonly string _arquivoResPerdasDiario = "perdasAlimDiario.txt";
        private readonly string _arquivoResPerdasMensal = "perdasAlimMensal.txt";
        private readonly string _arquivoResPerdasAnual = "perdasAlimAnual.txt";
        private readonly string _arquivoResAlimNaoConvergiram = "AlimentadoresNaoConvergentes.txt";
        private readonly string _arquivoDRPDRC = "AlimentadoresDRPDRC.txt";
        private readonly string _arqBarrasDRPDRC = "BarrasDRPDRC.txt";
        private readonly string _arqBarraTrafo = "BarrasTrafo.txt";
        private readonly string _arqTapsRTs = "TapsRTs.txt";
        private readonly string _arqDemandaMaxAlim = "demandaMaxAlim_" + "Jan" + ".xlsx";
        //private readonly string _arqTensoesBarramento = "tensoesBarramento.xlsx"; // TODO
        private readonly string _arqLoops = "loops.txt";
        private readonly string _arqEnergia = "energiaMesAlim.xlsx";
        private readonly string _arqCargaIsoladas = "CargasIsoladas.txt";
        private readonly string _arqCapLossesRed = "CapacitorLossesReduction.txt";
        private readonly string _arqResumoAlim = "ResumoAlim.txt";

        private readonly string _pathCurvasTxt; //caminhos
        public readonly string _AlgoritmoFluxo = "Normal"; //tipo fluxo Normal

        // paramentros da interface GUI
        public GUIParameters _parGUI;

        //
        public MainWindow _mWindow;

        // gerencia dados de medicao 
        public FeederMetering _medAlim;

        // construtor
        public GeneralParameters(MainWindow janelaPrincipal)
        {
            _mWindow = janelaPrincipal;

            //Paramentros Otimizacao.
            _parGUI = janelaPrincipal._parGUI;

            //Caminho curvas TXT 
            _pathCurvasTxt = _parGUI._pathRecursosPerm + "NovasCurvasTxt\\";

            // preenche variavel _objTipoDeDiasDoMes
            _objTipoDeDiasDoMes = new TipoDiasMes(_parGUI, janelaPrincipal);

            // 
            _medAlim = new FeederMetering(this);
        }

        // 
        public string GetNomeArqEnergia()
        {
            return _parGUI._pathRecursosPerm + _arqEnergia;
        }

        // 
        public string GetNomeArqDemandaMaxAlim()
        {
            return _parGUI._pathRecursosPerm + _arqDemandaMaxAlim;
        }

        public string GetNomeArqBarraTrafoLocal()
        {
            return _parGUI._pathRecursosPerm + _arqCargaIsoladas;
        }

        public string GetNomeCompArqLoops()
        {
            return _parGUI._pathRecursosPerm + _arqLoops;
        }

        public string GetNomeCompArqCapacitorLossesRed()
        {
            return _parGUI._pathRecursosPerm + _arqCapLossesRed;
        }
        public string GetNomeCompArqResumoAlim()
        {
            return _parGUI._pathRecursosPerm + _arqResumoAlim;
        }

        // grava _mapAlimLoadMult no arquivo excel
        public void GravaMapAlimLoadMultExcel()
        {
            //Novo arquivo de ajuste
            string arqAjuste = "Ajuste_" + _parGUI._mesAbrv3letras + ".xlsx";

            //Re-escreve xls de ajuste
            TxtFile.SafeDelete(_parGUI._pathRecursosPerm + arqAjuste);

            var file = new FileInfo(_parGUI._pathRecursosPerm + arqAjuste);

            //
            int mes = _parGUI.GetMes();

            // 
            Dictionary<string, double> mapAlimLoadMult = _medAlim._reqLoadMultMes._mapAlimLoadMult[mes];

            // Grava arquivo Excel
            XLSXFile.GravaDictionaryExcel(file, mapAlimLoadMult);
        }

        // obtem o nome do arquivo de perdas
        public string GetNomeArquivoPerdas()
        {
            string tipoFluxo = _parGUI._tipoFluxo;

            string arquivo;

            switch (tipoFluxo)
            {
                case "Snap":

                    arquivo = _parGUI._pathRecursosPerm + _arquivoResPerdasHorario;

                    break;

                case "Daily":

                    arquivo = _parGUI._pathRecursosPerm + _arquivoResPerdasDiario;

                    break;

                default:
                    arquivo = _parGUI._pathRecursosPerm + _arquivoResPerdasDiario;

                    break;
            }
            return arquivo;
        }

        // Deleta Arquivos Resultados
        public void DeletaArqResultados()
        {
            // modo otimiza nao deleta os arquivos de resutlados 
            if (!_parGUI._otmPorEnergia)
            {
                TxtFile.SafeDelete(GetNomeComp_arquivoResAlimNaoConvergiram());
                TxtFile.SafeDelete(GetNomeComp_arquivoResPerdasHorario());
                TxtFile.SafeDelete(GetNomeComp_arquivoResPerdasDiario());
                TxtFile.SafeDelete(GetNomeComp_arquivoResPerdasMensal());
                TxtFile.SafeDelete(GetNomeComp_arquivoResPerdasAnual());

                TxtFile.SafeDelete(GetNomeComp_arqBarrasDRPDRC());
                TxtFile.SafeDelete(GetNomeComp_arquivoDRPDRC());
                TxtFile.SafeDelete(GetNomeArqBarraTrafo());
                TxtFile.SafeDelete(GetNomeCompArqLoops());
                TxtFile.SafeDelete(GetNomeArqTapsRTs());

                TxtFile.SafeDelete(GetArqRmatrix());
                TxtFile.SafeDelete(GetArqXmatrix());
            }
        }

        public string GetNomeComp_arquivoResPerdasAnual()
        {
            return _parGUI._pathRecursosPerm + _arquivoResPerdasAnual;
        }

        public string GetNomeComp_arquivoResPerdasMensal()
        {
            return _parGUI._pathRecursosPerm + _arquivoResPerdasMensal;
        }

        private string GetNomeComp_arquivoResPerdasDiario()
        {
            return _parGUI._pathRecursosPerm + _arquivoResPerdasDiario;
        }

        private string GetNomeComp_arquivoResPerdasHorario()
        {
            return _parGUI._pathRecursosPerm + _arquivoResPerdasHorario;
        }

        public string GetNomeComp_arquivoResAlimNaoConvergiram()
        {
            return _parGUI._pathRecursosPerm + _arquivoResAlimNaoConvergiram;
        }

        public string GetNomeArqTapsRTs()
        {
            return _parGUI._pathRecursosPerm + _arqTapsRTs;
        }

        // nome arquivo Rmatrix
        public string GetArqRmatrix()
        {
            return _parGUI._pathRecursosPerm + "\\Rmatrix.txt";
        }

        // nome arquivo Xmatrix
        public string GetArqXmatrix()
        {
            return _parGUI._pathRecursosPerm + "\\Xmatrix.txt";
        }

        // nome arquivo condutor
        public string GetNomeArqCondutor()
        {
            return _parGUI._pathRecursosPerm + "Condutores.dss";
        }

        //
        public string GetNomeCargaBT_mes()
        {
            return _nomeAlimAtual + "CargaBT_" + _parGUI._mesAbrv3letras + ".dss";
        }

        public string GetNomeCargaMT_mes()
        {
            return _nomeAlimAtual + "CargaMT_" + _parGUI._mesAbrv3letras + ".dss";
        }

        //Seta nome do alim atual
        public void SetNomeAlimAtual(string nome)
        {
            _nomeAlimAtual = nome;
        }

        //Get _nomeAlimAtual 
        public string GetNomeAlimAtual()
        {
            return _nomeAlimAtual;
        }

        // Seta tensao saidaSE
        public void SetTensaoSaidaSE(string nome)
        {
            _parGUI._tensaoSaidaBarUsuario = nome;
        }

        // get nome do arquivo DRPDRC
        public string GetNomeComp_arquivoDRPDRC()
        {
            return _parGUI._pathRecursosPerm + _arquivoDRPDRC;
        }

        // get nome do arquivo DRPDRC
        public string GetNomeComp_arqBarrasDRPDRC()
        {
            return _parGUI._pathRecursosPerm + _arqBarrasDRPDRC;
        }

        // get nome do arquivo BarraTraf
        public string GetNomeArqBarraTrafo()
        {
            return _parGUI._pathRecursosPerm + _arqBarraTrafo;
        }

        // get dataPath OpenDSS
        public string GetDataPathAlimOpenDSS()
        {
            return _parGUI._pathRaizGUI + _nomeAlimAtual;
        }

        // get nome arquivo ajuste de acordo com o mes
        public string GetNomeArqAjuste(int mes)
        {
            return _parGUI._pathRecursosPerm + "Ajuste_" + TipoDiasMes.GetMesAbrv(mes) + ".xlsx";
        }

        // Get fedeer *.dss string directory
        private string GetDirAlimentadorDSS(string nomeAlm)
        {
            return _parGUI._pathRaizGUI + nomeAlm + "\\";
        }

        // get nome arquivo alimentador DSS 
        public string GetNomeArquivoAlimentadorDSS()
        {
            return GetDirAlimentadorDSS(_nomeAlimAtual) + _nomeAlimAtual + ".dss";
        }

        // nome arquivo gerador MT
        public string GetNomeGeradorMT_mes()
        {
            return _nomeAlimAtual + "GeradorMT_" + _parGUI._mesAbrv3letras + ".dss";
        }

        // nome arquivo gerador BT
        public string GetNomeGeradorBT_mes()
        {
            return _nomeAlimAtual + "GeradorBT_" + _parGUI._mesAbrv3letras + ".dss";
        }

        // get Nome e Path CurvasTxtCompleto
        public string GetNomeEPathCurvasTxtCompleto(string tipoDia)
        {
            return _pathCurvasTxt + _nomeArqCurvasDeCarga + tipoDia + ".dss";
        }

        // get nome arquivo capacitor
        public string GetNomeCapacitorMT()
        {
            return _nomeAlimAtual + "CapacitorMT.dss";
        }

        // nome completo arquivo AnualD.dss
        public string GetNomeArquivoB()
        {
            return _nomeAlimAtual + "AnualB.dss";
        }

        // diretorio alimenta
        public string GetDiretorioAlim()
        {
            return _parGUI._pathRaizGUI + _nomeAlimAtual + "\\";
        }

        // get OpenDSS LoadMult parameter
        public double GetLoadMultFromXlsxFile()
        {
            //
            if (_parGUI._usarLoadMult)
            {
                // loadMult inicial 
                return ( _medAlim._reqLoadMultMes.GetLoadMult() );
            }

            return 1.0;
        }
    }
}
