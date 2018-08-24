using ExecutorOpenDSS.Classes;
using ExecutorOpenDSS.Classes_Auxiliares;
using ExecutorOpenDSS.Classes_Principais;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ExecutorOpenDSS
{
    class ParamGeraisDSS
    {
        // nome alim atual
        private string _nomeAlimAtual;

        // mes abreviado em 3 letras
        private string _mesAbrv3letras;

        //
        public TipoDiasMes _objTipoDeDiasDoMes;

        //TODO private
        // arquivos resultados
        public string _nomeArqCurvasDeCarga = "CurvasDeCarga";
        public string _arquivoResPerdasHorario = "perdasAlimHorario.txt";
        public string _arquivoResPerdasDiario = "perdasAlimDiario.txt";
        public string _arquivoResPerdasMensal = "perdasAlimMensal.txt";
        public string _arquivoResPerdasAnual = "perdasAlimAnual.txt";
        public string _arquivoResAlimNaoConvergiram = "AlimentadoresNaoConvergentes.txt";
        private string _arquivoDRPDRC = "AlimentadoresDRPDRC.txt";
        private string _arqBarraTrafo = "BarrasTrafo.txt";

        //caminhos 
        private string _pathRecursosMensais;
        private string _pathCurvasTxt;

        // tipo fluxo Normal
        public string _tipoFluxo;

        // arquivos entrada
        private string _arqAjuste;

        // paramentros da interface GUI
        public ParametrosGUI _parGUI;
        public ParamAvancados _parAvan;

        // map demanda max X alim
        public Dictionary<string, double> _mapDadosDemanda;

        // map energia injetada X alim

        public EnergiaMes _reqEnergiaMes = new EnergiaMes();
        public LoadMultMes _reqLoadMultMes = new LoadMultMes();

        // map de tensoes do barramento
        public Dictionary<string, double> _mapTensaoBarramento;

        // TODO mover para classe parCurvaCarga
        public string _pathCurvasXLS;
        public Dictionary<string, string> _listaNomeArquivosCurvasCarga;
        
        // grava _mapAlimLoadMult no arquivo excel
        public void GravaMapAlimLoadMultExcel()
        {
            //Novo arquivo de ajuste
            _arqAjuste = "Ajuste_" + _mesAbrv3letras + ".xlsx";

            //Re-escreve xls de ajuste
            ArqManip.SafeDelete(_pathRecursosMensais + _arqAjuste);

            var file = new FileInfo(_pathRecursosMensais + _arqAjuste);

            //
            int mes = _parGUI._mesNum;

            // 
            Dictionary<string, double> mapAlimLoadMult = _reqLoadMultMes._mapAlimLoadMult[mes];

            // Grava arquivo Excel
            ArqManip.GravaDictionaryExcel(file, mapAlimLoadMult);

        }

        // set Mes 
        internal void setMes(int mes)
        {
            _parGUI._mesNum = mes;

            //atualiza abreviatua mes
            _mesAbrv3letras = TipoDiasMes.getMesAbrv(mes);
        }

        internal string getNomeCargaBT_mes()
        {
            return _nomeAlimAtual + "CargaBT_" + _mesAbrv3letras + ".dss";
        }

        internal string getNomeCargaMT_mes()
        {
            return _nomeAlimAtual + "CargaMT_" + _mesAbrv3letras + ".dss";
        }   

        //Seta nome do alim atual
        public void setNomeAlimAtual(string nome)
        {
            _nomeAlimAtual = nome;
        }

        //Get _nomeAlimAtual 
        public string getNomeAlimAtual()
        {
            return _nomeAlimAtual;
        }

        // obtem TensaoSaidaSE 
        public string GetTensaoSaidaSE()
        {
            // se modo usar tensoes barramento  
            if (_parGUI._usarTensoesBarramento)
            {
                // 
                if (_mapTensaoBarramento.ContainsKey(_nomeAlimAtual))
                {
                    // Obtem tensao de barramento do map
                    double tensaoPUstr = _mapTensaoBarramento[_nomeAlimAtual];

                    // Preenche variavel da interface
                    _parGUI._tensaoSaidaBarUsuario = tensaoPUstr.ToString();
                }
            }
            return _parGUI._tensaoSaidaBarUsuario;
        }

        // Seta tensao saidaSE
        public void SetTensaoSaidaSE(string nome)
        {
            _parGUI._tensaoSaidaBarUsuario = nome;
        }

        // get nome do arquivo gResPerdasDiario 
        public string getNomeComp_arquivoResPerdasDiario()
        {
            return _parGUI._pathRaizGUI + _arquivoResPerdasDiario;
        }

        // get nome do arquivo DRPDRC
        public string getNomeComp_arquivoDRPDRC()
        {
            return _parGUI._pathRaizGUI + _arquivoDRPDRC;
        }

        // get nome do arquivo BarraTraf
        public string getNomeArqBarraTrafo()
        {
            return _parGUI._pathRaizGUI + _arqBarraTrafo;
        }

        // get dataPath OpenDSS
        public string getDataPathAlimOpenDSS()
        {
            return _parGUI._pathRaizGUI + "ArquivosDssMTBT\\" + _nomeAlimAtual;
        }

        // construtor
        public ParamGeraisDSS(MainWindow janelaPrincipal)
        {
            // Define os parâmetros gerais
            janelaPrincipal.Disp("Inicializando parâmetros gerais...");

            //Verifica se foi solicitado o cancelamento.
            if (janelaPrincipal._cancelarExecucao)
            {
                janelaPrincipal.FinalizaProcesso();
                return;
            }

            //Paramentros Otimizacao.
            _parGUI = janelaPrincipal._parGUI;

            //Parametros Avancados
            _parAvan = janelaPrincipal._parAvan;

            //Pega âs três primeiras letras do mês
            _mesAbrv3letras = _parGUI._mes.Substring(0, 3);

            //Caminho dos recursos
            _pathRecursosMensais = _parGUI._pathRaizGUI + "Recursos\\";

            //Caminho curvas TXT 
            //_pathCurvasTxt = _parGUI._pathRecursosPerm + "CurvasTxt\\";
            _pathCurvasTxt = _parGUI._pathRecursosPerm + "NovasCurvasTxt\\";

            // TODO mover para classe especifica
            //Caminho curvas XLS 
            _pathCurvasXLS = _parGUI._pathRecursosPerm + "CurvasXls\\";

            //Tradução da função defCaracteristicasGeraisFP
            preencheTipoDia(janelaPrincipal._parGUI._tipoDia);

            // tipo fluxo
            _tipoFluxo = "Normal";

            //FIX ME 
            //TODO nao altera com a respectiva mudanca na interface. 
            //Nome do arquivo de ajuste
            _arqAjuste = "Ajuste_" + _mesAbrv3letras + ".xlsx";

            // Carrega map com valores de demanda maxima do alimentador
            carregaMapDemandaMaxAlim();

            // CArrega map com valores de energia mes do alimentador
            carregaMapEnergiaMesAlim();

            // Carrega map com os valores dos loadMult por alimentador
            carregaMapAjusteLoadMult();

            // Carrega map com os valores de tensao do barramento
            carregaMapTensaoBarramento();

            // preenche variavel _objTipoDeDiasDoMes
            _objTipoDeDiasDoMes = new TipoDiasMes(_parGUI, janelaPrincipal);
        }

        // traducao do tipo do dia da combobox para o codigo string necessario
        private void preencheTipoDia(string tipoDia)
        {
            switch (tipoDia)
            {
                case "Dia Útil":
                    _parGUI._tipoDia = "DU";
                    break;
                case "Domingo":
                    _parGUI._tipoDia = "DO";
                    break;
                case "Sábado":
                    _parGUI._tipoDia = "SA";
                    break;
            }
        }

        // obtem o nome do arquivo de perdas
        public string GetNomeArquivoPerdas()
        {
            string tipoFluxo = _parGUI._modoFluxo;

            string arquivo = null;

            switch (tipoFluxo)
            {
                case "Snap":

                    arquivo = _parGUI._pathRaizGUI + _arquivoResPerdasHorario;

                    break;

                case "Daily":

                    arquivo = _parGUI._pathRaizGUI + _arquivoResPerdasDiario;

                    break;

                default:
                    arquivo = _parGUI._pathRaizGUI + _arquivoResPerdasDiario;

                    break;
            }
            return arquivo;
        }

        // carrega map com informacoes do arquivo de tensao no barramento
        private void carregaMapTensaoBarramento()
        {
            // Nome arquivo tensoes barramento
            string nomeArquivo = _parGUI._pathRecursosPerm + "tensoesBarramento.xlsx";

            // Se arquivo existe
            if (File.Exists(nomeArquivo))
            {
                // 
                _mapTensaoBarramento = LeXLSX.XLSX2Dictionary(nomeArquivo);
            }
            else
            {
                throw new FileNotFoundException("Arquivo " + _pathRecursosMensais + nomeArquivo + " não encontrado.");
            }

        }

        //Tradução da função carregaMapDemandaMaxAlim
        private void carregaMapDemandaMaxAlim()
        {
            //TODO FIX ME 
            string nomeArquivo = "demandaMaxAlim_" + "Jan" + ".xlsx";

            string nomeArquivoCompleto = _pathRecursosMensais + nomeArquivo;

            _mapDadosDemanda = LeXLSX.XLSX2Dictionary(nomeArquivoCompleto);
        }

        //Tradução da função carregaMapDemandaMaxAlim
        private void carregaMapEnergiaMesAlim()
        {
            //
            string nomeArquivoEnergia = "energiaMesAlim.xlsx";
            string nomeArqEnergiaCompl = _pathRecursosMensais + nomeArquivoEnergia;

            // carrega requisitos para todos os meses
            for (int mes = 1; mes < 13; mes++)
            {
                int coluna = mes + 1;

                Dictionary<string, double> mapEnergiaMes2 = LeXLSX.XLSX2Dictionary(nomeArqEnergiaCompl, coluna);

                //adiciona na variavel da classe
                _reqEnergiaMes.Add(mes, mapEnergiaMes2);
            }
        }

        //Tradução da função leDadosAjustePerdas
        private void carregaMapAjusteLoadMult()
        {
            // carrega requisitos para todos os meses
            for (int mes = 1; mes < 13; mes++)
            {
                //obtem nome ajuste compl
                string arqAjusteCompl = getNomeArqAjuste(mes);

                //DEBUG
                //Dictionary<string, double> mapEnergiaMes2 = LeXLSX.XLSX2Dictionary(arqAjusteCompl);

                //adiciona na variavel da classe
                _reqLoadMultMes.Add(mes, LeXLSX.XLSX2Dictionary(arqAjusteCompl));
            }

        }

        private string getNomeArqAjuste(int mes)
        {
            return _pathRecursosMensais + "Ajuste_" + TipoDiasMes.getMesAbrv(mes) + ".xlsx";
        }

        //Tradução da função criaDiretorioDSS
        private string GetDirAlimentadorDSS(string nomeAlm)
        {
            string nomeDirAlm = _parGUI._pathRaizGUI + "ArquivosDssMTBT\\" + nomeAlm + "\\";
            return nomeDirAlm;
        }

        //
        public string getNomeArquivoAlimentadorDSS()
        {
            string nomeArqDSS = GetDirAlimentadorDSS(_nomeAlimAtual) + _nomeAlimAtual + ".dss";
            return nomeArqDSS;
        }

        // Deleta Arquivos Resultados
        internal void DeletaArqResultados()
        {
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + _arquivoResAlimNaoConvergiram);
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + _arquivoResPerdasHorario);
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + _arquivoResPerdasDiario);
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + _arquivoResPerdasMensal);
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + _arquivoResPerdasAnual);

            ArqManip.SafeDelete(getNomeArqBarraTrafo());

            //ArqManip.SafeDelete(AnaliseCargasIsoladas.getNomeArqBarraTrafo(this));

            ArqManip.SafeDelete(_parGUI._pathRaizGUI + "\\Rmatrix.txt");
            ArqManip.SafeDelete(_parGUI._pathRaizGUI + "\\Xmatrix.txt");        
        }

        // get Nome e Path CurvasTxtCompleto
        internal string getPathCurvasTxtCompleto()
        {
            return _pathCurvasTxt;
        }

        internal string getNomeArqPerdasAno()
        {
            return _parGUI._pathRaizGUI + _arquivoResPerdasAnual;
        }

        internal string getArqRmatrix()
        {
            return _parGUI._pathRaizGUI + "\\Rmatrix.txt";
        }

        internal string getArqXmatrix()
        {
            return _parGUI._pathRaizGUI + "\\Xmatrix.txt";
        }

        internal string getNomeGeradorMT_mes()
        {
            return _nomeAlimAtual + "GeradorMT_" + _mesAbrv3letras + ".dss";
        }

        // get Nome e Path CurvasTxtCompleto
        internal string getNomeEPathCurvasTxtCompleto()
        {
            return _pathCurvasTxt + _nomeArqCurvasDeCarga + _parGUI._tipoDia + ".dss";
        }

    }
}
