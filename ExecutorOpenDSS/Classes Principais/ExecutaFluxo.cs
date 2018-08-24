using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenDSSengine;
using System.IO;
using System.Windows.Controls;
using System.Threading;
using System.Windows.Threading;
using OfficeOpenXml;
using ExecutorOpenDSS.Classes;
using ExecutorOpenDSS.Classes_Principais;
using System.Runtime.InteropServices;
using ExecutorOpenDSS.Classes_Auxiliares;

namespace ExecutorOpenDSS
{
    class ExecutaFluxo
    {
        public static MainWindow _janela;
        public ParamGeraisDSS _paramGerais;
        private DSS _DSSObj;
        public ResultadoFluxo _resultadoFluxo = new ResultadoFluxo();
                
        public ExecutaFluxo(MainWindow janela, ParamGeraisDSS par)
        {
            //Executa o Fluxo
            janela.Disp("Executando Fluxo...");

            //Verifica se foi solicitado o cancelamento.
            if (janela._cancelarExecucao)
            {
                janela.FinalizaProcesso();
                return;
            }

            _janela = janela;
            _paramGerais = par;

            //Inicializa o servidor COM
            _DSSObj = new DSS();

            // Inicializa servidor COM
            _DSSObj.Start(0);

            // configuracoes gerais OpenDSS
            _DSSObj.AllowForms = par._parGUI._allowForms;
        }

        // executa modo Snap
        public void executaSnap(List<string> lstAlimentadoresCemig)
        {
            //Roda Fluxo para cada alimentador
            foreach (string nomeAlim in lstAlimentadoresCemig)
            {
                //Execução com otimização demanda energia
                if (_paramGerais._parGUI._otmPorDemMax)
                {
                    Otimiza(nomeAlim);
                }
                //Execução padrão
                else
                {
                    Snap(nomeAlim);
                }
            }
        }

        // executa modo Mensal
        public void executaMensal(List<string> lstAlimentadoresCemig)
        {
            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            //Roda fluxo para cada alimentador
            foreach (string alim in lstAlimentadoresCemig)
            {
                //Execução com otimização
                if (_paramGerais._parGUI._otmPorEnergia)
                {
                    // Otimiza por energia
                    Otimiza(alim);

                    continue;
                }
                // Otimiza PU saida
                if (_paramGerais._parAvan._otimizaPUSaidaSE)
                {
                    // Calcula nivel otimo PU
                    OtimizaPUSaidaSE(alim);

                    continue;
                }
                //Execução padrão
                Mensal(alim);
            }

            //Se modo otimiza grava arquivo load mult
            if (_paramGerais._parGUI._otmPorEnergia)
            {
                _paramGerais.GravaMapAlimLoadMultExcel();
            }
        }

        //Tradução da função executaFluxoSnapOpenDSS
        private void Snap(string nomeAlim)
        {
            //Limpa arquivos
            _paramGerais.DeletaArqResultados();

            Text DSSText = _DSSObj.Text;
            
            //Verifica se foi solicitado o cancelamento.
            if (_janela._cancelarExecucao)
            {
                return;
            }

            // atribui nomeAlim
            _paramGerais.setNomeAlimAtual(nomeAlim);

            // SnapPvt
            SnapPvt(DSSText);

            // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
            GravaPerdasArquivo();

            //informa usuario convergencia
            _janela.Disp(getMsgConvergencia(null, nomeAlim));

            //Plota perdas na tela
            _janela.Disp(_resultadoFluxo.getResultadoFluxoToConsole(
                _paramGerais.GetTensaoSaidaSE(), _paramGerais.getNomeAlimAtual()));
            //DEBUG
            //double somaCargas = OperacoesDSS.getSomaCargasAlim(DSSObj);
        }

        // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido
        private void GravaPerdasArquivo()
        {
            //Se modo otmiza não grava arquivo
            if (_paramGerais._parAvan._otimizaPUSaidaSE || _paramGerais._parGUI._otmPorEnergia || _paramGerais._parGUI._otmPorDemMax)
            {
                return;
            }

            // Se alim Nao Convergiu 
            if (!_resultadoFluxo._convergiuBool)
            {
                //Grava a lista de alimentadores não convergentes em um txt
                ArqManip.gravaLstAlimNaoConvergiram(_paramGerais, _janela);
            }
            // Grava Perdas de acordo com o tipo de fluxo
            else
            {
                string nomeAlim = _paramGerais.getNomeAlimAtual();

                // obtem o nome do arquivo de perdas, conforme o tipo do fluxo 
                string arquivo = _paramGerais.GetNomeArquivoPerdas();

                // Grava Perdas
                ArqManip.GravaPerdas(_resultadoFluxo, nomeAlim, arquivo, _janela);
            }
        }

        // SnapPvt
        private void SnapPvt(Text DSSText)
        {
            // get nome Alim
            string alimTmp = _paramGerais.getNomeAlimAtual();

            // Entra no diretório ArquivosDss e no subdiretorio 'nomeAlm'
            string nomeArqDSSCompleto = _paramGerais.getNomeArquivoAlimentadorDSS();

            // Verifica existencia xml
            if (File.Exists(nomeArqDSSCompleto))
            {
                // Carrega arquivos DSS
                CarregaArquivoDSS(DSSText, alimTmp);

                // Se nao ocorreu erro 
                if (DSSText.Result.Equals("\0\0"))
                {
                    // Interfaces
                    Solution DSSSolution = _DSSObj.ActiveCircuit.Solution;

                    // realiza ajuste das cargas 
                    DSSSolution.LoadMult = GetLoadMult();

                    // seta algorithm Normal ou Newton
                    DSSText.Command = "Set Algorithm = " + _paramGerais._tipoFluxo;

                    // seta modo snap.
                    DSSText.Command = "Set mode=snap";

                    // resolve circuito 
                    DSSSolution.Solve();

                    // Obtem valores de pot e energia dos medidores 
                    GetValoresEnergyMeter(DSSText, null);
                }
                // ocorreu erro
                else
                {
                    //adiciona alim como nao convergido
                    _resultadoFluxo.DefineComoNaoConvergiu();

                    //exibe o erro na janela
                    _janela.Disp(DSSText.Result);
                }
            }
            else
            {
                //adiciona alim como nao convergido
                _resultadoFluxo.DefineComoNaoConvergiu();

                //Exibe mensagem de erro, caso não encontrem os arquivos de resultados
                _janela.Disp(alimTmp + ": Arquivos *.dss não encontrados");
            }
        }

        //Tradução da função otimizaLoadMult
        private void Otimiza(string alim)
        {
            // Verifica se foi solicitado o cancelamento.
            if (_janela._cancelarExecucao)
            {
                return;
            }

            // otimiza PVT 
            double loadMult = OtimizaPvt(alim);

            // Atualiza map loadMultAlim
            _paramGerais._reqLoadMultMes.AtualizaMapAlimLoadMult(alim, loadMult, _paramGerais._parGUI._mesNum);
        }

        // Otimiza Pvt
        private double OtimizaPvt(string alimentador)
        {            
            //set Nome Alim atual 
            _paramGerais.setNomeAlimAtual(alimentador);

            // loadMult inicial 
            double loadMultInicial = 0;

            int mes = _paramGerais._parGUI._mesNum;

            // loadMultInicial
            if (_paramGerais._reqLoadMultMes._mapAlimLoadMult[mes].ContainsKey(alimentador))
            {
                loadMultInicial = _paramGerais._reqLoadMultMes._mapAlimLoadMult[mes][alimentador];
            }
            else
            {
                _janela.Disp(alimentador + ": Alimentador não encontrado no arquivo de ajuste");
                return loadMultInicial;
            }

            // OBS: condicao de retorno: Skipa alimentadores com loadMult igual a zero se loadMult == 0 
            if (loadMultInicial == 0)
            {
                _janela.Disp(alimentador + ": Alimentador não otimizado! LoadMult=0");
                return loadMultInicial;  
            }

            double loadMult;

            // modo otimiza LoadMult energia 
            if (_paramGerais._parGUI._otmPorEnergia)
            {
                loadMult = OtimizaLoadMultEnergiaPvt(alimentador, loadMultInicial);
            }
            //modo otimiza loadMult potencia
            else
            {
                loadMult = OtimizaLoadMultPvt(loadMultInicial);
            }
            return loadMult;   
        }

        //Tradução da função calculaFluxoDiario
        public void executaDiario(List<string> lstAlimentadoresCemig)
        {
            //Limpa arquivos
            _paramGerais.DeletaArqResultados();

            //Roda o fluxo para cada alimentador
            foreach (string nomeAlim in lstAlimentadoresCemig)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_janela._cancelarExecucao)
                {
                    return;
                }

                // cria arquivo e preenche Cabecalho, caso modo _calcDRPDRC
                if (_paramGerais._parAvan._calcDRPDRC)
                {
                    // Cria arquivo cabecalho
                    //AnaliseIndiceTensao.CriaArqCabecalho(_paramGerais, _janela);
                }

                // atribui nomeAlim
                _paramGerais.setNomeAlimAtual(nomeAlim);

                //Executa fluxo diário openDSS
                ExecutaFluxoDiarioOpenDSS();
            }
        }

        //Tradução da função calculaFluxoMensal
        private void Mensal(string alim)
        {
            //Verifica se foi solicitado o cancelamento.
            if (_janela._cancelarExecucao)
            {
                return;
            }

            // atribui nomeAlim
            _paramGerais.setNomeAlimAtual(alim);

            //Chama fluxo MensalPvt  
            MensalPvt();
        }

        // Fluxo anual
        public void executaAnual(List<string> lstAlimentadoresCemig)
        {
            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            //Seta modo anual
            _paramGerais._parGUI._modoAnual = true;

            //Vetor _resultadoFluxo
            List<ResultadoFluxo> lstResultadoFluxo = new List<ResultadoFluxo>();

            //Roda fluxo para cada alimentador
            foreach (string alim in lstAlimentadoresCemig)
            {
                // para cada mes executa um fluxo mensal
                for (int mes = 1; mes < 13; mes++)
                {
                    //atribui mes
                    _paramGerais._parGUI._mesNum = mes;

                    //Verifica se foi solicitado o cancelamento.
                    if (_janela._cancelarExecucao)
                    {
                        return;
                    }

                    // atribui alim 
                    _paramGerais.setNomeAlimAtual(alim);
 
                    // set mes
                    _paramGerais.setMes(mes);

                    //TODO TESTAR Otimiza
                    if (_paramGerais._parGUI._otmPorEnergia)
                    {
                        Otimiza(alim);
                    }
                    else 
                    {
                        //Chama fluxo MensalPvt  
                        MensalPvt();
                    }

                    ResultadoFluxo resTmp = new ResultadoFluxo(_resultadoFluxo);

                    //Armazena Resultado
                    lstResultadoFluxo.Add(resTmp);
                }
                
                //Calcula resultado Ano
                _resultadoFluxo.CalculaResAno(lstResultadoFluxo, alim, _paramGerais.getNomeArqPerdasAno(), _janela);
            }
        }

        //Tradução da função calculaFluxoMensal
        private void OtimizaPUSaidaSE(string alim)
        {
            // 
            double puSaidaSE = 0.96;

            while (puSaidaSE < 1.05)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_janela._cancelarExecucao)
                {
                    return;
                }

                // atribui pu Saida SE
                _paramGerais.SetTensaoSaidaSE(puSaidaSE.ToString());

                // atribui nomeAlim
                _paramGerais.setNomeAlimAtual(alim);

                //Chama fluxo MensalPvt  
                MensalPvt();

                // Atualiza PU
                puSaidaSE += 0.02;
            }            
        }

        // Calcula fluxo mensal 
        private void MensalPvt()
        {
            //Dia útil
            _paramGerais._parGUI._tipoDia = "DU";

            //Executa fluxo diário openDSS
            ExecutaFluxoDiarioOpenDSS();

            //Copia resultado do fluxo em variavel local
            ResultadoFluxo resDU = new ResultadoFluxo(_resultadoFluxo);

            // Se alimentador não convergiu, não calcula SA e DO
            if (!resDU._convergiuBool)
            {
                return;
            }

            //Sábado
            _paramGerais._parGUI._tipoDia = "SA";

            //Executa fluxo diário openDSS
            ExecutaFluxoDiarioOpenDSS();

            //Copia resultado do fluxo em variavel local
            ResultadoFluxo resSA = new ResultadoFluxo(_resultadoFluxo);

            // Se alimentador não convergiu, não calcula DO
            if (!resSA._convergiuBool)
            {
                return;
            }

            //Domingo
            _paramGerais._parGUI._tipoDia = "DO";

            //Executa fluxo diário openDSS
            ExecutaFluxoDiarioOpenDSS();

            //Copia resultado do fluxo em variavel local
            ResultadoFluxo resDO = new ResultadoFluxo(_resultadoFluxo);

            // Se alimentador não convergiu, 
            if (!resDO._convergiuBool)
            {
                return;
            }

            // calcula resultados mensal 
            _resultadoFluxo.CalculaResultadoFluxoMensal(resDU, resSA, resDO, _paramGerais, _janela);

            //Plota perdas na tela
            _janela.Disp(_resultadoFluxo.getResultadoFluxoToConsole(
                _paramGerais.GetTensaoSaidaSE(), _paramGerais.getNomeAlimAtual()));

        }

        // Tradução da função otimizaLoadMultPvt
        private double OtimizaLoadMultPvt(double loadMult)
        {
            // alimTmp
            string alimTmp = _paramGerais.getNomeAlimAtual();

            //referência de geração (semana típica ou medição do MECE)
            double refGeracao;
            if (_paramGerais._mapDadosDemanda.ContainsKey(alimTmp))
            {
                refGeracao = _paramGerais._mapDadosDemanda[alimTmp];
            }
            else
            {
                _janela.Disp(alimTmp + ": mapDadosDemanda não encontrado");
                return double.NaN;
            }

            //calcula geração inicial
            double geracao = ExecutaFluxoSnapOtm(loadMult);

            // contador fluxo de potência
            // OBS: considera a execução acima
            int contFP = 1;

            // Se geração inicial for maior que referência OU fluxo não convergiu
            // retorna função
            if (double.IsNaN(geracao) || geracao > refGeracao)
            {
                return _paramGerais._parGUI._loadMultAlternativo;
            }

            // Inicializa loadMult anterior com LoadMult
            double loadMultAnt = loadMult;

            //TODO testar
            while ( Math.Abs(geracao - refGeracao) > _paramGerais._parGUI.getPrecisao() )
            {
                //Verifica se foi solicitado o cancelamento.
                if (_janela._cancelarExecucao)
                {
                    return loadMult;
                }

                // atualiza loadMult
                loadMultAnt = loadMult;
                loadMult = _paramGerais._parGUI.getIncremento() * loadMult;

                // calcula geração
                geracao = ExecutaFluxoSnapOtm(loadMult);

                // incrementa contador fluxo
                contFP = contFP + 1;

                // se geracao nao eh vazia
                if (double.IsNaN(geracao))
                {
                    break;
                }
                // condicao saida while: se executou 50 FP
                if ((contFP == 10) || (geracao > refGeracao))
                {
                    break;
                }
            }
            // retorna loadmult anterior
            return loadMultAnt;
        }

        // Tradução da função otimizaLoadMultEnergiaPvt
        private double OtimizaLoadMultEnergiaPvt(string alimentador, double loadMult)
        {
            //referência de EnergiaMes (medição do MECE)
            double refEnergiaMes;

            // obtem ref de energia
            refEnergiaMes = _paramGerais._reqEnergiaMes.getRefEnergia(alimentador, _paramGerais._parGUI._mesNum);

            // verifica ausencia de medicao de energia
            if (refEnergiaMes.Equals(0))
            {
                _janela.Disp(alimentador + ": refEnergiaMes igual a 0");

                return 0;
            }

            //calcula energia inicial
            ExecutaFluxoMensalOtm(loadMult);

            // condicao de saida da funcao
            // energia simulada (energiaMes) esta proxima da energia medida (refEnergiaMes)
            if (verificaEnergiaSimulada(refEnergiaMes))
            {
                _janela.Disp(alimentador + " ajustado com precisão " + _paramGerais._parGUI.getPrecisao().ToString() + "KWh.");

                return loadMult;
            }

            // Se fluxo não convergiu retorna loadMult alternativo
            if ( ! _resultadoFluxo._convergiuBool )
            {
                _janela.Disp(alimentador + ": erro valor energia!");

                return _paramGerais._parGUI._loadMultAlternativo;
            }

            // 
            double energiaMesTmp = _resultadoFluxo.getEnergia();

            // Se geração inicial for MAIOR que referência
            if (energiaMesTmp > refEnergiaMes)
            {
                // sentido 
                int sentidoBusca = -1;

                // retroceder o loadmult
                loadMult = ajustaLoadMult(energiaMesTmp, refEnergiaMes, loadMult, sentidoBusca);
            }
            // Se geração inicial for MENOR que referência
            else
            {
                // avanca o loadmult
                loadMult = ajustaLoadMult(energiaMesTmp, refEnergiaMes, loadMult);
            }

            return loadMult;
        }

        // condicao de saida da funcao
        // energia simulada (energiaMes) esta proxima da energia medida (refEnergiaMes)
        private bool verificaEnergiaSimulada(double refEnergiaMes)
        {
            // se aproximacao de energia estiver dentro do limite de precisao ja calculado, ignora nova otimizacao
            if (getDiferencaEnergia(refEnergiaMes) < _paramGerais._parGUI.getPrecisao())
            {
                return true;
            }
            return false;
        }

        //obtem diferenca de energia de acordo com o modo
        private double getDiferencaEnergia(double refEnergiaMes)
        {
            double delta;

            delta = Math.Abs((_resultadoFluxo.getEnergia() - refEnergiaMes));

            return delta;
        }

        // ajusta loadMult incrementalmente
        private double ajustaLoadMult(double energiaMes, double refEnergiaMes, double loadMult, int sentidoBusca = 1)
        {
            // contador fluxo de potência
            // OBS: considera a execução acima
            int contFP = 1;

            // variavel que armazena o loadMult anterior
            double loadMultAnt = loadMult;
            int sentidoBuscaAnt = sentidoBusca;

            // alimTmp
            string alimTmp = _paramGerais.getNomeAlimAtual();

            // enquanto erro da energia for maior que precisao
            while (getDiferencaEnergia(refEnergiaMes) > _paramGerais._parGUI.getPrecisao())
            {
                //Verifica se foi solicitado o cancelamento.
                if (_janela._cancelarExecucao)
                {
                    return loadMultAnt;
                }

                // detecta se mudou sendido de busca 
                if (sentidoBuscaAnt != sentidoBusca)
                {
                    _janela.Disp(alimTmp + " mudança sentido de busca. Diminua o Incremento/Precisão!");

                    return loadMultAnt;
                }

                // condicao saida while: se executou 10 FP
                if (contFP == 10)
                {
                    _janela.Disp(alimTmp + " não ajustado em " + contFP.ToString() + " execuções de fluxo. Repita o processo!");

                    return loadMultAnt;
                }

                // 
                // atualiza loadMult, de acordo com o sentido de ajuste
                loadMultAnt = loadMult;
                if (sentidoBusca == 1)
                {
                    loadMult = loadMult * _paramGerais._parGUI.getIncremento();
                }
                // retrocede ajuste 
                else
                {
                    loadMult = loadMult / _paramGerais._parGUI.getIncremento();
                }

                // calcula energia mes. Se modo fluxo mensal Aproximado
                ExecutaFluxoMensalOtm(loadMult);

                // se nao convergiu dentro do loop, retorna o loadMult anterior
                if (!_resultadoFluxo._convergiuBool)
                {
                    //OBS: mensagem redundante.
                    //_janela.Disp(alimTmp + " problema otimiza loadMult!");
                    return loadMultAnt;
                }

                // incrementa contador fluxo
                contFP = contFP + 1;

                //atualiza sentido de busca
                if ((_resultadoFluxo.getEnergia() > refEnergiaMes) && (sentidoBusca == 1) ||
                     (_resultadoFluxo.getEnergia() < refEnergiaMes) && (sentidoBusca == -1))
                {
                    // guarda sentido busca anterior
                    sentidoBuscaAnt = sentidoBusca;

                    // inverte sentido 
                    sentidoBusca = -sentidoBusca;
                }
            }
            //saida 
            _janela.Disp(alimTmp + " ajustado nos limites escolhidos.");

            return loadMult;
        }

        // ExecutaFluxoMensalOtm
        private void ExecutaFluxoMensalOtm(double loadMult)
        {
            // fluxo mensal aproximado
            if (_paramGerais._parGUI._aproximaFluxoMensalPorDU)
            {
                ExecutaFluxoMensalOtmSimplificado(loadMult);
            }
            else
            {
                ExecutaFluxoMensalOtmCompleto(loadMult);
            }
        }

        // Funcao proxy, construida para Re-uso da funcao ExecutaFluxoDiarioOpenDSS
        private void ExecutaFluxoMensalOtmCompleto(double loadMult)
        {
            // adiciona loadMult ao paramento gerais           
            _paramGerais._parGUI.loadMultAtual = loadMult;

            // Executa fluxo mensal
            MensalPvt();
        }

        // Funcao proxy, construida para Re-uso da funcao ExecutaFluxoDiarioOpenDSS
        private void ExecutaFluxoMensalOtmSimplificado(double loadMult)
        {
            // adiciona loadMult ao paramento gerais           
            _paramGerais._parGUI.loadMultAtual = loadMult;

            // Executa fluxo diario
            ExecutaFluxoDiarioOpenDSS();

            //get Energia diaria
            double energiaDiaria = _resultadoFluxo.getEnergia();

            // mes
            int mes = _paramGerais._parGUI._mesNum;

            // calcula energia mensal
            double energiaMes = energiaDiaria * _paramGerais._objTipoDeDiasDoMes.GetNumDiasMes(mes);

            //preenchee variavel da classe
            _resultadoFluxo.setEnergia(energiaMes);
        }

        //Tradução da função executaFluxoSnapOpenDSSLoadMult
        private double ExecutaFluxoSnapOtm(double loadMult)
        {
            //alimTmp
            string alimTmp = _paramGerais.getNomeAlimAtual();

            // altera o loadMult
            _paramGerais._parGUI.loadMultAtual = loadMult;

            Text DSSText = _DSSObj.Text;

            // Chama fluxo snap
            SnapPvt(DSSText);

            //informa usuario convergencia
            _janela.Disp(getMsgConvergencia(null, alimTmp));

            //Plota perdas na tela
            _janela.Disp(_resultadoFluxo.getResultadoFluxoToConsole(
                _paramGerais.GetTensaoSaidaSE(), _paramGerais.getNomeAlimAtual()));

            // retorno
            return _resultadoFluxo.getMaxKW();
        }

        //Tradução da função executaFluxoDiarioOpenDSS
        private void ExecutaFluxoDiarioOpenDSS()
        {
            // nome alim
            string alimTmp = _paramGerais.getNomeAlimAtual();

            //Nome arquivo DSS completo 
            string nomeArqDSS = _paramGerais.getNomeArquivoAlimentadorDSS();

            //Verifica existencia do arquivo DSS
            if (File.Exists(nomeArqDSS))
            {
                // Executa fluxo diario nativo
                if (_paramGerais._parGUI._fluxoDiarioNativo)
                {
                    // SE modo _calcDRPDRC ou _calcTensaoBarTrafo necessario passar a hora 
                    if (_paramGerais._parAvan._calcDRPDRC || _paramGerais._parAvan._calcTensaoBarTrafo)
                    {
                         ExecutaFluxoDiarioOpenDSSNativo(_paramGerais._parGUI._hora);
                    }
                    else
                    {
                        ExecutaFluxoDiarioOpenDSSNativo(null);
                    }

                }
                // Executa fluxo diário - 24 vezes fluxo horário
                else
                {
                    ExecutaFluxoDiarioOpenDSS24vezesHorario();
                }

                //informa usuario convergencia
                if ( _resultadoFluxo._convergiuBool )
                {
                    _janela.Disp(getMsgConvergencia(null, alimTmp));

                    /*
                    //TODO somente plotar info se for modo DIARIO
                    if (_paramGerais._parGUI.m)
                    {
                    //Plota perdas na tela
                    _janela.Disp(_resultadoFluxo.getResultadoFluxoMensalConsole(
                        _paramGerais.GetTensaoSaidaSE(), _paramGerais.getNomeAlimAtual()));
                    }*/
                }
            }
            else
            {
                //
                _janela.Disp(alimTmp + ": Arquivos *.dss não encontrados");
            }
        }

        //Tradução da função executaFluxoDiarioOpenDSSNativoOpenDSS
        private void ExecutaFluxoDiarioOpenDSSNativo(string hora)
        {
            //alimTmp
            string alimTmp = _paramGerais.getNomeAlimAtual();

            Text DSSText = _DSSObj.Text;
       
            //% carrega arquivo DSS
            CarregaArquivoDSS(DSSText, alimTmp);            

            //% Interfaces
            Circuit DSSCircuit = _DSSObj.ActiveCircuit;
            Solution DSSSolution = DSSCircuit.Solution;

            //get loadMult da struct paramGerais
            DSSSolution.LoadMult = GetLoadMult();

            // Seta modo diario 
            if (hora == null)
            {
                DSSText.Command = "Set mode=daily  hour=0 number=24 stepsize=1h";
            }
            else
            {
                DSSText.Command = "Set mode=daily hour=" + hora + " number=1 stepsize=1h";
            }

            // OLD CODE
            //DSSText.Command = "Set Trapezoidal = yes";

            //% OBS: EH IMPRESCINDIVEL LIMPAR OS MEDIDORES DE ENERGIA ANTES DO SOLVE
            DSSText.Command = "energymeter.carga.action=Clear";

            //% resolve circuito 
            DSSSolution.Solve();

            // grava valores do EnergyMeter 
            GetValoresEnergyMeter(DSSText, null);

            // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
            GravaPerdasArquivo();

            // calcula DRP e DRC
            if (_paramGerais._parAvan._calcDRPDRC)
            {
                //
                CalculaDRPDRC();
            }

            // calcula tensao PU no primario trafos
            if (_paramGerais._parAvan._calcTensaoBarTrafo)
            {
                // obtem indices de tensao nos trafos

            }

            // verifica cargas isoladas
            if (_paramGerais._parAvan._verifCargaIsolada)
            {
                //analise cargas isoladas

            }

            // TODO criar interface
            //queryLineCode(_DSSObj.ActiveCircuit);

            // TODO criar interface 
            //queryLine(_DSSObj.ActiveCircuit);
        }

        // TODO refactory 
        private void queryLineCode(Circuit dssCircuit)
        {
            List<string> lstCabos = new List<string>();
            lstCabos.Add("CAB102");
            lstCabos.Add("CAB103");
            lstCabos.Add("CAB104");
            lstCabos.Add("CAB107");
            lstCabos.Add("CAB108");
            lstCabos.Add("CAB202");
            lstCabos.Add("CAB203");
            lstCabos.Add("CAB204");
            lstCabos.Add("CAB207");
            lstCabos.Add("CAB208");
            lstCabos.Add("CABA06");
            lstCabos.Add("CABA08");
            lstCabos.Add("CAB2021");
            lstCabos.Add("CAB1031");
            lstCabos.Add("CAB1021");
            lstCabos.Add("CAB2031");
            lstCabos.Add("CABA061");
            lstCabos.Add("CABBT106");
            lstCabos.Add("CABBT107");
            lstCabos.Add("CABBT108");
            lstCabos.Add("CABBT803");
            lstCabos.Add("CABBT805");
            lstCabos.Add("CABBT809");
            lstCabos.Add("CABBT810");
            lstCabos.Add("CABBT801");
            lstCabos.Add("CABBT807");
            lstCabos.Add("CABBT808");

            Text textDSS = _DSSObj.Text;

            List<string> resRmatrix = new List<string>();
            List<string> resXmatrix = new List<string>();

            foreach (string lineCode in lstCabos)
            {
                textDSS.Command = "? LineCode." + lineCode + ".Rmatrix";

                resRmatrix.Add(lineCode + "\tRmatrix=" + textDSS.Result);

                textDSS.Command = "? LineCode." + lineCode + ".Xmatrix";

                resXmatrix.Add("\tXmatrix=" + textDSS.Result);
            }
            
            ArqManip.GravaListArquivoTXT(resRmatrix, _paramGerais.getArqRmatrix() , _janela);

            ArqManip.GravaListArquivoTXT(resXmatrix, _paramGerais.getArqXmatrix() , _janela);

        }

        // TODO refactory 
        private void queryLine(Circuit dssCircuit)
        {
            List<string> lstLinhas = new List<string>();
            lstLinhas.Add("CAB102");
            lstLinhas.Add("CAB103");
            lstLinhas.Add("CAB104");
            lstLinhas.Add("CAB107");
            lstLinhas.Add("CAB108");
            lstLinhas.Add("CAB202");
            lstLinhas.Add("CAB203");
            lstLinhas.Add("CAB204");
            lstLinhas.Add("CAB207");
            lstLinhas.Add("CAB208");
            lstLinhas.Add("CABA06");
            lstLinhas.Add("CABA08");
            lstLinhas.Add("CABBT106");
            lstLinhas.Add("CABBT107");
            lstLinhas.Add("CABBT108");
            lstLinhas.Add("CABBT803");
            lstLinhas.Add("CABBT805");
            lstLinhas.Add("CABBT809");
            lstLinhas.Add("CABBT810");
            lstLinhas.Add("CABBT801");
            lstLinhas.Add("CABBT807");
            lstLinhas.Add("CABBT808");

            Text textDSS = _DSSObj.Text;

            List<string> resRmatrix = new List<string>();
            List<string> resXmatrix = new List<string>();

            foreach (string line in lstLinhas)
            {
                textDSS.Command = "? Line." + line + ".Rmatrix";

                resRmatrix.Add(line + "\tRmatrix=" + textDSS.Result);

                textDSS.Command = "? Line." + line + ".Xmatrix";

                resXmatrix.Add(line + "\tXmatrix=" + textDSS.Result);
            }

            ArqManip.GravaListArquivoTXT(resRmatrix, _paramGerais.getArqRmatrix(), _janela);

            ArqManip.GravaListArquivoTXT(resXmatrix, _paramGerais.getArqXmatrix(), _janela);

        }

        //Tradução da função executaFluxoDiarioOpenDSS24vezesHorario
        private void ExecutaFluxoDiarioOpenDSS24vezesHorario()
        {
            // nome alim
            string nomeAlim = _paramGerais.getNomeAlimAtual();

            Text DSSText = _DSSObj.Text;

            //get loadMult da struct paramGerais
            double loadMult = GetLoadMult();

            // OLD CODE TESTAR
            List<MyEnergyMeter> matrizPerdasHoraria = new List<MyEnergyMeter>();

            //Saida
            ResultadoFluxo saida24pat = new ResultadoFluxo();

            //executa fluxo horario 24 vezes
            //OBS: o for deve ser de 0 a 23 uma vez que a "integracao" do openDSS adiciona o stepsize (de 1h) ao valor
            //da hora "i" inicial.
            for (int i = 0; i <= 23; i++)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_janela._cancelarExecucao)
                {
                    return;
                }
                //transforma numero em string
                string hora = i.ToString();

                //carrega arquivo DSS  
                CarregaArquivoDSS(DSSText, nomeAlim);
                
                //executa fluxo horario openDSS
                ExecutaFluxoDiarioOpenDSS24vezesHorarioPvt(hora, DSSText, loadMult);

                //Se alimentador não convergiu, interrompe o fluxo
                if (!_resultadoFluxo._convergiuBool)
                {
                    break;
                }

                //Preenche a matriz de perdas
                matrizPerdasHoraria.Add(_resultadoFluxo._energyMeter);
            }

            // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
            verificaSaidaEGravaPerdasDiariaArquivo(_resultadoFluxo, nomeAlim, matrizPerdasHoraria);

        }

        // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
        private ResultadoFluxo verificaSaidaEGravaPerdasDiariaArquivo(ResultadoFluxo res1patamar, string nomeAlim, List<MyEnergyMeter> matrizPerdasHoraria)
        {
            // Se alim Nao Convergiu 
            if (!res1patamar._convergiuBool)
            {
                //Grava a lista de alimentadores não convergentes em um txt
                ArqManip.gravaLstAlimNaoConvergiram(_paramGerais, _janela);

                //return
                return res1patamar;
            }
            else
            {
                //grava perdas hora a hora
                GravaPerdasHoraAHora(nomeAlim, matrizPerdasHoraria);

                //totaliza perdas horarias dando o resultado diario
                ResultadoFluxo pPotEEnergia = TotalizaPerdasEnergiaDiaria(nomeAlim, matrizPerdasHoraria);

                // contatena nome alim com tipo do dia 
                string nomeComp = nomeAlim + _paramGerais._parGUI._tipoDia;

                // grava perdas alimentador em arquivo 
                ArqManip.GravaPerdasAlimArquivo(nomeComp, pPotEEnergia, _paramGerais.getNomeComp_arquivoResPerdasDiario(), _janela);

                //Saida
                ResultadoFluxo saida24pat = new ResultadoFluxo(pPotEEnergia);

                return saida24pat;
            }

        }

        //Tradução da função executaFluxoHorarioOpenDSSPvt
        private void ExecutaFluxoDiarioOpenDSS24vezesHorarioPvt(string hora, Text DSSText, double loadMult)
        {
            // Interfaces
            Circuit DSSCircuit = _DSSObj.ActiveCircuit;
            Solution DSSSolution = DSSCircuit.Solution;

            // seta loadMult 
            DSSSolution.LoadMult = loadMult;

            // seta modo diario
            DSSText.Command = "Set mode=daily hour=" + hora + " stepsize=1h number=1";

            //DSSText.Command = "Set Trapezoidal = yes";

            //% OBS: EH IMPRESCINDIVEL LIMPAR OS MEDIDORES DE ENERGIA ANTES DO SOLVE
            DSSText.Command = "energymeter.carga.action=Clear";

            // resolve circuito 
            DSSSolution.Solve();

            // grava valores do EnergyMeter 
            GetValoresEnergyMeter(DSSText,hora);
        }

        // TODO testar
        //Tradução da função gravaPerdasHoraAHora
        private void GravaPerdasHoraAHora(string nomeAlim, List<MyEnergyMeter> matrizPerdasHorariaAlim)
        {
            //OLD CODE
            //int colunas = matrizPerdasHorariaAlim[0].Count;
            int colunas = 12;

            //Perdas e perdas de potência hora a hora
            List<List<double>> perdasEnergiaHoraAHora = new List<List<double>>();
            List<List<double>> geracaoEPerdaHoraAHora = new List<List<double>>();

            /* TODO FIX ME 
            foreach (GeracaoPerdasPotEPerdasEnergiaType p in matrizPerdasHorariaAlim)
            {
                //Perdas de potência hora a hora
                geracaoEPerdaHoraAHora.Add(p.GetRange(0, 2));

                //Perdas hora a hora
                perdasEnergiaHoraAHora.AddEnergia(p.GetRange(2, colunas - 2));                
            }*/

            //Para cada hora
            for (int i = 1; i <= 24; i++)
            {
                //contatena nome alim com tipo do dia 
                string nomeComp = nomeAlim + _paramGerais._parGUI._tipoDia + i;
                
                //TODO FIX ME 
                ResultadoFluxo perdasPotEEnergia = new ResultadoFluxo();

                //TODO FIX ME 
                //grava perdas alimentador em arquivo
                ArqManip.GravaPerdasAlimArquivo(nomeComp, perdasPotEEnergia, _paramGerais._parGUI._pathRaizGUI + _paramGerais._arquivoResPerdasHorario, _janela);
            }
        }

        //Tradução da função getLoadMult
        private double GetLoadMult()
        {
            // alimTmp
            string alimTmp = _paramGerais.getNomeAlimAtual();

            // modo de otimizacao de energia OU potencia 
            if (_paramGerais._parGUI._otmPorEnergia || _paramGerais._parGUI._otmPorDemMax)
            {
                return _paramGerais._parGUI.loadMultAtual;
            }
            // obtem loadMult do map de loadMults
            else
            {
                int mes = _paramGerais._parGUI._mesNum;

                // se loadMult existe no map
                if (_paramGerais._reqLoadMultMes._mapAlimLoadMult[mes].ContainsKey(alimTmp))
                {
                    return _paramGerais._reqLoadMultMes._mapAlimLoadMult[mes][alimTmp];
                }
                else
                {
                    _janela.Disp(alimTmp + ": Alimentador não encontrado no arquivo de ajuste");

                    // retorna loadMult Default
                    return _paramGerais._parGUI.loadMultDefault;
                }
            }
        }

        //Tradução da função carregaArquivoDSS
        private void CarregaArquivoDSS(Text DSSText, string nomeAlim)
        {
            // Seta o dataPath igual ao path do arquivo
            // OBS: corrige um erro que acontecia apos a inicializacao do servidor COM, 
            // onde o mesmo apresentava o DataPath de uma execucao anterior
            _DSSObj.DataPath = _paramGerais.getDataPathAlimOpenDSS();

            // nome arquivo dss
            string nomeArqDSS = nomeAlim + ".dss";

            // limpa circuito de objeto recem criado
            DSSText.Command = "clear";

            // Obtem linhas do arquivo cabecalho
            string[] lines = System.IO.File.ReadAllLines(nomeArqDSS);

            // TODO refactory. Da pau caso tenha so uma linha. 
            // Obtem 2 linha do arquivo 
            string circuitHeader = lines[1];

            //% Redirect arquivo .dss 
            DSSText.Command = circuitHeader + _paramGerais.GetTensaoSaidaSE();

            //Anexa tipo de curva de carga, de acordo com o dia
            string nomeArqCurvaCarga = _paramGerais.getNomeEPathCurvasTxtCompleto();

            // redirect arquivo Curva de Carga
            DSSText.Command = "Redirect " + nomeArqCurvaCarga;

            // redirect arquivo .dss  
            DSSText.Command = "Redirect " + nomeAlim + "AnualA.dss";

            // redirect arquivo CargaBT
            if ( File.Exists( _paramGerais.getNomeCargaBT_mes() ))
            {
                DSSText.Command = "Redirect " + _paramGerais.getNomeCargaBT_mes();
            }

            // redirect arquivo CargaMT
            if ( File.Exists( _paramGerais.getNomeCargaMT_mes() ))
            {
                DSSText.Command = "Redirect " + _paramGerais.getNomeCargaMT_mes();
            }

            // redirect arquivo GeradorMT
            if (File.Exists( _paramGerais.getNomeGeradorMT_mes() ))
            {
                DSSText.Command = "Redirect " + _paramGerais.getNomeGeradorMT_mes();
            }
             
            // redirect arquivo .dss  
            DSSText.Command = "Redirect " + nomeAlim + "AnualB.dss";
        }

        // Calcula DRP e DRC
        private void CalculaDRPDRC()
        {
            // Interfaces
            Circuit DSSCircuit = _DSSObj.ActiveCircuit;
            Text DSSText = _DSSObj.Text;

            // se convergiu 
            if (DSSCircuit.Solution.Converged)
            {
                // cria objeto indice tensao

            }
        }

        //Tradução da Função gravaValoresEnergyMeter
        private void GetValoresEnergyMeter(Text DSSText, string hora)
        {
            // Interfaces
            Circuit DSSCircuit = _DSSObj.ActiveCircuit;

            // limpa variavel classe _resultadoFluxo
            _resultadoFluxo = new ResultadoFluxo();

            // se convergiu, obtem dados do medidor 
            if (DSSCircuit.Solution.Converged)
            {
                // Obtem dados para o medidor 
                DSSText.Command = "energymeter.carga.action=take";

                // preenche saida com as perdas do alimentador e verifica se dados estao corretos (ie. convergencia)
                _resultadoFluxo.GetPerdasAlim(DSSCircuit);
            }
            // se nao convergiu
            else 
            {
                _resultadoFluxo.DefineComoNaoConvergiu();
            }

            // TODO caso esteja no modo otimiza a mensagem deverá ser outra
            // caso nao tenha convergido, exibe msg de erro
            if ( ! _resultadoFluxo._convergiuBool )
            {
                string nomeAlim = _paramGerais.getNomeAlimAtual();

                _janela.Disp(nomeAlim + ": Alimentador não convergiu!");
            }
        }

        // get mensagem convergencia 
        private string getMsgConvergencia(string hora, string nomeAlim)
        {
            string str;

            if (hora != null)
            {
                str = nomeAlim + " Hour: " + add1toHour(hora) + " -> Solution Converged";
            }
            else
            {
                str = nomeAlim + " -> Solution Converged";
            }
            return str;
        }

        //adiciona 1hora a string hora
        private static string add1toHour(string hora)
        {
            int dHora = int.Parse(hora);
            dHora++;
            return dHora.ToString();
        }

        //Tradução da função totalizaPerdasEnergiaDiaria
        private ResultadoFluxo TotalizaPerdasEnergiaDiaria(string nomeAlim, List<MyEnergyMeter> resDU)
        {
            List<double> saida = new List<double>();

            // contatena nome alim com tipo do dia 
            string nomeComp = nomeAlim + _paramGerais._parGUI._tipoDia;

            //OLD CODE
            //int colunas = resDU[0].Count;
            int colunas = 12;

            //Perdas e perdas de potência hora a hora
            List<List<double>> perdasHoraAHora = new List<List<double>>();
            List<List<double>> perdasPotHoraAHora = new List<List<double>>();

            /* TODO FIX ME
            foreach (GeracaoPerdasPotEPerdasEnergiaType p in resDU)
            {
                //Perdas hora a hora
                perdasHoraAHora.Add(p.GetRange(2, colunas - 2));

                //Perdas de potência hora a hora
                perdasPotHoraAHora.Add(p.GetRange(0, 2));
            }*/

            // calcula as perdas energia 
            //Inicializa a variável, atribuindo o valor 0 para todos os elementos.
            List<double> perdasEnergia = new List<double>();
            for (int i = 0; i < perdasHoraAHora[0].Count; i++)
            {
                perdasEnergia.Add(0);
            }

            //Percorre cada elemento somando o valor atual com o valor acumulado
            foreach (List<double> p in perdasHoraAHora)
            {
                for (int c = 0; c < p.Count; c++)
                {
                    perdasEnergia[c] += p[c];
                }

            }

            // perdas potencia
            List<double> perdasDUMax = new List<double>();
            perdasDUMax.Add(0);
            perdasDUMax.Add(0);

            foreach (List<double> p in perdasPotHoraAHora)
            {
                for (int c = 0; c < 2; c++)
                {
                    perdasDUMax[c] = p[c] > perdasDUMax[c] ? p[c] : perdasDUMax[c];
                }
            }
            
            //TODO FIX ME
            //cria resPotenciaEEnergia conatenando as perdasDUMax e perdasEnergia 
            ResultadoFluxo res = new ResultadoFluxo();

            return res;
        }
    }
}
