//#define ENGINE
#if ENGINE
#else
using dss_sharp;
#endif

using ExecutorOpenDSS.MainClasses;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExecutorOpenDSS
{
    public class RunPowerFlow
    {
        public GeneralParameters _paramGerais;
        public DailyFlow _fluxoDiario;
        public MonthlyPowerFlow _fluxoMensal;
        private readonly List<string> _lstFeeders;

        public RunPowerFlow(GeneralParameters par)
        {
            _paramGerais = par;

            // load metering files (Excel)
            _paramGerais._medAlim.CarregaDados();

            try
            {
                //Lê os alimentadores e armazena a lista de alimentadores 
                _lstFeeders = CemigFeeders.GetTodos(_paramGerais._parGUI.GetArqLstAlimentadores());
            }
            catch (FileNotFoundException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return;
            }

            //Executa o Fluxo
            _paramGerais._mWindow.ExibeMsgDisplay("Executando Fluxo...");
        }

        // 
        public void GeraCargaReligadores()
        {
            // 
            RecloserTracer recloser = new RecloserTracer(_paramGerais);

            //Roda Fluxo para cada alimentador
            foreach (string nomeAlim in _lstFeeders)
            {
                // Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    break;
                }

                // atribui novo alimentador
                _paramGerais.SetNomeAlimAtual(nomeAlim);

                //cria objeto fluxo diario
                _fluxoDiario = new DailyFlow(_paramGerais);
                
                /* OLDCODE12
                // TODO testar
                bool ret = _fluxoDiario.LoadStringListwithDSSCommands();
                */

                //solves snap PF first
                Snap();

                // 
                bool ret = recloser.TraceAllReclosers(_fluxoDiario._oDSS);

                if (ret)
                {
                    recloser.GravaCargaMTBTIP_CSV();
                }
            }

        }

        // executa modo Snap
        public void ExecutesSnap()
        {
            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            //Roda Fluxo para cada alimentador
            foreach (string nomeAlim in _lstFeeders)
            {
                // Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    break;
                }

                // atribui novo alimentador
                _paramGerais.SetNomeAlimAtual(nomeAlim);

                //cria objeto fluxo diario
                _fluxoDiario = new DailyFlow(_paramGerais);

                //Execução com otimização demanda energia
                if (_paramGerais._parGUI._otmPorDemMax)
                {
                    Otimiza();
                }
                //Execução padrão
                else
                {
                    Snap();
                }                
            }

            // TODO testar
            //Se modo otimiza grava arquivo load mult
            if (_paramGerais._parGUI._otmPorDemMax)
            {
                _paramGerais.GravaMapAlimLoadMultExcel();
            }
        }

        // executes monthly power flow
        public void ExecutesMonthlyPowerFlow()
        {
            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            // return condition
            if (_lstFeeders.Count == 0)
            {
                _paramGerais._mWindow.ExibeMsgDisplay("Lista de alimentadores vazia!");
            }

            //Roda fluxo para cada alimentador
            foreach (string alim in _lstFeeders)
            {
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return;
                }

                // atribui nomeAlim
                _paramGerais.SetNomeAlimAtual(alim);

                //cria objeto fluxo diario
                _fluxoMensal = new MonthlyPowerFlow(_paramGerais);

                // if the dss files does not exist
                if (!_fluxoMensal.CheckIfDSSFilesExist())
                {
                    continue;
                }

                //Execução com otimização
                if (_paramGerais._parGUI._otmPorEnergia)
                {
                    // Otimiza por energia
                    Otimiza();

                    continue;
                }
                // Otimiza PU saida
                if (_paramGerais._parGUI._expanderPar._otimizaPUSaidaSE)
                {
                    // Calcula nivel otimo PU
                    OtimizaPUSaidaSE();

                    continue;
                }
                //Execução padrão
                Mensal();
            }

            //Se modo otimiza grava arquivo load mult
            if (_paramGerais._parGUI._otmPorEnergia)
            {
                _paramGerais.GravaMapAlimLoadMultExcel();
            }
        }

        // Run snap Power Flow 
        private void Snap()
        {
            //Verifica se foi solicitado o cancelamento.
            if (_paramGerais._mWindow._cancelarExecucao) { return; }

            _fluxoDiario.ExecutaFluxoSnap();

            /*
            // Entra no diretório ArquivosDss e no subdiretorio 'nomeAlm'
            string nomeArqDSSCompleto = _paramGerais.GetNomeArquivoAlimentadorDSS();

            // Verifica existencia xml
            if (File.Exists(nomeArqDSSCompleto))
            {                
                _fluxoDiario.ExecutaFluxoSnap();
            }
            else
            {
                //Exibe mensagem de erro, caso não encontrem os arquivos de resultados
                _paramGerais._mWindow.ExibeMsgDisplay(_paramGerais.GetNomeAlimAtual() + ": Arquivos *.dss não encontrados");
            }*/
        }

        // Optimise OpenDSS loadMult parameter 
        private void Otimiza()
        {
            // Verifica se foi solicitado o cancelamento.
            if (_paramGerais._mWindow._cancelarExecucao) { return; }

            // otimiza PVT 
            double loadMult = OtimizaPvt();

            // Atualiza map loadMultAlim
            _paramGerais._medAlim.AtualizaMapAlimLoadMult(loadMult);
        }

        // Otimiza Pvt
        private double OtimizaPvt()
        {
            //set Nome Alim atual 
            string alimTmp = _paramGerais.GetNomeAlimAtual();

            // loadMult inicial 
            double loadMultInicial = _paramGerais._medAlim._reqLoadMultMes.GetLoadMult();

            if (loadMultInicial.Equals(double.NaN))
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + ": Alimentador não encontrado no arquivo de ajuste");

                // sets alternative LM
                return _paramGerais._parGUI._loadMultAlternativo;
            }

            // OBS: condicao de retorno: Skipa alimentadores com loadMult igual a zero se loadMult == 0 
            if (loadMultInicial == 0)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + ": Alimentador não otimizado! LoadMult=0");

                return loadMultInicial;  
            }

            // OBS: condicao de retorno: Skipa alimentadores com loadMult > 10 
            if ((loadMultInicial > 10) || (loadMultInicial < 0.1))
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + ": Alimentador não otimizado! LoadMult = " + loadMultInicial);

                return 1;
                //return loadMultInicial;
            }

            double loadMult;

            // modo otimiza LoadMult energia 
            if (_paramGerais._parGUI._otmPorEnergia)
            {
                loadMult = OtimizaLoadMultEnergiaPvt(alimTmp, loadMultInicial);
            }
            //modo otimiza loadMult potencia
            else
            {
                loadMult = OtimizaLoadMultPvt(loadMultInicial);
            }
            return loadMult;
        }

        // Run daily Power Flow
        public void ExecutesDailyPowerFlow()
        {
            //Limpa arquivos
            _paramGerais.DeletaArqResultados();

            // cria arquivo e preenche Cabecalho, caso modo _calcDRPDRC
            if (_paramGerais._parGUI._expanderPar._calcDRPDRC)
            {
                // Cria arquivo cabecalho
                VoltageLevelAnalysis.CriaArqCabecalho(_paramGerais);
            }

            //Roda o fluxo para cada alimentador
            foreach (string nomeAlim in _lstFeeders)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return;
                }

                // atribui nomeAlim
                _paramGerais.SetNomeAlimAtual(nomeAlim);

                //cria objeto fluxo diario
                _fluxoDiario = new DailyFlow(_paramGerais);

                // Executa fluxo diário openDSS
                _fluxoDiario.ExecutaFluxoDiario();
            }
        }

        //Monthly power flow
        private bool Mensal()
        {
            // Executa Fluxo Mensal 
            bool ret = _fluxoMensal.ExecutaFluxoMensal();

            return ret;
        }

        // Fluxo anual
        public void ExecutesAnnualPowerFlow()
        {
            //Limpa Arquivos
            _paramGerais.DeletaArqResultados();

            //Seta modo anual
            _paramGerais._parGUI._modoAnual = true;

            //Roda fluxo para cada alimentador
            foreach (string alim in _lstFeeders)
            {
                // atribui alim 
                _paramGerais.SetNomeAlimAtual(alim);

                //Vetor _resultadoFluxo
                List<PFResults> lstResultadoFluxo = new List<PFResults>();

                // para cada mes executa um fluxo mensal
                for (int mes = 1; mes < 13; mes++)
                {
                    // Verifica se foi solicitado o cancelamento.
                    if (_paramGerais._mWindow._cancelarExecucao)
                    {
                        return;
                    }

                    // set mes
                    _paramGerais._parGUI.SetMes(mes);

                    //cria objeto fluxo diario
                    _fluxoMensal = new MonthlyPowerFlow(_paramGerais);

                    // Otimiza
                    if (_paramGerais._parGUI._otmPorEnergia)
                    {
                        Otimiza();
                    }
                    else
                    {
                        // Executa Fluxo Mensal 
                        _fluxoMensal.ExecutaFluxoMensal();
                    }
                    PFResults resTmp = new PFResults(_fluxoMensal._resFluxoMensal);

                    //Armazena Resultado
                    lstResultadoFluxo.Add(resTmp);
                    //}
                    /*
                    // FIX ME
                    // se nao carregou alimentador, forca mes = 13 para terminar o for 
                    else 
                    {
                        mes = 13;
                        break;
                    }*/
                }

                // se convergiu, Calcula resultado Ano 
                if (_fluxoMensal._resFluxoMensal._convergiuBool)
                {
                    //Calcula resultado Ano
                    _fluxoMensal._resFluxoMensal.CalculaResAno(lstResultadoFluxo, alim, _paramGerais.GetNomeComp_arquivoResPerdasAnual(), _paramGerais._mWindow);
                }
            }
        }

        //Optimise substation voltage level
        private void OtimizaPUSaidaSE()
        {
            // 
            double puSaidaSE = 0.96;

            while (puSaidaSE < 1.05)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return;
                }

                // atribui pu Saida SE
                _paramGerais.SetTensaoSaidaSE(puSaidaSE.ToString());

                // Executa Fluxo Mensal 
                _fluxoMensal.ExecutaFluxoMensal();

                // Atualiza PU
                puSaidaSE += 0.02;
            }
        }

        // Optimise OpenDSS loadMult parameter 
        private double OtimizaLoadMultPvt(double loadMult)
        {
            // alimTmp
            string alimTmp = _paramGerais.GetNomeAlimAtual();

            //referência de geração (semana típica ou medição do MECE)
            double refGeracao;
            if (_paramGerais._medAlim._mapDadosDemanda.ContainsKey(alimTmp))
            {
                refGeracao = _paramGerais._medAlim._mapDadosDemanda[alimTmp];
            }
            else
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + ": mapDadosDemanda não encontrado");
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

            // while difference of simulated energy and refEnergy is greater than precision
            while (Math.Abs(geracao - refGeracao) > _paramGerais._parGUI.GetPrecisao())
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return loadMult;
                }

                // atualiza loadMult
                loadMultAnt = loadMult;
                loadMult = _paramGerais._parGUI.GetIncremento() * loadMult;

                // calcula geração
                geracao = ExecutaFluxoSnapOtm(loadMult);

                // incrementa contador fluxo
                contFP++;

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

        // Optimise OpenDSS loadMult parameter accordingly with month energy measurement
        private double OtimizaLoadMultEnergiaPvt(string alimentador, double loadMult)
        {
            // obtem referência de EnergiaMes (medição do MECE)
            double refEnergiaMes = _paramGerais._medAlim._reqEnergiaMes.GetRefEnergia(alimentador, _paramGerais._parGUI.GetMes());

            // verifica ausencia de medicao de energia
            if (refEnergiaMes.Equals(0) || refEnergiaMes.Equals(double.NaN))
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimentador + ": Energia mansal igual a 0 ou NaN. Atribuindo loadMult alternativo.");

                return _paramGerais._parGUI._loadMultAlternativo;
            }

            //calcula energia inicial
            if (!ExecutaFluxoMensalOtm(loadMult))
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimentador + " alimentador não convergiu! Atribuindo loadMult alternativo.");

                return _paramGerais._parGUI._loadMultAlternativo;
            }

            // condicao de saida da funcao
            // energia simulada (energiaMes) esta proxima da energia medida (refEnergiaMes)
            if (VerificaEnergiaSimulada(refEnergiaMes))
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimentador + " ajustado com precisão " + _paramGerais._parGUI.GetPrecisao().ToString() + "KWh.");

                return loadMult;
            }

            // Se fluxo não convergiu retorna loadMult alternativo
            if (!_fluxoMensal._resFluxoMensal._convergiuBool)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(alimentador + ": erro valor energia!");

                return _paramGerais._parGUI._loadMultAlternativo;
            }

            // 
            double energiaMesTmp = _fluxoMensal._resFluxoMensal.GetEnergia();

            // sentido de busca 
            int sentidoBusca = 1;

            // Se geração inicial for MAIOR que referência
            if (energiaMesTmp > refEnergiaMes)
            {
                // sentido  // retroceder o loadmult
                sentidoBusca = -1;
            }

            // ajusta modelo de carga se opcao _paramGerais._expanderPar._modeloCargaCemig = true
            _fluxoMensal.AjustaModeloDeCargaCemig(sentidoBusca);

            // calcula novo LoadMult
            loadMult = AjustaLoadMult(refEnergiaMes, loadMult, sentidoBusca);

            return loadMult;
        }

        // condicao de saida da funcao
        // energia simulada (energiaMes) esta proxima da energia medida (refEnergiaMes)
        private bool VerificaEnergiaSimulada(double refEnergiaMes)
        {
            // se aproximacao de energia estiver dentro do limite de precisao ja calculado, ignora nova otimizacao
            if (GetDiferencaEnergia(refEnergiaMes) < _paramGerais._parGUI.GetPrecisao())
            {
                return true;
            }
            return false;
        }

        //obtem diferenca de energia de acordo com o modo
        private double GetDiferencaEnergia(double refEnergiaMes)
        {
            double delta;

            delta = Math.Abs((_fluxoMensal._resFluxoMensal.GetEnergia() - refEnergiaMes));

            return delta;
        }

        // ajusta loadMult incrementalmente
        private double AjustaLoadMult(double refEnergiaMes, double loadMult, int sentidoBusca = 1)
        {
            // contador fluxo de potência
            // OBS: considera a execução acima
            int contFP = 1;

            // variavel que armazena o loadMult anterior
            double loadMultAnt = loadMult;
            int sentidoBuscaAnt = sentidoBusca;

            // alimTmp
            string alimTmp = _paramGerais.GetNomeAlimAtual();

            // enquanto erro da energia for maior que precisao
            while (GetDiferencaEnergia(refEnergiaMes) > _paramGerais._parGUI.GetPrecisao())
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return loadMultAnt;
                }

                // detecta se mudou sendido de busca, interrompendo o processo e retornando loadMult apropriado. 
                if (sentidoBuscaAnt != sentidoBusca)
                {
                    _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + " mudança sentido de busca. Diminua o Incremento/Precisão!");

                    // se sentido de busca era decrescente, admite-se retornar loadMult 
                    if (sentidoBuscaAnt == -1)
                        return loadMult;
                    else
                        return loadMultAnt;
                }

                // condicao saida while: se executou 10 FP
                if (contFP == 10)
                {
                    _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + " não ajustado em " + contFP.ToString() + " execuções de fluxo. Repita o processo!");

                    return loadMultAnt;
                }

                // 
                // atualiza loadMult, de acordo com o sentido de ajuste
                loadMultAnt = loadMult;
                if (sentidoBusca == 1)
                {
                    loadMult *= _paramGerais._parGUI.GetIncremento();
                }
                // retrocede ajuste 
                else
                {
                    loadMult /= _paramGerais._parGUI.GetIncremento();
                }

                // calcula energia mes. Se modo fluxo mensal Aproximado com recarga=false
                if (!ExecutaFluxoMensalOtm(loadMult, false))
                {
                    _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + " não convergiu! Retornando loadMult anterior.");

                    return loadMultAnt;
                }

                // incrementa contador fluxo
                contFP++;

                //atualiza sentido de busca
                if ((_fluxoMensal._resFluxoMensal.GetEnergia() > refEnergiaMes) && (sentidoBusca == 1) ||
                     (_fluxoMensal._resFluxoMensal.GetEnergia() < refEnergiaMes) && (sentidoBusca == -1))
                {
                    // guarda sentido busca anterior
                    sentidoBuscaAnt = sentidoBusca;

                    // inverte sentido 
                    sentidoBusca = -sentidoBusca;
                }
            }
            //saida 
            _paramGerais._mWindow.ExibeMsgDisplay(alimTmp + " ajustado nos limites escolhidos.");

            return loadMult;
        }

        // ExecutaFluxoMensalOtm
        private bool ExecutaFluxoMensalOtm(double loadMult, bool recarga = true)
        {
            // fluxo mensal aproximado
            if (_paramGerais._parGUI.GetAproximaFluxoMensalPorDU())
            {
                // Executa fluxo diario
                return (_fluxoMensal.ExecutaFluxoMensalAproximacaoDU(loadMult, recarga));
            }
            else
            {
                // Executa fluxo mensal
                // OBS:FluxoMensal nao pode ser feito sem recarga  
                return (_fluxoMensal.ExecutaFluxoMensal(loadMult));
            }
        }

        // run snap power flow 
        private double ExecutaFluxoSnapOtm(double loadMult)
        {
            // Chama fluxo snap
            Snap();

            //alimTmp
            string alimTmp = _paramGerais.GetNomeAlimAtual();

            //informa usuario convergencia
            _paramGerais._mWindow.ExibeMsgDisplay(_fluxoDiario.GetMsgConvergencia(null, alimTmp));

            // retorno
            return _fluxoDiario._resFluxo.GetMaxKW();
        }
    }
}
