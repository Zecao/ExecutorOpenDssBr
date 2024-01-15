//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

using ExecutorOpenDSS.AuxClasses;
using System.Collections.Generic;
using System.IO;

namespace ExecutorOpenDSS.MainClasses
{
    public class DailyFlow
    {
        public GeneralParameters _paramGerais;
        public ObjDSS _oDSS;

        private readonly bool _soMT;
        private readonly string _nomeAlim;
        private readonly string _tipoDiaCrv;
        private List<string> _lstCommandsDSS;
        private bool _dssFound;
        public PFResults _resFluxo;
        private readonly List<int> _lstOfIndexModeloDeCarga = new List<int>();
        private readonly Dictionary<string, List<int>> _VRB_tapPerhour = new Dictionary<string, List<int>>();

        // constructor without ObjDSS parameter
        public DailyFlow(GeneralParameters paramGerais, bool soMT = false)
        {
            // seta variavel
            _soMT = soMT;

            // variaveis da classe
            _paramGerais = paramGerais;

            // nome alim
            _nomeAlim = paramGerais.GetNomeAlimAtual();

            // Sets day type (week day, saturday, sunday)
            _tipoDiaCrv = paramGerais._parGUI.GetCodTipoDia();

            //
            _oDSS = new ObjDSS(paramGerais);

            _dssFound = LoadStringListwithDSSCommands();
        }


        // constructor with ObjDSS parameter (used by monthly PF)
        public DailyFlow(GeneralParameters paramGerais, ObjDSS oDSS, string tipoDia = "DU", bool soMT = false)
        {
            // seta variavel
            _soMT = soMT;

            // variaveis da classe
            _paramGerais = paramGerais;

            // nome alim
            _nomeAlim = _paramGerais.GetNomeAlimAtual();

            // Sets day type (week day, saturday, sunday) 
            _tipoDiaCrv = tipoDia;

            //
            _oDSS = oDSS;

            _dssFound = LoadStringListwithDSSCommands();
        }

        /*
        Reduce {All | MeterName}
        Default is "All". Reduce the circuit according to circuit reduction options. See "Set ReduceOptions" and 
        "Set Keeplist" options The Energymeter objects actually perform the reduction. "All" causes all meters to reduce their
        zones.
        
        ReduceOption = { Default or [null] | Stubs [Zmag=nnn] | MergeParallel | BreakLoops | Switches | Laterals | Ends} Strategy for reducing feeders.
        Default is to eliminate all dangling end buses and buses without load, caps, or taps.
        "Stubs [Zmag=0.02]" merges short branches with impedance less than Zmag (default = 0.02 ohms
        "MergeParallel" merges lines that have been found to be in parallel
        "Breakloops" disables one of the lines at the head of a loop.
        "Ends" eliminates dangling ends only.
        "Switches" merges switches with downline lines and eliminates dangling switches.
        “Laterals [Keepload=YesNo]" uses the Remove command to eliminate all 1-phase laterals and optionally lump the load back to the parent 
        2- or 3-phase feeder bus (the default behavior).
        Marking buses with "Keeplist" will prevent their elimination.
        */
        public void RemoveMonophaseLineSegments()
        {
            _oDSS._DSSObj.Text.Command = "Set ReduceOption=Laterals KeepLoad=No";
            _oDSS._DSSObj.Text.Command = "Reduce";
            _oDSS._DSSObj.Text.Command = "BuildY";
        }

        //
        public void AjustaTapsRTs(int Vreg)
        {
            //iterador
            int iReg = _oDSS.GetActiveCircuit().RegControls.First;

            while (iReg != 0)
            {
                _oDSS.GetActiveCircuit().RegControls.ForwardVreg = Vreg;

                // itera
                iReg = _oDSS.GetActiveCircuit().RegControls.Next;
            }
        }

        // altera modelo de carga
        public void SetModeloDeCarga(List<string> novoModeloCarga)
        {
            //DEBUG
            //System.Predicate<string> modeloP = modPFunc;
            //string debug = _lstCommandsDSS.Find(modeloP);

            //DEBUG
            //string debug = _lstCommandsDSS.Find( str => str.Equals("new load.M2constZ pf = 0.92,Vminpu = 0.92,Vmaxpu = 1.5,model = 2,status = variable") );

            _lstCommandsDSS[_lstOfIndexModeloDeCarga[0]] = novoModeloCarga[0];
            _lstCommandsDSS[_lstOfIndexModeloDeCarga[1]] = novoModeloCarga[1];
        }

        // Load StringList with  Alimentador
        public bool LoadStringListwithDSSCommands()
        {
            bool ret;

            // cria arquivo
            if (_soMT)
            {
                ret = LoadDSSCommandStringList_MVFeeder();
            }
            else
            {
                ret = LoadDSSCommandStringList_CompleteFeeder();
            }
            return ret;
        }

        // Load DSS Comands string list
        private bool LoadDSSCommandStringList_CompleteFeeder()
        {
            _lstCommandsDSS = new List<string>
            {
                "clear" // limpa circuito de objeto recem criado
            };

            // nome arquivo dss completo
            string nomeArqDSScomp = _paramGerais.GetNomeArquivoAlimentadorDSS();

            //Verifica existencia do arquivo DSS
            if (File.Exists(nomeArqDSScomp))
            {
                // Obtem linhas do arquivo cabecalho
                string[] lines = System.IO.File.ReadAllLines(nomeArqDSScomp);

                // TODO refactory. Da pau caso tenha so uma linha. 
                _lstCommandsDSS.Add(lines[1]);
            }
            else
            {
                _paramGerais._mWindow.ExibeMsgDisplay(_nomeAlim + ": Arquivos *.dss não encontrados");

                return false;
            }

            // Redirect arquivo Curva de Carga, OBS: de acordo com o TIPO do dia 
            if (File.Exists(_paramGerais.GetNomeEPathCurvasTxtCompleto(_tipoDiaCrv)))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeEPathCurvasTxtCompleto(_tipoDiaCrv));
            }

            // condutores
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeArqCondutor());
            }

            string dir = _paramGerais.GetDiretorioAlim();

            // SegmentosMT
            string nArqSegmentosMT = _nomeAlim + "SegmentosMT.dss";
            if (File.Exists(dir + nArqSegmentosMT))
            {
                _lstCommandsDSS.Add("Redirect " + nArqSegmentosMT);
            }

            // ChavesMT
            string nArqChavesMT = _nomeAlim + "ChavesMT.dss";
            if (File.Exists(dir + nArqChavesMT))
            {
                _lstCommandsDSS.Add("Redirect " + nArqChavesMT);
            }

            // Reguladores.dss
            string nArqRTs = _nomeAlim + "Reguladores.dss";
            if (File.Exists(dir + nArqRTs))
            {
                _lstCommandsDSS.Add("Redirect " + nArqRTs);
            }

            // Transformadores
            string nArqTrafo = _nomeAlim + "Transformadores.dss";
            if (File.Exists(dir + nArqTrafo))
            {
                _lstCommandsDSS.Add("Redirect " + nArqTrafo);
            }

            // SegmentosBT
            string nArqSegmentosBT = _nomeAlim + "SegmentosBT.dss";
            if (File.Exists(dir + nArqSegmentosBT))
            {
                _lstCommandsDSS.Add("Redirect " + nArqSegmentosBT);
            }

            // Ramais
            string nArqRamais = _nomeAlim + "Ramais.dss";
            if (File.Exists(dir + nArqRamais))
            {
                _lstCommandsDSS.Add("Redirect " + nArqRamais);
            }

            // redirect arquivo CargaMT
            if (File.Exists(dir + _paramGerais.GetNomeCargaMT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCargaMT_mes());
            }

            /* // OLD CODE
            // se modelo carga Cemig 
            if (_paramGerais._parGUI._expanderPar._modeloCargaCemig)
            {
                // redirect arquivo CargaBT
                if (File.Exists(dir + _paramGerais.GetNomeCargaBTCemig_mes()))
                {
                    // armazena indice do modelo de Carga 
                    _lstOfIndexModeloDeCarga.Add(_lstCommandsDSS.Count);

                    // adiciona comando de definicao do MOdelo de Carga. Defaul 50% Z
                    _lstCommandsDSS.Add("new load.M2constZ pf = 0.92,Vminpu = 0.92,Vmaxpu = 1.5,model = 2,status = variable");

                    // armazena indice do modelo de Carga 
                    _lstOfIndexModeloDeCarga.Add(_lstCommandsDSS.Count);

                    // adiciona comando de definicao do MOdelo de Carga. Defaul 50% Z
                    _lstCommandsDSS.Add("new load.M3constPsqQ pf = 0.92,Vminpu = 0.92,Vmaxpu = 1.5,model = 3,status = variable");

                    _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCargaBTCemig_mes());
                }
            }
            else //MODELO antigo
            {
                // redirect arquivo CargaBT
                if (File.Exists(dir + _paramGerais.GetNomeCargaBT_mes()))
                {
                    _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCargaBT_mes());
                }
            }*/

            // redirect arquivo CargaBT
            if (File.Exists(dir + _paramGerais.GetNomeCargaBT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCargaBT_mes());

            }

            // redirect arquivo GeradorMT
            if (File.Exists(dir + _paramGerais.GetNomeGeradorMT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeGeradorMT_mes());
            }

            // redirect arquivo GeradorBT
            if (File.Exists(dir + _paramGerais.GetNomeGeradorBT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeGeradorBT_mes());
            }

            // redirect arquivo Capacitor
            if (_paramGerais._parGUI._expanderPar._incluirCapMT && File.Exists(dir + _paramGerais.GetNomeCapacitorMT()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCapacitorMT());
            }

            // nome arquivo dss completo
            string nomeArqBcomp = _paramGerais.GetNomeArquivoB();

            // adiciona comandos secundários
            if (File.Exists(dir + nomeArqBcomp))
            {
                // le linhas do arquivo "B" e adiciona na lista de comandos
                string[] tmp2 = System.IO.File.ReadAllLines(dir + nomeArqBcomp);

                // 
                for (int i = 0; i < tmp2.Length; i++)
                {
                    string tmp = tmp2[i];

                    // skipa linhas vazias
                    if (tmp.Equals(""))
                        continue;

                    // skipa linhas de comentario
                    if (tmp.StartsWith("!"))
                        continue;

                    // adiciona na lista de comandos
                    _lstCommandsDSS.Add(tmp);
                }
            }
            else
            {
                _paramGerais._mWindow.ExibeMsgDisplay("Arquivo " + nomeArqBcomp + " não encontrado");
                return false;
            }

            // BatchEdit
            if (!_paramGerais._parGUI._expanderPar._strBatchEdit.Equals(""))
            {
                // adiciona na lista de comandos
                _lstCommandsDSS.Add(_paramGerais._parGUI._expanderPar._strBatchEdit);
            }

            return true;
        }

        // Cria string com o arquivo DSS na memoria -> Feeder medium voltage part only!
        private bool LoadDSSCommandStringList_MVFeeder()
        {
            _lstCommandsDSS = new List<string>
            {
                // limpa circuito de objeto recem criado
                "clear"
            };

            // nome arquivo dss completo
            string nomeArqDSScomp = _paramGerais.GetNomeArquivoAlimentadorDSS();

            //Verifica existencia do arquivo DSS
            if (File.Exists(nomeArqDSScomp))
            {
                // Obtem linhas do arquivo cabecalho
                string[] lines = System.IO.File.ReadAllLines(nomeArqDSScomp);

                // TODO refactory. Da pau caso tenha so uma linha. 
                // Obtem 2 linha do arquivo  + nivel de tensao em PU
                _lstCommandsDSS.Add(lines[1]);
            }
            else
            {
                _paramGerais._mWindow.ExibeMsgDisplay(_nomeAlim + ": Arquivos *.dss não encontrados");

                return false;
            }

            // Redirect arquivo Curva de Carga, OBS: de acordo com o TIPO do dia 
            if (File.Exists(_paramGerais.GetNomeEPathCurvasTxtCompleto(_tipoDiaCrv)))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeEPathCurvasTxtCompleto(_tipoDiaCrv));
            }

            // condutores
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeArqCondutor());
            }

            string dir = _paramGerais.GetDiretorioAlim();

            // SegmentosMT
            string nArqSegmentosMT = _nomeAlim + "SegmentosMT.dss";
            if (File.Exists(dir + nArqSegmentosMT))
            {
                _lstCommandsDSS.Add("Redirect " + nArqSegmentosMT);
            }

            // ChavesMT
            string nArqChavesMT = _nomeAlim + "ChavesMT.dss";
            if (File.Exists(dir + nArqChavesMT))
            {
                _lstCommandsDSS.Add("Redirect " + nArqChavesMT);
            }

            // Reguladores.dss
            string nArqRTs = _nomeAlim + "Reguladores.dss";
            if (File.Exists(dir + nArqRTs))
            {
                _lstCommandsDSS.Add("Redirect " + nArqRTs);
            }

            /* //OBS: comentei ao debugar possivel erro em manobras em SEs inteira. A matriz de incidencia diminui em 1/6.
            // Transformadores
            string nArqTrafo = _nomeAlim + "Transformadores.dss";
            if (File.Exists(dir + nArqTrafo))
            {
                _lstCommandsDSS.Add("Redirect " + nArqTrafo);
            }*/

            // redirect arquivo CargaMT
            if (File.Exists(dir + _paramGerais.GetNomeCargaMT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCargaMT_mes());
            }

            // redirect arquivo GeradorMT
            if (File.Exists(dir + _paramGerais.GetNomeGeradorMT_mes()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeGeradorMT_mes());
            }

            // redirect arquivo Capacitor
            if (_paramGerais._parGUI._expanderPar._incluirCapMT && File.Exists(dir + _paramGerais.GetNomeCapacitorMT()))
            {
                _lstCommandsDSS.Add("Redirect " + _paramGerais.GetNomeCapacitorMT());
            }

            // nome arquivo dss completo
            string nomeArqBcomp = _paramGerais.GetNomeArquivoB();

            // le linhas do arquivo "B" e adiciona na lista de comandos
            string[] tmp2 = System.IO.File.ReadAllLines(dir + nomeArqBcomp);

            // 
            for (int i = 0; i < tmp2.Length; i++)
            {
                string tmp = tmp2[i];

                // skipa linhas vazias
                if (tmp.Equals(""))
                    continue;

                // skipa linhas de omentario
                if (tmp.StartsWith("!"))
                    continue;

                // adiciona na lista de comandos
                _lstCommandsDSS.Add(tmp);
            }

            //
            return true;
        }

        // Executa fluxo horario caso seja passado string hora
        public bool ExecutaFluxoHorario_SemRecarga(string hora = null)
        {
            // variavel de retorno;
            bool ret = ExecuteDailyPF_SemRecarga(hora);

            return ret;
        }

        // Executa comando no objDSS
        public bool LoadDSSObj()
        {
            //condicao de retorno
            if (_lstCommandsDSS.Count < 2)
            {
                return false;
            }

            // carrega objeto OpenDSS
            foreach (string comando in _lstCommandsDSS)
            {
                _oDSS._DSSObj.Text.Command = comando;
            }
            return true;
        }

        public bool ExecutaFluxoSnap()
        {
            //return cond.
            if (!_dssFound)
            {
                return false;
            }

            // carrega alimentador no objDSS
            bool ret = LoadDSSObj();

            if (ret)
            {
                ret = ExecutaFluxoSnapSemRecarga();
            }

            //
            return ret;
        }

        // Executa fluxo Snap
        public bool ExecutaFluxoSnapSemRecarga()
        {
            // Executa fluxo Snap PVT
            bool ret = ExecutaFluxoSnapPvt();

            //informa usuario convergencia
            if (ret)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(GetMsgConvergencia(null, _nomeAlim));
            }

            //Plota perdas na tela
            _paramGerais._mWindow.ExibeMsgDisplay(_resFluxo.GetResultadoFluxoToConsole(_paramGerais.GetNomeAlimAtual(), _oDSS, _tipoDiaCrv));

            return ret;
        }

        // Executa fluxo snap 
        private bool ExecutaFluxoSnapPvt()
        {
            // Interfaces
            Circuit DSSCircuit = _oDSS._DSSObj.ActiveCircuit;
            Solution DSSSolution = _oDSS._DSSObj.ActiveCircuit.Solution;

            // realiza ajuste das cargas 
            double loadMult = _paramGerais.GetLoadMultFromXlsxFile();

            DSSSolution.LoadMult = loadMult;

            // usuario escolheu tensao barramento
            if (_paramGerais._parGUI._usarTensoesBarramento)
            {
                DSSCircuit.Vsources.pu = double.Parse(_paramGerais._parGUI._tensaoSaidaBarUsuario);
            }

            // OLD CODE
            // seta algorithm Normal ou Newton
            //_oDSS._DSSText.Command = "Set Algorithm = " + _paramGerais._AlgoritmoFluxo;

            // seta modo snap.
            _oDSS._DSSObj.Text.Command = "Set mode=snap";

#if ENGINE
            DSSSolution.Solve();
#else
            try
            {
                DSSSolution.Solve();
            }
            catch (DSSException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return false;
            }
#endif

            if (DSSCircuit.Solution.Converged)
            {
                // TODO verificar se precisa do take p/ snsp
                // OBS: its not necessary to use take in energy meters.                 
                // Obtem dados para o medidor 
                _oDSS._DSSObj.Text.Command = "energymeter.carga.action=take";

                // Obtem valores de pot e energia dos medidores 
                bool ret = GetValoresEnergyMeter(loadMult);

                // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
                GravaPerdasArquivo();

                return true;
            }
            return false;
        }

        // Executa fluxo diario OU horario caso seja passado string hora
        public bool ExecutaFluxoDiario(double loadMult = 0, bool recarga = true, bool plot = true, string hora=null)
        {
            //return cond.
            if (!_dssFound)
            {
                return false;
            }

            //if user cancel the processing
            if (_paramGerais._mWindow._cancelarExecucao)
            {
                return false;
            }

            bool ret;

            // the flag recarga controls the reloading of dss txt files to OpenDSS obj. In the loadmult process does not need to reload.
            if (recarga)
            {
                // carrega objeto OpenDSS
                ret = LoadDSSObj();

                if (!ret)
                {
                    _paramGerais._mWindow.ExibeMsgDisplay("Erro carregamento alimentador " + _nomeAlim);
                }
            }

            // executes a daily PF hour by hour (to compute taps changes)
            if (_paramGerais._parGUI._expanderPar._verifTapsRTs)
            {
                ret = ExecutaDaily_PF_hourByhour();
            }
            else
            {
                // ExecutaFluxoDiario_SemRecarga
                ret = ExecuteDailyPF_SemRecarga(hora, loadMult);
            }

            if (ret && plot)
            {
                //Plota perdas na tela
                _paramGerais._mWindow.ExibeMsgDisplay(_resFluxo.GetResultadoFluxoToConsole(_paramGerais.GetNomeAlimAtual(), _oDSS, _tipoDiaCrv));
            }

            return ret;
        }

        // Run daily PowerFlow
        // 
        private bool ExecuteDailyPF_SemRecarga(string hora, double loadMult = 0)
        {
            //% Interfaces
            Circuit DSSCircuit = _oDSS._DSSObj.ActiveCircuit;
            Solution DSSSolution = _oDSS.GetActiveCircuit().Solution;

            // TODO separar as funcoes Otimiza e FluxoMensal p/ nao precisar confiar neste IF
            //gets loadMult from excel file 
            if (loadMult == 0)
            {
                loadMult = _paramGerais.GetLoadMultFromXlsxFile();

            }

            DSSSolution.LoadMult = loadMult;

            // usuario escolheu tensao barramento
            if (_paramGerais._parGUI._usarTensoesBarramento)
            {
                DSSCircuit.Vsources.pu = double.Parse(_paramGerais._parGUI._tensaoSaidaBarUsuario);
            }

            // TODO da erro no modulo reconfiguracao
            switch (_paramGerais._parGUI._tipoFluxo)
            {
                case "Hourly":
                    _oDSS._DSSObj.Text.Command = "Set mode=daily,hour=" + hora + ",number=1,stepsize=1h";
                    break;

                default: // "daily"
                    _oDSS._DSSObj.Text.Command = "Set mode=daily,hour=0,number=24,stepsize=1h";
                    break;
            }

#if ENGINE
            // resolve circuito 
            DSSSolution.Solve();
#else
            try
            {
                // resolve circuito 
                DSSSolution.Solve();
            }
            catch (DSSException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return false;
            }
#endif

            // se nao convergiu, retorna
            if (!DSSCircuit.Solution.Converged)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(_nomeAlim + " não convergiu!");
                return false;
            }

            // grava valores do EnergyMeter 
            bool ret = GetValoresEnergyMeter(loadMult);

            // se valores EnergyMeter estao consistentes
            if (ret)
            {
                // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
                GravaPerdasArquivo();

                // verifica geracao de relatorios
                GeraRelatorios();
            }

            //informa usuario convergencia
            if (ret)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(GetMsgConvergencia(null, _nomeAlim));
            }
            return ret;
        }
        
        // Run daily PowerFlow (pvt)
        private bool ExecutaDaily_PF_hourByhour()
        {
            //get loadMult
            double loadMult = _paramGerais.GetLoadMultFromXlsxFile();

            _oDSS.GetActiveCircuit().Solution.LoadMult = loadMult;

            // usuario escolheu tensao barramento
            if (_paramGerais._parGUI._usarTensoesBarramento)
            {
                _oDSS._DSSObj.ActiveCircuit.Vsources.pu = double.Parse(_paramGerais._parGUI._tensaoSaidaBarUsuario);
            }

            // Begin
            _oDSS._DSSObj.Text.Command = "Set mode=daily,hour=0,number=1,stepsize=1h";

            bool ret;

            for (int i = 0; i < 24; i++)
            {
#if ENGINE
                _oDSS.GetActiveCircuit().Solution.Solve();
#else
                try
                {
                    // resolve circuito 
                    _oDSS.GetActiveCircuit().Solution.Solve();
                }
                catch (DSSException e)
                {
                    _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                    return false;
                }
#endif

                // se nao convergiu, retorna
                if (!_oDSS.GetActiveCircuit().Solution.Converged)
                {
                    return false;
                }

                // Counts no. of taps
                ret = CountsNoTaps_VR(_VRB_tapPerhour);
            }
            // grava valores do EnergyMeter 
            ret = GetValoresEnergyMeter(loadMult);

            // se valores EnergyMeter estao consistentes
            if (ret)
            {
                // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido 
                GravaPerdasArquivo();

                // verifica geracao de relatorios
                GeraRelatorios();
            }

            //informa usuario convergencia
            if (ret)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(GetMsgConvergencia(null, _nomeAlim));
            }
            return ret;
        }

        public bool CountsNoTaps_VR(Dictionary<string, List<int>> VRB_tapPerhour)
        {
            int iRegControl = _oDSS._DSSObj.ActiveCircuit.RegControls.First;

            // para cada carga
            while (iRegControl != 0)
            {
                // nome trafo
                string trafoName = _oDSS._DSSObj.ActiveCircuit.RegControls.Name;

                int tap = _oDSS._DSSObj.ActiveCircuit.RegControls.TapNumber;

                // add taps to Dictionary
                if (VRB_tapPerhour.ContainsKey(trafoName))
                {
                    List<int> taps = VRB_tapPerhour[trafoName];

                    taps.Add(tap);

                    VRB_tapPerhour.Remove(trafoName);
                    VRB_tapPerhour.Add(trafoName, taps);
                }
                else
                {
                    VRB_tapPerhour.Add(trafoName, new List<int> { tap });
                }

                // itera
                iRegControl = _oDSS._DSSObj.ActiveCircuit.RegControls.Next;
            }
            return true;
        }

        // relatorios modo Hourly
        private void GeraRelatorios()
        {
            /* // FIX ME 
            // condicao de saida
            if (!_paramGerais._parGUI._tipoFluxo.Equals("Hourly") || )
            {
                return;
            }*/

            if (_paramGerais._parGUI._expanderPar._verifTapsRTs)
            {
                //
                VoltageReguladorAnalysis obj = new VoltageReguladorAnalysis(_oDSS._DSSObj.ActiveCircuit, _paramGerais, _VRB_tapPerhour);
                obj.PlotaTapRTs(_paramGerais._mWindow);
            }

            if (_paramGerais._parGUI._expanderPar._calcDRPDRC)
            {
                //
                CalculaDRPDRC();
            }

            // calcula tensao PU no primario trafos
            if (_paramGerais._parGUI._expanderPar._calcTensaoBarTrafo)
            {
                // obtem indices de tensao nos trafos
                TransformerVoltageLevelAnalysis obj2 = new TransformerVoltageLevelAnalysis(_oDSS._DSSObj.Text, _oDSS._DSSObj.ActiveCircuit, _paramGerais);
                obj2.PlotaNiveisTensaoBarras(_paramGerais._mWindow);
            }

            // verifica cargas isoladas
            if (_paramGerais._parGUI._expanderPar._verifCargaIsolada || _paramGerais._parGUI._expanderPar._isolatedLoadsReport)
            {
                //analise cargas isoladas
                IsolatedLoads obj = new IsolatedLoads(_oDSS._DSSObj.ActiveCircuit, _paramGerais);
                obj.PlotaCargasIsoladasArq(_paramGerais._mWindow);
            }

            // TODO criar interface
            //queryLineCode(_DSSObj.ActiveCircuit);

            // TODO criar interface 
            //queryLine( _oDSS._DSSCircuit);

            //IteraSobreLine(_oDSS._DSSCircuit);

            //IteraSobreLineCode(_oDSS._DSSCircuit);

        }

        // TODO criar interface
        public void IteraSobreLine()
        {
            Circuit dSSCircuit = _oDSS.GetActiveCircuit();

            //DEBUG
            int debug = dSSCircuit.Lines.First;

            do
            {
                string nome = dSSCircuit.Lines.Name;

                string lineCode = dSSCircuit.Lines.LineCode;

                int phases = dSSCircuit.Lines.Phases;

#if ENGINE
                //TODO
#else
                bool isSwitch = dSSCircuit.Lines.IsSwitch;
#endif

                /*
                double rho = dSSCircuit.Lines.Rho;
                double Xg = dSSCircuit.Lines.Xg;

                double r0 = dSSCircuit.Lines.R0;
                double r1 = dSSCircuit.Lines.R1;*/

            } while (dSSCircuit.Lines.Next != 0);

        }

        // TODO criar interface
        private void IteraSobreLineCode(Circuit dSSCircuit)
        {
            //DEBUG
            //int debug = dSSCircuit.LineCodes.First;

            do
            {
                string nome = dSSCircuit.LineCodes.Name;
                double r0 = dSSCircuit.LineCodes.R0; //0.1784  valores 336
                double x0 = dSSCircuit.LineCodes.X0; //0.4047
                double r1 = dSSCircuit.LineCodes.R1; //0.0580
                double x1 = dSSCircuit.LineCodes.X1; //0.1206

            } while (dSSCircuit.LineCodes.Next != 0);

        }

        // get mensagem convergencia 
        public string GetMsgConvergencia(string hora, string nomeAlim)
        {
            string str;

            if (hora != null)
            {
                str = nomeAlim + " Hour: " + Add1toHour(hora) + " -> Solução Convergiu";
            }
            else
            {
                str = nomeAlim + " -> Solução Convergiu";
            }
            return str;
        }

        //adiciona 1hora a string hora
        private static string Add1toHour(string hora)
        {
            int dHora = int.Parse(hora);
            dHora++;
            return dHora.ToString();
        }

        // verifica saida e grava perdas em arquivo OU alimentador que nao tenha convergido
        public void GravaPerdasArquivo()
        {
            //Se modo otmiza não grava arquivo
            if (_paramGerais._parGUI._expanderPar._otimizaPUSaidaSE || _paramGerais._parGUI._otmPorEnergia || _paramGerais._parGUI._otmPorDemMax)
            {
                return;
            }

            // Se alim Nao Convergiu 
            if (!_resFluxo._convergiuBool)
            {
                //Grava a lista de alimentadores não convergentes em um txt
                TxtFile.GravaLstAlimNaoConvergiram(_paramGerais);
            }
            // Grava Perdas de acordo com o tipo de fluxo
            else
            {
                string nomeAlim = _paramGerais.GetNomeAlimAtual();

                // obtem o nome do arquivo de perdas, conforme o tipo do fluxo 
                string arquivo = _paramGerais.GetNomeArquivoPerdas();

                // Grava Perdas
                TxtFile.GravaPerdas(_resFluxo, nomeAlim, arquivo, _paramGerais._mWindow);
            }
        }

        // TODO mover p/ classe VoltageLevelAnalysis
        // Calcula DRP e DRC
        private void CalculaDRPDRC()
        {
            // Interfaces
            Circuit DSSCircuit = _oDSS._DSSObj.ActiveCircuit;
            Text DSSText = _oDSS._DSSObj.Text;

            // se convergiu 
            if (DSSCircuit.Solution.Converged)
            {
                // cria objeto indice tensao
                VoltageLevelAnalysis indTensao = new VoltageLevelAnalysis(DSSCircuit, DSSText);

                // Calcula num Clientes com DRP e DRC 
                indTensao.CalculaNumClientesDRPDRC();

                // grava em arquivo
                indTensao.ImprimeNumClientesDRPDRC(_paramGerais);

                //
                indTensao.ImprimeBarrasDRPDRC(_paramGerais);
            }
        }

        // TODO mover p/ classe VoltageLevelAnalysis
        // Calcula DRP e DRC
        public List<string> GetBarrasDRPDRC()
        {
            // Interfaces
            Circuit DSSCircuit = _oDSS._DSSObj.ActiveCircuit;
            Text DSSText = _oDSS._DSSObj.Text;

            // se convergiu 
            if (DSSCircuit.Solution.Converged)
            {
                // cria objeto indice tensao
                VoltageLevelAnalysis indTensao = new VoltageLevelAnalysis(DSSCircuit, DSSText);

                // Calcula num Clientes com DRP e DRC 
                indTensao.CalculaNumClientesDRPDRC();
            }

            return VoltageLevelAnalysis._lstBarrasDRCeDRP;
        }

        /*
         *SampleEnergyMeters ={YES/TRUE | NO/FALSE} 
         
        Overrides default value for sampling EnergyMeter objects at the end of the solution loop. Normally Time and Duty modes do not
        sample EnergyMeters whereas Daily, Yearly, M!, M2, LD1 and LD2 modes do. 

        Use this Option to turn sampling on or off.
          */
        // Get EnergyMeter values
        public bool GetValoresEnergyMeter(double loadMult = 1)
        {
            // nem PFresults
            _resFluxo = new PFResults(loadMult);

            // preenche saida com as perdas do alimentador e verifica se dados estao corretos (ie. convergencia)
            bool ret = _resFluxo.GetPerdasAlim(_oDSS._DSSObj.ActiveCircuit);

            if (!ret)
            { 
                _paramGerais._mWindow.ExibeMsgDisplay(_nomeAlim + " resultados suspeitos EnergyMeter!");
            }

            /* // TODO verificar se precisa desta funcao
            // verifica geracao das usinas (i.e. se estao conectadas)
            VerifGerUsinasGDMT();
            */
            return ret;
        }

        // TODO verificar se estou usando esta funcao
        // verifica se usina do alimentador gerou energia, informando no display da tela.
        private void VerifGerUsinasGDMT()
        {
            // Se existe gerador, obtem gerador
            if (File.Exists(_paramGerais.GetDiretorioAlim() + _paramGerais.GetNomeGeradorMT_mes()))
            {
                // para cada gerador do alimentador
                do
                {
                    string[] nomeGen = _oDSS.GetActiveCircuit().Generators.AllNames;

                    // DEBUG
                    //string[] genRegisterNames = _oDSS.GetActiveCircuit().Generators.RegisterNames;

                    // para cada gerador
                    double[] genRegisterValues = _oDSS.GetActiveCircuit().Generators.RegisterValues;

                    if (genRegisterValues == null)
                    {
                        _paramGerais._mWindow.ExibeMsgDisplay("Usina desconectada!");
                    }

                } while (_oDSS.GetActiveCircuit().Generators.Next != 0);

            }
        }

        // TODO refactory 
        private void QueryLineCode()
        {
            List<string> lstCabos = new List<string>
            {
                "CAB102",
                "CAB103",
                "CAB104",
                "CAB107",
                "CAB108",
                "CAB202",
                "CAB203",
                "CAB204",
                "CAB207",
                "CAB208",
                "CABA06",
                "CABA08",
                "CAB2021",
                "CAB1031",
                "CAB1021",
                "CAB2031",
                "CABA061",
                "CABBT106",
                "CABBT107",
                "CABBT108",
                "CABBT803",
                "CABBT805",
                "CABBT809",
                "CABBT810",
                "CABBT801",
                "CABBT807",
                "CABBT808"
            };

            Text textDSS = _oDSS._DSSObj.Text;

            List<string> resRmatrix = new List<string>();
            List<string> resXmatrix = new List<string>();

            foreach (string lineCode in lstCabos)
            {
                textDSS.Command = "? LineCode." + lineCode + ".Rmatrix";

                resRmatrix.Add(lineCode + "\tRmatrix=" + textDSS.Result);

                textDSS.Command = "? LineCode." + lineCode + ".Xmatrix";

                resXmatrix.Add("\tXmatrix=" + textDSS.Result);
            }

            TxtFile.GravaListArquivoTXT(resRmatrix, _paramGerais.GetArqRmatrix(), _paramGerais._mWindow);

            TxtFile.GravaListArquivoTXT(resXmatrix, _paramGerais.GetArqXmatrix(), _paramGerais._mWindow);
        }

        // TODO refactory 
        private void QueryLine()
        {
            List<string> lstLinhas = new List<string>
            {
                "CAB102",
                "CAB103",
                "CAB104",
                "CAB107",
                "CAB108",
                "CAB202",
                "CAB203",
                "CAB204",
                "CAB207",
                "CAB208",
                "CABA06",
                "CABA08",
                "CABBT106",
                "CABBT107",
                "CABBT108",
                "CABBT803",
                "CABBT805",
                "CABBT809",
                "CABBT810",
                "CABBT801",
                "CABBT807",
                "CABBT808"
            };

            Text textDSS = _oDSS._DSSObj.Text;

            List<string> resRmatrix = new List<string>();
            List<string> resXmatrix = new List<string>();

            foreach (string line in lstLinhas)
            {
                textDSS.Command = "? Line." + line + ".Rmatrix";

                resRmatrix.Add(line + "\tRmatrix=" + textDSS.Result);

                textDSS.Command = "? Line." + line + ".Xmatrix";

                resXmatrix.Add(line + "\tXmatrix=" + textDSS.Result);
            }

            TxtFile.GravaListArquivoTXT(resRmatrix, _paramGerais.GetArqRmatrix(), _paramGerais._mWindow);

            TxtFile.GravaListArquivoTXT(resXmatrix, _paramGerais.GetArqXmatrix(), _paramGerais._mWindow);
        }
    }
}
