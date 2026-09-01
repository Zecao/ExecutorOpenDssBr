/* #if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif */

using ExecutorOpenDSS.Engine;
using System.Collections.Generic;
using System.Linq;
using ExecutorOpenDSS.AuxClasses;

namespace ExecutorOpenDSS.MainClasses
{
    class PlaceCapacitors
    {
        private readonly GeneralParameters _paramGerais;
        private DailyFlow _daily;
        private static int _contPF=0;

        private List<string> _lst_3PSwichBus;
        private List<string> _lst_BestResults;

        // aux lists
        private List<string> _lstInjEnergy = new List<string>();
        private List<double> _lstReducao = new List<double>();
        private List<string> _lstBus = new List<string>();

        // aux //OLD CODE 
        //private Dictionary<string, string> _dicAlimXBus = new Dictionary<string, string>();

        public PlaceCapacitors(GeneralParameters paramGerais)
        {
            _paramGerais = paramGerais;

            // Obs: necessario
            _paramGerais._medAlim.CarregaDados();

            //Lê os alimentadores e armazena a lista de alimentadores 
            List<string> alimentadores = CemigFeeders.GetTodos(_paramGerais._parGUI.GetArqLstAlimentadores());

            // Limpa Arquivos
            DeletaArqResultados();

            // analisa cada alimentador
            foreach (string nomeAlim in alimentadores)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return;
                }

                // 
                bool ret = PlaceCapacitorsPvt(nomeAlim);

                // saves results
                if (ret)
                {
                    SavesResults2File();
                }

                // TEST plot voltage profile
                // TODO criar nova funcao
                //PlotVoltageProfiles(nomeAlim);
            }
            //TODO saves results
            //SavesResults2File();            
        }

        public void DeletaArqResultados()
        {
            string nomeArq = _paramGerais.GetNomeCompArqCapacitorLossesRed();

            TxtFile.SafeDelete(nomeArq);
        }

        private bool PlaceCapacitorsPvt(string nomeAlim)
        {
            // atribui nomeAlim
            _paramGerais.SetNomeAlimAtual(nomeAlim);

            // Carrega arquivos DSS so MT
            //_daily = new DailyFlow(_paramGerais, true); //OLD CODE
            _daily = new DailyFlow(_paramGerais);

            bool ret = _daily.ExecutaFluxoDiario();            
         
            _contPF++; // power flow counter

            // Se executou fluxo
            if (ret)
            {
                //limpa 
                _lst_BestResults = new List<string>();

                // saves results in a list
                _lst_BestResults.Add("Original" + "\t" + _paramGerais.GetNomeAlimAtual() + "\t" + _daily._resFluxo.GetActiveAndReactiveEnergy() + "\t" + _daily._resFluxo.GetPerdasEnergia().ToString());

                // gets 3 phase switchs bus
                Get3PhaseSwitchBuses();

                // places capacitor and runs power flow
                PlaceCap_RunPowerFlow();

                // get best 
                GetBestBus();

                //plot PF counter
                _paramGerais._mWindow.ExibeMsgDisplay("N. Fluxos de potência: " + _contPF.ToString());
            }
            return ret;
        }

        // Gets the best bus and losses reduction and save to txt file
        private void GetBestBus()
        {
            // lower loss 
            double bestReduction = _lstReducao.Min();

            // index of lower loss //DEBUG
            int ind = _lstReducao.IndexOf(bestReduction);

            //bus of lower loss
            string bus = _lstBus[ind];
            string energy = _lstInjEnergy[ind].ToString();

            _lst_BestResults.Add(bus + "\t" + _paramGerais.GetNomeAlimAtual() + "\t" + energy + "\t" + bestReduction);
        }

        private bool PlaceCap_RunPowerFlow()
        {
            //voltage level
            double basekv = _daily._oDSS._DSSObj.ActiveCircuit.Vsources.BasekV;

            // counter for temporary capacitors
            int capCont = 0;

            // aux lists
            _lstReducao = new List<double>();
            _lstBus = new List<string>();
            _lstInjEnergy = new List<string>();

            bool ret;
            foreach (string bus in _lst_3PSwichBus)
            {
                //Verifica se foi solicitado o cancelamento.
                if (_paramGerais._mWindow._cancelarExecucao)
                {
                    return false;
                }

                //capacitor temporary name
                string cName = "c" + capCont.ToString();

                // gets bus for capacitor placemente
                string capCommand = "new capacitor." + cName + " bus1=" + bus + ",Phases=3,Conn=LN,Kvar=300,Kv=" + basekv;

                _daily._oDSS._DSSObj.Text.Command = capCommand;

                // run Power Flow
                ret = _daily.ExecutaFluxoDiario(1, false, false);//loadMUlt,reload
                _contPF++; // power flow counter

                if (ret)
                {
                    // OLD CODE
                    // saves results in a list
                    //_lst_Results.Add(bus + "\t" + _daily._resFluxo.GetInjectesEnergyAndLosses(_paramGerais.GetNomeAlimAtual()));

                    // saves Energy, loss reduction and bus in Aux Lists 
                    _lstInjEnergy.Add(_daily._resFluxo.GetActiveAndReactiveEnergy());

                    _lstReducao.Add(_daily._resFluxo.GetPerdasEnergia());
                    _lstBus.Add(bus);
                }

                //disable capacitors, accordingly to the OpenDSSengine
                if (EngineConfig.OpenDSSengine)
                {
                //#if ENGINE //OLD CODE
                    _daily._oDSS._DSSObj.Circuits[0].SetActiveElement(cName);
                    _daily._oDSS._DSSObj.Circuits[0].ActiveCktElement.Enabled = false;
                //#endif
                }
                else 
                {
                //#if ! ENGINE
                    _daily._oDSS._DSSObj.Circuits.SetActiveElement(cName);
                    _daily._oDSS._DSSObj.Circuits.ActiveCktElement.Enabled = false;                    
                //#endif
                }
                // capcont
                capCont++;

            }
            return true;
        }

        // Analise de Loops
        private void Get3PhaseSwitchBuses()
        {
            // Obtem switchs
            _lst_3PSwichBus = new List<string>();

            dynamic DSSCircuit = _daily._oDSS.GetActiveCircuit();

            // Obs: necessario setar classe ativa 
            DSSCircuit.SetActiveClass("line");

            //iterator
            int iter = DSSCircuit.Lines.First;

            while (iter != 0)
            {
                string nomeChave = DSSCircuit.Lines.Name;
                int phases = DSSCircuit.Lines.Phases;
                string bus = DSSCircuit.Lines.Bus1;
                bool isSwitch = false;

                if (EngineConfig.OpenDSSengine)
                {
                    dynamic dssText = _daily._oDSS._DSSObj.Text;
                    isSwitch = _daily.IsChave(dssText, nomeChave, DSSCircuit);
                }
                else
                {
                    //#if ! ENGINE OLD CODE

                    // TODO nao esta pegando o Switch corretamente
                    //isSwitch = DSSCircuit.Lines.IsSwitch;

                    dynamic dssText = _daily._oDSS._DSSObj.Text;
                    isSwitch = _daily.IsChave(dssText, nomeChave, DSSCircuit);
                    //#endif
                }
                // gets the bus of 3 phase switches
                if (isSwitch && phases == 3) //OLD CODE && !bus.Contains("#") -> skipa barra com # (representam lixo do Electric Office) 
                {
                    /* // DEBUG
                    string nome = dSSCircuit.Lines.Name;
                    int phases2 = dSSCircuit.Lines.Phases;
                    bool isSwitch2 = dSSCircuit.Lines.IsSwitch;
                    string bus = dSSCircuit.Lines.Bus1;*/

                    _lst_3PSwichBus.Add(bus);
                }

                // goes to next line
                iter = DSSCircuit.Lines.Next;
            }
        }

        //Plota niveis tensao nas barras dos trafos
        private void SavesResults2File()
        {
            TxtFile.GravaListArquivoTXT(_lst_BestResults, _paramGerais.GetNomeCompArqCapacitorLossesRed(), _paramGerais._mWindow);
        }
    }
}
