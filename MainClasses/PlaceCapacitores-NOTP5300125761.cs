//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

using System;
using System.Collections.Generic;
using System.Linq;
using ExecutorOpenDSS.AuxClasses;
using OfficeOpenXml.ConditionalFormatting;

namespace ExecutorOpenDSS.MainClasses
{
    class PlaceCapacitors
    {
        private readonly GeneralParameters _paramGerais;
        private DailyFlow _daily;

        private List<string> _lst_3PSwichBus;
        private List<string> _lst_Results;

        // aux lists
        private List<string> _lstInjEnergy;
        private List<double> _lstReducao;
        private List<string> _lstBus;

        public PlaceCapacitors(GeneralParameters paramGerais, List<string> alimentadores)
        {
            _paramGerais = paramGerais;

            // Obs: necessario
            _paramGerais._medAlim.CarregaDados();

            // Limpa Arquivos
            DeletaArqResultados();

            //TODO
            //_lst_Results = new List<string>();

            // analisa cada alimentador
            foreach (string nomeAlim in alimentadores)
            {
                // TODO 
                _lst_Results = new List<string>();

                PlaceCapacitorsPvt(nomeAlim);
                
                // TODO saves results
                SavesResults2File();
            }
            //TODO saves results
            //SavesResults2File();
        }

        private void DeletaArqResultados()
        {
            string nomeArq = _paramGerais.GetNomeCompArqCapacitorLossesRed();

            TxtFile.SafeDelete(nomeArq);
        }

        private bool PlaceCapacitorsPvt(string nomeAlim)
        {
            // atribui nomeAlim
            _paramGerais.SetNomeAlimAtual(nomeAlim);

            // Carrega arquivos DSS so MT
            _daily = new DailyFlow(_paramGerais, "DU", true);


            bool ret = _daily.LoadStringListwithDSSCommands();

            if (ret)
            {
                ret = _daily.ExecutaFluxoDiario();
            }  
            
            // Se executou fluxo
            if (ret)
            {
                // saves results in a list
                _lst_Results.Add("Original" + "\t" + _paramGerais.GetNomeAlimAtual() + "\t" + _daily._resFluxo.GetActiveAndReactiveEnergy() + "\t" + _daily._resFluxo.GetPerdasEnergia().ToString());

                // gets 3 phase switchs bus
                Get3PhaseSwitchBuses();

                // places capacitor and runs power flow
                PlaceCap_RunPowerFlow();

                // get best 
                GetBestBus();
            }
            return true;
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

            _lst_Results.Add(bus + "\t" + _paramGerais.GetNomeAlimAtual() + "\t" + energy + "\t" + bestReduction );
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
                //capacitor temporary name
                string cName = "c" + capCont.ToString();

                // gets bus for capacitor placemente
                string capCommand = "new capacitor." + cName + " bus1=" + bus + ",Phases=3,Conn=LN,Kvar=300,Kv=" + basekv;

                _daily._oDSS._DSSText.Command = capCommand;

                // run Power Flow
                ret = _daily.ExecutaFluxoDiario(1,false,false);//loadMUlt,reload

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

                //disable capacitors
                _daily._oDSS._DSSObj.Circuits.SetActiveElement(cName);
                _daily._oDSS._DSSObj.Circuits.ActiveCktElement.Enabled = false;
                
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

            Circuit dSSCircuit = _daily._oDSS.GetActiveCircuit();

            //iterator
            int iter = dSSCircuit.Lines.First;             
            while (iter != 0)
            {
                bool isSwitch = dSSCircuit.Lines.IsSwitch;
                int phases = dSSCircuit.Lines.Phases;
                string bus = dSSCircuit.Lines.Bus1;

                // gets the bus of 3 phase switches
                if (isSwitch && phases == 3 && ! bus.Contains("#")) //TODO skipa barra com #
                {
                    /* // DEBUG
                    string nome = dSSCircuit.Lines.Name;
                    int phases2 = dSSCircuit.Lines.Phases;
                    bool isSwitch2 = dSSCircuit.Lines.IsSwitch;
                    string bus = dSSCircuit.Lines.Bus1;*/

                    _lst_3PSwichBus.Add(bus);          
                }

                // goes to next line
                iter = dSSCircuit.Lines.Next;
            }
        }

        //Plota niveis tensao nas barras dos trafos
        private void SavesResults2File()
        {
            TxtFile.GravaListArquivoTXT(_lst_Results, _paramGerais.GetNomeCompArqCapacitorLossesRed(), _paramGerais._mWindow);
        }        
    }
}
