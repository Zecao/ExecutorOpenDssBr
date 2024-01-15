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

namespace ExecutorOpenDSS.MainClasses
{
    class PlaceCapacitors
    {
        private readonly GeneralParameters _paramGerais;
        private DailyFlow _daily;

        private List<string> _lst_3PSwichBus;
        private readonly List<string> _lst_Results; // TODO

        // aux lists
        private readonly List<string> _lstInjEnergy = new List<string>();
        private List<double> _lstReducao = new List<double>();
        private List<string> _lstBus = new List<string>();

        // aux 
        private Dictionary<string, string> _dicAlimXBus = new Dictionary<string, string>();

        public PlaceCapacitors(GeneralParameters paramGerais)
        {
            _paramGerais = paramGerais;

            // Obs: necessario
            _paramGerais._medAlim.CarregaDados();

            //Lê os alimentadores e armazena a lista de alimentadores 
            List<string> alimentadores = CemigFeeders.GetTodos(_paramGerais._parGUI.GetArqLstAlimentadores());

            // Limpa Arquivos
            DeletaArqResultados();

            //TODO
            _lst_Results = new List<string>();

            // analisa cada alimentador
            foreach (string nomeAlim in alimentadores)
            {
                // TODO 
                //_lst_Results = new List<string>();

                PlaceCapacitorsPvt(nomeAlim);
                
                // TODO saves results
                SavesResults2File();

                // TEST plot voltage profile
                // TODO criar nova funcao
                //PlotVoltageProfiles(nomeAlim);
            }
            //TODO saves results
            //SavesResults2File();            
        }

        private void PlotVoltageProfiles(string alim)
        {
            // atribui nomeAlim
            _paramGerais.SetNomeAlimAtual(alim);

            PreencheDic();

            // habilita fluxo horario //TODO
            _paramGerais._parGUI._tipoFluxo = "Hourly";

            // Carrega arquivos DSS so MT
            _daily = new DailyFlow(_paramGerais, true);

            bool ret = _daily.LoadDSSObj();            

            if (ret)
            {

                // Sig.ExecutaFluxoDiario(double loadMult = 0, bool recarga = true, bool plot = true)

                ret = _daily.ExecutaFluxoDiario(1, true, true, "13"); 
            }

            if (ret)
            {
                _daily._oDSS._DSSObj.Text.Command = "plot profile phases=all";
            }

            if (ret)
            {
                ret = _daily.ExecutaFluxoDiario(1, false, true, "19");
            }

            if (ret)
            {
                _daily._oDSS._DSSObj.Text.Command = "plot profile phases=all";
            }

            //
            string bus = _dicAlimXBus[alim];

            double basekv = _daily._oDSS._DSSObj.ActiveCircuit.Vsources.BasekV;
            _daily._oDSS._DSSObj.Text.Command = "new capacitor." + "c1" + " bus1=bmt" + bus + ".1.2.3.0,Phases=3,Conn=LN,Kvar=300,Kv=" + basekv;

            if (ret)
            {
                ret = _daily.ExecutaFluxoDiario(1, false, true, "13");
            }

            if (ret)
            {
                _daily._oDSS._DSSObj.Text.Command = "plot profile phases=all";
            }

            if (ret)
            {
                ret = _daily.ExecutaFluxoDiario(1, false, true, "19");
            }

            if (ret)
            {
                _daily._oDSS._DSSObj.Text.Command = "plot profile phases=all";
            }
        }

        private void PreencheDic()
        {
            _dicAlimXBus = new Dictionary<string, string>()
            {   { "ACSU114", "162580904" },
                { "CMID217","120967627"},
                { "UHSG24","183719757"},
                { "MCLQ408","174891249"},
                { "BCAD209","183357149"},
                { "ACSU108","98288213"},
                { "DVLD215","103499335"},
                { "SLAT313","151995067"},
                { "ITT17","101252510"},
                { "IUMD05","165417436"},
                { "PTC03","124528037"},
                { "MCLD202","123095116"},
                { "SLAT309","153994859"},
                { "PSAU17","180038887"},
                { "URAQ408","172852103"},
                { "ULAE707","154629529"},
                { "SALU09","165813762"},
                { "BNR04","120545205"},
                { "LGT15","158595411"},
                { "NVSU20","174732448"},
                { "BCSQ419","155853245"},
                { "PPS09","164347926"},
                { "SGTU10","125345826"},
                { "MAL09","101611652"},
                { "ITT16","101258285"},
                { "BETQ415","168987040"},
                { "JPIU08","159286730"},
                { "IBTM03","165680374"},
                { "PSAU14","178111783"},
                { "FRUD206","155280978"},
                { "IPC09","101239693"},
                { "IGPD207","176565065"},
                { "PAND214","162580314"},
                { "PDR105","124224069"},
                { "MDH06","179987766"},
                { "CETU04","160755414"},
                { "PMSU16","157747223"},
                { "IRO08","122045214"},
                { "SLAU06","103012310"},
                { "AVT06","144476987"},
                { "ULAE728","150506407"},
                { "BETD223","157256361"},
                { "SRS10","174159095"},
                { "BMO06","101335418"},
                { "SLAT315","171827297"},
                { "ARID212","119858666"},
                { "BSOT312","120654461"},
                { "ACSU115","173866760"},
                { "RCA11","124814582"},
                { "CUVD04","157113669"},
                { "PTID207","171500160"},
                { "IUAU06","178989100"},
                { "IGY05","175021986"},
                { "NVSU06","101994755"},
                { "IRO13","167876354"},
                { "GPED04","183773028"},
                { "SBAU02","102836855"},
                { "SLUU10","178895335"},
                { "ULAE721","153210753"},
                { "TTD521","160899331"},
                { "GPED12","155270405"},
                { "PSOU06","166706559"},
                { "CRM12","174340777"},
                { "CONU08","101957389"},
                { "BCSQ417","181192032"},
                { "MCLD214","123131460"},
                { "BCAD216","120149523"},
                { "AXAU04","119960383"},
                { "IBYD208","121816167"},
                { "ULAE705","126354813"},
                { "PTUU02","182460387"},
                { "PRSU07","165780928"},
                { "MCLD200","178453875"},
                { "URAU18","161301672"},
                { "IUAU03","178495563"},
                { "PSAD215","152924701"},
                { "CEMT14","165716296"},
                { "MHC05","181061332"},
                { "PRRU18","181838319"},
                { "BCAD208","120088350"},
                { "BETD209","98823419"},
                { "ULAD204","166398127"},
                { "PRTU05","165786324"},
                { "IGPU06","183701901"},
                { "BETC516","156856924"},
                { "PTUU09","166160305"},
                { "MAL16","181292900"},
                { "UHPI02","153811716"},
                { "CRM15","156432300"},
                { "FMA03","100453110"},
                { "PTC14","173129358"},
                { "PNV31","158112519"},
                { "SRS09","125584246"},
                { "PNV32","175351683"},
                { "RPA13","150717868"},
                { "PTID205","102599751"},
                { "IGPD208","166045427"},
                { "BETD222","165401208"},
                { "FRUU16","171096224"},
                { "CINC07","179174963"},
                { "PEB07","153457288"},
                { "BDPD203","177760415"},
                { "MUT14","173484284"},
                { "PRSD212","155058433"},
                { "SLAU12","160596531"},
                { "NLAQ414","169749988"},
                { "MCLD203","123080731"},
                { "PTC05","172863205"},
                { "CMH05","179685046"},
                { "CETU10","175570585"},
                { "IGPU03","175167651"},
                { "CAX01","181951201"},
                { "URAT313","141526165"},
                { "PLOT10","102299402"},
                { "SDEU17","182787128"},
                { "MAL12","174972013"},
                { "ULAE714","147925474"},
                { "SLAD210","102928009"},
                { "PZLU08","183262637"},
                { "ARID203","175426462"},
                { "SLAD212","172529157"},
                { "BHAT06","165865602"},
                { "UHGF07","181943190"},
                { "GPED09","179864625"},
                { "DVLD212","163631973"},
                { "GVST303","172540043"},
                { "SLAD206","102923543"},
                { "CITU04","165722862"},
                { "MAL15","178764973"},
                { "BSOT313","172452445"},
                { "UHGF16","103503017"},
                { "CUG08","104484102"},
                { "LAVD14","179683518"},
                { "SLUU15","177526346"},
                { "CNS04","165729592"},
                { "SAA12","124982918"},
                { "CEMT21","108507401"},
                { "CNLU07","158859573"},
                { "SLAT314","171651364"},
                { "PMSU20","124032997"},
                { "BHSN15","150342827"},
                { "SLUQ418","164914683"},
                { "SGS17","176839359"},
                { "IIGT309","100973231"}
            };
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
            _daily = new DailyFlow(_paramGerais, true);

            bool ret = _daily.ExecutaFluxoDiario();            

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

            _lst_Results.Add(bus + "\t" + _paramGerais.GetNomeAlimAtual() + "\t" + energy + "\t" + bestReduction);
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

            bool ret;
            foreach (string bus in _lst_3PSwichBus)
            {
                //capacitor temporary name
                string cName = "c" + capCont.ToString();

                // gets bus for capacitor placemente
                string capCommand = "new capacitor." + cName + " bus1=" + bus + ",Phases=3,Conn=LN,Kvar=300,Kv=" + basekv;

                _daily._oDSS._DSSObj.Text.Command = capCommand;

                // run Power Flow
                ret = _daily.ExecutaFluxoDiario(1, false, false);//loadMUlt,reload

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

#if ENGINE
                //disable capacitors
                throw new NotImplementedException();
#else
                //disable capacitors
                _daily._oDSS._DSSObj.Circuits.SetActiveElement(cName);
                _daily._oDSS._DSSObj.Circuits.ActiveCktElement.Enabled = false;

                // capcont
                capCont++;
#endif
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

#if ENGINE
                //bool isSwitch = false;
                throw new NotImplementedException();
#else
                bool isSwitch = dSSCircuit.Lines.IsSwitch;

                int phases = dSSCircuit.Lines.Phases;
                string bus = dSSCircuit.Lines.Bus1;

                // gets the bus of 3 phase switches
                if (isSwitch && phases == 3 && !bus.Contains("#")) //TODO skipa barra com #
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
#endif
            }
        }

        //Plota niveis tensao nas barras dos trafos
        private void SavesResults2File()
        {
            TxtFile.GravaListArquivoTXT(_lst_Results, _paramGerais.GetNomeCompArqCapacitorLossesRed(), _paramGerais._mWindow);
        }
    }
}
