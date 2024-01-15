//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

using ExecutorOpenDSS.AuxClasses;
using ExecutorOpenDSS.MainClasses;
using System.Collections.Generic;
using System.IO;

namespace ExecutorOpenDSS
{
    public class FeederSummary
    {
        private readonly GeneralParameters _paramGerais;
        private List<string> _lst_Results = new List<string> { "CodAlim\tnTrafo\tnVRB\tnCAP\tCAP_KVAr\tMVLoads\tMVLoads_kW\tLVLoads\tLVLoads_kW" +
            "\tnPV-MV\tPV-MV_kVA\tnPV-LV_PV\tPV-LV_kVA" +
            "\tnGerMV\tGerMV_kVA\tGerLV\tGerLV_kVA\t"};

        public FeederSummary(GeneralParameters paramGerais)
        {
            _paramGerais = paramGerais;
            List<string> alimentadores;

            try
            {
                alimentadores = CemigFeeders.GetTodos(_paramGerais._parGUI.GetArqLstAlimentadores());
            }
            catch (FileNotFoundException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return;
            }

            // Obs: necessario
            _paramGerais._medAlim.CarregaDados();

            // Limpa Arquivos
            DeletaArqResultados();

            // analisa cada alimentador
            foreach (string nomeAlim in alimentadores)
            {
                // atribui nomeAlim
                _paramGerais.SetNomeAlimAtual(nomeAlim);

                // Carrega arquivos DSS so MT
                DailyFlow _daily = new DailyFlow(_paramGerais, false);

                _daily.LoadDSSObj();

                /*
                if (ret)
                {
                    ret = _daily.ExecutaFluxoDiario_SemRecarga("13", 1);
                }*/
                                            
                // Transformers 
                int numTransformers = _daily._oDSS._DSSObj.ActiveCircuit.Transformers.Count;

                // VRBs
                int numVRBs = _daily._oDSS._DSSObj.ActiveCircuit.RegControls.Count;

                // Capacitors //1st number 2nd capacity
                List<double> capCount = Sum_CapacitorsCap(_daily._oDSS._DSSObj.ActiveCircuit.Capacitors);

                // Generators //1st number 2nd capacity
                List<double> genCap = Sum_GeneratorCap(_daily._oDSS._DSSObj.ActiveCircuit.Generators);

                // PVSystems //1st number 2nd capacity
                List<double> PVSystem_cap = Sum_PVSystemCap(_daily._oDSS._DSSObj.ActiveCircuit.PVSystems, _daily._oDSS._DSSObj.Text);

                // Loads 
                // int numCustomers = _daily._oDSS._DSSObj.ActiveCircuit.Loads.Count; //OBS includes public lights
                List<double> loads = Count_Loads(_daily._oDSS._DSSObj.ActiveCircuit.Loads);

                //
                string txt = nomeAlim + "\t" + numTransformers.ToString() + "\t" + numVRBs.ToString() + "\t" + capCount[0].ToString() + "\t" + capCount[1].ToString()
                    + "\t" + loads[0].ToString() + "\t" + loads[1].ToString() + "\t" + loads[2].ToString() + "\t" + loads[3].ToString() //loads
                    + "\t" + PVSystem_cap[0].ToString() + "\t" + PVSystem_cap[1].ToString() + "\t" + PVSystem_cap[2].ToString() + "\t" + PVSystem_cap[3].ToString() //PVSystem
                    + "\t" + genCap[0].ToString() + "\t" + genCap[1].ToString() + "\t" + genCap[2].ToString() + "\t" + genCap[3].ToString() ; //Generator

                _lst_Results.Add(txt);

                // TODO saves results
                //SavesResults2File();

            }
            //TODO saves results
            SavesResults2File();     
        }

        private List<double> Count_Loads(Loads loads)
        {
            double LV_count = 0.0;
            double MV_count = 0.0;
            double LV_kw = 0.0;
            double MV_kw = 0.0;

            int iter = loads.First;

            // para cada carga
            while (iter != 0)
            {
                if (loads.kV.Equals(13.8) || loads.kV.Equals(22.0) || loads.kV.Equals(34.5))
                {
                    MV_count += 1;
                    MV_kw += loads.kW;
                }
                else
                {
                    /* // OLD CODE excludes Public Lightnings
                    if (! loads.Name.Contains("IP"))
                    {
                        LV_count += 1;
                        LV_kw += loads.kW;
                    }*/
                    LV_count += 1;
                    LV_kw += loads.kW;
                }

                // itera
                iter = loads.Next;
            }

            List<double> ret = new List<double> { MV_count, MV_kw, LV_count, LV_kw };
            return ret;
        }

        //Plota niveis tensao nas barras dos trafos
        private void SavesResults2File()
        {
            TxtFile.GravaListArquivoTXT(_lst_Results, _paramGerais.GetNomeCompArqResumoAlim(), _paramGerais._mWindow);
        }

        private List<double> Sum_CapacitorsCap(Capacitors caps)
        {
            int numCap = caps.Count;

            if (numCap == 0)
            {
                new List<double> { 0.0, 0.0 };
            }

            double capKVAr = 0.0;

            int iter = caps.First;

            // para cada carga
            while (iter != 0)
            {

                capKVAr += caps.kvar;

                // itera
                iter = caps.Next;
            }

            return new List<double> { numCap, capKVAr };
        }

        private List<double> Sum_GeneratorCap(Generators gen)
        {
            //retorno
            if (gen.Count == 0)
            { 
                return new List<double> { 0.0, 0.0, 0.0, 0.0 };
            }

            //double LV_kW = 0.0;
            //double MV_kW = 0.0;
            double LV_kVA = 0.0;
            double MV_kVA = 0.0;
            double LV_count = 0.0;
            double MV_count = 0.0;

            int iter = gen.First;

            // para cada carga
            while (iter != 0)
            {
                if (gen.kV.Equals(13.8) || gen.kV.Equals(22.0) || gen.kV.Equals(34.5))
                {
                    //MV_kW += gen.kW;
                    MV_kVA += gen.kVArated;
                    MV_count++;
                }
                else
                {
                    //LV_kW += gen.kW;
                    LV_kVA += gen.kVArated;
                    LV_count++;
                }

                // itera
                iter = gen.Next;
            }
            List<double> ret = new List<double> { MV_count, MV_kVA, LV_count, LV_kVA };

            return ret;
        }

        private List<double> Sum_PVSystemCap(PVSystems pv, Text cl)
        {
            int numPVSystem = pv.Count;

            //retorno
            if (numPVSystem == 0)
            {
                new List<double> { 0.0, 0.0, 0.0, 0.0 };
            }

            double PV_MV_kVA = 0.0;
            double PV_MV_count = 0.0;
            double PV_LV_kVA = 0.0;
            double PV_LV_count = 0.0;

            int iter = pv.First;

            // para cada carga
            while (iter != 0)
            {
                cl.Command = "? PVSystem." + pv.Name + ".kV";
                string sKv = cl.Result;

                // TODO FIX whenfuture code for separate MV from LV PVSystems
                if (sKv.Equals("13.8") || sKv.Equals("22.0") || sKv.Equals("34.5"))
                {
                    PV_MV_kVA += pv.kVArated;
                    PV_MV_count++;
                }
                else
                {
                    PV_LV_kVA += pv.kVArated;
                    PV_LV_count++;
                }

                // itera
                iter = pv.Next;
            }

            return new List<double> { PV_MV_count, PV_MV_kVA, PV_LV_count, PV_LV_kVA };
        }

        public void DeletaArqResultados()
        {
            string nomeArq = _paramGerais.GetNomeCompArqResumoAlim();

            TxtFile.SafeDelete(nomeArq);
        }
    }
}
