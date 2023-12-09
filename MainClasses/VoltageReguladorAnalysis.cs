#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

using System.Collections.Generic;
using ExecutorOpenDSS.AuxClasses;
using System;

namespace ExecutorOpenDSS.MainClasses
{
    class VoltageReguladorAnalysis
    {
        private readonly Circuit _circuit;
        private readonly Transformers _trafosDSS;
        private readonly List<string> _tapsRT;
        private readonly GeneralParameters _param;
        private Dictionary<string, List<int>> _VRB_tapPerhour;
        private List<string> _VRBtapCounter;

        //constructor 
        public VoltageReguladorAnalysis(Circuit cir, GeneralParameters paramGerais, Dictionary<string, List<int>> VRB_tapPerhour)
        {
            _circuit = cir;
            _trafosDSS = cir.Transformers;
            _param = paramGerais;

            _VRB_tapPerhour = VRB_tapPerhour;

            _tapsRT = new List<string>();
        }

        //Plota niveis tensao nas barras dos trafos
        public void PlotaTapRTs(MainWindow janela)
        {
            // se convergiu 
            if (_circuit.Solution.Converged)
            {
                // Get taps Voltage Regulators
                GetTapRTs();

                // Calculates number of tap changinf
                CountTapChangings();

                // Grava arquivo
                GravaTapRTsArq(janela);
            }
        }

        private void CountTapChangings()
        {
            _VRBtapCounter = new List<string>();

            // for each Voltage regulator
            foreach (string key in _VRB_tapPerhour.Keys)
            {
                List<int> TapsHour = _VRB_tapPerhour[key];

                int hourInCheck = TapsHour[0]; // hour 0 is ther first hourInCheck
                int tapChanges = hourInCheck; // also, the first number of tap Changes 

                for (int i = 1; i < 24 - 1; i++)
                {
                    // se 2a hora em diante diferente 
                    if (TapsHour[i] != hourInCheck)
                    {
                        // 
                        int deltaTaps = Math.Abs(hourInCheck - TapsHour[i]);

                        // new hourInCheck
                        hourInCheck = TapsHour[i];

                        // 
                        tapChanges += deltaTaps;
                    }
                }

                // add tapChanges in the Dic.
                _VRBtapCounter.Add(_param.GetNomeAlimAtual() + "\t" + key + "\t" + tapChanges.ToString());
            }
        }

        // calcula tensao barra trafos 
        public void GetTapRTs()
        {
            int iTrafo = _trafosDSS.First;

            // para cada carga
            while (iTrafo != 0)
            {
                // nome trafo
                string trafoName = _trafosDSS.Name;

                //skipa banco de reguladores
                if (trafoName.Contains("rt"))
                {
                    //add
                    _tapsRT.Add(_param.GetNomeAlimAtual() + "\t" + trafoName + "\t" + _trafosDSS.Tap);
                }

                // itera
                iTrafo = _trafosDSS.Next;
            }
        }

        // 
        public void GravaTapRTsArq(MainWindow janela)
        {
            TxtFile.GravaListArquivoTXT(_tapsRT, _param.GetNomeArqTapsRTs(), janela);
            TxtFile.GravaListArquivoTXT(_VRBtapCounter, _param.GetNomeArqTapsRTs(), janela);
        }
    }
}
