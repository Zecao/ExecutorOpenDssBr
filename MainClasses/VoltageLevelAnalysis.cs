using System;
using System.Collections.Generic;
using ExecutorOpenDSS.AuxClasses;
using ExecutorOpenDSS.Engine;

namespace ExecutorOpenDSS.MainClasses
{
    class VoltageLevelAnalysis
    {
        private static int _numClientesDRP;
        private static int _numClientesDRC;
        private static int _numClientesOK;
        private static int _numClientesTotal;
        private static int _numClientesIP;
        private readonly dynamic _circuit;
        private readonly dynamic _DSSText;

        public static List<string> _lstBarrasDRCeDRP = new List<string>();

        public VoltageLevelAnalysis(dynamic cir, dynamic txt)
        {
            _circuit = cir;
            _DSSText = txt;

            _numClientesTotal = _circuit.Loads.Count;
            _numClientesDRP = 0;
            _numClientesDRC = 0;
            _numClientesIP = 0;
        }

        // Cria arquivo texto cabecalho DRP DRC 
        public static void CriaArqCabecalho(GeneralParameters paramGerais)
        {
            // 
            string nomeArq = paramGerais.GetNomeComp_arquivoDRPDRC();

            //Grava cabecalho
            string linha = "Alim\tF.A.\tDRP:\tDRC:\tTotal:";

            //Grava em arquivo
            TxtFile.GravaEmArquivo(linha, nomeArq, paramGerais._mWindow);
        }

        // obtem barra da carga por meio da interface de Text
        private string GetLoadBusName(string loadName)
        {
            // get load bus name
            _DSSText.Command = "? Load." + loadName + ".Bus1";

            // nome Barra
            string busName = _DSSText.Result;

            return busName;
        }

        // verificaNivelTensaoBarra
        // OBS: necessario receber o kVcarga, uma vez que o kvBase da barra eh alterado com a mudanca nos 
        // taps dos transformadores
        private void VerificaNivelTensaoBarra(string nomeBarra, double kVcarga, int numPhases)
        {
            // seta activebus a barra da carga 
            _circuit.SetActiveBus(nomeBarra);

            // obtem a barra, apos ativada
            dynamic barraDSS = _circuit.ActiveBus;

            // array de tensoes
            double[] tensaoPU = barraDSS.VMagAngle;
            double tensaoFaseA = tensaoPU[0];

            // se tensao igual igual a 0, retorna
            if (tensaoFaseA == 0)
            {
                return;
            }

            double tensoaFaseApu;

            // tensao fase A em pu
            if (numPhases == 1)
            {
                tensoaFaseApu = tensaoFaseA / (kVcarga * 1000 );
            }
            else
            {
                tensoaFaseApu = tensaoFaseA / (kVcarga * 1000 / Math.Sqrt(3));
            }         

            //BT OBS: assume que qquer tensao abaixo de 500Volts eh BT
            if (Math.Round(kVcarga, 3) <= 0.240)
            {
                VerificaNivelTensaoBarraBT(tensoaFaseApu, nomeBarra);
            }
            else
            {
                VerificaNivelTensaoBarraMT(tensoaFaseApu, nomeBarra);
            }
        }

        // verificaNivelTensaoBarraBT e incrementa contador de clientes
        private void VerificaNivelTensaoBarraBT(double tensaoPU, string nomeBarra)
        {
            //DRC < 191/220 V
            if (tensaoPU < 0.868)
            {
                _numClientesDRC++;
                _lstBarrasDRCeDRP.Add(nomeBarra + "\t" + tensaoPU.ToString());

                return;
            }
            //DRP para cada carga, checa o tensao
            if ((tensaoPU < 0.9213) && (tensaoPU >= 0.868))
            {
                _numClientesDRP++;
                _lstBarrasDRCeDRP.Add(nomeBarra + "\t" + tensaoPU.ToString());
            }
        }

        // verificaNivelTensaoBarraMT e incrementa contador de clientes
        private void VerificaNivelTensaoBarraMT(double tensaoPU, string nomeBarra)
        {
            //DRC para cada carga, checa o tensao
            if (tensaoPU < 0.90)
            {
                _numClientesDRC++;

                _lstBarrasDRCeDRP.Add(nomeBarra + "\t" + tensaoPU.ToString());

                return;
            }

            //DRP para cada carga, checa o tensao
            if ((tensaoPU >= 0.90) && (tensaoPU < 0.93))
            {
                _numClientesDRP++;

                _lstBarrasDRCeDRP.Add(nomeBarra + "\t" + tensaoPU.ToString());
            }
        }
        // TODO
        // imprime numero de clientes DPRPDR
        private void ImprimeNumClientesDRPDRC(MainWindow janela)
        {
            // numero clientes DRP
            janela.ExibeMsgDisplay("Clientes com DRP: " + _numClientesDRP.ToString());
            janela.ExibeMsgDisplay("Clientes com DRC: " + _numClientesDRC.ToString());
            janela.ExibeMsgDisplay("Clientes totais: " + _numClientesTotal.ToString());
        }

        // grava numero clientes com DRP e DRC no arquivo
        public void ImprimeNumClientesDRPDRC(GeneralParameters paramGerais)
        {
            //nome arquivo DRP e DRC
            string nomeArq = paramGerais.GetNomeComp_arquivoDRPDRC();

            // nome alim
            string nomeAlim = paramGerais.GetNomeAlimAtual();

            // linha //ALim DRP DRC totais    
            string linha = nomeAlim + "\t" + _numClientesOK.ToString() + "\t" + _numClientesDRP.ToString() + "\t" + _numClientesDRC.ToString() + "\t" + _numClientesTotal.ToString();

            //Grava em arquivo
            TxtFile.GravaEmArquivo(linha, nomeArq, paramGerais._mWindow);
        }

        // grava numero clientes com DRP e DRC no arquivo
        public void ImprimeBarrasDRPDRC(GeneralParameters paramGerais)
        {
            //nome arquivo DRP e DRC
            string nomeArq = paramGerais.GetNomeComp_arqBarrasDRPDRC();

            // nome alim
            string nomeAlim = paramGerais.GetNomeAlimAtual();

            // linha  
            List<string> lstStr = new List<string>();

            //
            foreach (string barras in _lstBarrasDRCeDRP)
            {
                lstStr.Add(nomeAlim + "\t" + barras);
            }

            //Grava em arquivo
            TxtFile.GravaListArquivoTXT(lstStr, nomeArq, paramGerais._mWindow);
        }

        // CalculaNumClientesDRPDRC
        public void CalculaNumClientesDRPDRC()
        {
            // nomeBarra
            string nomeBarra;

            // OBS: necessario p/ correto funcionamento do iterador, Next e etc.
            _circuit.SetActiveClass("load");

            int iter = _circuit.Loads.First;
            
            // work around
            for (int i = 1; i <= _circuit.Loads.Count; i++)
            {
                // go to the load "i"
                _circuit.Loads.idx = i;
            
            /*
            while ( iter != 0)
            { */
            
                // obtem nome da carga
                string loadName = _circuit.Loads.Name;

                // se load name comeca com LU = iluminacao publica
                if (loadName.Contains("ip"))
                {
                    // incrementa numero de clientes IP
                    _numClientesIP++;

                    continue;
                }
                int numFases;
                // numFases
                if (EngineConfig.OpenDSSengine)
                {
                    _DSSText.Command = "? Load." + loadName + ".Phases";
                    numFases = int.Parse(_DSSText.Result);
                }
                else 
                {
                    // TODO testar 
                    numFases = _circuit.Loads.Phases;

                    _DSSText.Command = "? Load." + loadName + ".Phases";
                    int numFases2 = int.Parse(_DSSText.Result);

                    if (numFases != numFases2)
                    {
                        // TODO tratar discrepancia
                        throw new Exception();
                    }
                }

                // nome Barra
                nomeBarra = GetLoadBusName(loadName);

                // verifica nivel tensao
                VerificaNivelTensaoBarra(nomeBarra, _circuit.Loads.kV, numFases);                
                
                // iter 
                iter = _circuit.Loads.Next;                              
            }

            // calcula numero de clientes faixa adequada
            CalculaNumClientesFaixaAdequada();
        }

        // calcula numclientes Faixa adequada
        private void CalculaNumClientesFaixaAdequada()
        {
            // ajusta numero de clientes excluindo pontos de IP
            _numClientesTotal -= _numClientesIP;

            //
            _numClientesOK = _numClientesTotal - _numClientesDRP - _numClientesDRC;
        }
    }
}
