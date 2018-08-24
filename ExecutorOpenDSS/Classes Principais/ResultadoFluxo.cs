using ExecutorOpenDSS.Classes_Auxiliares;
using ExecutorOpenDSS.Classes_Principais;
using OpenDSSengine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes
{ 
    class ResultadoFluxo
    {
        // membros
        public bool _convergiuBool = true;
        public MyEnergyMeter _energyMeter = new MyEnergyMeter();

        //Construtor por copia 
        public ResultadoFluxo(ResultadoFluxo rf)
        {
            _convergiuBool = rf._convergiuBool;
            _energyMeter = rf._energyMeter;
        }

        //Construtor
        public ResultadoFluxo()
        {

        }

        // formata resultado do fluco para console
        public string getResultadoFluxoToConsole(string tensao, string nomeAlim)
        {
            string tmp = "Alim\\PU SaidaSE:\\Energia(KWh)\\Perdas(KWh)\\Perdas(%)" +
                "\n" + nomeAlim +
                " \t" + tensao +
                " \t" + _energyMeter.kWh.ToString("0.00") +
                " \t" + _energyMeter.LossesKWh.ToString("0.00") +
                " \t" + CalcPercentualPerdas() + "%";
            return tmp;
        }

        //Tradução da função calculaResultadoFluxoMensal
        //O resultado no fluxo mensal eh armazenado na variavel da classe
        public void CalculaResultadoFluxoMensal(ResultadoFluxo perdasDU, ResultadoFluxo perdasSA, ResultadoFluxo perdasDO, ParamGeraisDSS paramGerais, MainWindow janela)
        {
            // calcula geracao e perdas maximas entre os 3 dias tipicos
            CalcGeracaoEPerdasMax(perdasDU, perdasSA, perdasDO);

            //Obtem mes
            int mes = paramGerais._parGUI._mesNum;

            // setMes
            _energyMeter.SetMes(mes, paramGerais._parGUI._modoAnual );

            // cria curva de carga dados: numero de dias do mes e matriz de consumo em PU
            Dictionary<string, int> numTipoDiasMes = paramGerais._objTipoDeDiasDoMes._qntTipoDiasMes[mes];
            
            // DIAS UTEIS
            perdasDU.MultiplicaEnergia(numTipoDiasMes["DU"]);

            // multiplica pelo Num dias
            perdasSA.MultiplicaEnergia(numTipoDiasMes["SA"]);

            // multiplica pelo Num dias
            perdasDO.MultiplicaEnergia(numTipoDiasMes["DO"]);

            // perdas energia
            SomaEnergiaDiasTipicos(perdasDU, perdasSA, perdasDO);

            // cria string com o formato de saida das perdas
            string conteudo = _energyMeter.formataResultado(paramGerais.getNomeAlimAtual());

            // se modo otimiza nao grava perdas arquivo
            if ( !paramGerais._parGUI._otmPorEnergia )
            {
                // grava perdas alimentador em arquivo
                ArqManip.GravaEmArquivo(conteudo, paramGerais._parGUI._pathRaizGUI + paramGerais._arquivoResPerdasMensal, janela);
            }
        }

        // Soma energia calculada dias tipicos 
        private void SomaEnergiaDiasTipicos(ResultadoFluxo perdasDU, ResultadoFluxo perdasSA, ResultadoFluxo perdasDO)
        {
            //
            List<double> perdasDUEnergia = perdasDU.getEnergiaEPerdasEnergiaSeg();

            List<double> perdasSAEnergia = perdasSA.getEnergiaEPerdasEnergiaSeg();

            List<double> perdasDOEnergia = perdasDO.getEnergiaEPerdasEnergiaSeg();

            // PREENCHE
            List<double> injecaoEPerdasEnergia = ListasOp.Soma(ListasOp.Soma(perdasDUEnergia, perdasSAEnergia), perdasDOEnergia);

            // adiciona a variave de classe
            _energyMeter.AddEnergia(injecaoEPerdasEnergia);
        }

        // TODO REFATORAR
        private void MultiplicaEnergia(int numDias)
        {
            //
            List<double> perdasEnergia = getEnergiaEPerdasEnergiaSeg();

            // multiplica pelo Num dias
            perdasEnergia = ListasOp.Multiplica(perdasEnergia, numDias);

            // perdasDU MAX - var temporaria
            List<double> perdasMax = getGeracaoEPerdaPotencia();

            // atualiza objeto com novos valores calculados.
            _energyMeter = new MyEnergyMeter();
            _energyMeter.MaxkW = perdasMax[0];
            _energyMeter.MaxkWLosses = perdasMax[1];

            _energyMeter.AddEnergia(perdasEnergia);
        }

        // calcula geracao e perdas maximas entre os 3 dias tipicos
        private void CalcGeracaoEPerdasMax(ResultadoFluxo perdasDU, ResultadoFluxo perdasSA, ResultadoFluxo perdasDO)
        {
            // DIAS UTEIS perdasDU MAX
            List<double> perdasDUMax = perdasDU.getGeracaoEPerdaPotencia();

            // SABADO MAX
            List<double> perdasSAMax = perdasSA.getGeracaoEPerdaPotencia();

            // DOMINGO MAX
            List<double> perdasDOMax = perdasDO.getGeracaoEPerdaPotencia();

            // geracao maxima
            double geracaoMax = Math.Max(Math.Max(perdasDUMax[0], perdasSAMax[0]), perdasDOMax[0]);

            // perdas maxima
            double perdasMax = Math.Max(Math.Max(perdasDUMax[1], perdasSAMax[1]), perdasDOMax[1]);

            // atribui variavel da classe 
            _energyMeter.MaxkW = geracaoMax;
            _energyMeter.MaxkWLosses = perdasMax;
        }

        // define alim como nao convergiu
        public void DefineComoNaoConvergiu()
        {
            _convergiuBool = false;
        }

        //get Energia
        public double getEnergia()
        {
            return _energyMeter.kWh;
        }

        // get Geracao
        public double getMaxKW()
        {
            return _energyMeter.MaxkW;
        }

        // TODO MUDAR RETORNO
        // TODO FIX gera excecao quando interrompe fluxo 24hrs customizado!
        // get Energia e perdas energia por segmento
        public List<double> getEnergiaEPerdasEnergiaSeg()
        {
            List<double> ret = new List<double>();
            ret.Add(_energyMeter.kWh);
            ret.Add(_energyMeter.kvarh);
            ret.Add(_energyMeter.LossesKWh);
            ret.Add(_energyMeter.TransformerLosses);
            ret.Add(_energyMeter.MTLineLosses);
            ret.Add(_energyMeter.BTLineLosses);
            ret.Add(_energyMeter.lineLossesPosMode);
            ret.Add(_energyMeter.lineLossesZeroMode);
            ret.Add(_energyMeter.NoLoadLosseskWh);
            ret.Add(_energyMeter.MTEnergy);
            ret.Add(_energyMeter.BTEnergy);

            return ret;
        }                      

        // get Geracao e Perda de Potencia
        public List<double> getGeracaoEPerdaPotencia()
        {
            // TODO refactory
            List<double> ret = new List<double>();
            ret.Add(_energyMeter.MaxkW);
            ret.Add(_energyMeter.MaxkWLosses);

            return ret;
        }

        // verifica se excedeu a geracao maxima
        private bool ExcedeuGeracaoMaxima()
        {
            //MAXENERGIA alim igual 40.000.000 kWh/mes 
            double MAXENERGIA = 40000000;

            // condicao de erro no fluxo diario 
            if (Math.Abs(getEnergia()) > MAXENERGIA)
            {
                return true;
            }
            return false;
        }

        // verifica se excedeu a geracao maxima
        private bool ExcedeuRequisitoMaximo()
        {
            //MAXREQUISITO alim igual 40.000 kWh 
            double MAXREQUISITO = 40000;

            // condicao de erro no fluxo diario 
            if (Math.Abs(getMaxKW()) > MAXREQUISITO)
            {
                return true;
            }
            return false;
        }

        // Calcula percentual de perdas
        public string CalcPercentualPerdas()
        {
            double energia = _energyMeter.kWh;
            double perdas = _energyMeter.LossesKWh;
            double percentual = perdas * 100 / energia;
            return percentual.ToString("0.000");
        }

        //Tradução da função getPerdasAlim
        public void GetPerdasAlim(Circuit DSSCircuit)
        {
            // TODO
            //string[] registersNames = DSSCircuit.Meters.RegisterNames;

            _energyMeter.MaxkW = DSSCircuit.Meters.RegisterValues[2];
            _energyMeter.MaxkWLosses = DSSCircuit.Meters.RegisterValues[14];

            // obtem medidores 
            _energyMeter.kWh = DSSCircuit.Meters.RegisterValues[0];
            _energyMeter.kvarh = DSSCircuit.Meters.RegisterValues[1];

            // aparentemente o reg 13 nao ontem perdas no trafo.
            _energyMeter.LossesKWh = DSSCircuit.Meters.RegisterValues[12];

            // perdas em transformadores
            double LineLosses = DSSCircuit.Meters.RegisterValues[22];
            _energyMeter.TransformerLosses = DSSCircuit.Meters.RegisterValues[23];

            // perdas seq. pos (25) e zero (26) e circuitos bi e monofasicos (28)
            _energyMeter.lineLossesPosMode = DSSCircuit.Meters.RegisterValues[24];
            _energyMeter.lineLossesZeroMode = DSSCircuit.Meters.RegisterValues[25];
            // lineLossesoneTwoPhase = DSSCircuit.Meters.RegisterValues[27];

            // perdas MT e BT
            // OBS: os registradores 33 e 34 mostram as perdas totais em cada segmento de tensao
            // Reg 33 = 13.8 kV Losses / Reg 34 = 0.22 kV Losses, 
            // sendo que queremos somente as perdas nas linhas por segmento de tensao
            // Reg 40 = 13.8 kV Line Loss / Reg 41 = 0.22 kV Line Loss / Reg 42 = 0.127 kV Line Loss
            _energyMeter.MTLineLosses  = DSSCircuit.Meters.RegisterValues[39];
            double BTLosses220 = DSSCircuit.Meters.RegisterValues[40];
            double BTLosses127 = DSSCircuit.Meters.RegisterValues[41];
            _energyMeter.BTLineLosses = BTLosses220 + BTLosses127;

            // Reg 19 NoLoadLosseskWh
            _energyMeter.NoLoadLosseskWh = DSSCircuit.Meters.RegisterValues[18];

            _energyMeter.MTEnergy = DSSCircuit.Meters.RegisterValues[60];
            _energyMeter.BTEnergy = DSSCircuit.Meters.RegisterValues[61];

            

            // verifica convergencia
            verificaConvergencia();
        }

        internal void verificaConvergencia()
        {
            // verifica se excedeu a geracao maxima
            if ( ExcedeuGeracaoMaxima() || ExcedeuRequisitoMaximo() || PerdaMaiorQueInjecao() )
            {
                DefineComoNaoConvergiu();
            }
        }

        // verifica se perda maior que a injecao de energia
        private bool PerdaMaiorQueInjecao()
        {
            
            //TODO verificar a existencia de geradores e adicionar a MaxkW do gerador ao do EM.
            return false;
            
            /*
            if ( _energyMeter.MaxkWLosses > _energyMeter.MaxkW)
                return true;
            else 
                return false;
            */
        }

        internal void setEnergia(double energiaMes)
        {
            _energyMeter.kWh = energiaMes;
        }

        //calcula resultado ano
        internal void CalculaResAno(List<ResultadoFluxo> lstResultadoFluxo, string alim, string arquivo, MainWindow jan)
        {
            // obtem 1mês
            ResultadoFluxo res1 = lstResultadoFluxo.First();

            //usa variavel da classe para armazenar a soma
            _energyMeter = res1._energyMeter;
                 
            //remove     
            lstResultadoFluxo.Remove(res1);

            // obtem medidores do 2mes em diante e soma com o 1mes
            foreach (ResultadoFluxo res in lstResultadoFluxo) 
            {
                //medidor 
                MyEnergyMeter emMes = res._energyMeter;
            
                //soma
                _energyMeter.Soma(emMes);
            }

            // cria string com o formato de saida das perdas
            string conteudo = _energyMeter.formataResultado(alim);

            // grava perdas alimentador em arquivo 
            ArqManip.GravaEmArquivo(conteudo, arquivo, jan);
        }
    }     
}

