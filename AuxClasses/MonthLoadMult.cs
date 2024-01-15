using System.Collections.Generic;
using System.IO;

namespace ExecutorOpenDSS.AuxClasses
{
    // This class holds one loadMult multiplier (double) per month per feeder
    // 
    public class MonthLoadMult
    {
        // map MONTH energia injetada X alim
        public Dictionary<int, Dictionary<string, double>> _mapAlimLoadMult = new Dictionary<int, Dictionary<string, double>>();
        private readonly GeneralParameters _paramGerais;

        // construtor
        public MonthLoadMult(GeneralParameters par)
        {
            _paramGerais = par;

            CarregaMapAjusteLoadMult();
        }

        public void CarregaMapAjusteLoadMult()
        {
             if (_paramGerais._parGUI._tipoFluxo.Equals("Anual"))
            {
                // carrega requisitos para todos os meses
                for (int mes = 1; mes < 13; mes++)
                {
                    //obtem nome ajuste compl
                    string arqAjusteCompl = _paramGerais.GetNomeArqAjuste(mes);

                    CarregaMapAjusteLoadMult_Pvt(arqAjusteCompl, mes);
                }
            }
            else
            {
                // get Mes
                int mes = _paramGerais._parGUI.GetMes();

                //obtem nome ajuste compl
                string arqAjusteCompl = _paramGerais.GetNomeArqAjuste(mes);

                CarregaMapAjusteLoadMult_Pvt(arqAjusteCompl, mes);
            }
        }

        // 
        private void CarregaMapAjusteLoadMult_Pvt(string arq, int mes)
        {
            try
            {
                //adiciona na variavel da classe
                _mapAlimLoadMult.Add(mes, XLSXFile.XLSX2Dictionary(arq));
            }
            catch (FileNotFoundException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return;
            }
        }

        // escreve map alim loadMult com novos loadMults
        public void AtualizaMapAlimLoadMult(string alim, double loadMult, int mes)
        {
            //atualiza alimentador
            _mapAlimLoadMult[mes][alim] = loadMult;
        }

        //
        public double GetLoadMult()
        {
            int mes = _paramGerais._parGUI.GetMes();

            string alimTmp = _paramGerais.GetNomeAlimAtual();

            if (_mapAlimLoadMult[mes].ContainsKey(alimTmp))
            {
                return _mapAlimLoadMult[mes][alimTmp];
            }
            else 
            {
                _paramGerais._mWindow.ExibeMsgDisplay(_paramGerais.GetNomeAlimAtual() + ": Alimentador não encontrado no arquivo de ajuste");

                // retorna loadMult Default
                return 1.0;
            }
        }
    }
}
