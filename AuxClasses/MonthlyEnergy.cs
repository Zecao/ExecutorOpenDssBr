using System.Collections.Generic;

namespace ExecutorOpenDSS.AuxClasses
{
    class MonthlyEnergy
    {
        public readonly GeneralParameters _paramGerais;

        // map (alim_mes,energia)       
        public Dictionary<string, double> _mapDadosEnergiaMes = new Dictionary<string, double>();

        //booleanas para o controle se arquivo Excel foi carregado       
        public bool _reqEnergiaMesCarregado = false;

        public MonthlyEnergy(GeneralParameters par)
        {
            _paramGerais = par;
        }

        // Load map from Excel file with month energy measurements 
        public void CarregaMapEnergiaMesAlim()
        {
            _paramGerais._mWindow.ExibeMsgDisplay("Carregando arquivo de requisitos de energia...");

            CarregaMapEnergiaMesAlimPvt();

            _reqEnergiaMesCarregado = true;
        }

        // Load map from Excel file with month energy measurements 
        internal void CarregaMapEnergiaMesAlimPvt()
        {
            string nomeArqEnergiaCompl = _paramGerais.GetNomeArqEnergia();

            string[,] energiaMes = XLSXFile.LeTudo(nomeArqEnergiaCompl);

            // para cada alim
            // OBS: nAlim comeca em 1 por causa do cabecalho
            for (int nAlim = 1; nAlim < energiaMes.GetLength(0); nAlim++)
            {
                // alim 
                string alim = energiaMes[nAlim, 0];

                // DEBUG
                /*
                if (alim.Equals("LAVD18"))
                {
                    int debug = 0;
                }*/

                // carrega requisitos para todos os meses
                for (int mes = 1; mes < 13; mes++)
                {
                    double energia = double.Parse(energiaMes[nAlim, mes]);

                    // chave tmp
                    string chave = mes.ToString() + "_" + alim;

                    //adiciona na variavel da classe
                    _mapDadosEnergiaMes.Add(chave, energia);
                }
            }

        }

        // obtem a referencia de energia para um dado mes
        public double GetRefEnergia(string nomeAlim, int mes)
        {
            // chave tmp
            string chave = mes.ToString() + "_" + nomeAlim;

            //verifica se nome 
            if (_mapDadosEnergiaMes.ContainsKey(chave))
            {
                return _mapDadosEnergiaMes[chave];
            }

            return double.NaN;
        }
    }
}
