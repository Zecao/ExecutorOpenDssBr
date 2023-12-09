using System.Collections.Generic;

namespace ExecutorOpenDSS.AuxClasses
{
    public class FeederMetering
    {
        public readonly GeneralParameters _paramGerais;

        //booleanas para o controle se arquivo Excel foi carregado
        public bool _reqDemandaMaxCarregado = false;

        // map demanda max X alim
        internal Dictionary<string, double> _mapDadosDemanda;

        // map energia injetada X alim
        internal MonthlyEnergy _reqEnergiaMes;

        // Load MultMes
        internal MonthLoadMult _reqLoadMultMes;

        //contrutor
        public FeederMetering(GeneralParameters par)
        {
            _paramGerais = par;

            _reqEnergiaMes = new MonthlyEnergy(par);
        }

        public void CarregaDados()
        {
            // Arquivo de loadMUlt sempre sera carregado
            _paramGerais._mWindow.ExibeMsgDisplay("Carregando arquivo de LoadMults...");

            // Carrega map com os valores dos loadMult por alimentador
            _reqLoadMultMes = new MonthLoadMult(_paramGerais);

            // carrega arquivo de requisito uma unica vez
            if ((_paramGerais._parGUI._otmPorEnergia) && (!_reqEnergiaMes._reqEnergiaMesCarregado))
            {
                // Carrega map com valores de energia mes do alimentador
                _reqEnergiaMes.CarregaMapEnergiaMesAlim();
            }

            // carrega maps conforme necessidade
            if ((_paramGerais._parGUI._otmPorDemMax) && (!_reqDemandaMaxCarregado))
            {
                // Carrega map com valores de Demanda mes do alimentador
                _paramGerais._medAlim.CarregaMapDemandaEnergiaMesAlim();
            }
        }

        // 
        internal void CarregaMapDemandaEnergiaMesAlim()
        {
            _paramGerais._mWindow.ExibeMsgDisplay("Carregando arquivo de LoadMults demanda...");

            // Carrega map com valores de demanda maxima do alimentador
            string nomeArquivoCompleto = _paramGerais.GetNomeArqDemandaMaxAlim();

            _mapDadosDemanda = XLSXFile.XLSX2Dictionary(nomeArquivoCompleto);

            _reqDemandaMaxCarregado = true;
        }

        //
        internal void AtualizaMapAlimLoadMult(double loadMult)
        {
            _reqLoadMultMes.AtualizaMapAlimLoadMult(_paramGerais.GetNomeAlimAtual(), loadMult, _paramGerais._parGUI.GetMes());
        }
    }
}
