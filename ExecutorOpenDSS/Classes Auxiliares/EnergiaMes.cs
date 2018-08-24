using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes_Auxiliares
{
    class EnergiaMes
    {
        // map energia injetada X alim       
        public Dictionary<int, Dictionary<string, double>> _mapDadosEnergiaMes = new Dictionary<int, Dictionary<string, double>>();

        // obtem a referencia de energia para um dado mes
        public double getRefEnergia(string nomeAlim, int mes)
        {
            //variavel de retorno
            Dictionary<string, double> mapEnergiaMes = _mapDadosEnergiaMes[mes];

            //verifica se nome 
            if (mapEnergiaMes.ContainsKey(nomeAlim))
            {
                return mapEnergiaMes[nomeAlim];
            }
            return double.NaN;
        }

        //
        internal void Add(int mes, Dictionary<string, double> mapEnergiaMes)
        {
            // TODO if not contains
            _mapDadosEnergiaMes.Add(mes, mapEnergiaMes);
        }

        internal Dictionary<string, double> GetMapEnergiaMesPvt(int mes)
        {
            Dictionary<string, double> mapLM = _mapDadosEnergiaMes[mes];
            return mapLM;
        }

    }
}
