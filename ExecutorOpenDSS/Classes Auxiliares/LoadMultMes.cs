using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes_Auxiliares
{
    class LoadMultMes
    {
        // map energia injetada X alim
        public Dictionary<int, Dictionary<string, double>> _mapAlimLoadMult = new Dictionary<int, Dictionary<string, double>>();

        // escreve map alim loadMult com novos loadMults
        public void AtualizaMapAlimLoadMult(string alim, double loadMult, int mes)
        {
            //atualiza alimentador
            _mapAlimLoadMult[mes][alim] = loadMult;
        }

        //
        internal void Add(int mes, Dictionary<string, double> mapLMMes)
        {
            _mapAlimLoadMult.Add(mes, mapLMMes);
        }
    }
}
