using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutorOpenDSS.Classes_Auxiliares
{
    public class ParamAvancados
    {
        public bool _otimizaPUSaidaSE = false; // otimizacao PU saida SE
        public bool _calcDRPDRC = false;
        public bool _calcTensaoBarTrafo = false;
        public bool _verifCargaIsolada = false;

        //construtor
        public ParamAvancados()
        { 
            
        }
    }
}
