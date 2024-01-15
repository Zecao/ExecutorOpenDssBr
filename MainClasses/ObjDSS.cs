//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

namespace ExecutorOpenDSS.MainClasses
{
    public class ObjDSS
    {
        public DSS _DSSObj;
        public GeneralParameters _paramGerais;

        public ObjDSS(GeneralParameters par)
        {
            //
            _paramGerais = par;

            //Inicializa o servidor COM
            _DSSObj = new DSS();

            // Inicializa servidor COM
            _DSSObj.Start(0);

            //
            _DSSObj.DataPath = par.GetDataPathAlimOpenDSS();

            /* TODO dss_sharp.DSSException: 'Cannot activate output with no console available! If you want to use a message output callback, register it before enabling AllowForms.'
            // configuracoes gerais OpenDSS
            _DSSObj.AllowForms = _paramGerais._parGUI._allowForms;
            */
        }

        // retorna o DSSCircuit 
        public Circuit GetActiveCircuit()
        {
            return _DSSObj.ActiveCircuit;
        }
    }
}
