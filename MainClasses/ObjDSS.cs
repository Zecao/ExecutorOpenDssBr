/* #if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif */
using ExecutorOpenDSS.Engine;

namespace ExecutorOpenDSS.MainClasses
{
    public class ObjDSS
    {
        public dynamic _DSSObj;
        public GeneralParameters _paramGerais;

        public ObjDSS(GeneralParameters par)
        {
            //
            _paramGerais = par;

            if (EngineConfig.OpenDSSengine)
            {
                _DSSObj = new OpenDSSengine.DSS();
            }
            else
            {
                //Inicializa o servidor COM
                _DSSObj = new dss_sharp.DSS();
            }


            // Inicializa servidor COM
            _DSSObj.Start(0);

            //
            try
            {
                _DSSObj.DataPath = par.GetDataPathAlimOpenDSS();
            }
            catch (dss_sharp.DSSException e)
            {
                _paramGerais._mWindow.ExibeMsgDisplay(e.Message);
                return;
            }

            _DSSObj.AllowForms = _paramGerais._parGUI._allowForms;
            /* TODO dss_sharp.DSSException: 'Cannot activate output with no console available! If you want to use a message output callback, register it before enabling AllowForms.'
            // configuracoes gerais OpenDSS
            _DSSObj.AllowForms = _paramGerais._parGUI._allowForms;
            */
        }

        // retorna o DSSCircuit 
        public dynamic GetActiveCircuit()
        {
            return _DSSObj.ActiveCircuit;
        }
    }
}
