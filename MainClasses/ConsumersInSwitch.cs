/* #if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif*/

using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Text;
using ExecutorOpenDSS.AuxClasses;

namespace ExecutorOpenDSS.MainClasses
{
    public class ConsumersInSwitch
    {
        private readonly string _banco = "GEO_SIGR_PERDAS"; //
        private readonly string _schemaDB = "GeoPerdas2023b.";

        private readonly string _host = @"PWNBS-PERTEC01\PTEC";

        private readonly string _arqSQL = "cargaAbaixoReligadorSQL.sql";
        private readonly string _nomeArqCargaMT = "cargaMT_Relig.csv";
        private readonly string _nomeArqCargaBT = "cargaBT_Relig.csv";
        private readonly string _nomeArqCargaIP = "cargaIP_Relig.csv";
        private readonly string _nomeArqGeradorMT = "geracaoMT_Relig.csv";
        private readonly string _nomeArqGeradorBT = "geracaoBT_Relig.csv";

        private ObjDSS _oDSS;
        private readonly GeneralParameters _par;
        private List<string> _lstReclosers;
        private List<string> _lstCons;

        private string _sqlCargaMT;
        private string _sqlCargaBT;
        private string _sqlIP;
        private string _sqlGeradorMT;
        private string _sqlGeradorBT;

        // public objects
        public StringBuilder _arqCargaMT;
        public StringBuilder _arqCargaBT;
        public StringBuilder _arqCargaIP;
        public StringBuilder _arqGeradorMT;
        public StringBuilder _arqGeradorBT;

        private static SqlConnectionStringBuilder _connBuilder;

        public ConsumersInSwitch(GeneralParameters par)
        {
            _par = par;

            TxtFile.SafeDelete(GetNomeArqCargaMT());
            TxtFile.SafeDelete(GetNomeArqCargaBT());
            TxtFile.SafeDelete(GetNomeArqCargaIP());
            TxtFile.SafeDelete(GetNomeArqGeradorMT());
            TxtFile.SafeDelete(GetNomeArqGeradorBT());

            // TODO
            _arqCargaMT = new StringBuilder();
            _arqCargaBT = new StringBuilder();
            _arqCargaIP = new StringBuilder();
            _arqGeradorMT = new StringBuilder();
            _arqGeradorBT = new StringBuilder();

            _connBuilder = new SqlConnectionStringBuilder();
            _connBuilder.DataSource = _host;
            _connBuilder.InitialCatalog = _banco;
            _connBuilder.IntegratedSecurity = false;
            //_connBuilder.UserID = "U_DBPERTEC01";
            //_connBuilder.Password = "294vd!@49s$$3208tD#SS";
            //_connBuilder.UserID = "U_DBPERTEC01";
            //_connBuilder.Password = "294vd!@49s$$3208tD#SS";
        }

        // 
        private bool Get_ConsumersInMeter(dynamic med)
        {
            //int numCli = med.NumSectionCustomers; //TODO 
            //int numCli2 = med.TotalCustomers;

            string[] pces = med.ZonePCE;

            _lstCons = new List<string>();

            foreach (string item in pces)
            {
                _lstCons.Add(item);
            }

            return true;
        }

        // sweep feeder returning all cliets below an equipment
        public bool GetClientsBelowSwitches(ObjDSS oDSS)
        {
            _oDSS = oDSS;

            // 1. Gets Energy Meters for reclosers
            bool ret = PutEnergyMetersOnReclosers();

            // 2. Solves again because new energyMeters
            _oDSS._DSSObj.Text.Command = "Solve";

            _lstReclosers = Get_ClosedSwitches();

            foreach (string recloser in _lstReclosers)
            {
                // 3.0 Gets losses below recloser 
                dynamic med = GetMeter(recloser);

                // 
                Get_ConsumersInMeter(med);

                //
                ConstructSQLQuery(recloser);

                // 4. Run queries
                ret = RunQueriesDB();

            }

            return ret;
        }

        dynamic GetMeter(string recloser)
        {
            //Obs: need this
            _oDSS._DSSObj.ActiveCircuit.SetActiveClass("energymeter");

            dynamic med = _oDSS._DSSObj.ActiveCircuit.Meters;

            //search recloser meter
            int iEM = med.First;

            while (iEM != 0)
            {
                if (med.Name.Equals(recloser))
                {
                    break;
                }
                iEM = med.Next;
            }
            return med;
        }

        private bool RunQueriesDB()
        {
            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

                /*
                if (!_sqlCargaMT.Equals(""))
                {
                    RunQuery_MVClients(conn);
                }
                if (!_sqlGeradorMT.Equals(""))
                {
                    RunQuery_MVGenerators(conn);
                }
                if (!_sqlCargaBT.Equals(""))
                {
                    RunQuery_LVClients(conn);
                }
                if (!_sqlGeradorBT.Equals(""))
                {
                    RunQuery_LVGenerators(conn);
                }
                if (!_sqlIP.Equals(""))
                {
                    RunQuery_Lamps(conn);
                }*/
                //fecha conexao
                conn.Close();
            }
            return true;
        }

        private void RunQuery_LVGenerators(SqlConnection conn)
        {
            using (SqlCommand command = conn.CreateCommand())
            {
                command.CommandText = _sqlGeradorBT;

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows) { return; }

                    while (rs.Read())
                    {
                        string nCons = rs["nUGBT"].ToString();
                        if (nCons.Equals("0")) { continue; }

                        string linha = rs["Relig"].ToString() + "\t";
                        linha += nCons + "\t";
                        linha += rs["S01_MWh"].ToString() + "\t";
                        linha += rs["S02_MWh"].ToString() + "\t";
                        linha += rs["S03_MWh"].ToString() + "\t";
                        linha += rs["S04_MWh"].ToString() + "\t";
                        linha += rs["S05_MWh"].ToString() + "\t";
                        linha += rs["S06_MWh"].ToString() + "\t";
                        linha += rs["S07_MWh"].ToString() + "\t";
                        linha += rs["S08_MWh"].ToString() + "\t";
                        linha += rs["S09_MWh"].ToString() + "\t";
                        linha += rs["S10_MWh"].ToString() + "\t";
                        linha += rs["S11_MWh"].ToString() + "\t";
                        linha += rs["S12_MWh"].ToString() + Environment.NewLine;

                        _arqGeradorBT.Append(linha);
                    }
                }
            }
        }

        private string GetNomeArqSQL()
        {
            return _par._parGUI._pathRecursosPerm + _arqSQL;
        }

        private string GetNomeArqCargaMT()
        {
            return _par._parGUI._pathRecursosPerm + _nomeArqCargaMT;
        }
        private string GetNomeArqCargaBT()
        {
            return _par._parGUI._pathRecursosPerm + _nomeArqCargaBT;
        }
        private string GetNomeArqCargaIP()
        {
            return _par._parGUI._pathRecursosPerm + _nomeArqCargaIP;
        }

        public void GravaCargaMTBTIP_CSV()
        {
            // grava em arquivo
            TxtFile.GravaEmArquivo2(_arqCargaMT.ToString(), GetNomeArqCargaMT(), _par._mWindow);
            TxtFile.GravaEmArquivo2(_arqCargaBT.ToString(), GetNomeArqCargaBT(), _par._mWindow);
            TxtFile.GravaEmArquivo2(_arqCargaIP.ToString(), GetNomeArqCargaIP(), _par._mWindow);
            TxtFile.GravaEmArquivo2(_arqGeradorMT.ToString(), GetNomeArqGeradorMT(), _par._mWindow);
            TxtFile.GravaEmArquivo2(_arqGeradorBT.ToString(), GetNomeArqGeradorBT(), _par._mWindow);
        }

        private string GetNomeArqGeradorBT()
        {
            return _par._parGUI._pathRecursosPerm + _nomeArqGeradorBT;
        }

        private string GetNomeArqGeradorMT()
        {
            return _par._parGUI._pathRecursosPerm + _nomeArqGeradorMT;
        }

        private List<string> Get_ClosedSwitches()
        {
            _lstReclosers = new List<string>();

            // Obs: avoids collaterals effects
            _oDSS._DSSObj.ActiveCircuit.SetActiveClass("line");

            // as recloser are still modelled as switches, its necessary do look in Lines
            dynamic lines = _oDSS._DSSObj.ActiveCircuit.Lines;

            int iLines = lines.First;

            while (iLines != 0)
            {
                string nome = lines.Name;

                //TODO // lines.Name.Contains("ctrr") &&
                // ctrr means recloser in Cemig feeders & is 3 phase & TODO is closed 
                if ((lines.Phases == 3) && (lines.IsSwitch))
                {
                    _lstReclosers.Add(nome);
                }
                iLines = lines.Next;
            }
            return _lstReclosers;
        }

        private bool PutEnergyMetersOnReclosers()
        {
            // CTRR44079 CTRR44105 CTRR45585

            // TODO
            //Reclosers recloser = _oDSS._DSSObj.ActiveCircuit.Reclosers;
            _lstReclosers = Get_ClosedSwitches();

            // Put EnergyMeters
            foreach (string recloser in _lstReclosers)
            {
                _oDSS._DSSObj.Text.Command = "New energymeter." + recloser + " element=Line." + recloser + ",terminal=1";

            }
            // DEBUG
            //Meters meters = _oDSS._DSSObj.ActiveCircuit.Meters;

            return true;
        }

        // constructs SQLs queries for MV loads, BT loads and public lamps
        private bool ConstructSQLQuery(string recloser)
        {
            _sqlCargaMT = "";
            _sqlCargaBT = "";
            _sqlIP = "";
            _sqlGeradorMT = "";
            _sqlGeradorBT = "";

            string lstCons = CemigFeeders.AddAposAndCommasForSQL(_lstCons);

            lstCons = lstCons.Replace("Load.bt_", "");
            lstCons = lstCons.Replace("Load.mt_", "");
            lstCons = lstCons.Replace("Generator.", "");
            lstCons = lstCons.Replace("Generator.", "");
            lstCons = lstCons.Replace("_m1", "");
            lstCons = lstCons.Replace("_m2", "");

            _sqlCargaMT = "select count(CodConsMT) as nConsMT, sum(EnerMedid12_MWh)as Consumo_MWh " +
                 "from " + _schemaDB + "StoredCargaMT " +
                 "where CodConsMT in (" + lstCons + ")" + Environment.NewLine;

            _sqlCargaBT = "select count(CodConsBT) as nConsBT, sum(EnerMedid12_MWh)as Consumo_MWh " +
                 "from " + _schemaDB + "StoredCargaBT " +
                 "where CodConsBT in (" + lstCons + ")" + Environment.NewLine;

            _sqlIP = "select count(CodConsBT) as nConsBT, sum(EnerMedid12_MWh)as Consumo_MWh " +
                 "from " + _schemaDB + "StoredCargaBT " +
                 "where CodConsBT in (" + lstCons + ")" + Environment.NewLine;

            _sqlGeradorBT = "select count(CodGerBT) as nGerBT, sum(EnerMedid12_MWh)as Consumo_MWh " +
                 "from " + _schemaDB + "StoredGeradorBT " +
                 "where CodGerBT in (" + lstCons + ")" + Environment.NewLine;

            _sqlGeradorMT = "select count(CodGerMT) as nGerMT, sum(EnerMedid12_MWh)as Consumo_MWh " +
                 "from " + _schemaDB + "StoredGeradorMT " +
                 "where CodGerMT in (" + lstCons + ")" + Environment.NewLine;

            return true;
        }
    }
}
