//#define ENGINE
#if ENGINE
using OpenDSSengine;
#else
using dss_sharp;
#endif

using ExecutorOpenDSS.AuxClasses;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ExecutorOpenDSS.MainClasses
{
    public class RecloserTracer
    {
        // servidor SGBD
        private readonly string _banco = "GEOPERDAS_2023"; // GEOPERDAS_2022 GEOPERDAS_2021 GEOPERDAS_2020 GEOPERDAS_2019 
        private readonly string _schemaDB = "geo2023.";
        private readonly string _dataSource = @"PWNBS-PERTEC01\PTEC"; //@"sa-corp-sql0";

        private readonly string _arqSQL = "cargaAbaixoReligadorSQL.sql";
        private readonly string _nomeArqCargaMT = "cargaMT_Relig.csv";
        private readonly string _nomeArqCargaBT = "cargaBT_Relig.csv";
        private readonly string _nomeArqCargaIP = "cargaIP_Relig.csv";
        private readonly string _nomeArqGeradorMT = "geracaoMT_Relig.csv";
        private readonly string _nomeArqGeradorBT = "geracaoBT_Relig.csv";

        private ObjDSS _oDSS;
        private List<string> _lstSegmentosMT;
        private List<string> _lstSegmentosBT;
        private List<string> _lstRamais;
        private string _sqlCargaMT;
        private string _sqlCargaBT;
        private string _sqlIP;
        private string _sqlGeradorMT;
        private string _sqlGeradorBT;
        private readonly GeneralParameters _par;
        private List<string> _lstReclosers;
        private StringBuilder _sqlQueries;

        // public objects
        public StringBuilder _arqCargaMT;
        public StringBuilder _arqCargaBT;
        public StringBuilder _arqCargaIP;
        public StringBuilder _arqGeradorMT;
        public StringBuilder _arqGeradorBT;

        private static SqlConnectionStringBuilder _connBuilder;

        public RecloserTracer(GeneralParameters par)
        {
            _par = par;

            TxtFile.SafeDelete(GetNomeArqCargaMT());
            TxtFile.SafeDelete(GetNomeArqCargaBT());
            TxtFile.SafeDelete(GetNomeArqCargaIP());
            TxtFile.SafeDelete(GetNomeArqGeradorMT());
            TxtFile.SafeDelete(GetNomeArqGeradorBT());

            _arqCargaMT = new StringBuilder();
            _arqCargaBT = new StringBuilder();
            _arqCargaIP = new StringBuilder();
            _arqGeradorMT = new StringBuilder();
            _arqGeradorBT = new StringBuilder();

            _connBuilder = new SqlConnectionStringBuilder();
            _connBuilder.DataSource = _dataSource;
            _connBuilder.InitialCatalog = _banco;
            _connBuilder.IntegratedSecurity = false;
            _connBuilder.UserID = "U_DBPERTEC01";
            _connBuilder.Password = "294vd!@49s$$3208tD#SS";
        }

        // sweep feeder returning all buses below an equipment
        public bool TraceAllReclosers(ObjDSS oDSS)
        {
            _oDSS = oDSS;

            // 1. Gets Energy Meters for reclosers
            bool ret = PutEnergyMetersOnReclosers();

            // 2. Solves again because new energyMeters
            _oDSS._DSSText.Command = "Solve";

            // 3. for each recloser
            foreach (string recloser in _lstReclosers)
            {
                // 3.0 Gets losses below recloser 

                // 3.1 Gets MV and LV Lines CodIDs below an recloser (or any other element).
                ret = Get_MVandLVLines_CodIDs(recloser);

                string recloserShort = recloser.Replace("ctr", "");

                // 3.2. Constructs the SQL query 
                ret = ConstructSQLQuery(recloserShort);

                // TODO flag to turn this function on
                // 4. Save queries results in csv
                GravaSQLQueries();

                // 4. Run queries
                ret = RunQueriesDB();
            }

            return ret;
        }

        private bool RunQueriesDB()
        {
            using (SqlConnection conn = new SqlConnection(_connBuilder.ToString()))
            {
                // abre conexao 
                conn.Open();

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
                }
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

        private void RunQuery_MVGenerators(SqlConnection conn)
        {
            using (SqlCommand command = conn.CreateCommand())
            {
                command.CommandText = _sqlGeradorMT;

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows) { return; }

                    while (rs.Read())
                    {
                        string nCons = rs["nUGMT"].ToString();
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

                        _arqGeradorMT.Append(linha);
                    }
                }
            }
        }

        private void RunQuery_Lamps(SqlConnection conn)
        {
            using (SqlCommand command = conn.CreateCommand())
            {
                command.CommandText = _sqlIP;

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows) { return; }

                    while (rs.Read())
                    {
                        string nCons = rs["nIP"].ToString();
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

                        _arqCargaIP.Append(linha);
                    }
                }
            }
        }

        private void RunQuery_LVClients(SqlConnection conn)
        {
            using (SqlCommand command = conn.CreateCommand())
            {
                command.CommandText = _sqlCargaBT;

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows) { return; }

                    while (rs.Read())
                    {
                        string nCons = rs["nUCBT"].ToString();
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

                        _arqCargaBT.Append(linha);
                    }
                }
            }
        }

        private void RunQuery_MVClients(SqlConnection conn)
        {
            using (SqlCommand command = conn.CreateCommand())
            {
                command.CommandText = _sqlCargaMT;
                //TODO 
                /*
                command.Parameters.AddWithValue("@codbase", _par._codBase);
                command.Parameters.AddWithValue("@CodAlim", _par._alim);*/

                using (var rs = command.ExecuteReader())
                {
                    // verifica ocorrencia de elemento no banco
                    if (!rs.HasRows) { return; }

                    // nUCMT',sum(EnerMedid01_MWh),sum(EnerMedid02_MWh),sum(EnerMedid03_MWh),sum(EnerMedid04_MWh),sum(EnerMedid05_MWh),sum(EnerMedid06_MWh),sum(EnerMedid07_MWh),sum(EnerMedid08_MWh),sum(EnerMedid09_MWh),sum(EnerMedid10_MWh),sum(EnerMedid11_MWh),sum(EnerMedid12_MWh)
                    while (rs.Read())
                    {
                        string nCons = rs["nUCMT"].ToString();
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

                        _arqCargaMT.Append(linha);
                    }
                }
            }
        }

        private string GetNomeArqSQL()
        {
            return _par._parGUI._pathRecursosPerm + _arqSQL;
        }

        // OLD CODE 
        // caso optemos por gravar os SQLs e executar manualmente.
        private void GravaSQLQueries()
        {
            _sqlQueries = new StringBuilder();

            // 3.3 Append 
            //_sqlQueries.Append("-- Recloser:" + recloser + Environment.NewLine);
            //_sqlQueries.Append("Select CodChvMT,Descr from StoredChaveMT where CodChvMT='" + recloserShort + "'" + Environment.NewLine);
            _sqlQueries.Append(_sqlCargaMT);
            _sqlQueries.Append(_sqlCargaBT);
            _sqlQueries.Append(_sqlIP);
            _sqlQueries.Append(_sqlGeradorMT);
            _sqlQueries.Append(_sqlGeradorBT);

            //TxtFile.SafeDelete(GetNomeArqSQL());

            // grava em arquivo
            TxtFile.GravaEmArquivo2(_sqlQueries.ToString(), GetNomeArqSQL(), _par._mWindow);
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

        private bool PutEnergyMetersOnReclosers()
        {
            // CTRR44079 CTRR44105 CTRR45585

            // TODO
            //Reclosers recloser = _oDSS._DSSObj.ActiveCircuit.Reclosers;

            // as recloser are still modelled as switches, its necessary do look in Lines
            Lines lines = _oDSS._DSSObj.ActiveCircuit.Lines;

            _lstReclosers = new List<string>();

            int iLines = lines.First;
            while (iLines != 0)
            {
                string nome = lines.Name;

                // ctrr means recloser in Cemig feeders & is 3 phase & TODO is closed
                if (lines.Name.Contains("ctrr") && lines.Phases == 3)
                {
                    _lstReclosers.Add(nome);
                }
                iLines = lines.Next;
            }

            // Put EnergyMeters
            foreach (string recloser in _lstReclosers)
            {
                _oDSS._DSSText.Command = "New energymeter." + recloser + " element=Line." + recloser + ",terminal=1";

                /*
                meters.Name = recloser.Replace("Line.ctrr", "");
                meters.MeteredElement = recloser;
                meters.MeteredTerminal = 1;
                */
            }
            // DEBUG
            //Meters meters = _oDSS._DSSObj.ActiveCircuit.Meters;

            return true;
        }

        // constructs SQLs queries for MV loads, BT loads and public lamps
        private bool ConstructSQLQuery(string recloser)
        {
            /* // DEBUG religador ULAE728
            if (recloser.Equals("r32433"))
            {
                int debug = 0;
            }*/

            _sqlCargaMT = "";
            _sqlCargaBT = "";
            _sqlIP = "";
            _sqlGeradorMT = "";
            _sqlGeradorBT = "";

            //
            if (_lstSegmentosMT.Count > 0)
            {
                string segMT = CemigFeeders.AddAposAndCommasForSQL(_lstSegmentosMT);

                _sqlCargaMT = "select '" + recloser + "' as Relig,sum(nUCMT)as'nUCMT',sum(S01_MWh)as'S01_MWh',sum(S02_MWh)as'S02_MWh',sum(S03_MWh)as'S03_MWh'," +
                    "sum(S04_MWh)as'S04_MWh',sum(S05_MWh)as'S05_MWh',sum(S06_MWh)as'S06_MWh',sum(S07_MWh)as'S07_MWh'," +
                    "sum(S08_MWh)as'S08_MWh',sum(S09_MWh)as'S09_MWh',sum(S10_MWh)as'S10_MWh',sum(S11_MWh)as'S11_MWh',sum(S12_MWh)as'S12_MWh' from" +
                    "( Select '" + recloser + "' as Relig,count(CodConsMT)as'nUCMT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredCargaMT as mt " +
                    "inner join " + _schemaDB + "StoredSegmentoMT as seg on seg.CodPonAcopl1 = mt.CodPonAcopl " +
                    "where seg.CodSegmMT in (" + segMT + ")" +
                    "UNION Select '" + recloser + "' as Relig,count(CodConsMT)as'nUCMT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredCargaMT as mt " +
                    "inner join " + _schemaDB + "StoredSegmentoMT as seg on seg.CodPonAcopl2 = mt.CodPonAcopl " +
                    "where seg.CodSegmMT in (" + segMT + ")" +
                    " ) as T group by T.Relig" + Environment.NewLine;

                _sqlGeradorMT = "select '" + recloser + "' as Relig,sum(nUGMT)as'nUGMT',sum(S01_MWh)as'S01_MWh',sum(S02_MWh)as'S02_MWh',sum(S03_MWh)as'S03_MWh'," +
                    "sum(S04_MWh)as'S04_MWh',sum(S05_MWh)as'S05_MWh',sum(S06_MWh)as'S06_MWh',sum(S07_MWh)as'S07_MWh'," +
                    "sum(S08_MWh)as'S08_MWh',sum(S09_MWh)as'S09_MWh',sum(S10_MWh)as'S10_MWh',sum(S11_MWh)as'S11_MWh',sum(S12_MWh)as'S12_MWh' from" +
                    "( Select '" + recloser + "' as Relig,count(CodGeraMT)as'nUGMT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredGeradorMT as mt " +
                    "inner join " + _schemaDB + "StoredSegmentoMT as seg on seg.CodPonAcopl1 = mt.CodPonAcopl " +
                    "where seg.CodSegmMT in (" + segMT + ")" +
                    "UNION Select '" + recloser + "' as Relig,count(CodGeraMT)as'nUGMT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredGeradorMT as mt " +
                    "inner join " + _schemaDB + "StoredSegmentoMT as seg on seg.CodPonAcopl2 = mt.CodPonAcopl " +
                    "where seg.CodSegmMT in (" + segMT + ")" +
                    " ) as T group by T.Relig" + Environment.NewLine;
            }
            if (_lstRamais.Count > 0)
            {
                string ramais = CemigFeeders.AddAposAndCommasForSQL(_lstRamais);

                //obs: o join tem q ser feito no rbt.CodPonAcopl2 pq os ramais de entrada foram exportados
                _sqlCargaBT = "Select '" + recloser + "' as Relig,count(CodConsBT) as 'nUCBT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredCargaBT as bt " +
                    "inner join " + _schemaDB + "StoredRamalBT as rbt on rbt.CodPonAcopl2 = bt.CodPonAcopl " +
                    "where rbt.CodRmlBT in (" + ramais + ")" + Environment.NewLine;

                _sqlGeradorBT = "Select '" + recloser + "' as Relig,count(CodGeraBT) as 'nUGBT',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredGeradorBT as bt " +
                    "inner join " + _schemaDB + "StoredRamalBT as rbt on rbt.CodPonAcopl2 = bt.CodPonAcopl " +
                    "where rbt.CodRmlBT in (" + ramais + ")" + Environment.NewLine;
            }
            if (_lstSegmentosBT.Count > 0)
            {
                string segmentosBT = CemigFeeders.AddAposAndCommasForSQL(_lstSegmentosBT);

                //obs: o join tem q ser feito no sbt.CodPonAcopl2 pq os cabos de IP foram exportados
                _sqlIP = "Select '" + recloser + "' as Relig,count(CodConsBT) as 'nIP',sum(EnerMedid01_MWh)as'S01_MWh',sum(EnerMedid02_MWh)as'S02_MWh'," +
                    "sum(EnerMedid03_MWh)as'S03_MWh',sum(EnerMedid04_MWh)as'S04_MWh',sum(EnerMedid05_MWh)as'S05_MWh',sum(EnerMedid06_MWh)as'S06_MWh'," +
                    "sum(EnerMedid07_MWh)as'S07_MWh',sum(EnerMedid08_MWh)as'S08_MWh',sum(EnerMedid09_MWh)as'S09_MWh'," +
                    "sum(EnerMedid10_MWh)as'S10_MWh',sum(EnerMedid11_MWh)as'S11_MWh',sum(EnerMedid12_MWh)as'S12_MWh' " +
                    "from " + _schemaDB + "StoredCargaBT as bt " +
                    "inner join " + _schemaDB + "StoredSegmentoBT as sbt on sbt.CodPonAcopl2 = bt.CodPonAcopl " +
                    "where sbt.CodSegmBT in (" + segmentosBT + ")" + Environment.NewLine;
            }
            return true;
        }

        /*
        // TODO funcao de CemigFeeder.cs
        private static string AddAposAndCommasForSQL(List<string> lst)
        {
            string retString;

            // inicializacao 
            retString = "'";

            // para cada alimentador da lista
            foreach (string alim in lst)
            {
                retString += alim;

                if (string.Equals(alim, lst.Last()))
                {
                    retString += "'";
                }
                else
                {
                    retString += "','";
                }
            }
            return retString;
        }*/

        //
        private bool Get_MVandLVLines_CodIDs(string recloser)
        {
            Meters med = _oDSS._DSSObj.ActiveCircuit.Meters;

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
            // Get MV, LV segments below the recloser 

            string[] allBranchs = med.AllBranchesInZone;

            _lstSegmentosMT = new List<string>();
            _lstSegmentosBT = new List<string>();
            _lstRamais = new List<string>();

            // extratifica segmentos MT e BT e ramais. 
            foreach (string branch in allBranchs)
            {
                string branch2;
                // MV line segment
                if (branch.Contains("smt"))
                {
                    branch2 = branch.Replace("Line.smt_", "");
                    _lstSegmentosMT.Add(branch2);
                    continue;
                }
                // if switch or recloser
                if (branch.Contains("ctr"))
                {
                    branch2 = branch.Replace("Line.ctr", "");
                    _lstSegmentosMT.Add(branch2);
                    continue;
                }
                // LV line segment
                if (branch.Contains("sbt"))
                {
                    branch2 = branch.Replace("Line.sbt_", "");
                    _lstSegmentosBT.Add(branch2);
                    continue;
                }
                if (branch.Contains("rbt"))
                {
                    branch2 = branch.Replace("Line.rbt_", "");
                    _lstRamais.Add(branch2);
                    continue;
                }
                /* //DEBUG
                if (branch.Contains("e"))
                {
                    continue;
                }*/

            }
            return true;
        }
    }
}

