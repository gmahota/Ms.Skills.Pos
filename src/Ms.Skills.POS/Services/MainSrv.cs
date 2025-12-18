using ErpBS100;
using StdBE100;
using StdBESql100;
using StdPlatBS100;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static StdBESql100.StdBESqlTipos;

namespace Ms.Skills.POS.Services
{
    public class MainSrv : DisposableBase
    {

        public MainSrv() { }

        public MainSrv(ErpBS BSO, StdBSInterfPub PSO)
        {
            PriEngine.Engine = BSO;

            PriEngine.Platform = PSO;
        }

        #region Configuration
        /// <summary>
        /// Método para resolução das assemblies.
        /// </summary>
        /// <param name="sender">Application</param>
        /// <param name="args">Resolving Assembly Name</param>
        /// <returns>Assembly</returns>
        public static System.Reflection.Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            string assemblyFullName;

            System.Reflection.AssemblyName assemblyName;

            assemblyName = new System.Reflection.AssemblyName(args.Name);
            assemblyFullName = System.IO.Path.Combine(Environment.GetEnvironmentVariable("PERCURSOSGP100", EnvironmentVariableTarget.Machine), assemblyName.Name + ".dll");

            if (System.IO.File.Exists(assemblyFullName))
                return System.Reflection.Assembly.LoadFile(assemblyFullName);
            else
                return null;
        }
        #endregion

        public ErpBS GetEngine()
        {
            return PriEngine.Engine;
        }

        #region System Coin Management
        /// <summary>
        /// Devolve a moeda do sistema
        /// </summary>
        /// <returns></returns>
        public string GetSystemCoin()
        {
            return GetEngine().Contexto.MoedaBase;
        }

        /// <summary>
        /// Devolve a lista de Moedas registradas no sistema
        /// </summary>
        /// <returns></returns>
        public DataTable GetCoins()
        {
            string strsql = "";
            try
            {
                strsql = strsql + "select moeda from moedas";

                return DataBaseRead(strsql);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Devolve o cambio de uma determinada moeda
        /// </summary>
        /// <param name="coin"></param>
        /// <param name="date"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public double GetExchange(string coin, DateTime date, string type = "P")
        {
            try
            {
                switch (type)
                {
                    case "R":
                        {
                            return GetEngine().Base.Moedas.DaCambioCompra(coin, date);
                        }

                    case "P":
                        {
                            return GetEngine().Base.Moedas.DaCambioVenda(coin, date);
                        }
                }

                return (
                    GetEngine().Base.Moedas.DaCambioVenda(coin, date) +
                        GetEngine().Base.Moedas.DaCambioCompra(coin, date))
                            / (double)2;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string GetSystemCoinAlternative()
        {
            try
            {
                return GetEngine().Contexto.MoedaAlternativa;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Database Queries
        public DataTable DataBaseRead(string querySql)
        {

            try
            {
                DataTable dt = new DataTable();

                string connectionString = PriEngine.Platform.BaseDados.DaConnectionStringNET(PriEngine.Platform.BaseDados.DaNomeBDdaEmpresa(PriEngine.Engine.Contexto.CodEmp),
                    "Default");

                SqlConnection con = new SqlConnection(connectionString);

                SqlDataAdapter da = new SqlDataAdapter(querySql, con);

                SqlCommandBuilder cb = new SqlCommandBuilder(da);

                da.Fill(dt);

                return dt;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //public DataTable DataBaseRead(string tabela, int maximo = 0, string campos = "", string filtros = "", string juncoes = "", string ordenacao = "")
        //{
        //    string strSql = "select ";

        //    try
        //    {

        //        DataTable dt;

        //        if (maximo > 0) strSql = strSql + string.Format("Top {0} ", maximo);

        //        if ((campos.Length > 0))
        //        {
        //            strSql = (strSql + campos);
        //        }
        //        else
        //        {
        //            strSql = (strSql + " * ");
        //        }

        //        strSql = (strSql + (" from " + tabela));
        //        if ((juncoes.Length > 0))
        //        {
        //            strSql = (strSql + (" " + juncoes));
        //        }

        //        if ((filtros.Length > 0))
        //        {
        //            strSql = (strSql + (" where " + filtros));
        //        }

        //        if ((ordenacao.Length > 0))
        //        {
        //            strSql = (strSql + (" order by " + ordenacao));
        //        }

        //        dt = this.DataBaseRead(strSql);

        //        return dt;
        //    }
        //    catch (Exception ex)
        //    {
        //        escreveErro(strSql);

        //        throw ex;
        //    }

        //}

        public DataTable DataBaseRead(string tabela, int maximo = 0, string campos = "", List<string> listFiltros = null, string juncoes = "", string groupBy = "", string ordenacao = "")
        {


            string strSql = "select ";

            string filtros = listFiltros is null ? "" : string.Join(" and ", listFiltros);
            try
            {

                DataTable dt;

                if (maximo > 0) strSql = strSql + string.Format("Top {0} ", maximo);

                if ((campos.Length > 0))
                {
                    strSql = (strSql + campos);
                }
                else
                {
                    strSql = (strSql + " * ");
                }

                strSql = (strSql + (" from " + tabela));
                if ((juncoes.Length > 0))
                {
                    strSql = (strSql + (" " + juncoes));
                }

                if ((filtros.Length > 0))
                {
                    strSql = (strSql + (" where " + filtros));
                }

                if ((groupBy.Length > 0))
                {
                    strSql = (strSql + (" group by " + groupBy));
                }

                if ((ordenacao.Length > 0))
                {
                    strSql = (strSql + (" order by " + ordenacao));
                }

                dt = this.DataBaseRead(strSql);

                return dt;
            }
            catch (Exception ex)
            {
                escreveErro(strSql);

                throw ex;
            }

        }
        public int DataBaseExecute(string querySQL)
        {
            try
            {
                DataTable dt = new DataTable();

                string connectionString = PriEngine.Platform.BaseDados.DaConnectionStringNET(PriEngine.Platform.BaseDados.DaNomeBDdaEmpresa(PriEngine.Engine.Contexto.CodEmp),
                    "Default");

                SqlConnection con = new SqlConnection(connectionString);

                SqlCommand command = new SqlCommand();
                command.CommandType = CommandType.Text;
                command.Connection = con;
                command.CommandText = querySQL;
                con.Open();

                return command.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public string GetParameter(string Name)
        {
            try
            {
                string result = "";
                string query = string.Format("select * from TDU_Parametros where CDU_Parametro = '{0}' ", Name);
                DataTable dt = DataBaseRead(query);

                if (dt.Rows.Count > 0)
                {
                    result = dt.Rows[0]["CDU_Valor"].ToString();
                }
                else
                {
                    new Exception("O parametro {0} não se encontra configurado na tabela de TDU_Parametros");
                }

                return result;
            }
            catch (Exception e)
            {
                throw e;
            }

        }

        public void DataTableNewLine(DataTable dt)
        {
            try
            {
                DataRow dr = dt.NewRow();
                dt.Rows.InsertAt(dr, 0);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
        public void escreveErro(string pastaLog, string name, string logMessage)
        {
            string ficheiro;
            try
            {
                ficheiro = pastaLog;

                using (StreamWriter w = File.AppendText(ficheiro + "\\" + string.Format("erro_{0}.log", name)))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void escreveErro(string logMessage)
        {
            string ficheiro;
            try
            {
                ficheiro = GetParameter("Log_PastaErro");

                using (StreamWriter w = File.AppendText(ficheiro + "\\" + string.Format("log_{0}.txt", DateTime.Now.ToString("ddMMyy"))))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void escreveLog(string pastaLog, string name, string logMessage)
        {
            string ficheiro;
            try
            {
                ficheiro = pastaLog;

                using (StreamWriter w = File.AppendText(ficheiro + "\\" + string.Format("log_{0}.log", name)))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
                //throw ex;
            }
        }

        public void escreveLog(string logMessage)
        {
            string ficheiro;
            try
            {
                ficheiro = GetParameter("Log_PastaErro");

                using (StreamWriter w = File.AppendText(ficheiro + "\\" + string.Format("log_{0}.txt", DateTime.Now.ToString("ddMMyy"))))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.WriteLine("{0} - {1}", DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss"), logMessage);
            }
            catch (Exception ex)
            {

            }
        }
        #region Gestao Excel
        public List<string> ListaFolhaExcel(string caminhoExcell)
        {
            try
            {

                if (!File.Exists(caminhoExcell))
                {
                    throw new Exception("File does not exists");
                }

                List<string> listSheet = new List<string>();

                OleDbConnection conn = new OleDbConnection();

                string connString = ExcelConnection(caminhoExcell);

                conn.ConnectionString = connString;
                conn.Open();

                DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                foreach (DataRow drSheet in dt.Rows)
                {
                    listSheet.Add(drSheet["TABLE_NAME"].ToString());
                }

                return listSheet;


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable CarregaListaExcel(string caminhoExcell, int sheetNumber, int linhaInicial)
        {

            DataTable dt;
            DataRow dr;

            try
            {

                if (!File.Exists(caminhoExcell))
                {
                    throw new Exception("File does not exists");
                }

                List<string> listSheet = new List<string>();

                OleDbConnection conn = new OleDbConnection();

                string connString = ExcelConnection(caminhoExcell);

                conn.ConnectionString = connString;
                conn.Open();

                var dtExcel = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                dt = ObtemDadosSheetExcell(conn, dtExcel.Rows[sheetNumber]["TABLE_NAME"].ToString());

                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return dt;

        }

        public DataTable ObtemDadosSheetExcell(OleDbConnection conn, string sheet)
        {
            try
            {
                DataTable dt = new DataTable();
                OleDbCommand cmd = new OleDbCommand();
                OleDbDataReader connReader;

                cmd.Connection = conn;
                cmd.CommandText = string.Format("Select * From [{0}]", sheet);

                connReader = cmd.ExecuteReader();
                dt.Load(connReader);

                connReader.Close();

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    object id = dt.Rows[i];
                    if (id != null && !String.IsNullOrEmpty(id.ToString().Trim()))
                    {

                    }
                    else
                    {
                        dt.Rows[i].Delete();
                    }
                }

                return dt;

            }
            catch (Exception ex)
            {
                throw new Exception("<ObtemDadosSheetExcell>_" + ex.Message);
            }
        }
        private string ExcelConnection(string fileName)
        {
            string provider = "Microsoft.ACE.OLEDB.12.0";
            string dataSource = fileName;
            string extendProperties = "'Excel 12.0;HDR=YES'";

            provider = provider.Length > 0 ? provider : "Microsoft.ACE.OLEDB.12.0";

            string conn = string.Format(
                   @"Provider={0};
                    Data Source ={1};
                    Extended Properties={2}",
                   provider, dataSource, extendProperties
                );
            return conn;
        }
        #endregion

        public object F4Event(string query, string Coluna, string nomeTabela)
        {
            try
            {
                return PriEngine.Platform.Listas.GetF4SQL(nomeTabela, query, Coluna);
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }




        public void DrillDownDocument(string documento)
        {
            try
            {
                documento = documento.Replace("Importado - ", "");

                DataTable dt;
                string query = string.Format("select TipoDoc, Serie, NumDoc,Filial from cabecdoc where tipodoc + ' '+ convert(nvarchar,numdoc) +'/'+serie ='{0}' and filial = '000' ", documento);
                string tipodoc, serie, filial, modulo = "V";
                int numdoc;

                dt = DataBaseRead(query);
                if (dt.Rows.Count > 0)
                {
                    tipodoc = dt.Rows[0]["TipoDoc"].ToString();
                    serie = dt.Rows[0]["Serie"].ToString();
                    filial = dt.Rows[0]["Filial"].ToString();
                    numdoc = Convert.ToInt32(dt.Rows[0]["NumDoc"]);

                    DrillDownDocument(modulo, filial, tipodoc, serie, numdoc);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void DrillDownDocument(string strModulo, string strFilial, string strTipodoc, string strSerie, int intNumDoc)
        {
            StdBESqlCampoDrillDown objCampoDrillDown = new StdBESqlCampoDrillDown();
            StdBEValoresStr objParam = new StdBEValoresStr();

            try
            {
                objCampoDrillDown.ModuloNotificado = ("GCP");
                objCampoDrillDown.Tipo = (EnumTipoDrillDownListas.tddlEventoAplicacao);
                objCampoDrillDown.Evento = ("GCP_EditarDocumento");

                objParam.InsereNovo("Modulo", strModulo);
                objParam.InsereNovo("Filial", strFilial);
                objParam.InsereNovo("Tipodoc", strTipodoc);
                objParam.InsereNovo("Serie", strSerie);
                objParam.InsereNovo("NumDocInt", intNumDoc.ToString());

                PriEngine.Platform.DrillDownLista(objCampoDrillDown, objParam);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        public void ExecutaDrillDownExploracaoCCT(string Exploracao, string TipoEntidade, string Entidade)
        {
            StdBESql100.StdBESqlCampoDrillDown objCampoDrillDown = new StdBESql100.StdBESqlCampoDrillDown
            {
                ModuloNotificado = "CCT",
                Tipo = StdBESql100.StdBESqlTipos.EnumTipoDrillDownListas.tddlEventoAplicacao,
                Evento = "GCP_MOSTRAEXPLORACAO"
            };

            StdBE100.StdBEValoresStr objParam = new StdBE100.StdBEValoresStr();
            objParam.InsereNovo("Exploracao", Exploracao);
            objParam.InsereNovo("TipoEntidade", TipoEntidade);
            objParam.InsereNovo("Entidade", Entidade);

            PriEngine.Platform.DrillDownLista(objCampoDrillDown, objParam);

            objCampoDrillDown = null;
            objParam = null;
        }

        public void ExecuteDrillDown(string Aplicacao, string Evento, string Param1, string Param2 = "", string Param3 = "", string Param4 = "", string Param5 = "")
        {
            StdBESqlCampoDrillDown campoDrillDown = new StdBESqlCampoDrillDown
            {
                ModuloNotificado = Aplicacao,
                Tipo = StdBESqlTipos.EnumTipoDrillDownListas.tddlEventoAplicacao,
                Evento = Evento
            };

            StdBEValoresStr param = new StdBEValoresStr();
            param.InsereNovo("Param1", Param1);

            if (!string.IsNullOrWhiteSpace(Param2))
                param.InsereNovo("Param2", Param2);

            if (!string.IsNullOrWhiteSpace(Param3))
                param.InsereNovo("Param3", Param3);

            if (!string.IsNullOrWhiteSpace(Param4))
                param.InsereNovo("Param4", Param4);

            if (!string.IsNullOrWhiteSpace(Param5))
                param.InsereNovo("Param5", Param5);

            PriEngine.Platform.DrillDownLista(campoDrillDown, param);
        }

        public string DaMesExtensoByNumero(int num)
        {
            string mes = "";

            switch (num)
            {
                case 1: mes = "Janeiro"; break;
                case 2: mes = "Fevereiro"; break;
                case 3: mes = "Março"; break;
                case 4: mes = "Abril"; break;
                case 5: mes = "Maio"; break;
                case 6: mes = "Junho"; break;
                case 7: mes = "Julho"; break;
                case 8: mes = "Agosto"; break;
                case 9: mes = "Setembro"; break;
                case 10: mes = "Outubro"; break;
                case 11: mes = "Novembro"; break;
                case 12: mes = "Dezembro"; break;
            }

            return mes;

        }
    }
}

