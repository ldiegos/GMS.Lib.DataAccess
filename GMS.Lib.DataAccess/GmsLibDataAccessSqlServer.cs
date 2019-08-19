using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// 
    /// </summary>
    public class GmsLibDataAccessSqlServer : IGmsLibDataAccess, IDisposable
    {
        // Flag: Has Dispose already been called?
        bool _disposed = false;
        // Instantiate a SafeHandle instance.
        readonly SafeHandle _handle = new SafeFileHandle(IntPtr.Zero, true);

        private SqlConnection _objCn;
        private OleDbConnection _objOleCn;

        private bool _conectado = false;

        /// <summary>
        /// 
        /// </summary>
        private string CadenaError { get; set; }


        /// <summary>
        /// 
        /// </summary>
        private SqlTransaction MyTrans { get; set; }

        private string StrConnectionString { get; set; }

        private int NumberMaxConnectionTries { get; set; } = 3;

        private int NumberMaxExecutionTries { get; set; } = 10;

        private int _countConnectionTries = 0;
        private int _countExcecutionTries = 0;

        private int WaitConnectionMillisecondsTimeout { get; set; } = 10000;
        private int WaitExecutionMillisecondsTimeout { get; set; } = 6000;

        private bool _bolConnect = false;

        #region "PrivateMethdos"

        private bool ReconexionSql()
        {
            bool bolconectado = false;
            try
            {
                this._objCn.Open();
                bolconectado = true;

            }
            catch (Exception)
            {

                bolconectado = false;


            }
            return bolconectado;

        }

        private bool ReconexionOleDb()
        {
            bool bolconectado = false;
            try
            {
                this._objOleCn.Open();
                bolconectado = true;

            }
            catch (Exception)
            {

                bolconectado = false;


            }
            return bolconectado;

        }

        //private void CheckResult(string query, ref DataTable dt, ref string queryResult, ref bool salida)
        //{
        //    string patterMatched = string.Empty;

        //    if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
        //    {
        //        if (_countExcecutionTries <= NumberMaxExecutionTries)
        //        {
        //            _countExcecutionTries++;
        //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
        //            salida = ExecuteStandAlone(query, ref dt, out queryResult);
        //        }
        //        else
        //        {
        //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
        //            salida = false;
        //        }
        //    }
        //}

        //private void CheckResult(string query, ref DataSet ds, ref string queryResult, ref bool salida)
        //{
        //    string patterMatched = string.Empty;

        //    if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
        //    {
        //        if (_countExcecutionTries <= NumberMaxExecutionTries)
        //        {
        //            _countExcecutionTries++;
        //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
        //            salida = ExecuteStandAlone(query, ref ds, out queryResult);
        //        }
        //        else
        //        {
        //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
        //            salida = false;
        //        }
        //    }
        //}

        //private void CheckResult(SqlCommand sqlCommand, ref DataTable dt, ref string queryResult, ref bool salida)
        //{
        //    string patterMatched = string.Empty;

        //    if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
        //    {
        //        if (_countExcecutionTries <= NumberMaxExecutionTries)
        //        {
        //            _countExcecutionTries++;
        //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
        //            salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
        //        }
        //        else
        //        {
        //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
        //            salida = false;
        //        }
        //    }
        //}

        //private void CheckResult(SqlCommand sqlCommand, ref DataSet ds, ref string queryResult, ref bool salida)
        //{
        //    string patterMatched = string.Empty;

        //    if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
        //    {
        //        if (_countExcecutionTries <= NumberMaxExecutionTries)
        //        {
        //            _countExcecutionTries++;
        //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
        //            salida = ExecuteStandAlone(sqlCommand, ref ds, out queryResult);
        //        }
        //        else
        //        {
        //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
        //            salida = false;
        //        }
        //    }
        //}

        //private void CheckResult(OleDbCommand sqlCommand, ref DataTable dt, ref string queryResult, ref bool salida)
        //{
        //    string patterMatched = string.Empty;

        //    if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
        //    {
        //        if (_countExcecutionTries <= NumberMaxExecutionTries)
        //        {
        //            _countExcecutionTries++;
        //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
        //            salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
        //        }
        //        else
        //        {
        //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
        //            salida = false;
        //        }
        //    }
        //}

        private void CheckResult(string queryResult, ref bool salida)
        {
            string patterMatched = string.Empty;

            if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
            {
                if (_countExcecutionTries <= NumberMaxExecutionTries)
                {
                    _countExcecutionTries++;
                    System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
                    salida = true;
                }
                else
                {
                    queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
                    salida = false;
                }
            }
        }

        #endregion "PrivateMethdos"

        #region "PublicMethdos"

        /// <summary>
        /// Constructor simple, sólo cadena de conexión, se cogen los valores de reconexión por defecto.
        /// </summary>
        /// <param name="connectionString">Cadenad de conexión</param>
        public GmsLibDataAccessSqlServer(string connectionString)
        {
            StrConnectionString = connectionString;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="connectionString">Connection string for SQLServer database</param>
        /// <param name="numMaxReconnections">Number of retries to connect to database</param>
        /// <param name="numSecondsBetweenReconnection">Seconds to wait for each retry</param>
        /// <param name="maxExecutionRetries">Number of retries to execute a query/command</param>
        /// <param name="numSecondsBetweenExecutionRetries">Seconds to wait for each execution retry</param>
        public GmsLibDataAccessSqlServer(string connectionString, int numMaxReconnections, int numSecondsBetweenReconnection
            , int maxExecutionRetries, int numSecondsBetweenExecutionRetries)
        {
            StrConnectionString = connectionString;
            NumberMaxConnectionTries = numMaxReconnections;
            WaitConnectionMillisecondsTimeout = numSecondsBetweenReconnection * 1000;
            NumberMaxExecutionTries = maxExecutionRetries;
            WaitExecutionMillisecondsTimeout = numSecondsBetweenExecutionRetries * 1000;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {

            // Dispose of unmanaged resources.
            Dispose(true);
        }

        // Protected implementation of Dispose pattern.
        /// <summary>
        /// 
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                _handle.Dispose();
                // Free any other managed objects here.
                //
            }

            if (_objOleCn != null)
            {
                CloseSqlConnection();
                _objOleCn = null;
            }

            // Free any unmanaged objects here.
            //
            _disposed = true;
        }


        /// <summary>
        /// Abrir conexión estandar tipo SQL Server
        /// </summary>
        /// <returns></returns>
        public bool OpenSqlConnection()
        {
            try
            {
                if (this._objCn == null)
                {
                    _objCn = new SqlConnection(this.StrConnectionString);
                }

                this._conectado = ReconexionSql();

                if (!this._conectado)
                {
                    System.Threading.Thread.Sleep(WaitConnectionMillisecondsTimeout);

                    _countConnectionTries++;
                    if (_countConnectionTries <= NumberMaxConnectionTries)
                    {
                        this.OpenSqlConnection();
                    }
                    else
                    {
                        _bolConnect = false;
                    }
                }
                else
                {
                    _bolConnect = true;
                }

                return _bolConnect;
            }
            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                this._conectado = false;
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool OpenOleDbConnection()
        {
            try
            {
                if (this._objOleCn == null)
                {
                    _objOleCn = new OleDbConnection(this.StrConnectionString);
                }

                this._conectado = ReconexionOleDb();

                if (!this._conectado)
                {
                    System.Threading.Thread.Sleep(WaitConnectionMillisecondsTimeout);

                    _countConnectionTries++;
                    if (_countConnectionTries <= NumberMaxConnectionTries)
                    {
                        this.OpenOleDbConnection();
                    }
                    else
                    {
                        _bolConnect = false;
                    }
                }
                else
                {
                    _bolConnect = true;
                }

                return _bolConnect;
            }
            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                this._conectado = false;
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool CloseSqlConnection()
        {

            try
            {
                if (_conectado)
                {
                    _objCn.Close();

                    _objCn = null;

                    _conectado = false;
                    this.MyTrans = null;
                }
                return true;
            }

            catch (Exception ex)
            {
                this._conectado = false;
                this.CadenaError = ex.Message;
                this.MyTrans = null;
                return false;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool CloseOleDbConnection()
        {

            try
            {
                if (_conectado)
                {
                    _objOleCn.Close();

                    _objOleCn = null;

                    _conectado = false;
                    this.MyTrans = null;
                }
                return true;
            }

            catch (Exception ex)
            {
                this._conectado = false;
                this.CadenaError = ex.Message;
                this.MyTrans = null;
                return false;
            }
        }


        /// <summary>
        /// Function: public bool EjecutaSP(string sProcedure, ref DataTable dt, ref string queryResult)
        ///     Ejecuta una cadena SQL y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sProcedure">Type string - Sentencia SQL que se desea ejecutar</param>
        /// <param name="dt">Type DataTable - Datatable donde se devuelven los datos de la ejecución</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecución</param> 
        /// <returns>Devuelve true/false en función de si se ejecuta correctamente.</returns>
        public bool Execute(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;
            try
            {


                if (
                    sProcedure.ToUpper().Contains("DELETE") ||
                    sProcedure.ToUpper().Contains("UPDATE") ||
                    sProcedure.ToUpper().Contains("INSERT")
                    )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    SqlCommand sqlCommand = new SqlCommand(sProcedure);
                
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(sProcedure, this._objCn);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }


                salida = true;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sProcedure, ref dt, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sProcedure, ref dt, out queryResult);
                }
            }
            return salida;
        }

        /// <summary>
        /// Function:  public bool EjecutaSP(SqlCommand sqlCommand, ref DataTable dt, ref string queryResult)
        ///     Ejecuta un comando SQL y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sqlCommand">Type SqlCommand - Comando SQLite que se desea ejecutar</param>
        /// <param name="dt">Type DataTable - Datatable donde se devuelven los datos de la ejecución</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecución</param> 
        /// <returns>Devuelve true/false en función de si se ejecuta correctamente.</returns>
        ///
        public bool Execute(SqlCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            bool salida;
            bool isNonQuery = false;
            try
            {
                sqlCommand.Connection = _objCn;


                if (
                    sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
                    sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                    sqlCommand.CommandText.ToUpper().Contains("INSERT")
                    )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                    dataAdapter = new SqlDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }

                salida = true;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }
            }
            finally
            {
                dataAdapter.Dispose();
            }

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(OleDbCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                sqlCommand.Connection = _objOleCn;


                if (
                    sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
                    sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                    sqlCommand.CommandText.ToUpper().Contains("INSERT")
                    )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                    dataAdapter = new OleDbDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }


                salida = true;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }
            }
            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand">OleDB command</param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(OleDbCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenOleDbConnection();

                sqlCommand.Connection = _objOleCn;


                if (
                    sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
                    sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                    sqlCommand.CommandText.ToUpper().Contains("INSERT")
                    )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                    dataAdapter = new OleDbDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }

                salida = true;

                CloseOleDbConnection();

                //foreach (string pattern in TimeOuts.ListTimeouts)
                //{
                //    if (queryResult.ToUpper().Contains(pattern))
                //    {
                //        if (_countExcecutionTries <= NumberMaxExecutionTries)
                //        {
                //            _countExcecutionTries++;
                //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
                //            salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                //        }
                //        else
                //        {
                //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
                //            salida = false;
                //        }

                //    }
                //}

                //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }

            }
            finally
            {
                CloseOleDbConnection();
            }

            return salida;
        }

        /// <summary>
        /// Return DataSet
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="ds"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataSet ds, out string queryResult)
        {
            SqlDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objCn;

                if (
                    sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
                    sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                    sqlCommand.CommandText.ToUpper().Contains("INSERT")
                )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    dataAdapter = new SqlDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(ds).ToString();
                }


                salida = true;

                CloseSqlConnection();

                //CheckResult(sqlCommand, ref ds, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref ds, out queryResult);
                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CloseSqlConnection();

                //foreach (string pattern in TimeOuts.ListTimeouts)
                //{
                //    if (queryResult.ToUpper().Contains(pattern))
                //    {
                //        if (_countExcecutionTries <= NumberMaxExecutionTries)
                //        {
                //            _countExcecutionTries++;
                //            System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
                //            salida = ExecuteStandAlone(sqlCommand, ref ds, out queryResult);
                //        }
                //        else
                //        {
                //            queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
                //            salida = false;
                //        }

                //    }
                //}

                //CheckResult(sqlCommand, ref ds, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref ds, out queryResult);
                }

            }
            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand">Sql Command</param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            SqlDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objCn;

                if (
                    sqlCommand.CommandText.ToUpper().Contains("DELETE") || 
                    sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                    sqlCommand.CommandText.ToUpper().Contains("INSERT")
                    ) 
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    dataAdapter = new SqlDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }
                
                salida = true;

                CloseSqlConnection();

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CloseSqlConnection();

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                }

            }
            finally
            {
                CloseSqlConnection();
                sqlCommand.Dispose();
            }

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure">string query</param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                if (
                    sProcedure.ToUpper().Contains("DELETE") ||
                    sProcedure.ToUpper().Contains("UPDATE") ||
                    sProcedure.ToUpper().Contains("INSERT")
                    )
                {
                    isNonQuery = true;
                }

                if (isNonQuery)
                {
                    SqlCommand sqlCommand= new SqlCommand(sProcedure,this._objCn);

                    queryResult = sqlCommand.ExecuteNonQuery().ToString();
                }
                else
                {
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                    SqlDataAdapter dataAdapter = new SqlDataAdapter(sProcedure, this._objCn);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }

                salida = true;

                CloseSqlConnection();

                //CheckResult(sProcedure, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sProcedure, ref dt, out queryResult);
                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                CloseSqlConnection();

                //CheckResult(sProcedure, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sProcedure, ref dt, out queryResult);
                }

            }

            return salida;
        }

        #endregion "PublicMethdos"


        #region Public_transaction_Methods

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool AbrirConexionTransaccion()
        {
            try
            {
                if (this._objCn == null)
                {
                    _objCn = new SqlConnection(this.StrConnectionString);
                }

                this._objCn.Open();
                MyTrans = this._objCn.BeginTransaction();


                this._conectado = true;
                return true;
            }


            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                this._conectado = false;
                return false;
            }
        }


        /// <summary>
        /// 
        /// </summary>
        public void commit_transacion()
        {
            MyTrans.Commit();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Rollback_transacion()
        {

            MyTrans.Rollback();
        }


        /// <summary>
        /// Este evento ejecuta en forma de Transacion todas las querys que recibe como parametro devuelve un bool indicandote si ha ido bien.
        /// </summary>
        /// <param name="arrayInsert"></param>
        /// <returns></returns>
        public bool Transaccion(string[] arrayInsert)
        {
            string consulta;
            bool error = false;

            SqlTransaction trans = this._objCn.BeginTransaction();


            for (int i = 0; i < arrayInsert.Length; i++)
            {
                consulta = arrayInsert[i];
                SqlCommand cmd = new SqlCommand(consulta, this._objCn, trans);
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch
                {
                    error = true;
                }
            } // fin del for

            if (error)
                trans.Rollback();
            else
                trans.Commit();

            return error;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <returns></returns>
        public bool EjecutaSP_command_transacion(string sProcedure)
        {
            bool salida;
            try
            {
                SqlCommand cmd = new SqlCommand(sProcedure, this._objCn);
                cmd.Transaction = MyTrans;
                cmd.ExecuteNonQuery();
                salida = true;
            }
            catch (Exception)
            {
                salida = false;
            }
            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="dt"></param>
        /// <returns></returns>
        public bool EjecutaSP_transaccion(string sProcedure, ref DataTable dt)
        {
            bool salida;
            try
            {
                //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
                SqlDataAdapter dataAdapter = new SqlDataAdapter(sProcedure, this._objCn);
                dataAdapter.SelectCommand.Transaction = this.MyTrans;
                dataAdapter.Fill(dt);
                salida = true;
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
            }
            return salida;
        }



        /// <summary>
        /// 
        /// </summary>
        /// <param name="tabla"></param>
        /// <returns></returns>
        public Int64 dame_insertado_transaccion(string tabla)
        {
            Int64 idInsert;
            string consulta = "select ident_current('" + tabla + "')";
            try
            {
                SqlCommand cmd = new SqlCommand(consulta, this._objCn, this.MyTrans);
                idInsert = System.Convert.ToInt64(cmd.ExecuteScalar());
                //ExecuteScalar. --> Ejecuta la consulta y devuelve la primera columna de la primera fila
                // Es bueno usarlo para recuperar un unico valor
            }
            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                idInsert = -1;
            }
            return idInsert;
        }
        #endregion  Public_transaction_Methods



        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <returns></returns>
        public bool EjecutaSP_command(string sProcedure)
        {
            bool salida;
            try
            {
                SqlCommand cmd = new SqlCommand(sProcedure, this._objCn);
                cmd.ExecuteNonQuery();//Ejecuta comandos como instrucciones INSERT, DELETE, UPDATE y SET de Transact-SQL.
                salida = true;
            }
            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                salida = false;
            }
            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tabla"></param>
        /// <returns></returns>
        public Int64 DameIdInsertado(string tabla)
        {
            Int64 idInsert;
            string consulta = "select ident_current('" + tabla + "')";
            try
            {
                SqlCommand cmd = new SqlCommand(consulta, this._objCn);
                idInsert = System.Convert.ToInt64(cmd.ExecuteScalar());

                //ExecuteScalar. --> Ejecuta la consulta y devuelve la primera columna de la primera fila

            }
            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                idInsert = -1;
            }
            return idInsert;
        }


        #region NotImplemented

        public bool FnCreateDataBase(string strConnectionString)
        {
            throw new NotImplementedException();
        }

        public bool ExecuteSqlQuery(string strSqlExec, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataSet dt, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(SQLiteCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            throw new NotImplementedException();
        }


        public bool BackupDataBase(string strPath2Exec, string strSqlite3Path)
        {
            throw new NotImplementedException();
        }


        public void ForceDataBaseUpdate(out string queryResult)
        {
            throw new NotImplementedException();
        }

        #endregion NotImplemented



    }
}
