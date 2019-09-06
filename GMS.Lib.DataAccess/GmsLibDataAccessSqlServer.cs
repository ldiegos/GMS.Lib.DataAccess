using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// 
    /// </summary>
    public class GmsLibDataAccessSqlServer : IGmsLibDataAccess
    {
        // Flag: Has Dispose already been called?
        bool _disposed = false;
        // Instantiate a SafeHandle instance.
        readonly SafeHandle _handle = new SafeFileHandle(IntPtr.Zero, true);

        private SqlConnection _objCn;
        private OleDbConnection _objOleCn;

        private bool _conectado = false;

        private readonly string _strConnectionString;

        /// <summary>
        /// 
        /// </summary>
        private string _error = string.Empty;


        /// <summary>
        /// 
        /// </summary>
        private SqlTransaction MyTrans { get; set; }

        //private string StrConnectionString { get; set; }

        private int NumberMaxConnectionTries { get; set; } = 3;

        private int NumberMaxExecutionTries { get; set; } = 3;

        private int _countConnectionTries = 0;
        private int _countExcecutionTries = 0;

        private int WaitConnectionMillisecondsTimeout { get; set; } = 10000;
        private int WaitExecutionMillisecondsTimeout { get; set; } = 10000;

        private bool _bolConnect = false;

        #region "PrivateMethdos"

        private bool ReconexionSql(out string result)
        {
            bool bolconectado = false;
            result = string.Empty;

            try
            {
                this._objCn.Open();
                bolconectado = true;

            }
            catch (Exception ex)
            {
                bolconectado = false;
                result = ex.Message;
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

        private void CheckResult(ref string queryResult, out bool timeoutPatterIsFound)
        {
            string patterMatched = string.Empty;
            timeoutPatterIsFound = false;

            if (TimeOuts.CheckTimeout(queryResult, out patterMatched))
            {
                if (_countExcecutionTries <= NumberMaxExecutionTries)
                {
                    _countExcecutionTries++;
                    System.Threading.Thread.Sleep(WaitExecutionMillisecondsTimeout);
                    timeoutPatterIsFound = true;
                }
                else
                {
                    queryResult += "#Number of retries exceed: " + NumberMaxExecutionTries;
                    timeoutPatterIsFound = false;
                }
            }
        }

        #endregion "PrivateMethdos"


        #region "PublicMethods"

        /// <summary>
        /// 
        /// </summary>
        public string GetError { get { return _error; } }

        /// <summary>
        /// Constructor simple, sólo cadena de conexión, se cogen los valores de reconexión por defecto.
        /// </summary>
        /// <param name="connectionString">Cadenad de conexión</param>
        public GmsLibDataAccessSqlServer(string connectionString)
        {
            _strConnectionString = connectionString;
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
            _strConnectionString = connectionString;
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
                    _objCn = new SqlConnection(this._strConnectionString);
                }

                this._conectado = ReconexionSql(out _error);

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
                        _error = $"Max reconnections({NumberMaxConnectionTries}) exceeded." + _error;
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
                this._error = ex.Message;
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
                    _objOleCn = new OleDbConnection(this._strConnectionString);
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
                this._error = ex.Message;
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
            bool result = false;
            
            try
            {
                if (_conectado)
                {
                    _objCn.Close();

                    _objCn = null;

                    _conectado = false;
                    this.MyTrans = null;

                    _error = $"Connection close...";

                    result = true;

                }
                else
                {
                    _error = $"Connection is not opened previously...";
                    result = false;
                }
            }
            catch (Exception ex)
            {
                this._conectado = false;
                this._error = ex.Message;
                this.MyTrans = null;
                result = false;
            }

            return result;

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
                this._error = ex.Message;
                this.MyTrans = null;
                return false;
            }
        }


        /// <summary>
        /// Function: public bool EjecutaSP(string sProcedure, ref DataTable dt, ref string queryResult)
        ///     Ejecuta una cadena SQL y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sProcedure">Type string - Sentencia SQL que se desea ejecutar</param>
        /// <param name="dataTable">Type DataTable - Datatable donde se devuelven los datos de la ejecución</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecución</param> 
        /// <returns>Devuelve true/false en función de si se ejecuta correctamente.</returns>
        public bool Execute(string sProcedure, ref DataTable dataTable, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;
            try
            {
                using (SqlCommand sqlCommand = new SqlCommand(sProcedure.ToString(), _objCn))
                {
                    string commandText = sProcedure.ToUpper();
                    string initialCommand = commandText.Trim().Substring(0, 6);
                    if (
                        initialCommand.Contains("DELETE") ||
                        initialCommand.Contains("UPDATE") ||
                        initialCommand.Contains("INSERT")
                        )
                    {
                        isNonQuery = true;
                    }

                    if (isNonQuery)
                    {
                        //Este tipo de consulta, sólo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqlCommand.ExecuteNonQuery().ToString();
                        salida = true;
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, sólo devuelve resultados en queryResult si es una SELECT.
                        using (SqlDataAdapter DataAdapter = new SqlDataAdapter(sqlCommand))
                        {
                            queryResult = DataAdapter.Fill(dataTable).ToString();
                            salida = true;
                        }
                    }

                    bool timeoutPatterIsFound = false;
                    CheckResult(ref queryResult, out timeoutPatterIsFound);

                    if (timeoutPatterIsFound)
                    {
                        salida = Execute(sProcedure, ref dataTable, out queryResult);
                    }


                }

            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sProcedure;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sProcedure, ref dataTable, out queryResult);
                }
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="dataSet"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(string sProcedure, ref DataSet dataSet, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;
            try
            {
                using (SqlCommand sqlCommand = new SqlCommand(sProcedure.ToString(), _objCn))
                {
                    string commandText = sProcedure.ToUpper();
                    string initialCommand = commandText.Trim().Substring(0, 6);
                    if (
                        initialCommand.Contains("DELETE") ||
                        initialCommand.Contains("UPDATE") ||
                        initialCommand.Contains("INSERT")
                        )
                    {
                        isNonQuery = true;
                    }

                    if (isNonQuery)
                    {
                        //Este tipo de consulta, sólo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqlCommand.ExecuteNonQuery().ToString();
                        salida = true;
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, sólo devuelve resultados en queryResult si es una SELECT.
                        using (SqlDataAdapter DataAdapter = new SqlDataAdapter(sqlCommand))
                        {
                            queryResult = DataAdapter.Fill(dataSet).ToString();
                            salida = true;
                        }
                    }

                    bool timeoutPatterIsFound = false;
                    CheckResult(ref queryResult, out timeoutPatterIsFound);

                    if (timeoutPatterIsFound)
                    {
                        salida = Execute(sProcedure, ref dataSet, out queryResult);
                    }


                }

            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sProcedure;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sProcedure, ref dataSet, out queryResult);
                }
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// Function:  public bool EjecutaSP(SqlCommand sqlCommand, ref DataTable dt, ref string queryResult)
        ///     Ejecuta un comando SQL y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sqlCommand">Type SqlCommand - Comando SQLite que se desea ejecutar</param>
        /// <param name="dataTable">Type DataTable - Datatable donde se devuelven los datos de la ejecución</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecución</param> 
        /// <returns>Devuelve true/false en función de si se ejecuta correctamente.</returns>
        ///
        public bool Execute(SqlCommand sqlCommand, ref DataTable dataTable, out string queryResult)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            bool salida;
            bool isNonQuery = false;
            try
            {
                sqlCommand.Connection = _objCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataTable).ToString();
                }

                salida = true;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataTable, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataTable, out queryResult);
                }
            }
            finally
            {
                dataAdapter.Dispose();
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataSet"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(SqlCommand sqlCommand, ref DataSet dataSet, out string queryResult)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter();
            bool salida;
            bool isNonQuery = false;
            try
            {
                sqlCommand.Connection = _objCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataSet).ToString();
                }

                salida = true;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataSet, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataSet, out queryResult);
                }
            }
            finally
            {
                dataAdapter.Dispose();
            }

            _error = queryResult;

            return salida;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataTable"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(OleDbCommand sqlCommand, ref DataTable dataTable, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                sqlCommand.Connection = _objOleCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataTable).ToString();
                }


                salida = true;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataTable, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataTable, out queryResult);
                }
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataSet"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(OleDbCommand sqlCommand, ref DataSet dataSet, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                sqlCommand.Connection = _objOleCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataSet).ToString();
                }


                salida = true;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataSet, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = Execute(sqlCommand, ref dataSet, out queryResult);
                }
            }

            _error = queryResult;

            return salida;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="query">string query</param>
        /// <param name="dataTable"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string query, ref DataTable dataTable, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {

                using (SqlConnection conn = new SqlConnection(_strConnectionString))
                {
                    conn.Open();

                    using (SqlCommand sqlCommand = new SqlCommand(query.ToString(), conn))
                    {
                        string commandText = sqlCommand.CommandText.ToUpper();
                        string initialCommand = commandText.Trim().Substring(0, 6);
                        if (
                            initialCommand.Contains("DELETE") ||
                            initialCommand.Contains("UPDATE") ||
                            initialCommand.Contains("INSERT")
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
                            //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet.
                            //Este tipo de consulta, sólo devuelve resultados en queryResult si es una SELECT.
                            using (SqlDataAdapter DataAdapter = new SqlDataAdapter(sqlCommand))
                            {
                                queryResult = DataAdapter.Fill(dataTable).ToString();
                            }
                        }

                        salida = true;

                    }

                    conn.Close();

                    bool timeoutPatterIsFound = false;
                    CheckResult(ref queryResult, out timeoutPatterIsFound);

                    if (timeoutPatterIsFound)
                    {
                        salida = ExecuteStandAlone(query, ref dataTable, out queryResult);
                    }
                }

            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + query;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(query, ref dataTable, out queryResult);
                }


            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="query"></param>
        /// <param name="dataSet"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string query, ref DataSet dataSet, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {

                using (SqlConnection conn = new SqlConnection(_strConnectionString))
                {
                    conn.Open();

                    using (SqlCommand sqlCommand = new SqlCommand(query.ToString(), conn))
                    {
                        string commandText = sqlCommand.CommandText.ToUpper();
                        string initialCommand = commandText.Trim().Substring(0, 6);
                        if (
                            initialCommand.Contains("DELETE") ||
                            initialCommand.Contains("UPDATE") ||
                            initialCommand.Contains("INSERT")
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
                            //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet.
                            //Este tipo de consulta, sólo devuelve resultados en queryResult si es una SELECT.
                            using (SqlDataAdapter DataAdapter = new SqlDataAdapter(sqlCommand))
                            {
                                queryResult = DataAdapter.Fill(dataSet).ToString();
                            }
                        }

                        salida = true;

                    }

                    conn.Close();

                    bool timeoutPatterIsFound = false;
                    CheckResult(ref queryResult, out timeoutPatterIsFound);

                    if (timeoutPatterIsFound)
                    {
                        salida = ExecuteStandAlone(query, ref dataSet, out queryResult);
                    }
                }

            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + query;

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(query, ref dataSet, out queryResult);
                }


            }

            _error = queryResult;

            return salida;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand">Sql Command</param>
        /// <param name="dataTable"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataTable dataTable, out string queryResult)
        {
            SqlDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataTable).ToString();
                }
                
                salida = true;

                CloseSqlConnection();

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataTable, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;
                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataTable, out queryResult);
                }
            }
            finally
            {
                CloseSqlConnection();
                sqlCommand.Dispose();
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand">Sql Command</param>
        /// <param name="dataSet"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataSet dataSet, out string queryResult)
        {
            SqlDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    queryResult = dataAdapter.Fill(dataSet).ToString();
                }

                salida = true;

                CloseSqlConnection();

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataSet, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;
                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataSet, out queryResult);
                }
            }
            finally
            {
                CloseSqlConnection();
                sqlCommand.Dispose();
            }

            _error = queryResult;

            return salida;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand">OleDB command</param>
        /// <param name="dataTable"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(OleDbCommand sqlCommand, ref DataTable dataTable, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objOleCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    dataAdapter = new OleDbDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dataTable).ToString();
                }

                salida = true;

                CloseSqlConnection();

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataTable, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;
                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataTable, out queryResult);
                }
            }
            finally
            {
                CloseSqlConnection();
                sqlCommand.Dispose();
            }

            _error = queryResult;

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataset"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(OleDbCommand sqlCommand, ref DataSet dataSet, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                OpenSqlConnection();

                sqlCommand.Connection = _objOleCn;

                string commandText = sqlCommand.CommandText.ToUpper();
                string initialCommand = commandText.Trim().Substring(0, 6);
                if (
                    initialCommand.Contains("DELETE") ||
                    initialCommand.Contains("UPDATE") ||
                    initialCommand.Contains("INSERT")
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
                    dataAdapter = new OleDbDataAdapter(sqlCommand);
                    queryResult = dataAdapter.Fill(dataSet).ToString();
                }

                salida = true;

                CloseSqlConnection();

                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataSet, out queryResult);
                }
            }
            catch (Exception ex)
            {
                salida = false;
                queryResult = ex.Message + "-" + sqlCommand.CommandText;
                bool timeoutPatterIsFound = false;
                CheckResult(ref queryResult, out timeoutPatterIsFound);

                if (timeoutPatterIsFound)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dataSet, out queryResult);
                }
            }
            finally
            {
                CloseSqlConnection();
                sqlCommand.Dispose();
            }

            _error = queryResult;

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
                    _objCn = new SqlConnection(this._strConnectionString);
                }

                this._objCn.Open();
                MyTrans = this._objCn.BeginTransaction();


                this._conectado = true;
                return true;
            }


            catch (Exception ex)
            {
                this._error = ex.Message;
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
                this._error = ex.Message;
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
                this._error = ex.Message;
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
                this._error = ex.Message;
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
                this._error = ex.Message;
                idInsert = -1;
            }
            return idInsert;
        }


        #region NotImplemented

        /// <summary>
        /// NotImplemented
        /// </summary>
        /// <param name="strSqlExec"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteSqlQuery(string strSqlExec, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// NotImplemented
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(SQLiteCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// NotImplemented
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataset"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(SQLiteCommand sqlCommand, ref DataSet dataset, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// NotImplemented
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
        /// NotImplemented
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
        /// NotImplemented
        /// </summary>
        /// <param name="strPath2Exec"></param>
        /// <param name="strSqlite3Path"></param>
        /// <returns></returns>
        public bool BackupDataBase(string strPath2Exec, string strSqlite3Path)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// NotImplemented
        /// </summary>
        /// <param name="strConnectionString"></param>
        /// <param name="queryResult"></param>
        public void ForceDataBaseUpdate(string strConnectionString, out string queryResult)
        {
            throw new NotImplementedException();
        }

        #endregion NotImplemented



    }
}
