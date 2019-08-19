using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using ADOX;

namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// 
    /// </summary>
    public class GmsLibDataAccessMsAccess : IGmsLibDataAccess, IDisposable
    {
        // Flag: Has Dispose already been called?
        bool _disposed = false;
        // Instantiate a SafeHandle instance.
        readonly SafeHandle _handle = new SafeFileHandle(IntPtr.Zero, true);

        private bool _conectado = false;

        private OleDbConnection ObjOleCn { get; set; }
        private SqlConnection ObjSqlCn { get; set; }
        
        /// <summary>
        /// 
        /// </summary>
        private string CadenaError { get; set; }
        
        /// <summary>
        /// 
        /// </summary>
        private SqlTransaction MyTrans { get; set; }

        private string StrConnectionString { get; set; }

        private int IntNumReint { get; set; } = 2;

        private int _intCountReintento = 0;

        private int IntPeriodTimer { get; set; } = 6000;

        private bool _bolConnect = false;

        #region "PublicMethods"

        /// <summary>
        /// Constructor simple, sólo cadena de conexión, se cogen los valores de reconexión por defecto.
        /// </summary>
        /// <param name="strConnectionStringBd">Cadenad de conexión</param>
        public GmsLibDataAccessMsAccess(string strConnectionStringBd, out bool bolFnResult)
        {
            StrConnectionString = strConnectionStringBd;

            bolFnResult = false;

            if (!OpenSqlConnection())
            {
                bolFnResult = CreateDatabase(strConnectionStringBd);
                CloseSqlConnection();
            }

        }


        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="strConnectionStringBd">Cadenad de conexión</param>
        /// <param name="intNumRecon">Numero de reintentos de conexion.</param>
        /// <param name="intNumSecs">Numero de segundos entre reconexion.</param>
        public GmsLibDataAccessMsAccess(string strConnectionStringBd, int intNumRecon, int intNumSecs)
        {
            StrConnectionString = strConnectionStringBd;
            ObjOleCn = new OleDbConnection();
            IntNumReint = intNumRecon;
            IntPeriodTimer = intNumSecs;
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose()
        {
            // Dispose of unmanaged resources.
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 
        /// </summary>
        ~GmsLibDataAccessMsAccess()
        {
            Dispose(false);
        }


        /// <summary>
        /// Abrir conexión estandar tipo SQL Server
        /// </summary>
        /// <returns></returns>
        public bool OpenSqlConnection()
        {
            try
            {
                if (this.ObjOleCn == null)
                {
                    ObjOleCn = new OleDbConnection(this.StrConnectionString);
                }


                this._conectado = Reconexion();

                if (!this._conectado)
                {
                    System.Threading.Thread.Sleep(IntPeriodTimer);

                    _intCountReintento++;
                    if (_intCountReintento <= IntNumReint)
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
        public bool CloseSqlConnection()
        {

            try
            {
                if (_conectado)
                {
                    ObjOleCn.Close();

                    ObjOleCn = null;

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
        public bool OpenOleDbConnection()
        {
            return OpenSqlConnection();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool CloseOleDbConnection()
        {
            return CloseSqlConnection();
        }

        /// <summary>
        /// Function to execute a SQLString as is without parameters, like and "select * from table where field='value'"
        /// This function need to has an open connection before use.
        /// </summary>
        /// <param name="sProcedure">SQL string with the sentences to be executed.</param>
        /// <param name="dt">Result stored in a DataTable.</param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool bolFnReturn;
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
                    OleDbDataAdapter dataAdapter = new OleDbDataAdapter(sProcedure, this.ObjOleCn);
                    queryResult = dataAdapter.Fill(dt).ToString();
                }
                bolFnReturn = true;
            }
            catch (Exception ex)
            {
                bolFnReturn = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;
            }
            return bolFnReturn;
        }

        /// <summary>
        /// Function:  public bool EjecutaSP(SqlCommand sqlCommand, ref DataTable dt, ref string queryResult)
        ///     Ejecuta un comando SQL y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sqlCommand">Type SqlCommand - Comando SQLite que se desea ejecutar</param>
        /// <param name="dt">Type DataTable - Datatable donde se devuelven los datos de la ejecución</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecución</param> 
        /// <returns>Devuelve true/false en función de si se ejecuta correctamente.</returns>
        public bool Execute(OleDbCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            OleDbDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;

            try
            {
                sqlCommand.Connection = this.ObjOleCn;

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
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;
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
        public bool Execute(SqlCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            SqlDataAdapter dataAdapter;
            bool salida;
            bool isNonQuery = false;
            try
            {
                sqlCommand.Connection = ObjSqlCn;

                if (sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
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
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;
            }
            return salida;
        }

        /// <summary>
        /// Function to execute a SQLString as is without parameters, like and "select * from table where field='value'"
        /// This function is autocontent, so the open and close connection is done inside.
        /// </summary>
        /// <param name="sProcedure">SQL string with the sentences to be executed.</param>
        /// <param name="dt">Result stored in a DataTable.</param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool bolFnReturn;
            try
            {
                using (OleDbConnection conn = new OleDbConnection(StrConnectionString))
                {
                    using (OleDbCommand oleDbCommand = new OleDbCommand(sProcedure.ToString(), conn))
                    {
                        conn.Open();

                        using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(oleDbCommand))
                        {
                            queryResult = oleDbDataAdapter.Fill(dt).ToString();
                            bolFnReturn = true;
                        }

                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                bolFnReturn = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;
            }
            return bolFnReturn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
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
                if (
                   sqlCommand.CommandText.ToUpper().Contains("DELETE") ||
                   sqlCommand.CommandText.ToUpper().Contains("UPDATE") ||
                   sqlCommand.CommandText.ToUpper().Contains("INSERT")
                   )
                {
                    isNonQuery = true;
                }

                using (OleDbConnection conn = new OleDbConnection(StrConnectionString))
                {
                    conn.Open();

                    sqlCommand.Connection = conn;
                    //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet

                    if (isNonQuery)
                    {
                        queryResult = sqlCommand.ExecuteNonQuery().ToString();
                    }
                    else
                    {
                        dataAdapter = new OleDbDataAdapter(sqlCommand);
                        queryResult = dataAdapter.Fill(dt).ToString();
                    }

                    salida = true;

                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;
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
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="ds"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataSet ds, out string queryResult)
        {
            throw new NotImplementedException();
        }

        #endregion "PublicMethods"

        #region "PublicStaticMethods"
        public static bool SearchElementInConnectionString(string strConnectionString, string element2Search, ref string resultSearch)
        {
            bool functionResult = false;

            string[] arrConnectionString = strConnectionString.Split(';');

            resultSearch = string.Empty;

            foreach (string strConnection in arrConnectionString)
            {
                if (strConnection != "" && strConnection.ToUpper().Contains(element2Search.ToUpper()))
                {
                    string[] arrDataSource;
                    arrDataSource = strConnection.Split('=');
                    resultSearch = arrDataSource[1];

                    functionResult = true;
                }
            }

            return functionResult;
        }
        #endregion "PublicStaticMethods"

        #region "PrivateMethods"

        ///
        ///  Protected implementation of Dispose pattern.
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

            if (ObjOleCn != null)
            {
                CloseSqlConnection();
                ObjOleCn = null;
            }

            // Free any unmanaged objects here.
            //
            _disposed = true;
        }

        private bool Reconexion()
        {
            bool bolconectado = false;
            try
            {
                this.ObjOleCn.Open();
                bolconectado = true;

            }
            catch (Exception ex)
            {

                bolconectado = false;


            }
            return bolconectado;

        }

        private bool CreateDatabase(string strConnectionString)
        {
            bool result = true;

            CatalogClass cat = new CatalogClass();
            cat.Create(strConnectionString);

            return result;

        }

        #endregion "PrivateMethods"


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
