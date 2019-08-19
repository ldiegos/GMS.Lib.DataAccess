using System;
using System.ComponentModel.Design;
using System.Data;
using System.Configuration;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Data.SQLite;
using System.Diagnostics;
using System.Timers;
using System.Windows.Forms;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;

//TreeViews


namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// Summary description for ClsBD
    /// </summary>
    public class GmsLibDataAcessSqlite : IGmsLibDataAccess, IDisposable
    {
        // Flag: Has Dispose already been called?
        bool _disposed = false;
        // Instantiate a SafeHandle instance.
        readonly SafeHandle _handle = new SafeFileHandle(IntPtr.Zero, true);

        private SQLiteConnection _objCn;

        /// <summary>
        /// 
        /// </summary>
        private string CadenaError { get; set; }

        /// <summary>
        /// 
        /// </summary>
        private SqlTransaction MyTrans { get; set; }

        private readonly string _strConnectionString;
        private bool _bolConnect = false;

        private int NumberMaxConnectionTries { get; set; } = 3;

        private int NumberMaxExecutionTries { get; set; } = 10;

        private int _countConnectionTries = 0;
        private int _countExcecutionTries = 0;

        private int WaitConnectionMillisecondsTimeout { get; set; } = 10000;
        private int WaitExecutionMillisecondsTimeout { get; set; } = 6000;

        #region "PrivateMethdos"

        /// <summary>
        /// FUNCION: private bool fnCreateDataBase(string strConnectionString)
        ///      Funcion privada. 
        ///      Funci�n que crea la base de datos de auditoria.
        /// </summary>
        /// <param name="strConnectionString">Type string - Cadena completa de conexi�n con la base de datos.</param>
        /// <returns>Devuelve true/false en funci�n de la creaci�n de la base de datos.</returns>
        private bool CreateDatabase(string strConnectionString)
        {
            bool bolResultadoFn = false;
            string[] arrConnectionString;
            string strDataBasePath = "";

            bool isFound = SearchElementInConnectionString(strConnectionString, "data source", ref strDataBasePath);

            if (isFound)
            {

                SQLiteConnection.CreateFile(strDataBasePath);

                bolResultadoFn = FnTestDataBase(strConnectionString);

                if (bolResultadoFn)
                {
                    string strResult = "";
                    ForceDataBaseUpdate(out strResult);
                }
            }
            else
            {
                bolResultadoFn = false;
            }

            return bolResultadoFn;

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
        
        //private void CheckResult(SQLiteCommand sqlCommand, ref DataTable dt, ref string queryResult, ref bool salida)
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

        //private void CheckResult(SQLiteCommand sqlCommand, ref DataSet ds, ref string queryResult, ref bool salida)
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



            /// <summary>
            /// FUNCION: private bool fnTestDataBase(string strConnectionString)
            ///      Funcion privada. 
            ///      Funci�n que comprueba si existe la base de datos de auditor�a.
            /// </summary>
            /// <param name="strConnectionString">Type string - Cadena completa de conexi�n con la base de datos.</param>
            /// <returns>Devuelve true/false en funci�n de la existencia de la base de datos.</returns>
            private bool FnTestDataBase(string strConnectionString)
        {
            string[] arrConnectionString;
            string strDataBasePath = "";
            string strResultadoFn = "";
            bool bolResultadoFn = false;
            arrConnectionString = strConnectionString.Split(';');

            //foreach (string strConnection in arrConnectionString)
            //{
            //    if (strConnection != "" && strConnection.Length >= "data source=".Length && strConnection.Substring(0, 12).ToUpper() == "data source=".ToUpper())
            //    {
            //        string[] arrDataSource;
            //        arrDataSource = strConnection.Split('=');
            //        strDataBasePath = arrDataSource[1];
            //    }
            //}

            bool isFound = SearchElementInConnectionString(strConnectionString, "data source", ref strDataBasePath);

            if (strDataBasePath != "")
            {
                bolResultadoFn = FnMngFFileExists(strDataBasePath, ref strResultadoFn);
            }
            else
            {
                bolResultadoFn = false;
            }

            return bolResultadoFn;

        }

        //private string FnExtractDatabasePathFromConnectionString(string strConnectionString)
        //{
        //    string[] arrConnectionString;
        //    string strDataBasePath = "";

        //    arrConnectionString = strConnectionString.Split(';');

        //    foreach (string strConnection in arrConnectionString)
        //    {
        //        if (strConnection != "" && (strConnection.Length >= "data source=".Length && strConnection.Substring(0, 12).ToUpper() == "data source=".ToUpper()))
        //        {
        //            string[] arrDataSource;
        //            arrDataSource = strConnection.Split('=');
        //            strDataBasePath = arrDataSource[1];
        //        }
        //    }

        //    return strDataBasePath;
        //}

        /// <summary>
        /// FUNCION: private bool fnMngFFileExists(string strFileName, ref string strResultado)
        ///      Funcion sobrecargada. 
        ///      Funci�n que comprueba si existe un fichero y escribre el resultado en una cadena de texto.    
        /// </summary>
        /// <param name="strFileName">Type string - Path y fichero del que se quiere saber si existe</param>
        /// <param name="strResultado">Type ref string - Devuelve un "fichero existe" o "fichero no existe".</param>
        /// <returns>Devuelve true/false en funci�n de la existencia del fichero.</returns>
        private bool FnMngFFileExists(string strFileName, ref string strResultado)
        {
            string strRutacompleta;
            string strDirectoryOrigin;
            int intIndexOfPuntosBarras;

            try
            {
                strRutacompleta = strFileName;

                intIndexOfPuntosBarras = strRutacompleta.LastIndexOf(":\\");

                strDirectoryOrigin = strRutacompleta.Substring(0, intIndexOfPuntosBarras - 1);

                if (strDirectoryOrigin == "file:\\")
                {
                    strRutacompleta = strRutacompleta.Substring(intIndexOfPuntosBarras - 1);
                }


                if (File.Exists(strRutacompleta))
                {
                    strResultado = "ruta existente.";
                    return true;

                }
                else
                {
                    strResultado = "ruta no existente.";
                    return false;

                }
            }
            catch (Exception ex)
            {
                StringBuilder stBResultado = new StringBuilder();

                //Se produce error, ya que el fichero viene directamente y no puede extraer el directorio
                //de la ruta, se asume que el resultado es false porque no sabe donde buscar el fichero.
                stBResultado.Append("Error grave en el path: ex.Source: ");
                stBResultado.Append(ex.Source);
                stBResultado.Append("; Mensaje error: ");
                stBResultado.Append(ex.Message);
                strResultado = stBResultado.ToString();
                return false;


            }
        }

        private void FnCreateTextFileAndExec(StringBuilder strCommand, string strPath2Exec)
        {

            var fullPath = strPath2Exec;

            var strLogPath = FnGlobAnalyzeLogPath("", fullPath);

            string strFicheroEjecutar = strLogPath + "\\SqliteBackup.bat";

            StringBuilder stBCodigoEjectutar = new StringBuilder();

            stBCodigoEjectutar.Append(strCommand)
                ;

            StreamWriter escribe;
            escribe = File.CreateText(strFicheroEjecutar); // A�ade al final del fichero
            escribe.WriteLine(strCommand);
            escribe.Close();

            FnFichExecFile(strFicheroEjecutar);

        }

        private bool FnFichExecFile(string strFicheroEjecutar)
        {
            ProcessStartInfo procInfo = new ProcessStartInfo();

            procInfo.UseShellExecute = true;

            procInfo.FileName = strFicheroEjecutar; //The file in that DIR.

            procInfo.WorkingDirectory = strFicheroEjecutar; //The working DIR.

            procInfo.Verb = "runas";

            Process.Start(procInfo);  //Start that process.


            return true;
        }

        /// <summary>
        /// Busqueda del "./" en el principo del Path que indica que el directorio est� dentro de la ruta actual.
        /// Sino, devolvemos la ruta completa sin el fichero del final, en el caso que lo hubiera.
        /// </summary>
        /// <param name="caller"></param>
        /// <param name="strOriginPath"></param>
        /// <returns></returns>
        public string FnGlobAnalyzeLogPath(string caller, string strOriginPath)
        {
            int intIndexOfPuntoBarra;
            string strDirectoryOrigin;

            string strReturnPath;

            intIndexOfPuntoBarra = strOriginPath.LastIndexOf(@"./");

            strDirectoryOrigin = strOriginPath.Substring(0, intIndexOfPuntoBarra + 1);

            if (strDirectoryOrigin == ".")
            {

                strReturnPath = caller + strOriginPath.Substring(intIndexOfPuntoBarra + 1, strOriginPath.Length - 1);
            }
            else
            {
                int intLastIndexOfBarra = strOriginPath.LastIndexOf(@"\");

                strReturnPath = strOriginPath.Substring(0, intLastIndexOfBarra + 1);
            }

            return strReturnPath;
        }

        #endregion "PrivateMethdos"
        
        #region "PublicMethods"

        /// <summary>
        /// Constructor simple, s�lo cadena de conexion, se cogen los valores de reconexion por defecto.
        /// </summary>
        /// <param name="strConnectionStringBd">Cadena de conexi�n.</param>
        /// <param name="bolFnResult"></param>
        public GmsLibDataAcessSqlite(string strConnectionStringBd, out bool bolFnResult)
        {
            bolFnResult = false;

            _strConnectionString = strConnectionStringBd;

            bool bolExisteDatabase = FnTestDataBase(_strConnectionString);

            if (bolExisteDatabase)
            {
                bolFnResult = true;
            }
            else
            {
                bolFnResult = CreateDatabase(_strConnectionString);
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="strConnectionStringBd">Cadena de conexi�n</param>
        /// <param name="intNumRecon">Numero de reintentos de conexion.</param>
        /// <param name="intNumSecs">Numero de segundos entre reconexion.</param>
        /// <param name="maxExecutionRetries"></param>
        /// <param name="bolFnResult"></param>
        public GmsLibDataAcessSqlite(string strConnectionStringBd, int intNumRecon, int intNumSecs, int maxExecutionRetries, out bool bolFnResult)
        {
            bolFnResult = false;

            _strConnectionString = strConnectionStringBd;

            NumberMaxConnectionTries = intNumRecon;
            WaitConnectionMillisecondsTimeout = intNumSecs;
            NumberMaxExecutionTries = maxExecutionRetries;

            bool bolExisteDatabase = FnTestDataBase(_strConnectionString);

            if (bolExisteDatabase)
            {
                bolFnResult = true;
            }
            else
            {
                bolFnResult = CreateDatabase(_strConnectionString);
            }
        }

        /// <summary>
        /// Destructo of the class.
        /// </summary>
        ~GmsLibDataAcessSqlite()
        {
            Dispose();
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

            if (_objCn != null)
            {
                _objCn.Close();
                _objCn = null;
            }

            // Free any unmanaged objects here.
            //
            _disposed = true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool OpenSqlConnection()
        {
            try
            {
                if (this._objCn == null)
                {
                    this._objCn = new SQLiteConnection(this._strConnectionString);
                    this._objCn.Open();
                }

                if (this._objCn == null)
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
                _objCn.Close();

                _objCn = null;

                this.MyTrans = null;
                return true;
            }

            catch (Exception ex)
            {
                this.CadenaError = ex.Message;
                this.MyTrans = null;
                return false;
            }
        }


        /// <summary>
        /// Search for elements value in the connection string. Every element is separated by the ;(semicolon) and the value with the =(equal)
        /// </summary>
        /// <param name="strConnectionString">Full connection string.</param>
        /// <param name="element2Search">Element to search, ex: data source, compress, synchronous, foreign keys, ect ect...</param>
        /// <param name="resultSearch">Value of the element in the connecton string </param>
        /// <returns>true = found the element and return the value.</returns>
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

        /// <summary>
        /// public void fnSQLiteForceDataBaseUpdate()
        /// function to force the SQLite system to shrink and compress the database.
        /// </summary>
        public void ForceDataBaseUpdate(out string queryResult)
        {
            SQLiteCommand sqlCommandSelect;
            DataTable dtInsert = new DataTable();
            queryResult = "";

            StringBuilder stBsqlSelect = new StringBuilder();

            stBsqlSelect.Append("vacuum");
            sqlCommandSelect = new SQLiteCommand(stBsqlSelect.ToString());

            Execute(sqlCommandSelect, ref dtInsert, out queryResult);
            //Close commands
            stBsqlSelect.Remove(0, stBsqlSelect.Length - 1);
            dtInsert.Clear();
            dtInsert.Dispose();
            sqlCommandSelect.Dispose();


        }

        /// <summary>
        /// Function to backup the sqlite file with datetime extension.
        /// </summary>
        /// <param name="strPath2Exec">Filesystem path to create the temp exec file.</param>
        /// <param name="strSqlite3Path">Filesystem path to the sqlite3.exe </param>
        /// <returns></returns>
        public bool BackupDataBase(string strPath2Exec, string strSqlite3Path)
        {
            StringBuilder stSb = new StringBuilder();
            //string strDataBase = FnExtractDatabasePathFromConnectionString(_strConnectionString);

            string strDataBase = string.Empty;

            bool isFound = SearchElementInConnectionString(_strConnectionString, "data source", ref strDataBase);

            stSb.Append(" For /f \"tokens=1-4 delims=/ \" %%a in ('date /t') do (set mydate=%%c-%%a-%%b) \n ")
            .Append(" For /f \"tokens=1-2 delims=/:\" %%a in ('time /t') do (set mytime=%%a%%b) \n ")
            .Append("\"" + strSqlite3Path + "\" " + strDataBase + " .dump > " + strDataBase + ".sql_%mydate%_%mytime% " + "\n")
            .Append("\"" + strSqlite3Path + "\" " + strDataBase + "_%mydate%_%mytime% < " + strDataBase + ".sql_%mydate%_%mytime% " + "\n")
            .Append("del " + strDataBase + ".sql_%mydate%_%mytime% " + "\n")
            //.Append(" pause \n")
            .Append(" exit 0 \n")
            ;

            FnCreateTextFileAndExec(stSb, strPath2Exec);

            return true;
        }
        
        public bool ExecuteSqlQuery(string strSqlExec, out string queryResult)
        {
            SQLiteCommand sqlCommandSelect;
            DataTable dtSelect = new DataTable();

            string strResultadoDb = "";
            bool bolBdResult = false;
            bool bolResult = false;

            StringBuilder stBsqlSelect = new StringBuilder();

            stBsqlSelect.Append(strSqlExec);

            sqlCommandSelect = new SQLiteCommand(stBsqlSelect.ToString());

            bolBdResult = ExecuteStandAlone(sqlCommandSelect, ref dtSelect, out strResultadoDb);
            
            if (bolBdResult)
            {
                if (dtSelect.Rows.Count >= 0)
                {
                    bolResult = true; //AWS has rows to show.
                }
                else
                {
                    bolResult = false;
                }
            }
            else
            {
                bolResult = false;
            }

            //CerrarConexion();
            //Close end commands
            stBsqlSelect.Remove(0, stBsqlSelect.Length - 1);
            dtSelect.Clear();
            dtSelect.Dispose();
            //dtrSelect.Close();
            //dtrSelect.Dispose();
            //dtSelectAllATT.Clear();
            //dtSelectAllATT.Dispose();
            sqlCommandSelect.Dispose();

            queryResult = strResultadoDb;

            return bolResult;

        }

        /// <summary>
        /// Function: public bool EjecutaSP(SQLiteCommand sqlCommand, ref DataTable dt, ref string queryResult)
        ///     Ejecuta un comando SQL de SQLite y devuelve datos en un DataTable. El resultado de la funcion se almacena en una cadena de texto.
        /// </summary>
        /// <param name="sqlCommand">Type SQLiteCommand - Comando SQLite que se desea ejecutar</param>
        /// <param name="dt">Type DataTable - Datatable donde se devuelven los datos de la ejecuci�n</param>
        /// <param name="queryResult">Type string - Cadena de texto con el restultado de la ejecuci�n</param> 
        /// <returns>Devuelve true/false en funci�n de si se ejecuta correctamente.</returns>        
        public bool Execute(SQLiteCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                //using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                //{
                using (SQLiteCommand sqLiteCommand = new SQLiteCommand())
                {
                    sqLiteCommand.CommandText = sqlCommand.CommandText;
                    sqLiteCommand.Connection = _objCn;

                    if (sqlCommand.Parameters.Count != 0)
                    {
                        SQLiteParameterCollection sp = sqlCommand.Parameters;

                        foreach (SQLiteParameter param in sp)
                        {
                            sqLiteCommand.Parameters.Add(param);
                        }

                        sp.Clear();
                    }


                    //------------------------------------
                    string commandText = sqLiteCommand.CommandText.ToUpper();
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
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                        using (SQLiteDataAdapter DataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = DataAdapter.Fill(dt).ToString();
                        }
                    }
                    //------------------------------------


                    //Close all objects.
                    sqlCommand.Dispose();
                    sqlCommand = null;
                }
                //}

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
                this.CadenaError = ex.Message + "-" + sqlCommand.CommandText;
                queryResult = this.CadenaError;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }

                sqlCommand.Dispose();
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
        public bool Execute(SQLiteCommand sqlCommand, ref DataSet dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                //using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                //{
                using (SQLiteCommand sqLiteCommand = new SQLiteCommand())
                {
                    sqLiteCommand.CommandText = sqlCommand.CommandText;
                    sqLiteCommand.Connection = _objCn;

                    if (sqlCommand.Parameters.Count != 0)
                    {
                        SQLiteParameterCollection sp = sqlCommand.Parameters;

                        foreach (SQLiteParameter param in sp)
                        {
                            sqLiteCommand.Parameters.Add(param);
                        }

                        sp.Clear();
                    }


                    //------------------------------------
                    string commandText = sqLiteCommand.CommandText.ToUpper();
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
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                        using (SQLiteDataAdapter DataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = DataAdapter.Fill(dt).ToString();
                        }
                    }
                    //------------------------------------

                    //Close all objects.
                    sqlCommand.Dispose();
                    sqlCommand = null;
                }
                //}

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
                this.CadenaError = ex.Message + "-" + sqlCommand.CommandText;
                queryResult = this.CadenaError;

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = Execute(sqlCommand, ref dt, out queryResult);
                }

                sqlCommand.Dispose();
            }


            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {

                //using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                //{
                using (SQLiteCommand sqLiteCommand = new SQLiteCommand(sProcedure.ToString(), _objCn))
                {
                    string commandText = sProcedure.ToUpper();
                    string initialCommand = commandText.Trim().Substring(0, 6);
                    if (
                        initialCommand.Contains("DELETE") ||
                        initialCommand.Contains("UPDATE") ||
                        initialCommand.Contains("INSERT")
                        )
                    {

                        //if (
                        //    sProcedure.ToUpper().Contains("DELETE") ||
                        //    sProcedure.ToUpper().Contains("UPDATE") ||
                        //    sProcedure.ToUpper().Contains("INSERT")
                        //    )
                        //{
                        isNonQuery = true;
                    }

                    if (isNonQuery)
                    {
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
                        salida = true;
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                        using (SQLiteDataAdapter DataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = DataAdapter.Fill(dt).ToString();
                            salida = true;
                        }
                    }

                    CheckResult(queryResult, ref salida);

                    if (salida)
                    {
                        salida = Execute(sProcedure, ref dt, out queryResult);
                    }
                }
                //}

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
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(string sProcedure, ref DataSet dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                using (SQLiteCommand sqLiteCommand = new SQLiteCommand(sProcedure.ToString(), _objCn))
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
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                        queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
                        salida = true;
                    }
                    else
                    {
                        //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                        //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                        using (SQLiteDataAdapter DataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = DataAdapter.Fill(dt).ToString();
                            salida = true;
                        }
                    }

                    CheckResult(queryResult, ref salida);

                    if (salida)
                    {
                        salida = Execute(sProcedure, ref dt, out queryResult);
                    }

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
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                {
                    using (SQLiteCommand sqLiteCommand = new SQLiteCommand())
                    {
                        sqLiteCommand.CommandText = sqlCommand.CommandText;
                        sqLiteCommand.Connection = conn;

                        conn.Open();

                        if (sqlCommand.Parameters.Count != 0)
                        {
                            SQLiteParameterCollection sp = sqlCommand.Parameters;

                            foreach (SQLiteParameter param in sp)
                            {
                                sqLiteCommand.Parameters.Add(param);
                            }

                            sp.Clear();
                        }


                        //------------------------------------
                        string commandText = sqLiteCommand.CommandText.ToUpper();
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
                            //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
                            queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
                        }
                        else
                        {
                            //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                            //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                            using (SQLiteDataAdapter DataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                            {
                                queryResult = DataAdapter.Fill(dt).ToString();
                            }
                        }
                        //------------------------------------

                        //Close all objects.
                        sqlCommand.Dispose();
                        sqlCommand = null;
                    }

                    //2016/10/03-GeMesoft: new
                    conn.Close();
                }

                salida = true;

                //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message + "-" + sqlCommand.CommandText;
                queryResult = this.CadenaError;

                //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                }

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
        public bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataSet dt, out string queryResult)
        {
            bool salida;
            bool isNonQuery = false;

            try
            {
                using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                {
                    using (SQLiteCommand sqLiteCommand = new SQLiteCommand())
                    {
                        sqLiteCommand.CommandText = sqlCommand.CommandText;
                        sqLiteCommand.Connection = conn;

                        conn.Open();

                        if (sqlCommand.Parameters.Count != 0)
                        {
                            SQLiteParameterCollection sp = sqlCommand.Parameters;

                            foreach (SQLiteParameter param in sp)
                            {
                                sqLiteCommand.Parameters.Add(param);
                            }

                            sp.Clear();
                        }

                        //------------------------------------
                        if (
                            sqLiteCommand.CommandText.ToUpper().Contains("DELETE") ||
                            sqLiteCommand.CommandText.ToUpper().Contains("UPDATE") ||
                            sqLiteCommand.CommandText.ToUpper().Contains("INSERT")
                            )
                        {
                            isNonQuery = true;
                        }

                        if (isNonQuery)
                        {
                            //Este tipo de consulta, s�lo devuelve resultados en queryResult si es un INSERT,DELETE,UPDATE.
#if DEBUG
                            //queryResult = "Timeout Expired";
                             queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
#else
                            queryResult = sqLiteCommand.ExecuteNonQuery().ToString();
#endif
                        }
                        else
                        {
                            //El objeto DataAdapter .NET de proveedor de datos est� ajustado para leer registros en un objeto DataSet.
                            //Este tipo de consulta, s�lo devuelve resultados en queryResult si es una SELECT.
                            using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                            {
#if DEBUG
                            //queryResult = "Timeout Expired";
                            queryResult = dataAdapter.Fill(dt).ToString();
#else
                            queryResult = dataAdapter.Fill(dt).ToString();
#endif
                            }
                        }
                        //------------------------------------

                        salida = true;

                        //Close all objects.
                        sqlCommand.Dispose();
                        sqlCommand = null;
                    }

                    //2016/10/03-GeMesoft: new
                    conn.Close();

                    //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);

                    CheckResult(queryResult, ref salida);

                    if (salida)
                    {
                        salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                    }

                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message + "-" + sqlCommand.CommandText;
                queryResult = this.CadenaError;

                //CheckResult(sqlCommand, ref dt, ref queryResult, ref salida);

                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sqlCommand, ref dt, out queryResult);
                }


            }

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string sProcedure, ref DataTable dt, out string queryResult)
        {
            bool salida;

            try
            {

                using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                {
                    using (SQLiteCommand sqLiteCommand = new SQLiteCommand(sProcedure.ToString(), conn))
                    {
                        conn.Open();

                        using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = dataAdapter.Fill(dt).ToString();
                            salida = true;
                        }

                        conn.Close();

                        //CheckResult(sProcedure, ref dt, ref queryResult, ref salida);
                        CheckResult(queryResult, ref salida);

                        if (salida)
                        {
                            salida = ExecuteStandAlone(sProcedure, ref dt, out queryResult);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                //CheckResult(sProcedure, ref dt, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sProcedure, ref dt, out queryResult);
                }

            }

            return salida;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sProcedure"></param>
        /// <param name="ds"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool ExecuteStandAlone(string sProcedure, ref DataSet ds, out string queryResult)
        {
            bool salida;

            try
            {

                using (SQLiteConnection conn = new SQLiteConnection(_strConnectionString))
                {
                    using (SQLiteCommand sqLiteCommand = new SQLiteCommand(sProcedure.ToString(), conn))
                    {
                        conn.Open();

                        using (SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(sqLiteCommand))
                        {
                            queryResult = dataAdapter.Fill(ds).ToString();
                            salida = true;
                        }

                        conn.Close();

                        //CheckResult(sProcedure, ref ds, ref queryResult, ref salida);
                        CheckResult(queryResult, ref salida);

                        if (salida)
                        {
                            salida = ExecuteStandAlone(sProcedure, ref ds, out queryResult);
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                salida = false;
                this.CadenaError = ex.Message;
                queryResult = this.CadenaError;

                //CheckResult(sProcedure, ref ds, ref queryResult, ref salida);
                CheckResult(queryResult, ref salida);

                if (salida)
                {
                    salida = ExecuteStandAlone(sProcedure, ref ds, out queryResult);
                }


            }

            return salida;
        }

        #endregion "PublicMethods"

        #region "NotImplemented"

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dt"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        public bool Execute(SqlCommand sqlCommand, ref DataTable dt, out string queryResult)
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

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool OpenOleDbConnection()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public bool CloseOleDbConnection()
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
        public bool Execute(OleDbCommand sqlCommand, ref DataTable dt, out string queryResult)
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
        public bool ExecuteStandAlone(OleDbCommand sqlCommand, ref DataTable dt, out string queryResult)
        {
            throw new NotImplementedException();
        }
        

        #endregion "NotImplemented"

    }
}