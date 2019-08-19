using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GMS.LIB.DataAccess
{
    /// <summary>
    /// 
    /// </summary>
    public interface IGmsLibDataAccess
    {
        /// <summary>
        /// 
        /// </summary>
        void Dispose();

        /// <summary>
        /// Open the conexion to the configured database. This method is valid to SQLServer, SQLite and MSAccess.
        /// </summary>
        /// <returns></returns>
        bool OpenSqlConnection();

        /// <summary>
        /// Close the conexion to the configured database. This method is valid to SQLServer, SQLite and MSAccess.
        /// </summary>
        /// <returns></returns>
        bool CloseSqlConnection();

        /// <summary>
        /// Open the conexion to the configured database. This method is valid to SQLServer and MSAccess.
        /// </summary>
        /// <returns></returns>
        bool OpenOleDbConnection();

        /// <summary>
        /// Close the conexion to the configured database. This method is valid to SQLServer and MSAccess.
        /// </summary>
        /// <returns></returns>
        bool CloseOleDbConnection();

        /// <summary>
        /// This method is not develope, the fact is that with the Execute and ExecuteStandAlone, is it possible to query almost everything.
        /// </summary>
        /// <param name="strSQLExec"></param>
        /// <param name="queryResult"></param>
        /// <returns></returns>
        bool ExecuteSqlQuery(string strSqlExec, out string queryResult);

        /// <summary>
        /// This method will execute a string query into the configured database. The method need an previouly opened connection.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="query"></param>
        /// <param name="datatable">Type DataTable</param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool Execute(string query, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQL command into the configured database. The method need an previouly opened connection.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool Execute(SqlCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a OLDEDB command into the configured database. The method need an previouly opened connection.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool Execute(OleDbCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQLite command into the configured database. The method need an previouly opened connection.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool Execute(SQLiteCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a string query into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="query"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(string query, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQL command into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQL command into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataSet
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataset"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(SqlCommand sqlCommand, ref DataSet dataset, out string queryResult);

        /// <summary>
        /// This method will execute a OLEDB command into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(OleDbCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQLite command into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataTable
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="datatable"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataTable datatable, out string queryResult);

        /// <summary>
        /// This method will execute a SQLite command into the configured database. The method will open a connection to database and close when finished.
        /// The result will be stored in a DataSet
        /// </summary>
        /// <param name="sqlCommand"></param>
        /// <param name="dataset"></param>
        /// <param name="queryResult">Normally filled with error</param>
        /// <returns></returns>
        bool ExecuteStandAlone(SQLiteCommand sqlCommand, ref DataSet dataset, out string queryResult);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="strPath2Exec"></param>
        /// <param name="strSqlite3Path"></param>
        /// <returns></returns>
        bool BackupDataBase(string strPath2Exec, string strSqlite3Path);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="queryResult"></param>
        void ForceDataBaseUpdate(out string queryResult);
    }
}
