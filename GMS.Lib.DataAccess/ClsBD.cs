using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;


/// <summary>
/// Summary description for ClsBD
/// </summary>
public class ClsBD
{
    //String strConnection = "User ID=sa;Initial Catalog=UTPBackup_Mad ;Data Source=PORTATIL\\SQLLOLO; password=sa;";

    private string _cadenaConexion = ConfigurationManager.ConnectionStrings["cadenaConexion"].ConnectionString;

    //Ordenador de Rafa
    //private string _cadenaConexion = "User ID=UtpLogin;Initial Catalog=UTPBackup_Mad ;Data Source=10.34.21.131; password=Utplogin;";
    
	//private string _cadenaConexion = "User ID=sa;Initial Catalog=UTPBackup_Mad ;Data Source=carlos; password=sa;";
    
	private SqlConnection _objCN;
    private bool _Conectado = false;
    public string _cadena_error;
    public SqlTransaction myTrans;
  

	public ClsBD()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    public bool AbrirConexion()
    {
        try
        {
            if (this._objCN == null)
            {
                _objCN = new SqlConnection(this._cadenaConexion);
            }

            this._objCN.Open();

            this._Conectado = true;
            return true;
        }
        catch (Exception ex)
        {
            this._cadena_error = ex.Message;
            this._Conectado = false;
            return false;
        }
    }



    public bool CerrarConexion()
    {
   
        try
        {
            if (_Conectado)
            {
                _objCN.Close();

                _objCN = null;

                _Conectado =false;
                this.myTrans = null;
            }
            return true;
        }

        catch (Exception ex)
        {
            this._Conectado = false;
            this._cadena_error = ex.Message;
            this.myTrans = null;
            return false;
        }
    }


    public bool EjecutaSP(string sProcedure,ref DataTable dt)
    {
        bool salida;
        try
        {
          //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
            SqlDataAdapter DataAdapter = new SqlDataAdapter(sProcedure, this._objCN);
            DataAdapter.Fill(dt);
            salida = true;
        }
        catch (Exception ex)
        {
            salida = false;
            this._cadena_error = ex.Message;
        }
        return salida;
    }

  public bool EjecutaSP_transaccion(string sProcedure, ref DataTable dt)
  {
    bool salida;
    try
    {
      //El objeto DataAdapter .NET de proveedor de datos está ajustado para leer registros en un objeto DataSet
      SqlDataAdapter DataAdapter = new SqlDataAdapter(sProcedure, this._objCN);
      DataAdapter.SelectCommand.Transaction = this.myTrans;
      DataAdapter.Fill(dt);
      salida = true;
    }
    catch (Exception ex)
    {
      salida = false;
      this._cadena_error = ex.Message;
    }
    return salida;
  }


    public bool EjecutaSP_command(string sProcedure)
    {
        bool salida;
        try
        {
            SqlCommand cmd = new SqlCommand(sProcedure, this._objCN);
            cmd.ExecuteNonQuery();//Ejecuta comandos como instrucciones INSERT, DELETE, UPDATE y SET de Transact-SQL.
            salida = true;
        }
        //catch (SqlException ex)
        //{
        //    //if (ex.ErrorCode == -2146232060)
        //    //    this._cadena_error = "ESTE REGISTRO NO PUEDE SER ELIMINADO,DEPENDE DE CONDUCTORES";
        //    //else
        //    //    this._cadena_error = ex.Message;
        //    salida = false;
        //}
        catch (Exception ex)
        {
            this._cadena_error = ex.Message;
            salida = false; 
        }
        return salida;
    }

  //********************************************************************************************
  //* Este evento ejecuta en forma de Transacion todas las querys que recibe como parametro de-*
  //* vuelve un bool indicandote si todo ha ido bien                                           *
  //********************************************************************************************
  public bool Transaccion (string [] ArrayInsert)
  {
    string consulta;
    bool error=false;

    SqlTransaction trans=this._objCN.BeginTransaction();

    
    for (int i=0; i < ArrayInsert.Length; i++)
    {
      consulta = ArrayInsert[i];
      SqlCommand cmd = new SqlCommand(consulta, this._objCN,trans);
      try
      {
        cmd.ExecuteNonQuery();
      }
      catch 
      {
        error =true;
      }
    } // fin del for

     if (error)
       trans.Rollback();
     else
       trans.Commit();

    return error;
  }

  public Int64 DameIdInsertado(string Tabla)
    {
      Int64 id_insert;
      string consulta= "select ident_current('"+Tabla+"')";
      try
      {
        SqlCommand cmd = new SqlCommand(consulta, this._objCN);
        id_insert = System.Convert.ToInt64(cmd.ExecuteScalar());
        //ExecuteScalar. --> Ejecuta la consulta y devuelve la primera columna de la primera fila
        // Es bueno usarlo para recuperar un unico valor
       // cmd.ExecuteScalar() + 1;
       
      }
      catch (Exception ex)
      {
        this._cadena_error = ex.Message;
        id_insert = -1;
      }
      return id_insert;
    }

  public bool AbrirConexionTransaccion()
  {
    try
    {
      if (this._objCN == null)
      {
        _objCN = new SqlConnection(this._cadenaConexion);
      }

      this._objCN.Open();
      myTrans = this._objCN.BeginTransaction();


      this._Conectado = true;
      return true;
    }


    catch (Exception ex)
    {
      this._cadena_error = ex.Message;
      this._Conectado = false;
      return false;
    }
  }


  public bool EjecutaSP_command_transacion(string sProcedure)
  {
    bool salida;
    try
    {
      SqlCommand cmd = new SqlCommand(sProcedure, this._objCN);
      cmd.Transaction = myTrans;
      cmd.ExecuteNonQuery();
      salida = true;
    }
    catch (Exception ex)
    {
      salida = false;
    }
    return salida;
  }

  public void commit_transacion()
  {
    myTrans.Commit();
  }

  public void Rollback_transacion()
  {

    myTrans.Rollback();
  }
  
  public Int64 dame_insertado_transaccion(string Tabla)
  {
    Int64 id_insert;
    string consulta = "select ident_current('" + Tabla + "')";
    try
    {
      SqlCommand cmd = new SqlCommand(consulta, this._objCN,this.myTrans);
      id_insert = System.Convert.ToInt64(cmd.ExecuteScalar());
      //ExecuteScalar. --> Ejecuta la consulta y devuelve la primera columna de la primera fila
      // Es bueno usarlo para recuperar un unico valor
    }
    catch (Exception ex)
    {
      this._cadena_error = ex.Message;
      id_insert = -1;
    }
    return id_insert;
  }

}
