using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsientoEspejo.clases
{
    class conexion
    {
        protected SqlConnection con_db;
        public conexion()
        {
            con_db = new SqlConnection();
            con_db.ConnectionString = ConfigurationManager.ConnectionStrings["pruebas"].ConnectionString;
            //con_db.ConnectionString = ConfigurationManager.ConnectionStrings["productiva"].ConnectionString;

        }
        public DataTable obtener_datos(string sql, string Tabla)
        {
            SqlCommand sqlAdap;
            SqlDataReader ds;
            DataTable retVal = new DataTable(Tabla);
            try
            {
                con_db.Open();
                sqlAdap = new SqlCommand(sql, con_db);
                ds = sqlAdap.ExecuteReader();
                retVal.Load(ds);
                ds.Close();
                sqlAdap.Dispose();
                con_db.Close();
            }
            catch (Exception ex)
            {
                throw new Exception("Error de conexion" + ex);
            }
            finally
            {
                sqlAdap = null;
                con_db.Close();

            }
            return retVal;
        }
        public Boolean insertar_datos(string query)
        {
            Boolean resp = false;
            SqlCommand sqlAdap;
            try
            {
                con_db.Open();
                sqlAdap = con_db.CreateCommand();
                sqlAdap.CommandType = CommandType.Text;
                sqlAdap.CommandText = query;
                int val = sqlAdap.ExecuteNonQuery();
                if (val > 0)
                    resp = true;

            }
            catch (Exception ex)
            {
                throw new Exception("Error de conexion" + ex);
            }
            finally
            {
                sqlAdap = null;
                con_db.Close();
            }
            return resp;
        }
    }
}
