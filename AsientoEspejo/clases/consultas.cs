using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsientoEspejo.clases
{
    class consultas
    {
        public bool existe_tabla(string name_table)
        {
            conexion con = new conexion();
            bool condicion = false;
            string SQL = App_GlobalResources.Resource1.consl_table;
            SQL = SQL.Replace("%0%", "@" + name_table);
            DataTable tabla = con.obtener_datos(SQL, "tabla");
            if (tabla.Rows.Count > 0)
                condicion = true;
            return condicion;
        }
        public bool existe_campo(string name_table, string name_campo)
        {
            conexion con = new conexion();
            bool condicion = false;
            string SQL = App_GlobalResources.Resource1.consl_table_campo;
            SQL = SQL.Replace("%1%", "@" + name_table);
            SQL = SQL.Replace("%0%", "U_" + name_campo);
            DataTable tabla = con.obtener_datos(SQL, "campos");
            if (tabla.Rows.Count > 0)
                condicion = true;
            return condicion;
        }
        public int consecutivo_log()
        {
            int total = 0;
            conexion con = new conexion();
            string sql = App_GlobalResources.Resource1.conse_log;
            DataTable tabla = con.obtener_datos(sql, "consecutivo");
            DataRow code = tabla.Rows[0];
            total = code.Field<int>("nro");
            return total;

        }
        public Boolean insert_log(log pr)
        {
            conexion conn = new conexion();
            string SQL = string.Format("insert into [dbo].[@ASIENTOS_LOG_ERROR] (Code,Name,U_AS_Descrip,U_AS_DocEntry,U_AS_No_Doc,U_AS_Fecha,U_AS_MsjError) values('{0}','{1}','{2}','{3}','{4}','{5}','{6}')", pr.codigo,pr.name, pr.descripcion, pr.docEntry, pr.docNum, pr.fecha, pr.msjError);
            return conn.insertar_datos(SQL);

        }
    }
}
