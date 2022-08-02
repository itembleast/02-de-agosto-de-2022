using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsientoEspejo.clases
{
    class funciones
    {
        public string inicio()
        {
            string resp = "";
            tabla_log_errores_asientos();
            conexionSAP con = new conexionSAP();
            try
            {
                /*Importante: el orden de los documentos, primero se validan las fac y NC que vienen de pedidos de 
                anfora el resto son asientos completos que pasan a IFRS.
                las NC de clientes que se hacen solo a intereses no se tienen en cuenta en base del pedido*/
                //

                //MessageBox.Show("creacion asientos fact de venta");
                con.fac();
                //MessageBox.Show("creacion asientos Notas credito");
                con.NC();
                //creacion NC parciales que vienen de anfora.
                //con.NCP();
                //MessageBox.Show("creacion asientos manuales");
                con.AsientosManuales();
                //MessageBox.Show("creacion asientos pagos efectuados");
                con.PagosEfectuados();
                //MessageBox.Show("creacion asientos pagos recibidos");
                con.PagosRecibidos();
                //MessageBox.Show("creacion asientos fact. prov.");
                con.facProv();
                //MessageBox.Show("creacion asientos todo el resto.");
                con.doc();
                
            }
            catch (Exception e)
            {
               resp= e.Message.ToString();
            }
            return resp;
            
        }
        public string tabla_log_errores_asientos()
        {
            string resp = "";
            string tabl = "ASIENTOS_LOG_ERROR";
            resp = existe_tabla(tabl, "LOG DE ERRORES", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            if (resp == "Y" || resp == "N")
            {
                List<string> res = new List<string>();
                res.Add(existe_campo(tabl, "AS_DocEntry", "Docentry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, 0));
                res.Add(existe_campo(tabl, "AS_No_Doc", "Numero Documento", SAPbobsCOM.BoFieldTypes.db_Alpha, 40, 0));
                res.Add(existe_campo(tabl, "AS_Descrip", "Descripcion", SAPbobsCOM.BoFieldTypes.db_Memo, 0, 0));
                res.Add(existe_campo(tabl, "AS_MsjError", "mensaje error", SAPbobsCOM.BoFieldTypes.db_Memo, 0,0));
                res.Add(existe_campo(tabl, "AS_Fecha", "Fecha", SAPbobsCOM.BoFieldTypes.db_Alpha, 40, 0));
                resp = "Se ha creado la tabla con sus campos";
            }
            return resp;
        }
        public string existe_tabla(string name_table, string description, SAPbobsCOM.BoUTBTableType type)
        {
            consultas consl = new consultas();
            string resp = "";
            if (consl.existe_tabla(name_table))
                resp = "N";
            else
            {
                conexionSAP add = new conexionSAP();
                try
                {
                    add.crear_tabla(name_table, description, type);
                    resp = "Y";
                }
                catch (Exception e)
                {
                    resp = e.Message.ToString();
                }
            }
            return resp;
        }
        public string existe_campo(string name_table, string name_field, string description, SAPbobsCOM.BoFieldTypes type, int tam, int extra)
        {
            string resp = "";
            consultas consl = new consultas();
            if (consl.existe_tabla(name_table))
            {
                if (consl.existe_campo(name_table, name_field))
                    resp = "N";
                else
                {
                    conexionSAP add = new conexionSAP();
                    try
                    {
                        add.crear_campos_tabla(name_table, name_field, description, type, tam, extra);
                        resp = "Y";
                    }
                    catch (Exception e)
                    {
                        resp = e.Message.ToString();
                    }
                }
            }
            else
            {
                resp = "Debe crear primero la tabla";
            }
            return resp;
        }
    }
}
