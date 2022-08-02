using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsientoEspejo.clases
{
    class conexionSAP
    {
        protected SAPbobsCOM.Company Empresa;
        public void crear_tabla(string Name, string Description, SAPbobsCOM.BoUTBTableType tipo)
        {
            if (conexion())
            {
                int lRetCode;
                SAPbobsCOM.UserTablesMD userTables;
                try
                {

                    userTables = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                    userTables.TableName = Name;
                    userTables.TableDescription = Description;
                    userTables.TableType = tipo;
                    lRetCode = userTables.Add();
                    if (lRetCode != 0)
                    {
                        throw new Exception(Empresa.GetLastErrorDescription());

                    }


                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                finally
                {
                    userTables = null;
                    GC.Collect();
                    Empresa.Disconnect();
                }
            }
            else
            {
                GC.Collect();
                Empresa.Disconnect();
                throw new Exception("error de conexion");
            }

        }
        public void crear_campos_tabla(string name_table, string name_campo, string descripcion, SAPbobsCOM.BoFieldTypes tipo, int tam, int extra)
        {
            if (conexion())
            {
                int lRetCode;
                SAPbobsCOM.UserFieldsMD oUserFieldsMD;
                try
                {
                    oUserFieldsMD = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUserFieldsMD.TableName = name_table;
                    oUserFieldsMD.Name = name_campo;
                    oUserFieldsMD.Description = descripcion;
                    oUserFieldsMD.Type = tipo;
                    if (tam != 0)
                        oUserFieldsMD.EditSize = tam;
                    if (extra == 1)
                    {
                        oUserFieldsMD.ValidValues.Value = "Y";
                        oUserFieldsMD.ValidValues.Description = "SI";
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.ValidValues.Value = "N";
                        oUserFieldsMD.ValidValues.Description = "NO";
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.DefaultValue = "N";
                    }
                    else if (extra == 2)
                    {
                        oUserFieldsMD.ValidValues.Value = "1";
                        oUserFieldsMD.ValidValues.Description = "Adquiriente";
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.ValidValues.Value = "2";
                        oUserFieldsMD.ValidValues.Description = "Factura de Venta";
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.ValidValues.Value = "3";
                        oUserFieldsMD.ValidValues.Description = "Nota Credito";
                        oUserFieldsMD.ValidValues.Add();
                        oUserFieldsMD.ValidValues.Value = "4";
                        oUserFieldsMD.ValidValues.Description = "Nota Debito";
                        oUserFieldsMD.ValidValues.Add();
                    }

                    lRetCode = oUserFieldsMD.Add();
                    if (lRetCode != 0)
                    {

                        throw new Exception(Empresa.GetLastErrorDescription());

                    }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                }
                catch (Exception ex)
                {

                    throw new Exception(ex.Message);
                }
                finally
                {
                    Empresa.Disconnect();
                }
            }
            else
            {
                Empresa.Disconnect();
                throw new Exception("error en la conexion");
               // MessageBox.Show("error en la conexión");
            }
        }
        public conexionSAP()
        {
            Empresa = new SAPbobsCOM.Company();
        }
        public bool conexion()
        {
            bool condicion = false;
            try
            {
                Empresa.Server = ConfigurationManager.AppSettings.Get("Server").ToString();
                Empresa.LicenseServer = ConfigurationManager.AppSettings.Get("LicenseServer").ToString();
                Empresa.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                Empresa.DbUserName = ConfigurationManager.AppSettings.Get("ServerDBUser").ToString();
                Empresa.DbPassword = ConfigurationManager.AppSettings.Get("ServerDBPass");
                Empresa.CompanyDB = ConfigurationManager.AppSettings.Get("BD2").ToString();
                Empresa.UserName = ConfigurationManager.AppSettings.Get("UserSAP1").ToString();
                Empresa.Password = ConfigurationManager.AppSettings.Get("PasswordSAP1").ToString();
                Empresa.UseTrusted = false;
                int estado = Empresa.Connect();
                if (estado != 0)
                {
                    //MessageBox.Show("conexion fallo");
                    throw new Exception(Empresa.GetLastErrorDescription());
                }
                else
                //MessageBox.Show("conexion realizada correctamente");
                condicion = true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show("desconexión " + ex.Message);
                Empresa.Disconnect();
                throw new Exception(ex.Message);
            }
            return condicion;

        }
        public string doc()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.doc_pendientes;
                Query = Query.Replace("%0%", "");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;//52355;
                    Number = oRecordset.Fields.Item("DocNum").Value; //2068;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.cabecera;
                    Query1 = Query1.Replace("%0%", Number.ToString());
                    Query1 = Query1.Replace("%1%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);

                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("BaseRef").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.lineas;
                    Query2 = Query2.Replace("%0%", Number.ToString());
                    Query2 = Query2.Replace("%1%", DocEntry.ToString());
                    Query2 = Query2.Replace("%2%", TransID.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_CTA_IFRS").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_CTA_IFRS").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value;
                            JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                            JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                       // MessageBox.Show("Creado el asiento" + Number);
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();
                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }
        public string fac()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines, oRecordsetLinesSinPedido;
            string Query, Query1, Query2, Query3;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.doc_pendientes;
                Query = Query.Replace("%0%", "where Z0.Tipo='01 - Factura de Venta'");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.cabecera;
                    Query1 = Query1.Replace("%0%", Number.ToString());
                    Query1 = Query1.Replace("%1%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);
                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("BaseRef").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLinesSinPedido = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query3 = App_GlobalResources.Resource1.FactCantLineasSinPedido;
                    Query3 = Query3.Replace("%0%", DocEntry.ToString());
                    oRecordsetLinesSinPedido = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLinesSinPedido.DoQuery(Query3);
                    if (oRecordsetLinesSinPedido.Fields.Item("cant").Value>0)
                    { 
                        oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        Query2 = App_GlobalResources.Resource1.lineasfac2;
                        //Query2 = Query2.Replace("%0%", Number.ToString());
                        Query2 = Query2.Replace("%0%", DocEntry.ToString());
                        //Query2 = Query2.Replace("%2%", TransID.ToString());
                        oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                        oRecordsetLines.DoQuery(Query2);

                        while (!oRecordsetLines.EoF)
                        {
                            double debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                            double credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                            double sumdebitcredit = debit + credit;
                            string CtaIfrs = oRecordsetLines.Fields.Item("U_AcctCode").Value;
                            if (sumdebitcredit != 0)
                            {
                                JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_AcctCode").Value);
                                JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                                JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                                JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("InfoCo01").Value;
                                JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                                //JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                                JournalEntry.Lines.Add();
                            }
                            oRecordsetLines.MoveNext();
                        }
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                        //MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription(); 
                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }

        public string facProv()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines, oRecordsetLinesSinPedido;
            string Query, Query1, Query2, Query3;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.Doc_Pendientes_FacCompra;
                Query = Query.Replace("%0%", "where Z0.Tipo='06 - Factura de Proveedor'");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.cabecera;
                    Query1 = Query1.Replace("%0%", Number.ToString());
                    Query1 = Query1.Replace("%1%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);
                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    var refdate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("BaseRef").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.Lin_Fac_Compra;
                    Query2 = Query2.Replace("%0%", Number.ToString());
                    Query2 = Query2.Replace("%1%", DocEntry.ToString());
                    Query2 = Query2.Replace("%2%", TransID.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_CTA_IFRS").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_CTA_IFRS").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value;
                            JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                            JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                        //MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();

                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }

        public string NC()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            decimal Perc;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.NC_Pendientes;
                Query = Query.Replace("%0%", "where Z0.Tipo='02 - Nota Crédito Clientes' ");
                Query = Query.Replace("%1%", "=");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.Cab_NC;
                    Query1 = Query1.Replace("%0%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);
                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = 23;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("RefDate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("Ref1").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;//oRecordsetCab.Fields.Item("Ref2").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    Perc = Convert.ToDecimal(oRecordsetCab.Fields.Item("Perc").Value);
                    
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.Lin_NC;
                    Query2 = Query2.Replace("%0%", TransID.ToString());
                    Query2 = Query2.Replace("%1%", Perc.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_AcctCode").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_AcctCode").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value; ;
                            JournalEntry.Lines.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                            //JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                       // MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();

                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }


        public string NCP()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            decimal Perc;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.NC_Pendientes;
                Query = Query.Replace("%0%", "where Z0.Tipo='02 - Nota Crédito Clientes' ");
                Query = Query.Replace("%1%", "<>");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.Cab_NCParcial_anfora;
                    Query1 = Query1.Replace("%0%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);
                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = 23;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("RefDate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("Ref1").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;//oRecordsetCab.Fields.Item("Ref2").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    Perc = Convert.ToDecimal(oRecordsetCab.Fields.Item("Perc").Value);

                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.Lin_NC;
                    Query2 = Query2.Replace("%0%", TransID.ToString());
                    Query2 = Query2.Replace("%1%", Perc.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_AcctCode").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_AcctCode").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("U_Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("U_Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value; ;
                            JournalEntry.Lines.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                            //JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                        // MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();
                        
                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }

        public string AsientosManuales()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry;
            int RetVal;
            string resp = "", TransID;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.doc_pendientes;
                Query = Query.Replace("%0%", "where z0.Tipo='10 - Asientos a partir de Asientos Manuales'");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.CabeceraAsientosManuales;
                    Query1 = Query1.Replace("%0%", DocEntry.ToString());
                    oRecordsetCab.DoQuery(Query1);
                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = Number.ToString();
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";
                    

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.LineasAsientosManuales;
                    //Query2 = Query2.Replace("%0%", Number.ToString());
                    //Query2 = Query2.Replace("%1%", DocEntry.ToString());
                    Query2 = Query2.Replace("%2%", TransID.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_CTA_IFRS").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_CTA_IFRS").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value; ;
                            JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                            JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.CostingCode = oRecordsetLines.Fields.Item("Norma_Reparto").Value;
                            JournalEntry.Lines.CostingCode2 = oRecordsetLines.Fields.Item("Norma_Reparto2").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                      //  MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                        //MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();
                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }

        public string PagosRecibidos()
        {   
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry, TransID_primer_asiento;
            int RetVal;
            string resp = "", TransID, CardCode;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.Pagos_Rec_Pendientes;
                Query = Query.Replace("%0%", "");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;
                    Number = oRecordset.Fields.Item("DocNum").Value;
                    TransID_primer_asiento = oRecordset.Fields.Item("TransId_1erAsiento").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.CabeceraRecibosPago;
                    Query1 = Query1.Replace("%0%", Number.ToString());
                    Query1 = Query1.Replace("%1%", TransID_primer_asiento.ToString());
                    oRecordsetCab.DoQuery(Query1);

                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("BaseRef").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    CardCode = oRecordsetCab.Fields.Item("CardCode").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";
                    JournalEntry.UserFields.Fields.Item("U_WP_Espejo").Value = "Y";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.lineas;
                    Query2 = Query2.Replace("%0%", Number.ToString());
                    Query2 = Query2.Replace("%1%", DocEntry.ToString());
                    Query2 = Query2.Replace("%2%", TransID.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_CTA_IFRS").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_CTA_IFRS").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("Credit").Value;
                            if (string.IsNullOrEmpty(oRecordsetLines.Fields.Item("U_InfoCo01").Value))
                            {
                                JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = CardCode;
                            }
                            else
                            {
                                JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value;
                            }

                            JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                            JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }

                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);
                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    
                    if (RetVal == 0)
                    {
                       // MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                        
                    }
                    else
                    {
                       // MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();
                    }
                    //guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
              //  MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }
        
        public string PagosEfectuados()
        {
            Recordset oRecordset, oRecordsetCab, oRecordsetLines;
            string Query, Query1, Query2;
            int Number, DocEntry,TransID_primer_asiento;
            int RetVal;
            string resp = "", TransID;
            log errores = new log();
            consultas consl = new consultas();
            try
            {
                //conexion();
                oRecordset = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Query = App_GlobalResources.Resource1.Pagos_Efec_Pendientes;
                Query = Query.Replace("%0%", "");
                oRecordset.DoQuery(Query);
                while (!oRecordset.EoF)
                {
                    DocEntry = oRecordset.Fields.Item("DocEntry").Value;//52355;
                    Number = oRecordset.Fields.Item("DocNum").Value; //2068;
                    TransID_primer_asiento= oRecordset.Fields.Item("TransId_1erAsiento").Value;

                    oRecordsetCab = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query1 = App_GlobalResources.Resource1.Cabecera_Pagos_Efec;
                    Query1 = Query1.Replace("%0%", Number.ToString());
                    Query1 = Query1.Replace("%1%", TransID_primer_asiento.ToString());
                    oRecordsetCab.DoQuery(Query1);

                    SAPbobsCOM.JournalEntries JournalEntry = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                    //Campos de cabecera
                    JournalEntry.Series = oRecordsetCab.Fields.Item("series").Value;
                    JournalEntry.TaxDate = oRecordsetCab.Fields.Item("TaxDate").Value;
                    JournalEntry.DueDate = oRecordsetCab.Fields.Item("DueDate").Value;
                    JournalEntry.ReferenceDate = oRecordsetCab.Fields.Item("refdate").Value;
                    JournalEntry.Memo = oRecordsetCab.Fields.Item("Memo").Value;
                    JournalEntry.ProjectCode = oRecordsetCab.Fields.Item("Project").Value;
                    JournalEntry.Reference = oRecordsetCab.Fields.Item("BaseRef").Value;
                    JournalEntry.Reference2 = oRecordsetCab.Fields.Item("TransID").Value;
                    TransID = oRecordsetCab.Fields.Item("TransID").Value;
                    JournalEntry.Reference3 = oRecordsetCab.Fields.Item("ref3").Value;
                    //campos de usuario para cabecera
                    JournalEntry.UserFields.Fields.Item("U_OK1_IFRS").Value = "I";

                    //Lineas
                    oRecordsetLines = Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    Query2 = App_GlobalResources.Resource1.lineas;
                    Query2 = Query2.Replace("%0%", Number.ToString());
                    Query2 = Query2.Replace("%1%", DocEntry.ToString());
                    Query2 = Query2.Replace("%2%", TransID.ToString());
                    oRecordsetLines = ((SAPbobsCOM.Recordset)(Empresa.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
                    oRecordsetLines.DoQuery(Query2);

                    while (!oRecordsetLines.EoF)
                    {
                        double debit = oRecordsetLines.Fields.Item("Debit").Value;
                        double credit = oRecordsetLines.Fields.Item("Credit").Value;
                        double sumdebitcredit = debit + credit;
                        string CtaIfrs = oRecordsetLines.Fields.Item("U_CTA_IFRS").Value;
                        if (sumdebitcredit != 0)
                        {
                            JournalEntry.Lines.AccountCode = (oRecordsetLines.Fields.Item("U_CTA_IFRS").Value);
                            JournalEntry.Lines.Debit = oRecordsetLines.Fields.Item("Debit").Value;
                            JournalEntry.Lines.Credit = oRecordsetLines.Fields.Item("Credit").Value;
                            JournalEntry.Lines.UserFields.Fields.Item("U_InfoCo01").Value = oRecordsetLines.Fields.Item("U_InfoCo01").Value;
                            JournalEntry.Lines.ProjectCode = oRecordsetLines.Fields.Item("Project").Value;
                            JournalEntry.Lines.LineMemo = oRecordsetLines.Fields.Item("LineMemo").Value;
                            JournalEntry.Lines.Add();
                        }
                        oRecordsetLines.MoveNext();
                    }
                    RetVal = JournalEntry.Add();
                    //liberacion de uso.
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(JournalEntry);


                    errores.docNum = Number.ToString();
                    errores.docEntry = DocEntry.ToString();
                    if (RetVal == 0)
                    {
                       // MessageBox.Show("Creado el asiento" + Number);
                        errores.descripcion = "creación exitosa";
                        errores.msjError = "Se  ha creado  correctamente el asiento con referencia 1: " + Number;
                    }
                    else
                    {
                       // MessageBox.Show("error al Crear el asiento" + Number);
                        errores.descripcion = "error";
                        resp = "Error al crear el documento, DocEntry: " + errores.docEntry + " - DocNum: " + errores.docNum + " " + Empresa.GetLastErrorDescription();
                    }
                    // guardar las variables del log y insertar la informacion en la tabla
                    errores.fecha = errores.fecha_hoy();
                    errores.msjError = resp;
                    errores.codigo = errores.consecutivo().ToString();
                    errores.name = errores.codigo;
                    consl.insert_log(errores);

                    oRecordset.MoveNext();
                }
            }
            catch (Exception ex)
            {
              // MessageBox.Show("error " + ex.Message.ToString());
                errores.descripcion = "error";
                errores.fecha = errores.fecha_hoy();
                errores.msjError = ex.Message.ToString();
                errores.codigo = errores.consecutivo().ToString();
                errores.name = errores.codigo;
                consl.insert_log(errores);
                throw new Exception(ex.Message);
            }
            return resp;
        }
    }
}
