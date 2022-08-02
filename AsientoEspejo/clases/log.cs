using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsientoEspejo.clases
{
    class log
    {
        public string fecha { set; get; }
        public string descripcion { set; get; }
        public string docNum { set; get; }
        public string docEntry { set; get; }
        public string msjError { set; get; }
        public string name { set; get; }
        public string codigo { set; get; }
        public log()
        {
            docNum = "";
            docEntry = "";
            descripcion = "";
            fecha = "";
            msjError = "";
        }
        public string fecha_hoy()
        {
            DateTime thisDay = DateTime.Now;
            return thisDay.ToString();
        }
        public int consecutivo()
        {
            int consecutivo = 0;
            consultas con = new consultas();
            if (!string.IsNullOrEmpty(con.consecutivo_log().ToString()))
                consecutivo = con.consecutivo_log();
            else
                consecutivo = -1;
            return (consecutivo + 1);
        }

    }
}
