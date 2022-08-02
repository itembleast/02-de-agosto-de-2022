using AsientoEspejo.clases;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AsientoEspejo
{
    public partial class principal : Form
    {
        public principal()
        {
            InitializeComponent();
           
        }

        private void principal_Load(object sender, EventArgs e)
        {
            funciones func = new funciones();
            func.inicio();
            Close();
        }
    }
}