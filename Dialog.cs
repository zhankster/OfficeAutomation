using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeAutomation
{
    public partial class Dialog : Form
    {
        public string Title { get; set; }
        public string Msg { get; set; }

        public Dialog()
        {
            InitializeComponent();
        }

        private void Dialog_Load(object sender, EventArgs e)
        {
            this.Text = Title;
            this.lbMsg.Text = Msg;
        }
    }
}
