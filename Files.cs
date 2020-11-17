using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace OfficeAutomation
{

    public partial class Files : Form
    {
        public string Facility_code { get; set; }
        public string Billing_folder { get; set; }

        public DataTable dtFiles;
        public Files()
        {
            InitializeComponent();
            //this.Text = "Files to add for " + facility_code;
        }

        private void Files_Load(object sender, EventArgs e)
        {
            this.Text = "Files to add for " + Facility_code;
            lbFolder.Text = "Folder: '" + Billing_folder + "'";
            Load_Files();
        }

        private void Load_Files()
        {
            String[] files = Directory.GetFiles(Billing_folder);
            dtFiles = new DataTable();
            dtFiles.Columns.Add("File Name");

            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = new FileInfo(files[i]);
                dtFiles.Rows.Add(file.Name);
            }

            gvFiles.DataSource = dtFiles;

        }

        private void txtFacFilter_TextChanged(object sender, EventArgs e)
        {
            dtFiles.DefaultView.RowFilter = string.Format("[{0}] LIKE '%{1}%'", "File Name", txtFacFilter.Text);
        }

        private void gvFiles_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;
            string name = gvFiles.Rows[row].Cells["File Name"].Value.ToString();
            string name_new = System.IO.Path.GetFileNameWithoutExtension(Billing_folder + name) 
                + "_" + Facility_code + ".pdf";
            DialogResult dr = new DialogResult();
            Dialog frm = new Dialog();
            frm.Title = "Rename File";
            frm.Msg = "Do you want to rename the file '" + name + "' to '" + name_new + "'";
            frm.Top = this.Top + ((this.Height / 2) - (frm.Height / 2));
            frm.Left = this.Left + ((this.Width / 2) - (frm.Width / 2));
            frm.StartPosition = FormStartPosition.CenterParent;
            dr = frm.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                File.Move(Billing_folder + name, Billing_folder + name_new);
                Load_Files();
            }
            else if (dr == DialogResult.Cancel)
            {
                return;
            }

        }
    }
}

