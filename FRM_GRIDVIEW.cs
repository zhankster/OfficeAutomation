﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing.Printing;

namespace OfficeAutomation
{
    public partial class FRM_GRIDVIEW : Form
    {
        public FRM_GRIDVIEW()
        {
            InitializeComponent();
        }

        DataTable dataTable = new DataTable();
        SqlDataAdapter dataAdapter;
        BindingSource bindingSource1 = new BindingSource();
        string conn = Properties.Settings.Default.RxBackend;
        //private export2Excel export2XLS;

        private bool LOAD_GRID()
        {
            try
            {
                Utility.WriteActivity("Loading data grid.");
                //string sql = "SELECT * FROM MANUAL_CHARGES";

                dataTable.Clear();
                dataAdapter = new SqlDataAdapter(Tag.ToString(), conn);
                //dataAdapter = new SqlDataAdapter(sql, conn);
                using (SqlConnection myConnection = new SqlConnection(conn))
                {

                    using (SqlCommandBuilder commandBuilder = new SqlCommandBuilder(dataAdapter))
                    {

                        dataTable.Locale = System.Globalization.CultureInfo.InvariantCulture;
                        dataAdapter.Fill(dataTable);

                        bindingSource1.DataSource = dataTable;

                        DGV.DataSource = bindingSource1;


                    }
                }


                splitContainer2.Panel2Collapsed = true;
                toolStripButton4.Enabled = false;

                bool EXIST = false;

                foreach (DataGridViewColumn COL in DGV.Columns)
                {
                    if (COL.HeaderText.ToUpper() == "IMAGE")
                    {
                        COL.Visible = false;
                        EXIST = true;

                        break;
                    }
                }


                if (EXIST)
                {
                    toolStripButton4.Enabled = true;
                    BTN_RESET.Enabled = true;
                    BTN_IMAGE_CHANGE.Enabled = true;
                    for (int x = 0; x != DGV.RowCount; x++)
                    {
                        if (DGV["IMAGE", x].Value.ToString() == "System.Byte[]")
                        {
                            toolStripButton4.PerformClick();
                            break;
                        }
                    }

                }




                foreach (DataGridViewColumn DC in DGV.Columns)
                {
                    CBO_COLUMN.Items.Add(DC.HeaderText);
                }

                CBO_COLUMN.SelectedIndex = 0;
                Utility.WriteActivity("Data grid loaded.");
                return true;

            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.Message);
                return false;

            }
        }

        private void FRM_GRIDVIEW_Load(object sender, EventArgs e)
        {
            try
            {
                splitContainer1.Panel1Collapsed = true;

                if (LOAD_GRID())
                {
                    CBO_OPERATOR.Text = "<>";

                    saveToolStripButton.Enabled = (!DGV.ReadOnly);
                    toolStripButton2.Enabled = (!DGV.ReadOnly);
                    DGV.AllowUserToAddRows = (!DGV.ReadOnly);
                    DGV.AllowUserToDeleteRows = (!DGV.ReadOnly);

                    this.DialogResult = DialogResult.OK;
                }



            }

            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }

            Utility.WriteActivity(Text + " loaded.");
        }


        private void saveToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                TXT_FILTER.Focus();

                if (dataTable.GetChanges() == null)
                {
                    throw new Exception("No changes exists to be saved.");
                }
                else if (MessageBox.Show("Update " + Text + "?", "CONFIRM", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    saveToolStripButton.Enabled = false;



                    Cursor = Cursors.WaitCursor;

                    using (SqlConnection myConnection = new SqlConnection(conn))
                    {
                        myConnection.Open();

                        SqlCommandBuilder cmdConfig = new SqlCommandBuilder(dataAdapter);
                        dataTable.GetChanges();
                        dataAdapter.Update(dataTable);
                        dataTable.AcceptChanges();

                        myConnection.Close();
                    }
                    Cursor = Cursors.Default;
                    MessageBox.Show("Save completed.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
            Cursor = Cursors.Default;
            saveToolStripButton.Enabled = true;
        }
        private void TXT_FILTER_TextChanged(object sender, EventArgs e)
        {
            try
            {

                APPLY_FILTER();

            }
            catch (Exception)
            {
                //PUB.RAISE_ERROR(EX);
            }

            Cursor = Cursors.Default;
        }

        private void APPLY_FILTER()
        {
            try
            {

                Cursor = Cursors.WaitCursor;

                bindingSource1.RemoveFilter();

                if (toolStripButton3.Checked)
                {

                    if (CBO_OPERATOR.Text.ToUpper().Contains("LIKE"))
                    {

                        bindingSource1.Filter = CBO_COLUMN.Text + " " + CBO_OPERATOR.Text + " '*" + TXT_FILTER.Text + "*'";

                    }
                    else
                    {

                        bindingSource1.Filter = CBO_COLUMN.Text + " " + CBO_OPERATOR.Text + " '" + TXT_FILTER.Text + "'";

                    }
                }

            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }

            Cursor = Cursors.Default;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                //PrintDGV.Print_DataGridView(DGV, Text);
            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
        }

        private void FRM_GRIDVIEW_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (dataTable.GetChanges() != null && MessageBox.Show("Changes have been made without saving.\r\nExit without saving changes?", "CONFIRM EXIT", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                if (MessageBox.Show("Reject all changes and refresh data?", "CONFIRM", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {

                    dataTable.RejectChanges();
                    if (LOAD_GRID())
                    {
                        MessageBox.Show("All changes have been rejected and data has been refreshed.", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            try
            {
                saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (!saveFileDialog1.FileName.Equals(String.Empty))
                    {
                        System.IO.FileInfo f = new System.IO.FileInfo(saveFileDialog1.FileName);
                        if (f.Extension.Equals(".xls"))
                        {

                            StartExport(saveFileDialog1.FileName);
                        }
                        else
                        {
                            MessageBox.Show("Invalid file type");
                        }
                    }
                    else
                    {
                        MessageBox.Show("You did pick a location to save file to");
                    }
                }
            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }

        }
        private void StartExport(String filepath)
        {
            //toolStripButton5.Enabled = false;
            //BackgroundWorker bg = new BackgroundWorker();
            //bg.DoWork += new DoWorkEventHandler(bg_DoWork);
            //bg.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bg_RunWorkerCompleted);
            //bg.RunWorkerAsync(filepath);
            //export2XLS = new export2Excel();

        }


        private void bg_DoWork(object sender, DoWorkEventArgs e)
        {
            //FRM_INPUTBOX frm = new FRM_INPUTBOX();
            //frm.LBL_STATIC.Text = "Enter a name for this worksheet.";
            //frm.Text = "Worksheet Name";

            //if (frm.ShowDialog() == DialogResult.OK)
            //{
            //    export2XLS.ExportToExcel(dataTable.DefaultView, (String)e.Argument, frm.TXT_RESULT.Text);
            //}
        }


        private void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            toolStripButton5.Enabled = true;
        }

        private void toolStripSeparator2_Click(object sender, EventArgs e)
        {

        }

        private void DGV_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            groupBox1.Text = "Data (" + Convert.ToString(DGV.RowCount - 1) + ")";
        }

        private void DGV_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            groupBox1.Text = "Data (" + Convert.ToString(DGV.RowCount - 1) + ")";
        }

        private void BTN_IMAGE_CHANGE_Click(object sender, EventArgs e)
        {
            try
            {
                if (DLG_IMAGECHANGE.ShowDialog() == DialogResult.OK)
                {
                    Cursor.Current = Cursors.WaitCursor;
                    //load file
                    PB_USER_IMAGE.ImageLocation = DLG_IMAGECHANGE.FileName;

                    Image img = Image.FromFile(DLG_IMAGECHANGE.FileName);
                    DGV["IMAGE", DGV.SelectedRows[0].Index].Value = img;

                }

            }

            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
            Cursor.Current = Cursors.Default;
        }

        private void DGV_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (BTN_IMAGE_CHANGE.Enabled)
                {
                    //PB_USER_IMAGE.Image = Properties.Resources.no_image;

                    //if (DGV.SelectedRows.Count > 0 && !Convert.IsDBNull(DGV["IMAGE", DGV.SelectedRows[0].Index].Value))
                    //{
                    //    try
                    //    {
                    //        byte[] bytes = (byte[])(DGV["IMAGE", DGV.SelectedRows[0].Index].Value);
                    //        System.IO.MemoryStream ms = new System.IO.MemoryStream(bytes, false);
                    //        Image img = Image.FromStream(ms);
                    //        this.PB_USER_IMAGE.Image = img;
                    //        bytes = null;
                    //        ms = null;
                    //        img = null;
                    //    }
                    //    catch (Exception EX)
                    //    {
                    //        Utility.WriteActivity(EX.ToString());
                    //    }
                    //}
                }

            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
        }

        private void BTN_RESET_Click(object sender, EventArgs e)
        {
            //Cursor = Cursors.WaitCursor;
            //try
            //{
            //    Image img = Properties.Resources.no_image;
            //    DGV["IMAGE", DGV.SelectedRows[0].Index].Value = img;
            //    PB_USER_IMAGE.Image = img;

            //}
            //catch (Exception EX)
            //{
            //    PUB.RAISE_ERROR(EX);
            //}
            //Cursor = Cursors.Default;
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                splitContainer1.Panel1Collapsed = (!toolStripButton3.Checked);
                APPLY_FILTER();
            }
            catch (Exception EX)
            {
                Utility.WriteActivity(EX.ToString());
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            splitContainer2.Panel2Collapsed = (!toolStripButton4.Checked);
        }

        private void CBO_COLUMN_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (CBO_COLUMN.Text.ToUpper().Contains("NUM"))
                {
                    TXT_FILTER.Text = "0";
                }
                APPLY_FILTER();
            }
            catch (Exception)
            {
                //PUB.RAISE_ERROR(EX);
            }

            Cursor = Cursors.Default;
        }

        private void FRM_GRIDVIEW_Activated(object sender, EventArgs e)
        {
            //PUB.ME.TXT_SQL.Text = Tag.ToString();
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            LOAD_GRID();
        }

        private void DGV_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mStaged = new ContextMenu();

                mStaged.MenuItems.Add(new MenuItem("Delete Selected Rows", (o, ev) =>
                {
                    if (DGV.SelectedRows.Count > 0)
                    {
                        foreach (DataGridViewRow item in DGV.SelectedRows)
                        {
                            DGV.Rows.RemoveAt(item.Index);
                        }
                    }
                    else
                    {  //optional    
                        MessageBox.Show("Please select a row");
                    }
                }));

                mStaged.MenuItems.Add(new MenuItem("Copy Selected Rows", (o, ev) =>
                {
                    if (DGV.GetCellCount(DataGridViewElementStates.Selected) > 0)
                    {
                        try
                        {
                            Clipboard.SetDataObject(
                                DGV.GetClipboardContent());
                        }
                        catch (System.Runtime.InteropServices.ExternalException)
                        {
                            // "The Clipboard could not be accessed. Please try again.";
                        }
                    }
                }));

                mStaged.Show(DGV, new Point(e.X, e.Y));

            }
        }
    }
}
