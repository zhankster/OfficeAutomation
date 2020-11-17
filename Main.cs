using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Diagnostics;
using CrystalReportsNinja;
using System.Security.Cryptography;
using System.IO;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Net.Mail;
using S22.Imap;
//using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Drawing.Imaging;
using Ghostscript.NET.Rasterizer;
using ImageMagick;
using OfficeAutomation.Resource;
using Microsoft.Office.Interop;

namespace OfficeAutomation
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            CONN_CIPS = prop.CIPS;
            GetSettings();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            gvFac.DataSource = bsFac;
            gvRpt.DataSource = bsRpt;
            gvAccList.DataSource = bsAcc;
            gvCodes.DataSource = bsCodes;
            gvMC.DataSource = bsMC;
            dpBilling.Value = DateTime.Now.AddMonths(-1);
            dpMC_Export.Value = DateTime.Now.AddMonths(-1);
            GetUpdate();
            GetRoles(userName);

            foreach (var item in roles)
            {
                Utility.WriteActivity("Role Included: " + item);
            }


            if (roles.Count < 1)
            {
                tabControl1.TabPages.Remove(tpManualCharges);
                MessageBox.Show("Your Windows user account must be setup for this application", "User Not Setup");
                //Application.Exit();
            }

            if (!roles.Contains("Administrator") && !roles.Contains("Billing"))
            {
                tabControl1.TabPages.Remove(tpBilling);
                tabControl1.TabPages.Remove(tpEmail);
                tabControl1.TabPages.Remove(tpAttach);
                tabControl1.TabPages.Remove(tpAccountList);
                gbMC_Import.Visible = false;
                btnBulkEdit.Visible = false;
            }

            if (!roles.Contains("Administrator") && !roles.Contains("Billing") && !roles.Contains("BillTech"))
            {
                tabControl1.TabPages.Remove(tpFac);
            }

            if (!roles.Contains("Administrator") && !roles.Contains("Billing") && !roles.Contains("PharmTech"))
            {
                tabControl1.TabPages.Remove(tpManualCharges);
            }

        }

        #region Global Vars
        readonly string userName = Environment.UserName;
        OfficeAutomation.Properties.Settings prop = OfficeAutomation.Properties.Settings.Default;
        private BindingSource bsFac = new BindingSource();
        private SqlDataAdapter daFac = new SqlDataAdapter();
        private BindingSource bsBill = new BindingSource();
        private SqlDataAdapter daBill = new SqlDataAdapter();
        private BindingSource bsRpt = new BindingSource();
        private SqlDataAdapter daRpt = new SqlDataAdapter();
        private BindingSource bsAcc = new BindingSource();
        private SqlDataAdapter daAcc = new SqlDataAdapter();
        private BindingSource bsCodes = new BindingSource();
        private SqlDataAdapter daCodes = new SqlDataAdapter();
        private BindingSource bsMC = new BindingSource();
        private SqlDataAdapter daMC = new SqlDataAdapter();
        DataTable dtBilling;
        private BindingSource bsBillSent = new BindingSource();
        private SqlDataAdapter daBillSent = new SqlDataAdapter();
        DataTable dtBillingSent;
        DataTable dtFacEmail;
        BindingSource bsStates = new BindingSource();
        static string CONN_CIPS = "";
        static string CONN_RX = "";
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        static string outputFolder = "";
        static string outputFolderNew = "";
        string cropImgIn = "";
        string cropImgOut = "";
        string DATA_RPT = "";
        string SEND_RPT = "";
        bool DataSet = false;
        bool AccInsert = true;
        string MC_ID = "";
        string Acc_ID = "";
        string Code_ID = "";
        List<string> roles = new List<string>();

        #region SQL for Billing Monthly Report check
        string BILLING_CS = @"SELECT 
	            COUNT(MANUAL_CHARGES.ACCT) As CNT
            FROM
            CIPS.DBO.MANUAL_CHARGES 
	            INNER JOIN  CIPS.DBO.ACCOUNT_LIST ON
                    MANUAL_CHARGES.ACCT  = ACCOUNT_LIST.ACCOUNT 
                 LEFT OUTER JOIN CIPS.DBO.SUMMARY_CS ON
                    MANUAL_CHARGES.ACCT = SUMMARY_CS.ACCT 
            WHERE
                (MANUAL_CHARGES.CAGEGORY NOT LIKE 'PAYMENT%' AND
                MANUAL_CHARGES.CAGEGORY NOT LIKE 'BALANCE FROM QS1%')  AND
                MANUAL_CHARGES.DESCRIPTION  NOT LIKE 'DIS%' AND
            DATEPART(M, MANUAL_CHARGES.DATE) = DATEPART(M, DATEADD(M, -1, GETDATE()))
            AND DATEPART(YYYY, MANUAL_CHARGES.DATE) = DATEPART(YYYY, DATEADD(M, -1, GETDATE()))
            AND MANUAL_CHARGES.ACCT  LIKE ";
        string BILLING_CIPS_WS = @"SELECT 
	            COUNT(FIL.ID)  As CNT
            FROM CIPS_WHOLESALE.DBO.FIL 
	            INNER JOIN CIPS_WHOLESALE.DBO.DRG ON
		            FIL.DRG_ID = DRG.ID
	            INNER JOIN CIPS_WHOLESALE.DBO.FAC ON
		            FIL.FAC_ID = FAC.ID
            WHERE DATEPART(M, FIL.FIL_DATE) = DATEPART(M, DATEADD(M, -1, GETDATE()))
	            AND DATEPART(YYYY, FIL.FIL_DATE) = DATEPART(YYYY, DATEADD(M, -1, GETDATE()))
		            AND FIL.QTY_DSP <> 0.00
		            AND DRG.BILL_FLAG = 'T'
		            AND FAC.DCODE LIKE ";
        string BILLING_CIPS = @"SELECT 
	            COUNT(FIL.ID) As CNT
            FROM CIPS.dbo.FIL 
	            INNER JOIN CIPS.dbo.DRG ON
		            FIL.DRG_ID = DRG.ID
	            INNER JOIN CIPS.dbo.FAC ON
		            FIL.FAC_ID = FAC.ID
            WHERE DATEPART(M, FIL.FIL_DATE) = DATEPART(M, DATEADD(M, -1, GETDATE()))
	            AND DATEPART(YYYY, FIL.FIL_DATE) = DATEPART(YYYY, DATEADD(M, -1, GETDATE()))
		            AND FIL.QTY_DSP <> 0.00
		            AND DRG.BILL_FLAG = 'T'
		            AND FAC.DCODE LIKE ";
        #endregion END SQL Billing

        List<Facility> facilities = new List<Facility>();
        public struct Fac
        {
            public string name { get; set; }
            public string code { get; set; }
            public string email { get; set; }
            public string fax { get; set; }
            public string phone { get; set; }
            public string notify_type { get; set; }
        }
        #endregion END Global Vars

        #region Database Functions
        public bool GetRoles(string username)
        {
            using (SqlConnection conn = new SqlConnection(CONN_RX))
            {
                string sql = @"SELECT U.USERNAME, R.DESCRIPTION FROM RPT_USERS U 
                    INNER JOIN RPT_USER_ROLE T ON U.ID = T.USER_ID
                    INNER JOIN RPT_ROLES R ON T.ROLE_ID = R.ID
                    WHERE U.USERNAME = @user";
                SqlCommand command = new SqlCommand(sql, conn);
                command.Parameters.AddWithValue("@user", username);

                try
                {
                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        roles.Add(reader["DESCRIPTION"].ToString());
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }

            }
            return true;
        }

        public bool ExportManualCharges()
        {
            DateTime dt = dpMC_Export.Value;
            int year = dt.Year;
            int month = dt.Month;
            int last_day = DateTime.DaysInMonth(year, month);
            string bill_month = year.ToString() + "-" + month.ToString() + "-" + last_day.ToString();
            string monthName = dpMC_Export.Value.ToString("MMMM", CultureInfo.InvariantCulture);
            Utility.WriteActivity(monthName);
            string sql_suffix = "GROUP BY [ACCT] ORDER BY [ACCT]";
            string sql_suffix_detail = " ORDER BY [DATE]";
            var csv = new StringBuilder();
            Utility.WriteActivity(month.ToString() + "/" + year.ToString());
            string exp_type = "";
            try
            {
                exp_type = ddMC_Export.SelectedItem.ToString();
            }
            catch 
            {
            }

            if (exp_type == "")
            {
                return false;
            }

            string sql = "", category_label = "" ;

            string sql_manual = "SELECT [ACCT],'" + bill_month + "' as [DATE], SUM([PRICE]) [PRICE] FROM MANUAL_CHARGES  WHERE ";
            sql_manual += " DATEPART(month, [DATE]) = " + month.ToString() + " and DATEPART(year, [DATE])=" + year.ToString();

            string sql_detail = "SELECT ACCT, [DATE], CATEGORY, DESCRIPTION, QTY, PRICE  FROM MANUAL_CHARGES  WHERE ";
            sql_detail += " DATEPART(month, [DATE]) = " + month.ToString() + " and DATEPART(year, [DATE])=" + year.ToString();

            string sql_cips = "SELECT CHG.DCODE [ACCT],'" + bill_month + "' as [DATE],";
                   sql_cips += @"SUM(FIL.COPAY_PRICE) - SUM(ISNULL(RTN.CREDIT, 0)) as PRICE
                            FROM
	                            CIPS.dbo.FIL FIL INNER JOIN CIPS.dbo.PAT PAT ON
                                    FIL.PAT_ID = PAT.ID
                                 INNER JOIN CIPS.dbo.FAC FAC ON
                                    FIL.FAC_ID = FAC.ID
                                 INNER JOIN CIPS.dbo.DRG DRG ON
                                    FIL.DRG_ID = DRG.ID
                                 LEFT OUTER JOIN CIPS.dbo.RTN RTN ON
                                    FIL.ID = RTN.FIL_ID
                                 LEFT OUTER JOIN CIPS.dbo.CHG CHG ON
                                    FIL.CHG_ID = CHG.ID
                            WHERE
                                FIL.QTY_DSP <> 0 AND";
            sql_cips += "( DATEPART(month, [FIL_DATE]) = " + month.ToString() + " and DATEPART(year, [FIL_DATE])=" + year.ToString()+ ") ";
            sql_cips += @"AND FIL.STATUS <> 'V'
                            GROUP BY CHG.DCODE
                            ORDER BY CHG.DCODE";

            string sql_cips_wh = "SELECT CHG.DCODE [ACCT],'" + bill_month + "' as [DATE],";
                   sql_cips_wh += @"SUM(FIL.COPAY_PRICE) - SUM(ISNULL(RTN.CREDIT, 0)) as PRICE
                            FROM
	                            CIPS_WHOLESALE.dbo.FIL FIL INNER JOIN CIPS_WHOLESALE.dbo.PAT PAT ON
                                    FIL.PAT_ID = PAT.ID
                                 INNER JOIN CIPS_WHOLESALE.dbo.FAC FAC ON
                                    FIL.FAC_ID = FAC.ID
                                 INNER JOIN CIPS_WHOLESALE.dbo.DRG DRG ON
                                    FIL.DRG_ID = DRG.ID
                                 LEFT OUTER JOIN CIPS_WHOLESALE.dbo.RTN RTN ON
                                    FIL.ID = RTN.FIL_ID
                                 LEFT OUTER JOIN CIPS_WHOLESALE.dbo.CHG CHG ON
                                    FIL.CHG_ID = CHG.ID
                            WHERE
                                FIL.QTY_DSP <> 0 AND";
            sql_cips_wh += "( DATEPART(month, [FIL_DATE]) = " + month.ToString() + " and  DATEPART(year, [FIL_DATE])=" + year.ToString() + ") ";
            sql_cips_wh += @"AND FIL.STATUS <> 'V'
                            GROUP BY CHG.DCODE
                            ORDER BY CHG.DCODE";

            Clipboard.SetText(sql_cips + " \n " + sql_cips_wh);

            using (SqlConnection conn = new SqlConnection(CONN_RX))
            {
                switch (exp_type.Trim())
                {
                    case "Jail Stock":
                        sql = sql_manual + " AND CATEGORY LIKE '%Stock%' " + sql_suffix;
                        category_label = "Stock";
                        break;
                    case "Local Pharmacy":
                        sql = sql_manual + " AND CATEGORY LIKE '%Local Pharmacy%' " + sql_suffix;
                        category_label = "Local Pharmacy";
                        break;
                    case "Jail Stock (details)":
                        sql = sql_detail + " AND CATEGORY LIKE '%Stock%' " + sql_suffix_detail;
                        category_label = "Stock (details)";
                        break;
                    case "Local Pharmacy (details)":
                        sql = sql_detail + " AND CATEGORY LIKE '%Local Pharmacy%' " + sql_suffix_detail;
                        category_label = "Local Pharmacy (details)";
                        break;
                    case "Wholesale":
                        sql = sql_cips_wh;
                        category_label = "Wholesale";
                        break;
                    case "CIPS":
                        sql = sql_cips;
                        category_label = "Reg Meds";
                        break;
                }

                SqlCommand command = new SqlCommand(sql, conn);

                try
                {
                    conn.Open();
                    SqlDataReader dr = command.ExecuteReader();
                    if (category_label == "Stock (details)" || category_label == "Local Pharmacy (details)")
                    {
                        csv.AppendLine("ACCT,DATE,CATEGORY,DESCRIPTION,QTY,PRICE");
                    }
                    else
                    {
                        csv.AppendLine("Acct,Date,Description,Price,Inv #");
                    }
                    while (dr.Read())
                    {
                        var acct = dr["ACCT"].ToString();
                        var date = DateTime.Parse(dr["DATE"].ToString()).ToShortDateString();
                        var price = dr["PRICE"].ToString();

                        string category_detail = "", description = "", qty = "";
                        if (category_label == "Stock (details)" || category_label == "Local Pharmacy (details)")
                        {
                            category_detail = dr["CATEGORY"].ToString();
                            description = dr["DESCRIPTION"].ToString();
                            qty = dr["QTY"].ToString();
                        }
 
                        var category = category_label + " " + month.ToString() + "/" + year.ToString();
                        if (exp_type.Trim() == "CIPS")
                        {
                            category = monthName + " " + year.ToString() + " " + category_label;
                        }

                        var newLine = string.Format("{0},{1},{2},{3},", acct, date, category, price);
                        if (category_label == "Stock (details)" || category_label == "Local Pharmacy (details)")
                        {
                            newLine = string.Format("{0},{1},{2},{3},{4},{5}", acct, date, category_detail, description.Replace(","," "), qty, price);
                        }
                        csv.AppendLine(newLine);
                    }
                    dr.Close();

                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "CSV files (*.csv)|*.csv";
                    //File.WriteAllText(@"C:\Temp\test.csv", csv.ToString());
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllText(sfd.FileName, csv.ToString());
                    }
                    else
                    {
                        return false;
                    }

                    Utility.WriteActivity("Export file for '" + exp_type + "' created at " + sfd.FileName);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return false;
                }

            }
            return true;
        }

        public Fac GetFacility(string fac_code)
        {
            Fac fac = new Fac
            {
                code = fac_code,
                name = "null",
                email = "null",
                fax = "null"
            };

            string sql = @"SELECT DNAME
                , F.DCODE
                , ISNULL(A.EMAIL, '') as EMAIL
                , ISNULL(A.NOTIFY_TYPE, '') as NOTIFY_TYPE
                , ISNULL(F.PHONE1, '') as PHONE
                , ISNULL(F.FAX1, '') as FAX
                FROM CIPS.dbo.FAC F 
                LEFT JOIN RXBackend.dbo.FAC_ALT A
	                ON F.DCODE = A.DCODE
                WHERE F.DCODE = @dcode";

            using (SqlConnection conn = new SqlConnection(CONN_CIPS))
            {
                SqlCommand command = new SqlCommand(sql, conn);
                command.Parameters.AddWithValue("@dcode", fac_code);

                try
                {
                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        fac.name = reader["DNAME"].ToString();
                        //fac.email = reader["EMAIL"].ToString();
                        fac.fax = reader["FAX"].ToString().StartsWith("1") ? reader["FAX"].ToString() : "1-" + reader["FAX"].ToString();
                        fac.phone = reader["PHONE"].ToString().StartsWith("1") ? reader["PHONE"].ToString() : "1-" + reader["PHONE"].ToString();
                        fac.notify_type = reader["NOTIFY_TYPE"].ToString();
                    }
                    reader.Close();
                    fac.email = GetArxAddresses(fac_code);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }

            return fac;
        }

        public void LoadFacilities()
        {
            try
            {
                string selectCommand = @"SELECT 
		                A.DCODE as [Group_Code]
                        ,F.DNAME as [Facility_Name]
                        ,ISNULL(
                        STUFF((SELECT ';' + ADDRESS + ':' + CAST(Billing as VARCHAR)
                        FROM
                        FAC_EMAIL E
                            WHERE A.DCODE = E.FAC_CODE
                            FOR XML PATH('')), 1, 1, ''), '') [Email_Addresses]
                        ,ISNULL(USER1, '') Comments
                        FROM
                        RXBackend.dbo.FAC_ALT A 
                        LEFT JOIN CIPS.dbo.FAC F 
                            ON A.DCODE = F.DCODE";

                daFac = new SqlDataAdapter(selectCommand, CONN_RX);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daFac);

                System.Data.DataTable table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };

                daFac.Fill(table);
                bsFac.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public void LoadAccounts()
        {
            try
            {
                string selectCommand = @"SELECT * FROM ACCOUNT_LIST";

                daAcc = new SqlDataAdapter(selectCommand, CONN_RX);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daFac);

                System.Data.DataTable table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };

                daAcc.Fill(table);
                bsAcc.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public void LoadManualCharges()
        {
            try
            {
                DateTime dt = dpMC_Trans.Value;
                int year = dt.Year;
                int month = dt.Month;
                Utility.WriteActivity(month.ToString() + "/" + year.ToString());

                string selectCommand = "SELECT * FROM MANUAL_CHARGES  where";
                selectCommand += " DATEPART(month, [DATE]) = " + month.ToString() + " and DATEPART(year, [DATE]) = " + year.ToString();

                daMC = new SqlDataAdapter(selectCommand, CONN_RX);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daMC);

                System.Data.DataTable table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };

                daMC.Fill(table);
                bsMC.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public void SaveFaclity()
        {
            if (string.IsNullOrWhiteSpace(txtGroupCode.Text))
            {
                MessageBox.Show("You must select a Group Code", "No Group Code", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var conn = new SqlConnection(CONN_RX);
            conn.Open();
            var sql = "";
            if (btnAddNew.Text == "Add New")
            {
                sql = "UPDATE [FAC_ALT]";
                sql += " SET [USER1] = '" + txtFacUser.Text.Trim() + "'";
                sql += " WHERE [DCODE] = '" + txtGroupCode.Text.Trim() + "'";
            }
            else
            {
                sql = @"INSERT INTO [FAC_ALT]
                        ([DCODE]
                        ,[USER1])";
                sql += " VALUES ('" + txtGroupCode.Text.Trim();
                sql += "','" + txtFacUser.Text.Trim() + "')";
            }
            //txtInfo.Text = sql;
            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
                Utility.WriteActivity("Save successful for " + txtFacilityName.Text);
                MessageBox.Show("Saved...");
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Saved");
            }
            finally
            {
                conn.Close();
                LoadFacilities();
            }
        }

        public void SaveReport(string data, string send)
        {
            var conn = new SqlConnection(CONN_RX);
            conn.Open();
            var sql = "";
            if (btnAddRpt.Text == "Update")
            {
                sql = "UPDATE [FAC_REPORTS]";
                sql += " SET [FAC_DATA] = '" + txtDataRpt.Text.Trim() + "'";
                sql += ",[FAC_SEND] = '" + txtSendRpt.Text.Trim() + "'";
                sql += ",[UPDATED] = GETDATE() ";
                sql += " WHERE [FAC_DATA] = '" + data + "' AND [FAC_SEND] = '" + send + "'";
            }
            else
            {
                sql = @"INSERT INTO [FAC_REPORTS]
                        ([FAC_DATA]
                        ,[FAC_SEND])";
                sql += " VALUES ('" + txtDataRpt.Text.Trim();
                sql += "','" + txtSendRpt.Text.Trim() + "')";
            }

            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
                Utility.WriteActivity("Save successful");
                MessageBox.Show("Saved...");
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Saved");
            }
            finally
            {
                conn.Close();
                LoadReporting();
            }
        }

        public void SaveCode()
        {
            var conn = new SqlConnection(CONN_RX);
            conn.Open();
            var sql = "";
            if (btnAddCode.Text == "Update")
            {
                sql = "UPDATE [BILLING_CODES]";
                sql += " SET [CATEGORY] = '" + ddBilling_Codes.SelectedValue.ToString() + "'";
                sql += ",[DESCRIPTION] = '" + txtBilling_Code.Text.Trim() + "'";
                sql += " WHERE [ID] = " + Code_ID;
            }
            else
            {
                sql = @"INSERT INTO [dbo].[BILLING_CODES]
                       ([CATEGORY]
                       ,[DESCRIPTION])";
                sql += " VALUES ('" + ddBilling_Codes.SelectedValue.ToString();
                sql += "','" + txtBilling_Code.Text.Trim() + "')";
            }

            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
                Utility.WriteActivity("Save successful");
                MessageBox.Show("Saved...");
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Saved");
            }
            finally
            {
                conn.Close();
                LoadCodes();
            }
        }

        public bool Execute_Sql(string sql, string sql_conn)
        {
            bool success = false;
            var conn = new SqlConnection(sql_conn);
            conn.Open();

            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
                success = true;
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                success = false;
            }
            finally
            {
                conn.Close();
            }
            return success;
        }

        public void LoadReporting()
        {
            try
            {
                string selectCommand = @"SELECT 
                    [FAC_DATA] as [Data_Facility_Code]
	                ,f.[DNAME] as [Data_Facility_Name]
                    ,[FAC_SEND] as [Send_Facility_Code]
	                ,fs.[DNAME] as [Send_Facility_Name]
	                ,r.[ID]
                FROM [dbo].[FAC_REPORTS] r
                LEFT JOIN [CIPS].[dbo].[FAC] f
	                on r.FAC_DATA = f.DCODE
                LEFT JOIN [CIPS].[dbo].[FAC] fs
	                on r.FAC_SEND = fs.DCODE 
                ORDER BY [FAC_DATA]";

                daRpt = new SqlDataAdapter(selectCommand, CONN_RX);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daRpt);

                System.Data.DataTable table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };

                daRpt.Fill(table);
                bsRpt.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public void LoadCodes()
        {
            try
            {
                string selectCommand = @"SELECT * FROM BILLING_CODES";

                daCodes = new SqlDataAdapter(selectCommand, CONN_RX);

                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daCodes);

                System.Data.DataTable table = new DataTable
                {
                    Locale = CultureInfo.InvariantCulture
                };

                daCodes.Fill(table);
                bsCodes.DataSource = table;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public void InsertFAC_TRANS(string trans_type, string fac_code, string documents, string email_sent, string notes, string created_by)
        {
            var conn = new SqlConnection(CONN_RX);
            conn.Open();

            var sql = @"INSERT INTO [dbo].[FAC_TRANS]
                        ([TRANS_DATE]
                        ,[TRANS_TYPE]
                        ,[FAC_CODE]
                        ,[DOCUMENTS]
                        ,[EMAIL_SENT]
                        ,[NOTES]
                        ,[CREATED_BY])";
            sql += " VALUES (GETDATE(),'" + trans_type + "','" + fac_code + "','";
            sql += documents + "','" + email_sent + "','" + notes + "','";
            sql += created_by + "')";

            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Added");
            }
            finally
            {
                conn.Close();
            }
        }

        public void LogActivity(string activity, int user, string description, string item)
        {
            var sql = @" INSERT INTO [RPT_ACTIVITY]
                    ([ACTIVITY]
                    ,[USER]
                    ,[DESCRIPTION]
                    ,[ITEM]
                    ,[ALT_USER]
                    ,[ALT_ID])
                    VALUES
                    (@activity, @user , @description, @item, 0, 0)";

            var conn = new SqlConnection(CONN_RX);
            conn.Open();
            //txtInfo.Text = sql;
            var com = new SqlCommand(sql, conn);
            com.Parameters.AddWithValue("@activity", activity);
            com.Parameters.AddWithValue("@user", user);
            com.Parameters.AddWithValue("@description", description);
            com.Parameters.AddWithValue("@item", item);

            try
            {
                com.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Saved");
            }
            finally
            {
                conn.Close();
                LoadFacilities();
            }
        }

        public string GetArxAddresses(string fac_code)
        {
            string addresses = "";
            using (SqlConnection conn = new SqlConnection(CONN_RX))
            {
                string sql = "SELECT ADDRESS FROM FAC_EMAIL WHERE ARX = 1 AND FAC_CODE = @dcode";
                SqlCommand command = new SqlCommand(sql, conn);
                command.Parameters.AddWithValue("@dcode", fac_code);

                try
                {
                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        addresses += reader["ADDRESS"].ToString() + ";";
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
            return addresses;
        }

        public bool GetFacilityBIlling()
        {
            string notes = dpBilling.Value.ToString("yyyy-MM");
            dtBilling = new DataTable
            {
                Locale = CultureInfo.InvariantCulture
            };
            string sql = @"SELECT 
                    CAST(0 as BIT) Send
                    ,A.DCODE as Code
                    ,ISNULL(F.DNAME, '') as Facility
                    ,ISNULL(
			         STUFF((SELECT 
						';' +
						CASE WHEN C.DCODE IS NOT NULL THEN C.DCODE
						WHEN H.DCODE IS NOT NULL THEN H.DCODE
						END
			        FROM
			             CIPS.dbo.FAC FS 
			             LEFT JOIN CIPS.dbo.FAC_CHG G 
			             ON FS.ID = G.FAC_ID
			             LEFT JOIN CIPS.dbo.CHG C
			             ON G.CHG_ID = C.ID
						LEFT JOIN CIPS.dbo.CHG H
                        ON FS.CHG_ID = H.ID
			             WHERE A.DCODE = FS.DCODE
			             FOR XML PATH('')), 1, 1, ''), '') [Accounts]
                    ,ISNULL(
                    STUFF((SELECT ';' + ADDRESS
                    FROM
                    FAC_EMAIL E
                        WHERE A.DCODE = E.FAC_CODE AND Billing = 1
                        FOR XML PATH('')), 1, 1, ''), '') [Email]
                    ,'' as Documents
                    ,CASE WHEN ISNULL(O.NOTES, 'NA') = 'NA' THEN 'N' ELSE 'Y' END as S
                    FROM
                    RXBackend.dbo.FAC_ALT A 
                    LEFT JOIN CIPS.dbo.FAC F 
                        ON A.DCODE = F.DCODE
                    ";
            sql += @"OUTER APPLY ( 
	                SELECT TOP 1 NOTES FROM FAC_TRANS T 
	                WHERE T.FAC_CODE = A.DCODE";
            sql += " AND NOTES = '" + notes + "') O ORDER BY A.DCODE";
            Clipboard.SetText(sql);

            try
            {
                gvStaged.DataSource = bsBill;
                string connection = CONN_RX;
                daBill = new SqlDataAdapter(sql, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daBill);

                daBill.Fill(dtBilling);
                bsBill.DataSource = dtBilling;

                gvStaged.Columns["Send"].Width = 50;
                gvStaged.Columns["Code"].Width = 50;
                gvStaged.Columns["Facility"].Width = 200;
                gvStaged.Columns["Accounts"].Width = 120;
                gvStaged.Columns["Email"].Width = 120;
                gvStaged.Columns["Documents"].Width = 120;
                gvStaged.Columns["S"].Width = 30;

            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
            }

            UpdateBillingDatatable(dtBilling);

            return true;
        }

        public bool GetSentBIlling()
        {
            string notes = dpBilling.Value.ToString("yyyy-MM");

            dtBillingSent = new DataTable
            {
                Locale = CultureInfo.InvariantCulture
            };

            string sql = @"SELECT 
                           [TRANS_DATE]
                          ,[TRANS_TYPE]
                          ,[FAC_CODE]
                          ,F.DNAME NAME
                          ,[DOCUMENTS]
                          ,[EMAIL_SENT]
                          ,[NOTES]
                          ,[CREATED_BY]
                        FROM [dbo].[FAC_TRANS] T 
                        LEFT JOIN CIPS.dbo.FAC F
                        ON T.FAC_CODE = F.DCODE
                    ";
            sql += " WHERE TRANS_TYPE ='BILLING_EMAIL' AND NOTES = '" + notes + "'";
            try
            {
                gvBillingSent.DataSource = bsBillSent;
                string connection = CONN_RX;
                daBillSent = new SqlDataAdapter(sql, connection);
                SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daBillSent);

                daBillSent.Fill(dtBillingSent);
                bsBillSent.DataSource = dtBillingSent;

                //gvBillingSent.Columns["Send"].Width = 50;
                //gvBillingSent.Columns["Code"].Width = 50;
                //gvBillingSent.Columns["Facility"].Width = 200;
                //gvBillingSent.Columns["Accounts"].Width = 140;
                //gvBillingSent.Columns["Email"].Width = 140;
                //gvBillingSent.Columns["Documents"].Width = 140;

            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
            }

            return true;
        }

        public bool CreateBillingReports()
        {
            LoadReporting();
            string fac_data = "", fac_send = "", data_name = "", send_name = "", r_data = "";
            var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
            string output = "Billing Export for " + report_date + "\r\n\r\n";
            string export_path = prop.BillingExports;
            string dsn = txtDSN_CIPS.Text;
            try
            {
                for (int i = 0; i + 1 < gvRpt.Rows.Count; i++)
                {
                    fac_data = gvRpt.Rows[i].Cells["Data_Facility_Code"].Value.ToString();
                    fac_send = gvRpt.Rows[i].Cells["Send_Facility_Code"].Value.ToString();
                    data_name = gvRpt.Rows[i].Cells["Data_Facility_Name"].Value.ToString();
                    send_name = gvRpt.Rows[i].Cells["Send_Facility_Name"].Value.ToString();

                    string sql = BILLING_CIPS + " '" + fac_data + "%'";
                    string report = prop.BillingRptFolder + "BILLING_CIPS.rpt";
                    int r_count = GetRowCount(sql);
                    dsn = txtDSN_CIPS.Text;

                    if (r_count > 0)
                    {
                        ExportBillingReport(report, export_path, "pdf", fac_data, fac_send, dsn);
                        r_data = "CIPS: " + r_count.ToString() + " record(s) found for '" + fac_data + "'";
                        output += r_data + "\r\n";
                        Utility.WriteActivity(r_data);
                    }
                    else
                    {
                        r_data = "CIPS: " + r_count.ToString() + " records found for '" + fac_data + "'";
                        output += r_data + "\r\n";
                        Utility.WriteActivity(r_data);
                    }

                    sql = BILLING_CS + " '" + fac_data + "%'";
                    report = prop.BillingRptFolder + "BILLING_CS.rpt";
                    r_count = GetRowCount(sql);
                    dsn = txtDSN_CIPS.Text;

                    if (r_count > 0)
                    {
                        ExportBillingReport(report, export_path, "pdf", fac_data, "CS_" + fac_send, dsn);
                        r_data = "Cover Sheet: " + r_count.ToString() + " record(s) found for '" + fac_data + "'";
                        output += r_data + "\r\n";
                        Utility.WriteActivity(r_data);
                    }
                    else
                    {
                        r_data = "Cover Sheet: " + r_count.ToString() + " records found for '" + fac_data + "'";
                        output += r_data + "\r\n";
                        Utility.WriteActivity(r_data);
                    }

                    sql = BILLING_CIPS_WS + " '" + fac_data + "%'";
                    report = prop.BillingRptFolder + "BILLING_CIPS_WS.rpt";
                    r_count = GetRowCount(sql);
                    dsn = txtDSN_WS.Text;

                    if (r_count > 0)
                    {
                        ExportBillingReport(report, export_path, "pdf", fac_data, "WS_" + fac_send, dsn);
                        r_data = "CIPS WS: " + r_count.ToString() + " record(s) found for '" + fac_data + "'";
                        output += r_data + "\r\n\r\n";
                        Utility.WriteActivity(r_data);
                    }
                    else
                    {
                        r_data = "CIPS WS: " + r_count.ToString() + " records found for '" + fac_data + "'";
                        output += r_data + "\r\n\r\n";
                        Utility.WriteActivity(r_data);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            string filename = export_path + report_date + "_Billing.txt";
            try
            {
                File.WriteAllText(filename, output);
                Process.Start("notepad.exe", filename);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return true;
        }

        public static int GetRowCount(string sql)
        {
            int cnt = 0;
            using (SqlConnection conn = new SqlConnection(CONN_RX))
            {
                SqlCommand command = new SqlCommand(sql, conn);
                try
                {
                    conn.Open();
                    SqlDataReader reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        cnt = reader.GetInt32(reader.GetOrdinal("CNT"));
                    }
                    reader.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return 0;
                }
                return cnt;
            }
        }

        public bool UpdateAddressList(bool insert)
        {
            bool success = false;
            SqlCommand cmd;
            var conn = new SqlConnection(prop.RxBackend);
            string sql_insert = @"INSERT INTO [dbo].[ACCOUNT_LIST]
               ([ID]
               ,[AccountName]
               ,[GroupCode]
               ,[Contact1]
               ,[Contact2]
               ,[Address1]
               ,[Address2]
               ,[Address3]
               ,[City]
               ,[City2]
               ,[State]
               ,[Zip]
               ,[Zip2]
               ,[Phone]
               ,[Email]
               ,[Email2]
               ,[Email3]
               ,[Type]
               ,[Terms]
               ,[Rep]
               ,[Stmts]
               ,[ShipTo]
               ,[Tax]
               ,[EmailStmts]
               ,[InvoiceNumber])
            VALUES
               (@ID
		       ,@AccountName
		       ,@GroupCode
		       ,@Contact1
		       ,@Contact2
		       ,@Address1
		       ,@Address2
		       ,@Address3
		       ,@City
		       ,@City2
		       ,@State
		       ,@Zip
		       ,@Zip2
		       ,@Phone
		       ,@Email
		       ,@Email2
		       ,@Email3
		       ,@AccType
		       ,@Terms
		       ,@Rep
		       ,@Statements
		       ,@ShipTo
		       ,@Tax
		       ,@EmailStmts
		       ,@InvoiceNumber
		       )";
            string sql_update = @"UPDATE [dbo].[ACCOUNT_LIST]
            SET [ID] = @ID
              ,[AccountName] = @AccountName
              ,[GroupCode] = @GroupCode
              ,[Contact1] = @Contact1
              ,[Contact2] = @Contact2
              ,[Address1] = @Address1
	          ,[Address2] = @Address2
              ,[Address3] = @Address3
              ,[City] = @City
              ,[City2] = @City2
              ,[State] = @State
              ,[Zip] = @Zip
              ,[Zip2] = @Zip2
              ,[Phone] = @Phone
              ,[Email] = @Email
              ,[Email2] = @Email2
              ,[Email3] = @Email3
              ,[Type] = @AccType
              ,[Terms] = @Terms
              ,[Rep] = @Rep
              ,[Stmts] = @Statements
              ,[ShipTo] = @ShipTo
              ,[Tax] = @Tax
              ,[EmailStmts] = @EmailStmts
              ,[InvoiceNumber] = @InvoiceNumber
            WHERE [ID] = @ID_Update";

            if (insert)
                cmd = new SqlCommand(sql_insert, conn);
            else
                cmd = new SqlCommand(sql_update, conn);
            string ID = txtAccID.Text;

            if (!insert && ID != Acc_ID)
            {
                ID = Acc_ID;
            }

            try
            {
                conn.Open();

                cmd.Parameters.Add("@ID", SqlDbType.VarChar).Value = txtAccID.Text;
                cmd.Parameters.Add("@ID_Update", SqlDbType.VarChar).Value = ID;
                cmd.Parameters.Add("@AccountName", SqlDbType.VarChar).Value = txtAccName.Text;
                cmd.Parameters.Add("@GroupCode", SqlDbType.VarChar).Value = txtAccGroupCode.Text;
                cmd.Parameters.Add("@Address1", SqlDbType.VarChar).Value = txtAccAddress1.Text;
                cmd.Parameters.Add("@Address2", SqlDbType.VarChar).Value = txtAccAddress2.Text;
                cmd.Parameters.Add("@Address3", SqlDbType.VarChar).Value = txtAccAddress3.Text;
                cmd.Parameters.Add("@City", SqlDbType.VarChar).Value = txtAccCity.Text;
                cmd.Parameters.Add("@City2", SqlDbType.VarChar).Value = txtAccCity2.Text;
                cmd.Parameters.Add("@State", SqlDbType.VarChar).Value = ddAccStates.SelectedValue.ToString();
                cmd.Parameters.Add("@Zip", SqlDbType.VarChar).Value = txtAccZip.Text;
                cmd.Parameters.Add("@Zip2", SqlDbType.VarChar).Value = txtAccZip2.Text;
                cmd.Parameters.Add("@Phone", SqlDbType.VarChar).Value = txtAccPhone.Text;
                cmd.Parameters.Add("@Email", SqlDbType.VarChar).Value = txtAccEmail.Text;
                cmd.Parameters.Add("@Email2", SqlDbType.VarChar).Value = txtAccEmail2.Text;
                cmd.Parameters.Add("@Email3", SqlDbType.VarChar).Value = txtAccEmail3.Text;
                cmd.Parameters.Add("@AccType", SqlDbType.VarChar).Value = ddAccType.SelectedValue.ToString();
                cmd.Parameters.Add("@Terms", SqlDbType.VarChar).Value = ddAccTerms.SelectedValue.ToString();
                cmd.Parameters.Add("@Rep", SqlDbType.VarChar).Value = txtAccRep.Text;
                cmd.Parameters.Add("@Statements", SqlDbType.VarChar).Value = txtAccStatements.Text;
                cmd.Parameters.Add("@ShipTo", SqlDbType.VarChar).Value = txtAccShipTo.Text;
                cmd.Parameters.Add("@Tax", SqlDbType.VarChar).Value = txtAccTax.Text;
                cmd.Parameters.Add("@EmailStmts", SqlDbType.VarChar).Value = txtAccEmailStmts.Text;
                cmd.Parameters.Add("@InvoiceNumber", SqlDbType.VarChar).Value = txtAccInvoiceNumber.Text;
                cmd.Parameters.Add("@Contact1", SqlDbType.VarChar).Value = txtAccContact1.Text;
                cmd.Parameters.Add("@Contact2", SqlDbType.VarChar).Value = txtAccContact2.Text;

                cmd.ExecuteNonQuery();
                success = true;
                ClearAccList();
                if (insert)
                    Utility.WriteActivity("New Account added");
                else
                    Utility.WriteActivity("Account updated");
                LoadAccounts();
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show(ex.Message);
                success = false;
            }
            finally
            { 

                conn.Close();
            }
            return success;
        }

        public bool UpdateManualCharges(bool insert, decimal qty, decimal price)
        {
            bool success = false;
            SqlCommand cmd;
            var conn = new SqlConnection(prop.RxBackend);
            string sql_insert = @"INSERT INTO [dbo].[MANUAL_CHARGES]
                        ([ACCT]
                        ,[DATE]
                        ,[CATEGORY]
                        ,[DESCRIPTION]
                        ,[QTY]
                        ,[PRICE]
                        ,[TECH])
                    VALUES
                       (@ACCT
		               ,@DATE
		               ,@CATEGORY
                       ,@DESCRIPTION
		               ,@QTY
		               ,@PRICE
		               ,@TECH
		               )";
            string sql_update = @"
                UPDATE [dbo].[MANUAL_CHARGES]
                   SET [ACCT] = @ACCT
                      ,[DATE] = @DATE
                      ,[CATEGORY] = @CATEGORY
                      ,[DESCRIPTION] = @DESCRIPTION
                      ,[QTY] = @QTY
                      ,[PRICE] = @PRICE     
                      ,[UPDATED] = GETDATE()
                      ,[UPDATED_BY] = @TECH
                 WHERE [ID] = @ID";

            if (insert)
                cmd = new SqlCommand(sql_insert, conn);
            else
                cmd = new SqlCommand(sql_update, conn);
            string ID = txtAccID.Text;

            if (!insert && ID != Acc_ID)
            {
                ID = Acc_ID;
            }

            try
            {
                conn.Open();

                if (!insert)
                    cmd.Parameters.Add("@ID", SqlDbType.Int).Value = int.Parse(MC_ID);
                cmd.Parameters.Add("@ACCT", SqlDbType.VarChar).Value = ddMC_Account.SelectedValue.ToString();
                cmd.Parameters.Add("@DATE", SqlDbType.DateTime).Value = dpMC_Date.Value.AddDays(1).AddSeconds(-1); ;
                cmd.Parameters.Add("@CATEGORY", SqlDbType.VarChar).Value = ddMC_Category.SelectedValue.ToString();
                cmd.Parameters.Add("@DESCRIPTION", SqlDbType.VarChar).Value = txtMC_Desc.Text;
                cmd.Parameters.Add("@QTY", SqlDbType.Decimal).Value = qty;
                cmd.Parameters.Add("@PRICE", SqlDbType.Decimal).Value = price;
                cmd.Parameters.Add("@TECH", SqlDbType.VarChar).Value = txtMC_Tech.Text;

                cmd.ExecuteNonQuery();
                if (!insert)
                    LogActivity("MAN_CHARGE", 0, 
                    MC_ID + "- Q: " + qty.ToString() + ", P: " + price.ToString(), txtMC_Tech.Text);
                success = true;

            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                success = false;
            }
            finally
            {
                conn.Close();
            }
            if (insert)
                Utility.WriteActivity("New Manual Charge added");
            else
                Utility.WriteActivity("Manual Charge updated");

            ClearManualCharges();
            LoadManualCharges();

            return success;

        }

        public void ImportManualCharges(string pth)
        {
            Utility.WriteActivity("Reading: " + pth);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@pth);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string acct = "", dte = "", category = "", description = "",qty = "",price = "", tech = "";
            SqlCommand cmd;
            var conn = new SqlConnection(prop.RxBackend);
            string sql_insert = @"INSERT INTO [dbo].[MANUAL_CHARGES]
                        ([ACCT]
                        ,[DATE]
                        ,[CATEGORY]
                        ,[DESCRIPTION]
                        ,[QTY]
                        ,[PRICE]
                        ,[TECH])
                    VALUES
                       (@ACCT
		               ,@DATE
		               ,@CATEGORY
                       ,@DESCRIPTION
		               ,@QTY
		               ,@PRICE
		               ,@TECH
		               )";
            cmd = new SqlCommand(sql_insert, conn);

            for (int i = 1; i <= rowCount; i++)
            {

                try
                {
                    if(i > 1)
                    {
                        acct = xlRange.Cells[i, 1].Value2.ToString().Trim();
                        dte = xlRange.Cells[i, 2].Value.ToString().Trim();
                        category = xlRange.Cells[i, 3].Value.ToString().Trim();
                        description = xlRange.Cells[i, 4].Value2.ToString().Trim();
                        qty = xlRange.Cells[i, 5].Value2.ToString().Trim();
                        price = xlRange.Cells[i, 6].Value2.ToString().Trim();
                        tech = xlRange.Cells[i, 7].Value2.ToString().Trim();                       

                        try
                        {
                            conn.Open();
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add("@ACCT", SqlDbType.VarChar).Value = acct;
                            cmd.Parameters.Add("@DATE", SqlDbType.DateTime).Value = DateTime.Parse(dte); 
                            cmd.Parameters.Add("@CATEGORY", SqlDbType.VarChar).Value = category;
                            cmd.Parameters.Add("@DESCRIPTION", SqlDbType.VarChar).Value = description;
                            cmd.Parameters.Add("@QTY", SqlDbType.Decimal).Value = Decimal.Parse(qty);
                            cmd.Parameters.Add("@PRICE", SqlDbType.Decimal).Value = Decimal.Parse(price);
                            cmd.Parameters.Add("@TECH", SqlDbType.VarChar).Value = tech;

                            cmd.ExecuteNonQuery();

                            Utility.WriteActivity("Record added: " + acct + ": " + description);
                        }
                        catch (Exception ex)
                        {
                            Utility.WriteActivity(ex.Message);
                        }
                        finally
                        {
                            conn.Close();
                        }

                    }

                }
                catch (Exception ex)
                {
                Utility.WriteActivity(ex.Message);
                }

            }

        }

        public void FillDropDownList(string Query, ComboBox DropDownName, string CONNECTION_STRING)
        {
            // If you use a DataTable (or any object which implmenets IEnumerable)
            // you can bind the results of your query directly as the 
            // datasource for the ComboBox. 
            DataTable dt = new DataTable();

            // Where possible, use the using block for data access. The 
            // using block handles disposal of resources and connection 
            // cleanup for you:
            using (var cn = new SqlConnection(CONNECTION_STRING))
            {
                using (var cmd = new SqlCommand(Query, cn))
                {
                    cn.Open();

                    try
                    {
                        dt.Load(cmd.ExecuteReader());
                    }
                    catch (SqlException e)
                    {
                        // Do some logging or something. 
                        MessageBox.Show("There was an error accessing your data. DETAIL: " + e.ToString());
                    }
                }
            }

            DropDownName.DataSource = dt;
            DropDownName.ValueMember = dt.Columns[0].ColumnName;
            DropDownName.DisplayMember = dt.Columns[1].ColumnName;
        }

        #endregion  --END Database Functions

        #region Reporting Funcions

        public void ExportReport(string report, string export_path, string export_type, string[] parms, string dsn)
        {
            var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");

            string code = "", typ = "", valid = "";
            for (int i = 0; i + 1 < gvNotiifications.Rows.Count; i++)
            {
                code = gvNotiifications.Rows[i].Cells[0].Value.ToString();
                typ = gvNotiifications.Rows[i].Cells[5].Value.ToString();
                valid = gvNotiifications.Rows[i].Cells[6].Value.ToString();

                string[] rpt = { "-S", dsn,
                "-F", report,
                "-O", export_path + report_date + "_" + code + ".pdf",
                "-E", export_type};
                string[] p = { "-A", "Facility:" + code, "-A" };

                var rpt_data = rpt.Concat(p).ToArray();


                if (valid == "True")
                {
                    Utility.WriteActivity("Running report: " + report);
                    RunReport(rpt_data);
                }

            }

            Utility.WriteActivity("Report export transactions complete");
        }

        public void ExportBillingReport(string report, string export_path, string export_type,
            string data, string send, string dsn)
        {
            var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");

            string[] rpt = { "-S", dsn,
            "-F", report,
            "-O", export_path + report_date + "_" + data + "_" + send,
            "-E", export_type,
            "-A", "Facility:" + data};
            //string[] p = { "-A", "Facility:" + data, "-A" };

            //var rpt_data = rpt.Concat(p).ToArray();

            Utility.WriteActivity("Running report: " + report);
            RunReport(rpt);

            Utility.WriteActivity("Report export transactions complete for " + data);
        }

        public void FaxReports(string report, string export_path, string export_type, string[] parms, string dsn)
        {

            string code = "", typ = "", valid = "", name = "";
            var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
            for (int i = 0; i + 1 < gvNotiifications.Rows.Count; i++)
            {
                code = gvNotiifications.Rows[i].Cells[0].Value.ToString();
                name = gvNotiifications.Rows[i].Cells[1].Value.ToString();
                typ = gvNotiifications.Rows[i].Cells[5].Value.ToString();
                valid = gvNotiifications.Rows[i].Cells[6].Value.ToString();

                if (typ == "Fax" || typ == "Both")
                {
                    string[] rpt = { "-S", dsn,
                        "-F", report,
                        "-E", export_type};
                    string[] p = { "-N", prop.FaxPrinter, "-A", "Facility:" + code };

                    var rpt_data = rpt.Concat(p).ToArray();

                    if (valid == "True")
                    {
                        Utility.WriteActivity(code + "-" + name + " Notify Type: " + typ);

                        bool sent = RunReport(rpt_data);
                        if (sent)
                        {
                            LogActivity("FAC_FAX", 0, name, report_date);
                            gvNotiifications.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                        }
                    }
                }
                else
                {
                    continue;
                }

            }

            //RunReport(rpt_data);
        }

        public void FaxExportToFolder()
        {
            string code = "", typ = "", valid = "", name = "", fax = "";
            var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
            var export_folder = prop.NotifyExports;
            string doc_path = "";
            var fax_folder = prop.FaxFolder;
            string export_path = "";
            for (int i = 0; i + 1 < gvNotiifications.Rows.Count; i++)
            {
                code = gvNotiifications.Rows[i].Cells[0].Value.ToString();
                //name = gvNotiifications.Rows[i].Cells[1].Value.ToString();
                typ = gvNotiifications.Rows[i].Cells[5].Value.ToString();
                valid = gvNotiifications.Rows[i].Cells[6].Value.ToString();

                if (typ == "Fax" || typ == "Both")
                {
                    var fac = GetFacility(code);
                    name = fac.name;
                    fax = fac.fax.StartsWith("1") ? fac.fax : "1-" + fac.fax;

                    doc_path = export_folder + report_date + "_" + code + ".pdf";
                    export_path = fax_folder + report_date + "_" + code + "@F201 " + name + "@@F211 " + fax + "@.pdf";
                    bool found = File.Exists(doc_path);

                    if (!found)
                    {
                        MessageBox.Show("The file for facility code '" + code + "' was not found", "File not found",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        continue;
                    }

                    if (valid == "True")
                    {
                        Utility.WriteActivity("File: " + doc_path + " sent to " + export_path);

                        if (File.Exists(export_path))
                        {
                            File.Delete(export_path);
                        }

                        System.IO.File.Move(doc_path, export_path);

                        LogActivity("FAC_FAX", 0, name, report_date);
                        gvNotiifications.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
                else
                {
                    continue;
                }
            }

            Utility.WriteActivity("Fax transactions complete");
        }

        public bool RunReport(string[] args)
        {
            try
            {
                // read program arguments into Argument Container
                ArgumentContainer argContainer = new ArgumentContainer();
                argContainer.ReadArguments(args);

                if (argContainer.GetHelp)
                    Helper.ShowHelpMessage();
                else
                {
                    string _logFilename = string.Empty;

                    if (argContainer.EnableLog)
                        _logFilename = "ninja-" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".log";

                    ReportProcessor reportNinja = new ReportProcessor(_logFilename)
                    {
                        ReportArguments = argContainer,
                    };

                    reportNinja.Run();
                }
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(string.Format("Exception: {0}", ex.Message));
                Utility.WriteActivity(string.Format("Inner Exception: {0}", ex.InnerException));
                return false;
            }
            return true;
        }
        #endregion

        #region Utility Functions
        public void GetCalNotifications()
        {
            UserCredential credential;

            try
            {
                using (var stream =
                    new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    // The file token.json stores the user's access and refresh tokens, and is created
                    // automatically when the authorization flow completes for the first time.
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        System.Threading.CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                    Utility.WriteActivity("Credential file saved to: " + credPath);
                }

                // Create Google Calendar API service.
                var service = new CalendarService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });

                // Define parameters of request.
                EventsResource.ListRequest request = service.Events.List(prop.CalendarID);
                request.TimeMin = dptFacExport.Value.Date;
                request.TimeMax = dptFacExport.Value.Date.AddDays(1);
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.MaxResults = 1000;
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                // List events.
                Events events = request.Execute();
                Console.WriteLine("Upcoming notifications:");
                Utility.WriteActivity("Upcoming notifications:");
                facilities.Clear();
                if (events.Items != null && events.Items.Count > 0)
                {
                    foreach (var eventItem in events.Items)
                    {
                        var fac = new Fac();
                        string when = eventItem.Start.DateTime.ToString();
                        string ev = eventItem.Summary;
                        if (String.IsNullOrEmpty(when))
                        {
                            when = eventItem.Start.Date;
                            Utility.WriteActivity(ev);

                            var h = ev.Split('-')[0]; //Hyphen split
                            var s = ev.Split(' ')[0];//Space split
                            if (h.Length < 5)
                            {
                                //Utility.WriteActivity(h.Trim());
                                fac = GetFacility(h.Trim());
                                Utility.WriteActivity(fac.name + ": imported");
                            }
                            else if (s.Length < 5)
                            {
                                //Utility.WriteActivity(s.Trim());
                                fac = GetFacility(s.Trim());
                                Utility.WriteActivity(fac.name + ": imported");
                            }
                            else
                            {
                                //Utility.WriteActivity(strVal);
                                fac.code = ev;
                                fac.name = "none";
                                fac.email = "none";
                                fac.fax = "none";
                                fac.phone = "none";
                                fac.notify_type = "none";
                                Utility.WriteActivity(fac.name + ": imported");
                            }
                            var valid = fac.name != "none" ? true : false;
                            var notify_type = ev.Contains("(e") && valid ? "Email" : "Fax";
                            if (ev.Contains("{"))
                            {
                                var str = GetStringBetweenCharacters(ev, '{', '}').Replace(" ", "").ToLower();

                                if (str.Contains("ef") || str.Contains("fe"))
                                {
                                    notify_type = "Both";
                                }
                                else if (str.Contains("e"))
                                {
                                    notify_type = "Email";
                                }
                                else if (str.Contains("f"))
                                {
                                    notify_type = "Fax";
                                }
                            }

                            notify_type = fac.notify_type != "" ? fac.notify_type : notify_type;
                            Facility facility = new Facility(fac.code, fac.name, fac.phone, fac.fax, fac.email, notify_type, valid);
                            facilities.Add(facility);

                        }
                        Console.WriteLine("{0} ({1})", eventItem.Summary, when);
                        Utility.WriteActivity(eventItem.Summary + ":" + when);
                    }
                }
                else
                {
                    Console.WriteLine("No upcoming notifications found.");
                    Utility.WriteActivity("No upcoming notifications found.");
                }

                if (facilities.Count > 0)
                {
                    DataTable dt = new DataTable();
                    dt.Columns.Add("Code", typeof(string));
                    dt.Columns.Add("Facility Name", typeof(string));
                    dt.Columns.Add("Email Addresses", typeof(string));
                    dt.Columns.Add("Fax", typeof(string));
                    dt.Columns.Add("Phone", typeof(string));
                    dt.Columns.Add("Notify Type", typeof(string));
                    dt.Columns.Add("Valid", typeof(string));
                    foreach (var fac in facilities)
                    {
                        DataRow dr = dt.NewRow();
                        dr["Code"] = fac.code;
                        dr["Facility Name"] = fac.name;
                        dr["Email Addresses"] = fac.email;
                        dr["Fax"] = fac.fax;
                        dr["Phone"] = fac.phone;
                        dr["Notify Type"] = fac.notify_type;
                        dr["Valid"] = fac.valid_code.ToString();
                        dt.Rows.Add(dr);
                        if (!fac.valid_code)
                        {
                            txtInfo.Text += fac.code + "\r\n";
                        }
                    }

                    gvNotiifications.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }

        public static void WriteConfig(string key, string val)
        {
            try
            {
                Properties.Settings.Default[key] = val;
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static string ReadConfig(string key)
        {
            try
            {
                return Properties.Settings.Default[key].ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return "Not Found";
            }
        }

        public void GetSettings()
        {
            try
            {
                txtMC_Tech.Text = userName;
                txtNotifyReport.Text = ReadConfig("NotifyReport");
                txtNotifyExports.Text = ReadConfig("NotifyExports");
                txtCIPS.Text = ReadConfig("CIPS");
                txtRxBackend.Text = ReadConfig("RxBackend");
                CONN_CIPS = ReadConfig("CIPS");
                CONN_RX = ReadConfig("RxBackend");
                txtAddress.Text = ReadConfig("EmailAddress");
                txtPassword.Text = Decrypt(ReadConfig("EmailPassword"));
                txtMailbox.Text = ReadConfig("Mailbox");
                txtEmailServer.Text = ReadConfig("EmailServer");
                txtDSN_CIPS.Text = ReadConfig("DSN_CIPS");
                txtDSN_WS.Text = ReadConfig("DSN_WS");
                txtWS_ID.Text = ReadConfig("WS_ID");
                txtDSN_RxBackend.Text = ReadConfig("DSN_RxBackend");
                txtEmailPort.Text = ReadConfig("EmailPort");
                txtFaxPrinter.Text = ReadConfig("FaxPrinter");
                txtCalendarID.Text = ReadConfig("CalendarID");
                txtFaxFolder.Text = ReadConfig("FaxFolder");
                txtForward.Text = ReadConfig("ForwardAddress");
                txtBillingRptFolder.Text = ReadConfig("BillingRptFolder");
                txtBillingExports.Text = ReadConfig("BillingExports");
                txtCE_Report.Text = ReadConfig("CE_Report");
                txtCE_Export.Text = ReadConfig("CE_Export");
                //Attachment Processing
                txtAddressAttachment.Text = ReadConfig("AddressAttachment");
                txtPasswordAttachment.Text = Decrypt(ReadConfig("PasswordAttachment"));
                txtDownloadFolder.Text = ReadConfig("download");
                txtProcessFolder.Text = ReadConfig("process");
                txtRenamedFolder.Text = ReadConfig("renamed");
                txtPythonFolder.Text = ReadConfig("python");
                txtFrom.Text = prop.fromCrop.X.ToString() + "," + prop.fromCrop.Y.ToString();
                txtTo.Text = prop.toCrop.X.ToString() + "," + prop.toCrop.Y.ToString();
                txtDpi.Text = prop.dpi.ToString();

                txtEmailMessage.Text = prop.SendEmailMessage;
                txtEmailSubject.Text = prop.SendEmailSubject;
                txtUpdateFolder.Text = prop.update;

                outputFolder = txtProcessFolder.Text;
                outputFolderNew = txtRenamedFolder.Text;

                //Drop Downs
                var stateList = new StateArray();
                var states = stateList.ListOfStates();

                var bindingSource1 = new BindingSource();
                bindingSource1.DataSource = states;

                ddAccStates.DataSource = bindingSource1.DataSource;

                ddAccStates.DisplayMember = "Name";
                ddAccStates.ValueMember = "Abbreviations";
                DataSet = true;

                string sql = "SELECT '' as DESC_, '' as DESCRIPTION UNION  ";
                sql += @"SELECT DESCRIPTION as DESC_, DESCRIPTION FROM BILLING_CODES 
                            WHERE CATEGORY = 'ACC_TYPE' ORDER BY DESCRIPTION ";
                FillDropDownList(sql, ddAccType, prop.RxBackend);

                sql = "SELECT '' as DESCRIPTION, '' as DESCRIPTION UNION ALL ";
                sql += "SELECT DESCRIPTION, DESCRIPTION FROM BILLING_CODES WHERE CATEGORY = 'TERMS' ";
                FillDropDownList(sql, ddAccTerms, prop.RxBackend);

                sql = "SELECT '' as DESCRIPTION, '' as DESCRIPTION UNION ALL ";
                sql += "SELECT DESCRIPTION, DESCRIPTION FROM BILLING_CODES WHERE CATEGORY = 'MC_CATEGORY' ";
                FillDropDownList(sql, ddMC_Category, prop.RxBackend);

                this.ddMC_Account.SelectedIndexChanged -= new System.EventHandler(this.ddMC_Account_SelectedIndexChanged);
                sql = "SELECT '' as ID_VAL, '' as ID UNION ";
                sql += "SELECT ID as ID_VAL, ID FROM ACCOUNT_LIST ORDER BY ID ";
                FillDropDownList(sql, ddMC_Account, prop.RxBackend);
                this.ddMC_Account.SelectedIndexChanged += new System.EventHandler(this.ddMC_Account_SelectedIndexChanged);

                this.ddBilling_Codes.SelectedIndexChanged -= new System.EventHandler(this.ddBilling_Codes_SelectedIndexChanged);
                sql = "SELECT '' as CATEGORY, '' as CATEGORY UNION ALL ";
                sql += "SELECT DISTINCT CATEGORY AS CAT_VAL, CATEGORY FROM BILLING_CODES  ";
                FillDropDownList(sql, ddBilling_Codes, prop.RxBackend);
                this.ddBilling_Codes.SelectedIndexChanged += new System.EventHandler(this.ddBilling_Codes_SelectedIndexChanged);

                var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                txtBuildInfo.Text = String.Format("Application Build: {0}", version);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void UpdateSettings(string configName, string newText, string activityText, bool includeValues)
        {
            try
            {
                string prevText = ReadConfig(configName);
                string values = includeValues ? " [" + prevText + "] to [" + newText + "]" : "";
                WriteConfig(configName, newText);
                Utility.WriteActivity(activityText + values);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void readExcelFile(string pth)
        {
            Utility.WriteActivity("Reading: " + pth);
            facilities.Clear();

            //Create COM Objects. Create a COM object for everything that is referenced
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@pth);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = prop.FAC_COLUMN; //xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            var strVal = "";

            for (int i = 1; i <= rowCount; i++)
            {
                //for (int j = 1; j <= colCount; j++)
                if (colCount == prop.FAC_COLUMN && i > 1)
                {
                    //new line
                    //if (j == 1)
                    //{
                    //    //Console.Write("\r\n");
                    //    //Utility.WriteActivity("");
                    //}

                    try
                    {
                        if (xlRange.Cells[i, colCount] != null && xlRange.Cells[i, colCount].Value2 != null)
                        {
                            var fac = new Fac();
                            //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
                            strVal = xlRange.Cells[i, colCount].Value2.ToString();
                            if (strVal.Trim() != "")
                            {
                                var h = strVal.Split('-')[0]; //Hyphen split
                                var s = strVal.Split(' ')[0];//Space split
                                if (h.Length < 4)
                                {
                                    fac = GetFacility(h.Trim());
                                    Utility.WriteActivity(fac.name + ": imported");
                                }
                                else if (s.Length < 5)
                                {
                                    fac = GetFacility(s.Trim());
                                    Utility.WriteActivity(fac.name + ": imported");
                                }
                                else
                                {
                                    fac.code = strVal;
                                    fac.name = "none";
                                    fac.email = "none";
                                    fac.fax = "none";
                                    fac.phone = "none";
                                    fac.notify_type = "none";
                                    Utility.WriteActivity(fac.name + ": imported");
                                }
                                var valid = fac.name != "none" ? true : false;
                                var notify_type = strVal.Contains("(e") && valid ? "Email" : "Fax";
                                if (strVal.Contains("{"))
                                {
                                    var str = GetStringBetweenCharacters(strVal, '{', '}').Replace(" ", "").ToLower();

                                    if (str.Contains("ef") || str.Contains("fe"))
                                    {
                                        notify_type = "Both";
                                    }
                                    else if (str.Contains("e"))
                                    {
                                        notify_type = "Email";
                                    }
                                    else if (str.Contains("f"))
                                    {
                                        notify_type = "Fax";
                                    }
                                }
                                notify_type = fac.notify_type != "" ? fac.notify_type : notify_type;
                                Facility facility = new Facility(fac.code, fac.name, fac.phone, fac.fax, fac.email, notify_type, valid);
                                facilities.Add(facility);
                            }
                        }


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }

            DataTable dt = new DataTable();
            dt.Columns.Add("Code", typeof(string));
            dt.Columns.Add("Facility Name", typeof(string));
            dt.Columns.Add("Email Addresses", typeof(string));
            dt.Columns.Add("Fax", typeof(string));
            dt.Columns.Add("Phone", typeof(string));
            dt.Columns.Add("Notify Type", typeof(string));
            dt.Columns.Add("Valid", typeof(string));
            foreach (var fac in facilities)
            {
                DataRow dr = dt.NewRow();
                dr["Code"] = fac.code;
                dr["Facility Name"] = fac.name;
                dr["Email Addresses"] = fac.email;
                dr["Fax"] = fac.fax;
                dr["Phone"] = fac.phone;
                dr["Notify Type"] = fac.notify_type;
                dr["Valid"] = fac.valid_code.ToString();
                dt.Rows.Add(dr);
                if (!fac.valid_code)
                {
                    txtInfo.Text += fac.code + "\r\n";
                }
            }

            gvNotiifications.DataSource = dt;
            Process[] excelProcesses = Process.GetProcessesByName("excel");
            foreach (Process p in excelProcesses)
            {
                if (string.IsNullOrEmpty(p.MainWindowTitle)) // use MainWindowTitle to distinguish this excel process with other excel processes 
                {
                    p.Kill();
                }
            }
        }

        public void readExcelImportCharges(string pth)
        {
            Utility.WriteActivity("Reading: " + pth);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@pth);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = prop.FAC_COLUMN; //xlRange.Columns.Count;

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            string acct = "", date = "", first_name = "", last_name = "", cost = "", qty="", med = "", description = "" ;
            decimal qty_d, price_d;

            SqlCommand cmd;
            var conn = new SqlConnection(prop.RxBackend);
            string sql_insert = @"INSERT INTO [dbo].[MANUAL_CHARGES]
                        ([ACCT]
                        ,[DATE]
                        ,[CATEGORY]
                        ,[DESCRIPTION]
                        ,[QTY]
                        ,[PRICE]
                        ,[TECH])
                    VALUES
                       (@ACCT
		               ,@DATE
		               ,@CATEGORY
                       ,@DESCRIPTION
		               ,@QTY
		               ,@PRICE
		               ,@TECH
		               )";
            cmd = new SqlCommand(sql_insert, conn);

            DataTable dt = new DataTable();
            dt.Columns.Add("Acct", typeof(string));
            dt.Columns.Add("Date", typeof(DateTime));
            dt.Columns.Add("Category", typeof(string));
            dt.Columns.Add("Description", typeof(string));
            dt.Columns.Add("Qty", typeof(decimal));
            dt.Columns.Add("Price", typeof(decimal));
            dt.Columns.Add("Tech", typeof(string));

            for (int i = 1; i <= rowCount; i++)
            {
                if (i > 1)
                {
                    try
                    {
                        if ( xlRange.Cells[i, 1].Value2 != null)
                        {
                            first_name = ConvertTo_ProperCase(xlRange.Cells[i, 1].Value2.ToString());
                            last_name = ConvertTo_ProperCase(xlRange.Cells[i, 2].Value2.ToString());
                            acct = xlRange.Cells[i, 5].Value2.ToString();
                            date = xlRange.Cells[i, 7].Value.ToString();
                            med = ConvertTo_ProperCase(xlRange.Cells[i, 11].Value2.ToString());
                            qty = xlRange.Cells[i, 12].Value2.ToString();
                            cost = xlRange.Cells[i, 13].Value2.ToString();
                            //Utility.WriteActivity(acct + " :: " + date + " :: " + first_name + " :: " + last_name + " :: " + med + " :: " + qty + " :: " + cost);

                            if (Decimal.TryParse(qty.Trim(), out qty_d) ) {

                            }
                            else
                            {
                                Utility.WriteActivity("Could not convert qty value on line " + i);
                                continue;
                            }

                            if (Decimal.TryParse(cost.Trim(), out price_d))
                            {

                            }
                            else
                            {
                                Utility.WriteActivity("Could not convert cost value on line " + i);
                                continue;
                            }
                            description = first_name + " " + last_name + " ; " + med;

                            DataRow dr = dt.NewRow();
                            dr["Acct"] = acct;
                            dr["Date"] = DateTime.Parse(date);
                            dr["Category"] = "Local Pharmacy";
                            dr["Description"] = first_name + " " + last_name + " ; " + med;
                            dr["Qty"] = qty_d;
                            dr["Price"] = price_d;
                            dr["Tech"] = Environment.UserName;
                            dt.Rows.Add(dr);

                            try
                            {
                                conn.Open();
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add("@ACCT", SqlDbType.VarChar).Value = acct;
                                cmd.Parameters.Add("@DATE", SqlDbType.DateTime).Value = DateTime.Parse(date);
                                cmd.Parameters.Add("@CATEGORY", SqlDbType.VarChar).Value = "Local Pharmacy";
                                cmd.Parameters.Add("@DESCRIPTION", SqlDbType.VarChar).Value = description;
                                cmd.Parameters.Add("@QTY", SqlDbType.Decimal).Value = Decimal.Parse(qty);
                                cmd.Parameters.Add("@PRICE", SqlDbType.Decimal).Value = Decimal.Parse(cost);
                                cmd.Parameters.Add("@TECH", SqlDbType.VarChar).Value = Environment.UserName;

                                cmd.ExecuteNonQuery();

                                Utility.WriteActivity("Record added: " + acct + " :: " + description);
                            }
                            catch (Exception ex)
                            {
                                Utility.WriteActivity(ex.Message);
                            }
                            finally
                            {
                                conn.Close();
                            }

                        }


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }
            StringBuilder sb = new StringBuilder();
            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().                                             Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "CSV files (*.csv)|*.csv";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText(sfd.FileName, sb.ToString());
            }
            //else
            //{
            //    return false;
            //}



            //gvNotiifications.DataSource = dt;
            //Process[] excelProcesses = Process.GetProcessesByName("excel");
            //foreach (Process p in excelProcesses)
            //{
            //    if (string.IsNullOrEmpty(p.MainWindowTitle)) // use MainWindowTitle to distinguish this excel process with other excel processes 
            //    {
            //        p.Kill();
            //    }
            //}
        }

        public static string ConvertTo_ProperCase(string text)
        {
            TextInfo myTI = new CultureInfo("en-US", false).TextInfo;
            return myTI.ToTitleCase(text.ToLower());
        }

        public static string Encrypt(string textToEncrypt)
        {
            try
            {
                string ToReturn = "";
                string _key = "ay$a5%&jwrtmnh;lasjdf98787";
                string _iv = "abc@98797hjkas$&asd(*$%";
                byte[] _ivByte = { };
                _ivByte = System.Text.Encoding.UTF8.GetBytes(_iv.Substring(0, 8));
                byte[] _keybyte = { };
                _keybyte = System.Text.Encoding.UTF8.GetBytes(_key.Substring(0, 8));
                MemoryStream ms = null; CryptoStream cs = null;
                byte[] inputbyteArray = System.Text.Encoding.UTF8.GetBytes(textToEncrypt);
                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    ms = new MemoryStream();
                    cs = new CryptoStream(ms, des.CreateEncryptor(_keybyte, _ivByte), CryptoStreamMode.Write);
                    cs.Write(inputbyteArray, 0, inputbyteArray.Length);
                    cs.FlushFinalBlock();
                    ToReturn = Convert.ToBase64String(ms.ToArray());
                }
                return ToReturn;
            }
            catch (Exception ae)
            {
                throw new Exception(ae.Message, ae.InnerException);
            }
        }

        public static string Decrypt(string textToDecrypt)
        {
            try
            {
                string ToReturn = "";
                string _key = "ay$a5%&jwrtmnh;lasjdf98787";
                string _iv = "abc@98797hjkas$&asd(*$%";
                byte[] _ivByte = { };
                _ivByte = System.Text.Encoding.UTF8.GetBytes(_iv.Substring(0, 8));
                byte[] _keybyte = { };
                _keybyte = System.Text.Encoding.UTF8.GetBytes(_key.Substring(0, 8));
                MemoryStream ms = null; CryptoStream cs = null;
                byte[] inputbyteArray = new byte[textToDecrypt.Replace(" ", "+").Length];
                inputbyteArray = Convert.FromBase64String(textToDecrypt.Replace(" ", "+"));
                using (DESCryptoServiceProvider des = new DESCryptoServiceProvider())
                {
                    ms = new MemoryStream();
                    cs = new CryptoStream(ms, des.CreateDecryptor(_keybyte, _ivByte), CryptoStreamMode.Write);
                    cs.Write(inputbyteArray, 0, inputbyteArray.Length);
                    cs.FlushFinalBlock();
                    Encoding encoding = Encoding.UTF8;
                    ToReturn = encoding.GetString(ms.ToArray());
                }
                return ToReturn;
            }
            catch (Exception ae)
            {
                throw new Exception(ae.Message, ae.InnerException);
            }

        }

        public void ClearFacTextBoxes()
        {
            txtGroupCode.Text = "";
            txtFacUser.Text = "";
            txtFacUser.Text = "";
            txtFacilityName.Text = "";
            txtFacEmail.Text = "";
        }

        public void ClearCodes()
        {
            txtBilling_Code.Text = "";
            ddBilling_Codes.SelectedValue = "";
            btnAddCode.Text = "Add";
        }

        public bool SendEmail(string msg, string subject, string recip, string from, string from_name, string[] attachments)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(prop.EmailServer);

                mail.From = new MailAddress(from, from_name);
                var recipients = recip.Split(';');
                foreach (var r in recipients)
                {
                    mail.To.Add(r);
                }
                mail.Subject = subject;
                mail.Body = msg;
                if (attachments != null && attachments.Length > 0)
                {
                    for (int i = 0; i < attachments.Length; i++)
                    {
                        if (attachments[i].Trim().Length > 2)
                        {
                            System.Net.Mail.Attachment attachment;
                            attachment = new System.Net.Mail.Attachment(attachments[i]);
                            mail.Attachments.Add(attachment);
                        }
                    }

                }

                SmtpServer.Port = prop.EmailPort;
                SmtpServer.Credentials =
                new System.Net.NetworkCredential(prop.EmailAddress, Decrypt(prop.EmailPassword));
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);
                SmtpServer.Dispose();
                mail.Dispose();
                Utility.WriteActivity("Mail Sent to " + recip);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

            return true;

        }

        public bool ValidateEmail(string email)
        {
            if (string.IsNullOrWhiteSpace(email.Trim()))
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(email,
                    @"^(|([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5}){1,25})+([;.](([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5}){1,25})+)*$");
            }
            catch (RegexMatchTimeoutException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public static void LogError(string msg)
        {
            SimpleLogger sl = new SimpleLogger();
            sl.Error(msg);
            //Utility.WriteActivity(msg);
        }

        public string DownloadAttachments()
        {
            var title = this.Text;
            this.Text = title + " - Downloading...";
            string hostname = txtEmailServer.Text,
            username = txtAddressAttachment.Text,
            password = txtPasswordAttachment.Text,
            mailbox = txtMailbox.Text,
            attachmentDirectory = txtDownloadFolder.Text,
            msg_ret = "";
            int att_count = 0;

            try
            {
                using (ImapClient client = new ImapClient(hostname, 993, username, password, AuthMethod.Login, true))
                {
                    Utility.WriteActivity("Connected to server '" + hostname + "'");
                    client.DefaultMailbox = mailbox;

                    // Returns a collection of identifiers of all mails matching the specified search criteria.
                    IEnumerable<uint> uids;
                    if (chkDate.Checked)
                    {
                        uids = client.Search(SearchCondition.SentSince(dptDownload.Value.Date));
                    }
                    else
                    {
                        uids = client.Search(SearchCondition.All());
                    }

                    // Download mail messages from the default mailbox.
                    IEnumerable<MailMessage> messages = client.GetMessages(uids);

                    for (int i = 0; i < messages.Count(); i++)
                    {
                        string str1 = messages.ElementAt(i).Subject.ToString();
                        int cnt_attach = messages.ElementAt(i).Attachments.Count;
                        if (cnt_attach > 0)
                        {
                            var att = messages.ElementAt(i).Attachments[0];

                            //string filename = string.Format(@"{0}{1}_{2}{3}", attachmentDirectory,
                            //    Path.GetFileNameWithoutExtension(att.Name), DateTime.Now.AddSeconds(i).ToString("MMddyyyyhhmmss"), Path.GetExtension(att.Name));
                            string filename = string.Format(@"{0}{1}{2}{3}", attachmentDirectory,
                            "DOC", DateTime.Now.AddSeconds(i).ToString("MMddyyyyhhmmss"), Path.GetExtension(att.Name));
                            var file_att = att.ContentStream;
                            using (System.IO.FileStream output = new System.IO.FileStream(filename, FileMode.Create))
                            {
                                file_att.CopyTo(output);
                                Utility.WriteActivity(Path.GetFileName(att.Name) + " downloaded to " + attachmentDirectory);
                            }

                        }
                        att_count = i;
                    }
                    msg_ret = att_count.ToString() + " attachment(s) downloaded from mailbox " + mailbox + " to folder " + attachmentDirectory;
                }
            }
#pragma warning disable CA1031 // Do not catch general exception types


            catch (Exception e)
#pragma warning restore CA1031 // Do not catch general exception types
            {
                LogError(e.Message);
                MessageBox.Show(e.Message);
            }
            finally
            {
                this.Text = title;
                //write_activity(msg_ret);
            }
            return (msg_ret);
        }

        public void DirectorySearch(string dir)
        {
            try
            {
                foreach (string f in Directory.GetFiles(dir))
                {
                    Utility.WriteActivity(Path.GetFileName(f) + "####");
                }
                foreach (string d in Directory.GetDirectories(dir))
                {
                    Utility.WriteActivity(Path.GetFileName(d));
                    DirectorySearch(d);
                }
            }
            catch (System.Exception ex)
            {
                LogError(ex.Message);
            }
        }

        void PdfToPng(string inputFile, string outputFileName)
        {
            string msg = "Invalid numeric value";
            int dpi = ParseInt(txtDpi.Text, out bool valid);
            if (!valid)
            {
                LogError(msg);
                return;
            }
            var xDpi = dpi; //set the x DPI
            var yDpi = dpi; //set the y DPI
            var pageNumber = 1; // the pages in a PDF document

            using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
            {
                rasterizer.Open(inputFile); //opens the PDF file for rasterizing
                Console.WriteLine("In Path: " + inputFile);
                Console.WriteLine("In Count: " + rasterizer.PageCount.ToString());

                //set the output image(png's) complete path
                var outputPNGPath = Path.Combine(outputFolder, string.Format("{0}.png", outputFileName));
                Console.WriteLine("Out: " + Path.Combine(outputFolder, string.Format("{0}.png", outputFileName)));

                //converts the PDF pages to png's 
                var pdf2PNG = rasterizer.GetPage(xDpi, yDpi, pageNumber);

                //save the png's
                pdf2PNG.Save(outputPNGPath, ImageFormat.Png);

                Console.WriteLine("Saved " + outputPNGPath);
            }
        }

        void PdfToJpg(string inputFile, string outputFileName)
        {
            string msg = "Invalid numeric value";
            int dpi = ParseInt(txtDpi.Text, out bool valid);
            if (!valid)
            {
                LogError(msg);
                return;
            }
            var xDpi = dpi; //set the x DPI
            var yDpi = dpi; //set the y DPI
            var pageNumber = 1; // the pages in a PDF document

            using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
            {
                rasterizer.Open(inputFile); //opens the PDF file for rasterizing
                Console.WriteLine("In Path: " + inputFile);
                Console.WriteLine("In Count: " + rasterizer.PageCount.ToString());

                //set the output image(png's) complete path
                var outputPNGPath = Path.Combine(outputFolder, string.Format("{0}.jpg", outputFileName));
                Console.WriteLine("Out: " + Path.Combine(outputFolder, string.Format("{0}.jpg", outputFileName)));

                //converts the PDF pages to png's 
                var pdf2TIF = rasterizer.GetPage(xDpi, yDpi, pageNumber);

                //save the png's
                pdf2TIF.Save(outputPNGPath, ImageFormat.Jpeg);

                Console.WriteLine("Saved " + outputPNGPath);
            }
        }

        public void Crop(string imagePath, string outputPath)
        {
            if (imagePath == null || outputPath == null)
            {
                return;
            }

            bool valx, valy;
            string msg = "Invalid numeric value";
            Point ptFrom = ParsePoint(txtFrom.Text, out valx);
            Point ptTo = ParsePoint(txtTo.Text, out valy);
            if (!valx || !valy)
            {
                LogError(msg);
                return;
            }

            Tuple<int, int> from = new Tuple<int, int>(ptFrom.X, ptFrom.Y);
            Tuple<int, int> to = new Tuple<int, int>(ptTo.X, ptTo.Y);

            using (MagickImage image = new MagickImage(imagePath))
            {
                image.Crop(new MagickGeometry(from.Item1, from.Item2, to.Item1 - from.Item1, to.Item2 - from.Item2));
                image.Grayscale();
                image.Write(outputPath);
            }
        }

        public string TextFromImage(string pth)
        {
            string result = "";
            try
            {
                ProcessStartInfo start = new ProcessStartInfo();
                start.FileName = txtPythonFolder.Text + @"python.exe";
                start.Arguments = "\"" + Application.StartupPath + "\\scripts\\ocr.py\" " + "\"" + pth + "\"";
                Utility.WriteActivity("SP: " + start.Arguments.ToString());
                start.UseShellExecute = false;
                start.WindowStyle = ProcessWindowStyle.Hidden;
                start.CreateNoWindow = true;
                start.RedirectStandardOutput = true;
                using (Process process = Process.Start(start))
                {
                    using (StreamReader reader = process.StandardOutput)
                    {
                        result = reader.ReadToEnd();
                        Utility.WriteActivity("Text: " + result);
                    }
                }
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                return ("");
            }
            return result;
        }

        public static bool CheckNonAlpha(string str)
        {
            if (string.IsNullOrEmpty(str))
                return false;

            for (int i = 0; i < str.Length; i++)
            {
                if (!(char.IsLetter(str[i])) && (!(char.IsNumber(str[i]))))
                    return false;
            }

            return true;
        }

        public static Point ParsePoint(string strPt, out bool valid)
        {
            string msg;
            valid = true;
            Point pt = new Point(10000, 10000);
            if (strPt == null)
            {
                valid = false;
                return pt;
            }

            var ptsArr = strPt.Split(',');
            int x, y;
            if (int.TryParse(ptsArr[0], out x))
            {

            }
            else
            {
                msg = "Invalid Size value";
                MessageBox.Show(msg);
                LogError(msg);
                valid = false;
                return pt;
            }

            if (int.TryParse(ptsArr[1], out y))
            {

            }
            else
            {
                msg = "Invalid Size value";
                MessageBox.Show(msg);
                LogError(msg);
                valid = false;
                return pt;
            }

            pt.X = x;
            pt.Y = y;

            return pt;
        }

        public static int ParseInt(string num, out bool valid)
        {
            int x = 0;
            string msg;
            if (int.TryParse(num, out x))
            {
                valid = true;
            }
            else
            {
                msg = "Invalid Size value";
                MessageBox.Show(msg);
                LogError(msg);
                valid = false;
                return x;
            }

            return x;
        }

        public static bool NameHasInvalidChars(string path)
        {
            return (!string.IsNullOrEmpty(path) && path.IndexOfAny(System.IO.Path.GetInvalidPathChars()) >= 0);
        }

        public void UpdateBillingDatatable(DataTable table)
        {
            foreach (DataRow row in table.Rows)
            {
                string code = row["Code"].ToString();
                string email = row["Email"].ToString();
                string[] accounts;

                if (row["Accounts"].ToString().Length > 1)
                {
                    accounts = row["Accounts"].ToString().Split(';');
                }
                else
                {
                    accounts = new string[] { " " };
                }
                string docs = "";
                docs = GetBillingDocuments(txtBillingExports.Text, code, accounts);
                row["Documents"] = docs;
                if (email.Trim().Length > 2)
                {
                    row["Send"] = true;
                }
            }
        }

        public string GetBillingDocuments(string folder, string facility, string[] accounts)
        {
            DirectoryInfo d = new DirectoryInfo(@folder);
            FileInfo[] Files = d.GetFiles("*.pdf"); //Getting Text files
            string docs = "", identity = "", acct = "", doc = "";
            foreach (FileInfo file in Files)
            {
                string no_ext = System.IO.Path.GetFileNameWithoutExtension(file.Name);
                doc = file.Name.Trim();
                if (no_ext.Contains("_"))
                {
                    identity = no_ext.Substring(no_ext.LastIndexOf('_') + 1).Trim();
                }
                else
                {
                    identity = "invalid";
                }

                acct = accounts.Where(s => file.Name.Contains(s)).FirstOrDefault();

                if (acct != null && acct != facility.Trim())
                {
                    if (acct.Trim() != "")
                    {
                        Utility.WriteActivity("Report Acct: " + identity + ": " + facility.Trim() + ": " + file.Name);
                        docs += doc + ";";
                    }

                }
                else if (identity.Trim() == facility.Trim())
                {
                    Utility.WriteActivity("Report Fac: " + identity + ": " + facility.Trim() + ": " + file.Name);
                    docs += doc + ";";
                }
            }
            return docs;
        }

        public bool SendBillingDocs(string code, string docs, string email)
        {
            string user = Environment.UserName;  //System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string notes = dpBilling.Value.ToString("yyyy-MM");
            string msg = prop.SendEmailMessage;
            string subject = prop.SendEmailSubject;


            //Utility.WriteActivity("Code: " + user + code + " Email: " + email + " Docs: " + docs);
            bool sent;
            string[] documents = docs.Trim().Split(';');
            for (int i = 0; i < documents.Length; i++)
            {
                if (documents[i].Trim() != "")
                    documents[i] = txtBillingExports.Text + documents[i];
            }

            if (cbNotifyOnly.Checked)
            {
                Array.Clear(documents, 0, documents.Length);
                documents = null;
                msg = prop.NotifyEmailMessage;
                subject = prop.NotifyEmailSubject;
            }

            string[] addresses = email.Trim().Split(';');
            for (int i = 0; i < addresses.Length; i++)
            {
                if (addresses[i].Trim() != "")
                {
                    sent = SendEmail(msg, subject, addresses[i].Trim(), "operations@ihspharmacy.com", "IHS Pharmacy", documents);
                    if (sent)
                    {
                        InsertFAC_TRANS("BILLING_EMAIL", code, docs, addresses[i].Trim(), notes, user);

                    }
                }
            }

            return true;

        }

        public void ClearRpt()
        {
            txtDataRpt.Text = "";
            txtSendRpt.Text = "";
        }

        public void ClearAccList()
        {
            txtAccID.Text = "";
            txtAccName.Text = "";
            txtAccGroupCode.Text = "";
            txtAccContact1.Text = "";
            txtAccContact2.Text = "";
            txtAccAddress1.Text = "";
            txtAccAddress2.Text = "";
            txtAccAddress3.Text = "";
            txtAccCity.Text = "";
            txtAccCity2.Text = "";
            ddAccStates.SelectedValue = "";
            txtAccZip.Text = "";
            txtAccZip2.Text = "";
            txtAccPhone.Text = "";
            txtAccEmail.Text = "";
            txtAccEmail2.Text = "";
            txtAccEmail3.Text = "";
            ddAccType.SelectedValue = "";
            ddAccTerms.SelectedValue = "";
            txtAccRep.Text = "";
            txtAccStatements.Text = "";
            txtAccShipTo.Text = "";
            txtAccTax.Text = "";
            txtAccEmailStmts.Text = "";
            txtAccInvoiceNumber.Text = "";

            AccInsert = true;
            Acc_ID = "";
            btnAccUpdate.Text = "Add Account";
        }

        public bool GetUpdate()
        {
            string update_dir = prop.update;
            //string update_dir = @"\\192.168.50.202\dev\update\";
            string app_folder = Application.StartupPath + @"\";
            string local_file = "", remote_file = "";

            try
            {
                FileVersionInfo RemoteFileInfo;
                FileVersionInfo.GetVersionInfo(Path.Combine(Application.StartupPath, "OfficeAutomation.exe"));
                FileVersionInfo LocalFileInfo = FileVersionInfo.GetVersionInfo
                    (Application.StartupPath + "\\OfficeAutomation.exe");
                local_file = LocalFileInfo.FileVersion;

                if (File.Exists(update_dir + @"OfficeAutomation.exe"))
                {
                    FileVersionInfo.GetVersionInfo(Path.Combine(update_dir, "OfficeAutomation.exe"));
                    RemoteFileInfo = FileVersionInfo.GetVersionInfo
                        (update_dir + "\\OfficeAutomation.exe");
                    remote_file = RemoteFileInfo.FileVersion;
                }

                System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(update_dir);
                int count = dir.GetFiles().Length;
                if (String.Equals(local_file, remote_file) || remote_file == "")
                {
                    string msg = "No Update Available";
                    Utility.WriteActivity(msg);
                    //MessageBox.Show(msg);
                }
                else
                {
                    if (MessageBox.Show("An update is available\nDo you want to apply it?", "Update Available", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return false;
                    }
                    Utility.WriteActivity("Update for " + count.ToString() + " file(s)");
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = "AutoUpdate.exe";
                    startInfo.Arguments = update_dir;
                    Process.Start(startInfo);
                }
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show(ex.Message);
                return false;
            }

            return true;
        }

        public void ClearManualCharges()
        {
            txtMC_Desc.Text = "";
            txtMC_Price.Text = "";
            txtMC_Qty.Text = "";
            ddMC_Category.SelectedValue = "";
            ddMC_Account.SelectedValue = "";
            btnMC_Update.Text = "Save";
            lbMC_AccCode.Text = "";
            lbMC_AccName.Text = "";
            dpMC_Date.Value = DateTime.Now;
        }

        public static string GetStringBetweenCharacters(string input, char charFrom, char charTo)
        {
            int posFrom = input.IndexOf(charFrom);
            if (posFrom != -1) //if found char
            {
                int posTo = input.IndexOf(charTo, posFrom + 1);
                if (posTo != -1) //if found char
                {
                    return input.Substring(posFrom + 1, posTo - posFrom - 1);
                }
            }

            return string.Empty;
        }

        #endregion --- End Utility Functions

        #region Click Events
        private void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                string typ = cbImportType.GetItemText(cbImportType.SelectedItem);
                if (typ == "File")
                {
                    OpenFileDialog fd = new OpenFileDialog();
                    fd.Filter = "Excel Files | *.xlsx; *.xls";
                    Utility.WriteActivity("Open File Dialog");

                    if (fd.ShowDialog() == DialogResult.OK)
                    {
                        readExcelFile(fd.FileName);
                    }
                }
                else
                {
                    if (typ == "Remote")
                    {
                        var current_date = DateTime.Now.ToString("MM-dd-yyyy");
                        var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
                        if (current_date == report_date)
                        {
                            DialogResult result = MessageBox.Show("The select date is the same as the current date\nDo you want to use it?",
                                "Use Current Date", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (result == DialogResult.No)
                            {
                                return;
                            }
                        }

                        GetCalNotifications();
                    }
                    else
                    {
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
            }
        }

        private void btnFileExport_Click(object sender, EventArgs e)
        {
            try
            {
                string[] parms = { "-A", "Facility:DJ", "-A", "DateAfter:05-01-2020" };
                ExportReport(txtNotifyReport.Text, txtNotifyExports.Text, "pdf", parms, txtDSN_CIPS.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
                return;
            }
        }

        private void btnRptFile_Click(object sender, EventArgs e)
        {
            try
            {
                var btnName = ((sender as System.Windows.Forms.Button).Name);
                (sender as System.Windows.Forms.Button).BackColor = Color.Yellow;
                var folder = "";
                OpenFileDialog fbd = new OpenFileDialog();
                fbd.Filter = "Report Files | *.rpt";

                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    if (string.IsNullOrEmpty(fbd.FileName.ToString()))
                    {
                        return;
                    }

                    folder = fbd.FileName;
                    //txtNotifyReport.Text = folder;
                    //WriteConfig("NotifyReport", folder);
                    //Utility.WriteActivity("Notify Report updated");
                }
                else
                {
                    return;
                }

                switch (btnName)
                {
                    case "btnNotifyReport":
                        txtNotifyReport.Text = folder;
                        WriteConfig("NotifyReport", folder);
                        Utility.WriteActivity("Notify report updated");
                        break;
                    case "btnCE_Report":
                        txtCE_Report.Text = folder;
                        WriteConfig("CE_Report", folder);
                        Utility.WriteActivity("Cotrolled Export report updated");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                (sender as System.Windows.Forms.Button).BackColor = Color.Transparent;
            }
        }

        private void btnFolder_Click(object sender, EventArgs e)
        {
            try
            {
                var btnName = ((sender as System.Windows.Forms.Button).Name);
                (sender as System.Windows.Forms.Button).BackColor = Color.Yellow;
                var folder = "";
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    if (string.IsNullOrEmpty(fbd.SelectedPath.ToString()))
                    {
                        return;
                    }

                    folder = fbd.SelectedPath + @"\";
                }

                else
                {
                    return;
                }
                switch (btnName)
                {
                    case "btnNotifyExports":
                        txtNotifyExports.Text = folder;
                        WriteConfig("NotifyExports", folder);
                        Utility.WriteActivity("Notify Exports folder updated");
                        break;
                    case "btnBillingRptFolder":
                        txtBillingRptFolder.Text = folder;
                        WriteConfig("BillingRptFolder", folder);
                        Utility.WriteActivity("Billing Reports folder updated");
                        break;
                    case "btnAltRptFolder":
                        txtAltRptFolder.Text = folder;
                        break;
                    case "btnBillingExports":
                        txtBillingExports.Text = folder;
                        WriteConfig("BillingExports", folder);
                        Utility.WriteActivity("Billing Exports folder updated");
                        break;
                    case "btnDownloadFolder":
                        txtDownloadFolder.Text = folder;
                        WriteConfig("download", folder);
                        Utility.WriteActivity("Download folder updated");
                        break;
                    case "btnProcessFolder":
                        txtProcessFolder.Text = folder;
                        WriteConfig("download", folder);
                        Utility.WriteActivity("Process folder updated");
                        break;
                    case "btnRenamedFolder":
                        txtRenamedFolder.Text = folder;
                        WriteConfig("renamed", folder);
                        Utility.WriteActivity("Renamed folder updated");
                        break;
                    case "btnPythonFolder":
                        txtPythonFolder.Text = folder;
                        WriteConfig("python", folder);
                        Utility.WriteActivity("Renamed folder updated");
                        break;
                    case "btnUpdateFolder":
                        txtUpdateFolder.Text = folder;
                        WriteConfig("update", folder);
                        Utility.WriteActivity("Update folder updated");
                        break;
                    case "btnCE_Export":
                        txtCE_Export.Text = folder;
                        WriteConfig("CE_Export", folder);
                        Utility.WriteActivity("Controlled Export folder updated");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                (sender as System.Windows.Forms.Button).BackColor = Color.Transparent;
            }
        }

        private void btnTextBox_Click(object sender, EventArgs e)
        {
            try
            {
                var btnName = ((sender as System.Windows.Forms.Button).Name);
                (sender as System.Windows.Forms.Button).BackColor = Color.Yellow;

                if (MessageBox.Show("Do you want to update the selected value?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }
                switch (btnName)
                {
                    case "btnCIPS":
                        UpdateSettings("CIPS", txtCIPS.Text, "CIPS connection changed from ", true);
                        CONN_CIPS = txtCIPS.Text;
                        break;
                    case "btnRxBackend":
                        UpdateSettings("RxBackend", txtRxBackend.Text, "RxBackend connection changed from ", true);
                        CONN_RX = txtRxBackend.Text;
                        break;
                    case "btnDSN_CIPS":
                        UpdateSettings("DSN_CIPS", txtDSN_CIPS.Text, "CIPS DSN changed from ", true);
                        break;
                    case "btnDSN_WS":
                        UpdateSettings("DSN_WS", txtDSN_WS.Text, "CIPS Wholesale DSN changed from ", true);
                        break;
                    case "btnWS_ID":
                        UpdateSettings("WS_ID", txtWS_ID.Text, "CIPS Wholesale identifier changed from ", true);
                        break;
                    case "btnDSN_RxBackend":
                        UpdateSettings("DSN_RxBackend", txtDSN_RxBackend.Text, "RxBackend DSN changed from ", true);
                        break;
                    case "btnAddress":
                        UpdateSettings("EmailAddress", txtAddress.Text, "Email Address changed from ", true);
                        CONN_RX = txtRxBackend.Text;
                        break;
                    case "btnPassword":
                        UpdateSettings("EmailPassword", Encrypt(txtPassword.Text), "Email Password changed", false);
                        break;
                    case "btnMailbox":
                        UpdateSettings("Mailbox", txtMailbox.Text, "Mailbox changed from ", true);
                        break;
                    case "btnEmailServer":
                        UpdateSettings("EmailServer", txtEmailServer.Text, "Email server changed from ", true);
                        break;
                    case "btnEmailPort":
                        Utility.WriteActivity("SMTP port changed from " + ReadConfig("EmailPort") + " to " + txtEmailPort.Text);
                        Properties.Settings.Default.EmailPort = Int16.Parse(txtEmailPort.Text);
                        Properties.Settings.Default.Save();
                        break;
                    case "btnForward":
                        UpdateSettings("ForwardAddress", txtForward.Text, "Forwarding Email Address changed from ", true);
                        break;
                    case "btnFaxPrinter":
                        UpdateSettings("FaxPrinter", txtFaxPrinter.Text, "Fax Printer changed from ", true);
                        break;
                    case "btnCalendarID":
                        UpdateSettings("CalendarID", txtCalendarID.Text, "Calendar ID changed from ", true);
                        break;
                    case "btnFaxFolder":
                        UpdateSettings("FaxFolder", txtFaxFolder.Text, "Fax Folder ID changed from ", true);
                        break;
                    case "btnAddressAttachment":
                        UpdateSettings("AddressAttachment", txtAddressAttachment.Text, "Attachment Address changed from ", true);
                        break;
                    case "btnPasswordAttachment":
                        UpdateSettings("PasswordAttachment", Encrypt(txtPasswordAttachment.Text), "Attachment Email Password changed", false);
                        break;
                    case "btnSaveEmail":
                        if (rbEmailSend.Checked)
                        {
                            UpdateSettings("SendEmailSubject", txtEmailSubject.Text, "Send Email Password changed", false);
                            UpdateSettings("SendEmailMessage", txtEmailMessage.Text, "Send Email Message changed", false);
                        }
                        else
                        {
                            UpdateSettings("NotifyEmailSubject", txtEmailSubject.Text, "Notify Email Password changed", false);
                            UpdateSettings("NotifyEmailMessage", txtEmailMessage.Text, "Notify Email Message changed", false);
                        }
                        break;
                }


            }
            catch (Exception ex)
            {
                //LogError(ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                (sender as System.Windows.Forms.Button).BackColor = Color.Transparent;
            }

        }

        private void btnTest_Click(object sender, EventArgs e)
        {
            //string[] arr = { "AU99999999", "AUBOP99999", "AUF9999999", "AUS9999999" };
            //string facility = "AU";
            //string docs = GetBillingDocuments(txtRenamedFolder.Text, facility, arr);
            //txtInfo.Text = docs;
            //string[] attachments = { @"C:\Files\renamed\1.pdf" };
            //var sent = SendEmail("Your ARX Report from IHS Pharmacy is attached", "Your ARX Report is attached ", "hank@dekalbal.com", "operations@ihspharmacy.com", "IHS Pharmacy", attachments);

            FRM_GRIDVIEW frm = new FRM_GRIDVIEW();
            frm.Tag = "SELECT TOP 10 * FROM MANUAL_CHARGES";
            frm.ShowDialog();

        }

        private void gvFac_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowindex = gvFac.CurrentCell.RowIndex;
            string[] emails;
            gvFacEmail.DataSource = null;
            try
            {
                txtGroupCode.Text = gvFac.Rows[rowindex].Cells[0].Value.ToString();
                txtFacilityName.Text = gvFac.Rows[rowindex].Cells[1].Value.ToString();
                txtFacUser.Text = gvFac.Rows[rowindex].Cells["Comments"].Value.ToString();

                if (gvFac.Rows[rowindex].Cells[2].Value.ToString().Length > 4)
                {

                    emails = gvFac.Rows[rowindex].Cells[2].Value.ToString().Split(';');

                    //gvFacEmail.Columns.Add("Address", "Address");
                    //gvFacEmail.Columns.Add("Use", typeof(bool));

                    dtFacEmail = new DataTable();
                    dtFacEmail.Columns.Add("Address", typeof(string));
                    dtFacEmail.Columns.Add("Use", typeof(bool));

                    for (int i = 0; i < emails.Length; i++)
                    {
                        var email = emails[i].Split(':');
                        //gvFacEmail.Rows.Add(new object[] { email[0], email[1].Equals("1")? true: false });
                        dtFacEmail.Rows.Add(email[0], email[1].Equals("1") ? true : false);
                    }
                    gvFacEmail.DataSource = dtFacEmail;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void btnAddNew_Click(object sender, EventArgs e)
        {
            if (btnAddNew.Text == "Add New") {
                txtGroupCode.ReadOnly = false;
                btnAddNew.Text = "Clear\\Update";
                lbUpdate.Text = "Add New";
                ClearFacTextBoxes();
            }
            else
            {
                txtGroupCode.ReadOnly = true;
                txtGroupCode.Text = "";
                btnAddNew.Text = "Add New";
                lbUpdate.Text = "Update";
                ClearFacTextBoxes();
            }
        }

        private void btnCheckGC_Click(object sender, EventArgs e)
        {
            Fac fac = GetFacility(txtGroupCode.Text);
            txtFacilityName.Text = fac.name;
        }

        private void btnFacSave_Click(object sender, EventArgs e)
        {
            SaveFaclity();
        }

        private void btnFacilityEmail_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to send all the email notiifications displayed?",
                "Send Email Notification", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }

            try
            {
                string code = "", typ = "", valid = "", fac_name = "", email = "";
                var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
                for (int i = 0; i + 1 < gvNotiifications.Rows.Count; i++)
                {
                    code = gvNotiifications.Rows[i].Cells[0].Value.ToString();
                    typ = gvNotiifications.Rows[i].Cells[5].Value.ToString();
                    valid = gvNotiifications.Rows[i].Cells[6].Value.ToString();
                    fac_name = gvNotiifications.Rows[i].Cells[1].Value.ToString();
                    email = gvNotiifications.Rows[i].Cells[2].Value.ToString();

                    var att_file = report_date + "_" + code;
                    var att_path = prop.NotifyExports + att_file + ".pdf";

                    bool file_exists = File.Exists(att_path);
                    bool sent = false;

                    if (valid == "True" && (typ == "Email" || typ == "Both"))
                    {
                        if (file_exists)
                        {
                            //email = "dekalb.hda@gmail.com;hank@dekalbal.com;zrefugee@gmail.com;";
                            if (prop.ForwardAddress.ToString() != "")
                            {
                                email = email + prop.ForwardAddress.ToString();
                            }
                            string[] attachments = { att_path };
                            Utility.WriteActivity(fac_name + ": " + email + ": " + att_path);
                            sent = SendEmail("Your ARX Report from IHS Pharmacy is attached", "Your ARX Report is attached ", email, "operations@ihspharmacy.com", "IHS Pharmacy", attachments);

                            if (sent)
                            {
                                LogActivity("FAC_EMAIL", 0, fac_name, report_date);
                                if (typ == "Email")
                                {
                                    gvNotiifications.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                                    File.Delete(att_path);
                                }
                                else
                                {
                                    gvNotiifications.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                                }
                            }

                        }
                        else
                        {
                            Utility.WriteActivity("The file for [" + fac_name + "] does not exist");
                        }
                    }


                }
                Utility.WriteActivity("Email transactions complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
                return;
            }


        }

        private void btnFacilityFax_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to send all the fax notiifications displayed?",
            "Send Fax Notification", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }

            try
            {

                FaxExportToFolder();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
                return;
            }

        }

        private void btnExportReports_Click(object sender, EventArgs e)
        {
            if (rbBilling.Checked == true)
            {
                CreateBillingReports();
            }
            else
            {
                try
                {
                    txtReportIssues.Text = "";
                    var folder = prop.BillingRptFolder;
                    var report_date = dpBilling.Value.ToString("yyyy-MM");
                    if (txtAltRptFolder.Text.Trim() != "")
                    {
                        folder = txtAltRptFolder.Text;
                    }

                    //ExportReport(txtNotifyReport.Text, txtNotifyExports.Text, "pdf", parms, txtDSN_CIPS.Text);

                    foreach (string file in Directory.EnumerateFiles(folder, "*.rpt"))
                    {
                        string contents = file.ToString();
                        Utility.WriteActivity(contents);

                        string dsn = file.ToString().Contains(txtWS_ID.Text.Trim()) ? txtDSN_WS.Text : txtDSN_CIPS.Text;

                        string[] rpt = { "-S", dsn,
                    "-F", file,
                    "-O", txtBillingExports.Text + report_date + "_" + Path.GetFileNameWithoutExtension(file) + ".pdf",
                    "-E", "pdf"};

                        bool success = RunReport(rpt);

                        if (!success)
                        {
                            txtReportIssues.Text += file + "\r\n";
                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    Utility.WriteActivity(ex.Message);
                    return;
                }
            }
        }

        //Attachment Processing
        private void btnDownload_Click(object sender, EventArgs e)
        {
            try
            {
                var msg_ret = DownloadAttachments();
                Utility.WriteActivity(msg_ret);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
                return;
            }
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            Stopwatch watch = new Stopwatch();
            watch.Start();
            Utility.WriteActivity("File renaming started");
            var title = this.Text;
            var txt = "";
            var outPath = "";
            var cnt = 0;
            this.Text = title + " - Renaming...";

            try
            {
                var pdfFiles = Directory.GetFiles(txtProcessFolder.Text, "*.pdf");
                //process each PDF file
                foreach (var pdfFile in pdfFiles)
                {

                    var fileName = Path.GetFileNameWithoutExtension(pdfFile);
                    Utility.WriteActivity("Renaming file: " + pdfFile);
                    PdfToJpg(pdfFile, fileName);
                    cropImgIn = outputFolder + fileName + ".jpg";
                    cropImgOut = txtRenamedFolder.Text;

                    Crop(cropImgIn, cropImgIn);
                    var outName = TextFromImage(cropImgIn).Trim();
                    if (string.IsNullOrEmpty(outName))
                    {
                        Utility.WriteActivity("No text for file found");
                        continue;
                    }

                    outName = outName.Replace(" ", "");

                    if (CheckNonAlpha(outName))
                    {
                        outPath = outputFolderNew + outName + ".pdf";
                    }
                    else
                    {
                        outName = "INVALID" + cnt.ToString();
                        outPath = outputFolderNew + outName + ".pdf";
                        Utility.WriteActivity("Illegal characters in file name");
                    }

                    if (File.Exists(outPath))
                    {
                        var msg = "The file " + outPath + " already exists";
                        Utility.WriteActivity(msg);
                        LogError(msg);
                        string fileAppend = DateTime.Now.AddSeconds(cnt).ToString("MMddyyyyhhmmss") + "_";
                        var existingFile = outputFolderNew + fileAppend + outName + ".pdf";
                        File.Copy(pdfFile, existingFile);
                        File.Delete(cropImgIn);
                        Utility.WriteActivity("File written as: " + existingFile);
                        cnt++;
                        continue;
                    }

                    try
                    {
                        File.Copy(pdfFile, outPath);
                        File.Delete(cropImgIn);
                        Utility.WriteActivity("File: " + outPath + " written");
                    }
                    catch (Exception ex)
                    {
                        LogError(ex.Message);
                        Utility.WriteActivity(ex.Message);
                        continue;
                    }
                    finally
                    {
                        cnt++;
                    }
                }

            }
            catch (Exception ex)
            {
                LogError(ex.Message);
                Utility.WriteActivity(ex.Message);
            }
            finally
            {
                this.Text = title;
            }

            watch.Stop();
            txt = watch.Elapsed.TotalSeconds.ToString();
            Utility.WriteActivity("Process Time: " + txt);
        }

        private void btnSingleFile_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            var pth = "";
            ofd.Filter = "PDF files|*.pdf";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pth = ofd.FileName;
            }

            var fileName = Path.GetFileNameWithoutExtension(pth);
            PdfToPng(pth, fileName);
            cropImgIn = outputFolder + fileName + ".png";
            cropImgOut = txtRenamedFolder.Text;

            Crop(cropImgIn, cropImgIn);
            TextFromImage(cropImgIn);
            var outName = TextFromImage(cropImgIn);
            var outPath = outputFolderNew + outName + ".pdf";

            if (File.Exists(outPath))
            {
                var msg = "The file " + outPath + " already exists";
                MessageBox.Show(msg, "File Exits", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Utility.WriteActivity(msg);
            }
            File.Copy(pth, outPath);
            File.Delete(cropImgIn);
            Utility.WriteActivity("File: " + outPath + " written");
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            var ofd = new OpenFileDialog();
            var pth = "";
            ofd.Filter = "PDF files|*.pdf";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                pth = ofd.FileName;
            }

            PdfToPng(pth, pth);
            //cropImgIn = outputFolder + fileName + ".png";
            cropImgIn = pth + ".png";

            Crop(cropImgIn, cropImgIn);
            Utility.WriteActivity(cropImgIn);
            txtCheck.Text = cropImgIn;
            pbCheck.Image = Image.FromFile(cropImgIn);
        }

        private void btnPreview_Click(object sender, EventArgs e)
        {
            try
            {
                GetFacilityBIlling();
                Utility.WriteActivity("Preview Loading Complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
            }
        }

        private void gvStaged_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int cell = gvStaged.CurrentCell.ColumnIndex;
            int row = e.RowIndex;
            string val = "", facility = "";
            string[] vals;
            if (e.RowIndex == -1) return; //check if row index is not selected
            {
                if (cell.Equals(3) || cell.Equals(4) || cell.Equals(5))
                    if (gvStaged.CurrentCell != null && gvStaged.CurrentCell.Value != null)
                    {
                        val = gvStaged.CurrentCell.Value.ToString();
                        vals = gvStaged.CurrentCell.Value.ToString().Split(';');
                        facility = gvStaged.Rows[row].Cells[2].Value.ToString();
                        lbPreview.Text = facility;
                        txtPreview.Text = "";

                        for (int i = 0; i < vals.Length; i++)
                        {
                            txtPreview.Text += vals[i] + Environment.NewLine;
                        }

                    }
            }

        }

        private void gvBillingSent_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;
            string code = gvBillingSent.Rows[row].Cells[1].Value.ToString();
            //Mess
            Utility.WriteActivity(code);
        }

        private void gvStaged_RowHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;

            string email = gvStaged.Rows[row].Cells["Email"].EditedFormattedValue.ToString();
            string docs = gvStaged.Rows[row].Cells["Documents"].Value.ToString();
            string code = gvStaged.Rows[row].Cells["Code"].Value.ToString();
            string send = gvStaged.Rows[row].Cells["Send"].Value.ToString();
            string name = gvStaged.Rows[row].Cells["Facility"].Value.ToString();

            DialogResult result = MessageBox.Show("Do you want to send the billing documents for\n" + name + "?",
                "Send Billing Documents", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }

            //if (send == "True" && email.Trim().Length > 2 && docs.Trim().Length > 2)
            if (email.Trim().Length > 2 && (docs.Trim().Length > 2 || cbNotifyOnly.Checked))
            {
                SendBillingDocs(code, docs, email);
            }

        }

        private void gvStaged_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //int row = e.RowIndex;

            //string docs = gvStaged.Rows[row].Cells["Documents"].Value.ToString();
            //string accounts = gvStaged.Rows[row].Cells["Accounts"].EditedFormattedValue.ToString();
            //string name = gvStaged.Rows[row].Cells["Facility"].Value.ToString();

            //DialogResult result = MessageBox.Show("Do you want to add the accounts to\n" + name + "?",
            //    "Add Accounts", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (result == DialogResult.No)
            //{
            //    return;
            //}
            ////if (docs.Trim().Length > 0)
            ////{
            ////    accounts = ";" + accounts;
            ////}

            //gvStaged.Rows[row].Cells["Documents"].Value = docs + accounts + ".pdf";
        }

        private void gvStaged_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int cell = gvStaged.CurrentCell.ColumnIndex;
            int row = e.RowIndex;
            string docs = "";
            string name = "";
            string val = "";
            string[] vals;
            if (e.RowIndex == -1) return; //check if row index is not selected
            {
                if (cell.Equals(3))
                {
                    val = gvStaged.CurrentCell.EditedFormattedValue.ToString() + ".pdf";
                    vals = gvStaged.CurrentCell.EditedFormattedValue.ToString().Split(';');
                    docs = gvStaged.Rows[row].Cells["Documents"].Value.ToString();
                    name = gvStaged.Rows[row].Cells["Facility"].Value.ToString();

                    if (vals.Length > 1)
                    {
                        val = "";
                        for (int i = 0; i < vals.Length; i++)
                        {
                            if (i != vals.Length)
                            {
                                val += vals[i] + ".pdf;";
                            }
                            else
                            {
                                val += vals[i] + ".pdf";
                            }
                        }
                    }
                }
                else
                {
                    return;
                }
            }
            DialogResult result = MessageBox.Show("Do you want to add the accounts to\n" + name + "?",
                "Add Accounts", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }

            gvStaged.Rows[row].Cells["Documents"].Value = docs + val;
            gvStaged.Rows[row].Cells["Accounts"].Selected = false;
            gvStaged.Rows[row].Cells["Documents"].Selected = true;

        }

        private void gvRpt_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;

            DATA_RPT = gvRpt.Rows[row].Cells["Data_Facility_Code"].Value.ToString();
            SEND_RPT = gvRpt.Rows[row].Cells["Send_Facility_Code"].Value.ToString();
            txtDataRpt.Text = DATA_RPT;
            txtSendRpt.Text = SEND_RPT;

            btnAddRpt.Text = "Update";
        }

        private void btnSendSelected_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to send emails to the selected facilities?", "Send Email", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            string email, docs, code, send;
            try
            {
                if (dtBilling != null)
                {
                    foreach (DataGridViewRow row in gvStaged.Rows)
                    {
                        email = row.Cells["Email"].EditedFormattedValue.ToString();
                        docs = row.Cells["Documents"].Value.ToString();
                        code = row.Cells["Code"].Value.ToString();
                        send = row.Cells["Send"].Value.ToString();

                        if (send == "True" && email.Trim().Length > 2 && (docs.Trim().Length > 2 || cbNotifyOnly.Checked))
                        {
                            SendBillingDocs(code, docs, email);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
            }

        }

        private void btnRefreshSent_Click(object sender, EventArgs e)
        {
            try
            {
                GetSentBIlling();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
            }
        }

        private void gvBillingSent_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int cell = gvBillingSent.CurrentCell.ColumnIndex;
            int row = e.RowIndex;
            string val = "", facility = "";
            string[] vals;
            if (e.RowIndex == -1) return; //check if row index is not selected
            {
                if (cell.Equals(4) || cell.Equals(5))
                    if (gvBillingSent.CurrentCell != null && gvBillingSent.CurrentCell.Value != null)
                    {
                        val = gvBillingSent.CurrentCell.Value.ToString();
                        vals = gvBillingSent.CurrentCell.Value.ToString().Split(';');
                        facility = gvBillingSent.Rows[row].Cells[3].Value.ToString();
                        lbSent.Text = facility;
                        txtSent.Text = "";

                        for (int i = 0; i < vals.Length; i++)
                        {
                            txtSent.Text += vals[i] + Environment.NewLine;
                        }

                    }
            }
        }

        private void btnEmailExports_Click(object sender, EventArgs e)
        {

        }

        private void btnRefreshFacSettings_Click(object sender, EventArgs e)
        {
            LoadFacilities();
        }

        private void btnUserGuide_Click(object sender, EventArgs e)
        {
            try
            {
                var pth = Application.StartupPath + "\\scripts\\Office Automation Guide.docx";
                Process.Start(pth);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Utility.WriteActivity(ex.Message);
            }
        }

        private void btnFrom_Click(object sender, EventArgs e)
        {
            try
            {
                var btnName = ((sender as Button).Name);
                bool valid;
                string msg = "Invalid numeric value";
                (sender as Button).BackColor = Color.Yellow;
                Point pt = new Point(10000, 10000);
                if (MessageBox.Show("Do you want to update the selected value?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                switch (btnName)
                {
                    case "btnFrom":
                        pt = ParsePoint(txtFrom.Text, out valid);
                        if (!valid)
                        {
                            LogError(msg);
                            return;
                        }
                        Properties.Settings.Default["fromCrop"] = pt;
                        Properties.Settings.Default.Save();
                        Utility.WriteActivity("Crop From updated");
                        break;
                    case "btnTo":
                        pt = ParsePoint(txtTo.Text, out valid);
                        if (!valid)
                        {
                            LogError(msg);
                            return;
                        }
                        Properties.Settings.Default["toCrop"] = pt;
                        Properties.Settings.Default.Save();
                        Utility.WriteActivity("Crop To updated");
                        break;
                    case "btnDpi":
                        int x = ParseInt(txtDpi.Text, out valid);
                        if (!valid)
                        {
                            LogError(msg);
                            return;
                        }
                        Properties.Settings.Default["dpi"] = x;
                        Properties.Settings.Default.Save();
                        Utility.WriteActivity("Image resolution updated");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogError(ex.Message);
                MessageBox.Show(ex.Message);
            }
            finally
            {
                (sender as Button).BackColor = Color.Transparent;
            }
        }

        private void btnAddEmail_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtGroupCode.Text))
            {
                Utility.WriteActivity("A Group Code must be selected before adding an email address");
                return;
            }

            if (!ValidateEmail(txtFacEmail.Text.Trim()))
            {
                MessageBox.Show("The Email Address entered is not valid", "Invalid Email Address", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var conn = new SqlConnection(CONN_RX);
            conn.Open();

            var sql = @"INSERT INTO [FAC_EMAIL]
                        ([FAC_CODE]
                        ,[ADDRESS]
                        ,[Billing])";
            sql += " VALUES ('" + txtGroupCode.Text.Trim();
            sql += "','" + txtFacEmail.Text.Trim();
            sql += "', 1 )";
            var com = new SqlCommand(sql, conn);
            try
            {
                com.ExecuteNonQuery();
                MessageBox.Show("Saved...");
            }
            catch (Exception ex)
            {
                Utility.WriteActivity(ex.Message);
                MessageBox.Show("Not Saved");
            }
            finally
            {
                conn.Close();
                ClearFacTextBoxes();
                LoadFacilities();

                try
                {
                    String searchValue = txtGroupCode.Text;
                    int rowIndex = -1;
                    foreach (DataGridViewRow row in gvFac.Rows)
                    {
                        if (row.Cells[0].Value.ToString().Equals(searchValue))
                        {
                            rowIndex = row.Index;
                            break;
                        }
                    }

                    gvFac.CurrentCell = gvFac.Rows[rowIndex].Cells[0];
                }
                catch
                {

                }
            }
        }

        private void btnUpdateAddresses_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtGroupCode.Text) || gvFacEmail.Rows.Count < 1)
            {
                Utility.WriteActivity("A Group Code with email addresses must be selected before updating");
                return;
            }
            var code = txtGroupCode.Text;
            string address = "", use = "";
            var conn = new SqlConnection(CONN_RX);
            conn.Open();

            foreach (DataGridViewRow row in gvFacEmail.Rows)
            {
                address = row.Cells["Address"].Value.ToString();
                use = row.Cells["Use"].Value.ToString();
                use = use == "True" ? "1" : "0";

                var sql = "UPDATE [FAC_EMAIL] SET [Billing] = " + use;
                sql += " WHERE [FAC_CODE]= '" + code + "' AND [ADDRESS]= '" + address + "'";

                var com = new SqlCommand(sql, conn);
                try
                {
                    use = use == "1" ? "True" : "False";
                    com.ExecuteNonQuery();
                    Utility.WriteActivity("Setting for " + address + " saved as " + use);
                }
                catch (Exception ex)
                {
                    Utility.WriteActivity(ex.Message);
                    MessageBox.Show("Not Saved");
                }
                finally
                {

                }
            }
            conn.Close();
            LoadFacilities();
        }

        private void btnRefreshRpt_Click(object sender, EventArgs e)
        {
            LoadReporting();
            ClearRpt();
        }

        private void btnAddRpt_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtDataRpt.Text) || string.IsNullOrWhiteSpace(txtSendRpt.Text))
            {
                MessageBox.Show("You must have a Group Code for data and sending", "No Group Codes", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveReport(DATA_RPT, SEND_RPT);
            ClearRpt();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearRpt();
            btnAddRpt.Text = "Add";
        }

        private void gvRpt_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mRpt = new ContextMenu();
                string fac_data = "", fac_send = "", data_name = "", send_name = "", id = "";

                int currentMouseOverRow = gvRpt.HitTest(e.X, e.Y).RowIndex;

                if (currentMouseOverRow >= 0)
                {
                    fac_send = gvRpt.Rows[currentMouseOverRow].Cells["Send_Facility_Code"].Value.ToString();
                    fac_data = gvRpt.Rows[currentMouseOverRow].Cells["Data_Facility_Code"].Value.ToString();
                    data_name = gvRpt.Rows[currentMouseOverRow].Cells["Data_Facility_Name"].Value.ToString();
                    send_name = gvRpt.Rows[currentMouseOverRow].Cells["Send_Facility_Name"].Value.ToString();
                    id = gvRpt.Rows[currentMouseOverRow].Cells["ID"].Value.ToString();
                }

                mRpt.MenuItems.Add(new MenuItem("Copy Row", (o, ev) =>
                {
                    //MessageBox.Show((o as MenuItem).Text);
                    Clipboard.SetText(fac_data + ", " + data_name + ", " + fac_send + ", " + send_name);
                }));

                mRpt.MenuItems.Add(new MenuItem("Delete", (o, ev) =>
                {
                    string msg = "Data Facility: " + fac_data
                    + data_name + ", Send Facility: " + fac_send + ", " + send_name;
                    DialogResult result = MessageBox.Show("Do you want to delete the record [" + msg + "]?",
                    "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        string sql = "DELETE FROM FAC_REPORTS WHERE ID =" + id;
                        bool success = Execute_Sql(sql, CONN_RX);
                        if (success)
                            LoadReporting();
                    }
                    else
                    {
                        return;
                    }
                }));

                mRpt.Show(gvRpt, new Point(e.X, e.Y));

            }
        }

        private void gvStaged_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mStaged = new ContextMenu();
                string fac_code = "", fac_name = "";

                int currentMouseOverRow = gvStaged.HitTest(e.X, e.Y).RowIndex;

                if (currentMouseOverRow >= 0)
                {
                    fac_code = gvStaged.Rows[currentMouseOverRow].Cells["Code"].Value.ToString();
                    fac_name = gvStaged.Rows[currentMouseOverRow].Cells["Facility"].Value.ToString();
                }

                //mRpt.MenuItems.Add(new MenuItem("Copy Row", (o, ev) =>
                //{
                //    //MessageBox.Show((o as MenuItem).Text);
                //    Clipboard.SetText(fac_code + ", " + fac_name);
                //}));

                mStaged.MenuItems.Add(new MenuItem("Open Billing Exports Folder for " + fac_code, (o, ev) =>
                {
                    Utility.WriteActivity(fac_code);
                    Files frm = new Files();
                    frm.Facility_code = fac_code;
                    frm.Billing_folder = prop.BillingExports;
                    frm.StartPosition = FormStartPosition.CenterParent;
                    frm.Show(this);
                    frm.Top = this.Top + ((this.Height / 2) - (frm.Height / 2));
                    frm.Left = this.Left + ((this.Width / 2) - (frm.Width / 2));
                }));

                mStaged.Show(gvStaged, new Point(e.X, e.Y));

            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            GetUpdate();
        }

        private void btnAccRefresh_Click(object sender, EventArgs e)
        {
            LoadAccounts();
        }

        private void gvAccList_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;

            txtAccID.Text = gvAccList.Rows[row].Cells["ID"].Value.ToString();
            txtAccName.Text = gvAccList.Rows[row].Cells["AccountName"].Value.ToString();
            txtAccGroupCode.Text = gvAccList.Rows[row].Cells["GroupCode"].Value.ToString();
            txtAccContact1.Text = gvAccList.Rows[row].Cells["Contact1"].Value.ToString();
            txtAccContact2.Text = gvAccList.Rows[row].Cells["Contact2"].Value.ToString();
            txtAccAddress1.Text = gvAccList.Rows[row].Cells["Address1"].Value.ToString();
            txtAccAddress2.Text = gvAccList.Rows[row].Cells["Address2"].Value.ToString();
            txtAccAddress3.Text = gvAccList.Rows[row].Cells["Address3"].Value.ToString();
            txtAccCity.Text = gvAccList.Rows[row].Cells["City"].Value.ToString();
            txtAccCity2.Text = gvAccList.Rows[row].Cells["City2"].Value.ToString();
            ddAccStates.SelectedValue = gvAccList.Rows[row].Cells["State"].Value.ToString();
            txtAccZip.Text = gvAccList.Rows[row].Cells["Zip"].Value.ToString();
            txtAccZip2.Text = gvAccList.Rows[row].Cells["Zip2"].Value.ToString();
            txtAccPhone.Text = gvAccList.Rows[row].Cells["Phone"].Value.ToString();
            txtAccEmail.Text = gvAccList.Rows[row].Cells["Email"].Value.ToString();
            txtAccEmail2.Text = gvAccList.Rows[row].Cells["Email2"].Value.ToString();
            txtAccEmail3.Text = gvAccList.Rows[row].Cells["Email3"].Value.ToString();
            ddAccType.SelectedValue = gvAccList.Rows[row].Cells["Type"].Value.ToString();
            ddAccTerms.SelectedValue = gvAccList.Rows[row].Cells["Terms"].Value.ToString();
            txtAccRep.Text = gvAccList.Rows[row].Cells["Rep"].Value.ToString();
            txtAccStatements.Text = gvAccList.Rows[row].Cells["Stmts"].Value.ToString();
            txtAccShipTo.Text = gvAccList.Rows[row].Cells["ShipTo"].Value.ToString();
            txtAccTax.Text = gvAccList.Rows[row].Cells["Tax"].Value.ToString();
            txtAccEmailStmts.Text = gvAccList.Rows[row].Cells["EmailStmts"].Value.ToString();
            txtAccInvoiceNumber.Text = gvAccList.Rows[row].Cells["InvoiceNumber"].Value.ToString();

            AccInsert = false;
            Acc_ID = gvAccList.Rows[row].Cells["ID"].Value.ToString();
            btnAccUpdate.Text = "Update Account";
        }

        private void gvMC_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;

            txtMC_Desc.Text = gvMC.Rows[row].Cells["DESCRIPTION"].Value.ToString();
            txtMC_Price.Text = gvMC.Rows[row].Cells["PRICE"].Value.ToString();
            txtMC_Qty.Text = gvMC.Rows[row].Cells["QTY"].Value.ToString();
            ddMC_Account.SelectedValue = gvMC.Rows[row].Cells["ACCT"].Value.ToString();
            ddMC_Category.SelectedValue = gvMC.Rows[row].Cells["CATEGORY"].Value.ToString();
            dpMC_Date.Value = DateTime.Parse(gvMC.Rows[row].Cells["DATE"].Value.ToString());
            MC_ID = gvMC.Rows[row].Cells["ID"].Value.ToString();

            btnMC_Update.Text = "Update";
        }

        private void btnAccClear_Click(object sender, EventArgs e)
        {
            ClearAccList();
        }

        private void btnAccUpdate_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(txtAccID.Text) || String.IsNullOrWhiteSpace(txtAccName.Text)
                || String.IsNullOrWhiteSpace(txtAccGroupCode.Text))
            {
                MessageBox.Show("An ID, Account Name and Group Code are required", "Missing Item",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (String.IsNullOrWhiteSpace(txtAccAddress1.Text))
            {
                MessageBox.Show("The first Address line is required", "Missing Item",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (String.IsNullOrWhiteSpace(txtAccCity.Text))
            {
                MessageBox.Show("The first City line is required", "Missing Item",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (String.IsNullOrWhiteSpace(ddAccStates.SelectedValue.ToString()))
            {
                MessageBox.Show("A State is required", "Missing Item",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (String.IsNullOrWhiteSpace(txtAccZip.Text))
            {
                MessageBox.Show("The first Zip code line is required", "Missing Item",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            UpdateAddressList(AccInsert);
        }

        private void btnRefreshCodes_Click(object sender, EventArgs e)
        {
            LoadCodes();
        }

        private void btnMC_Update_Click(object sender, EventArgs e)
        {
            
            decimal q,p;

            if (ddMC_Account.SelectedValue.ToString().Trim() == "")
            {
                MessageBox.Show("An Account ID is required", "Missing Account");
                return;
            }

            if (txtMC_Desc.Text.Trim() == "" || ddMC_Category.SelectedValue.ToString().Trim() == "")
            {
                MessageBox.Show("A Descrition and Category are required", "Missing Item");
                return;
            }

            if (decimal.TryParse(txtMC_Qty.Text, out q))
            {
                //valid
            }
            else
            {
                MessageBox.Show("Please enter a valid number for the Quantity", "Invalid Number");
                return;
            }

            if (decimal.TryParse(txtMC_Price.Text, out p))
            {
                //valid
            }
            else
            {
                MessageBox.Show("Please enter a valid number for the Price", "Invalid Number");
                return;
            }

            bool insert = false;

            if (btnMC_Update.Text == "Save")
                insert = true;

            UpdateManualCharges(insert, q, p);
        }

        private void btnMC_Clear_Click(object sender, EventArgs e)
        {
            ClearManualCharges();
        }

        private void btnMC_Refresh_Click(object sender, EventArgs e)
        {
            LoadManualCharges();
            
        }

        private void gvMC_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mMC = new ContextMenu();
                string  id = "", acct = "", price="", qty = "";

                int currentMouseOverRow = gvMC.HitTest(e.X, e.Y).RowIndex;

                if (currentMouseOverRow >= 0)
                {
                    id = gvMC.Rows[currentMouseOverRow].Cells["ID"].Value.ToString();
                    acct = gvMC.Rows[currentMouseOverRow].Cells["ACCT"].Value.ToString();
                    price = gvMC.Rows[currentMouseOverRow].Cells["PRICE"].Value.ToString();
                    qty = gvMC.Rows[currentMouseOverRow].Cells["QTY"].Value.ToString();
                }


                mMC.MenuItems.Add(new MenuItem("Delete (Account: " + acct + " )", (o, ev) =>
                {
                    string msg = "Account: " + acct;
                    DialogResult result = MessageBox.Show("Do you want to delete the record for account id [" + msg + "]?",
                    "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        string sql = "DELETE FROM MANUAL_CHARGES WHERE ID =" + id;
                        bool success = Execute_Sql(sql, CONN_RX);
                        if (success)
                            LogActivity("MAN_CHARGE DEL", 0,
                            acct + "- Q: " + qty.ToString() + ", P: " + price.ToString(), txtMC_Tech.Text);
                            Utility.WriteActivity("Record for Account:[" + acct + "] deleted");
                            LoadManualCharges();
                        return;
                    }
                    else
                    {
                        return;
                    }
                }));

                mMC.Show(gvMC, new Point(e.X, e.Y));

            }
        }

        private void gvCodes_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int row = e.RowIndex;

            ddBilling_Codes.SelectedValue = gvCodes.Rows[row].Cells["CATEGORY"].Value.ToString();
            txtBilling_Code.Text = gvCodes.Rows[row].Cells["DESCRIPTION"].Value.ToString();
            Code_ID = gvCodes.Rows[row].Cells["ID"].Value.ToString();

            btnAddCode.Text = "Update";
        }

        private void btnAddCode_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtBilling_Code.Text) || string.IsNullOrWhiteSpace(ddBilling_Codes.SelectedValue.ToString()))
            {
                MessageBox.Show("You must have a Category and Description", "Missing Item", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SaveCode();
            ClearCodes();
        }

        private void btnClearCode_Click(object sender, EventArgs e)
        {
            ClearCodes();
        }

        private void gvCodes_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu mCode = new ContextMenu();
                string id = "", desc = "", category = "";

                int currentMouseOverRow = gvCodes.HitTest(e.X, e.Y).RowIndex;

                if (currentMouseOverRow >= 0)
                {
                    id = gvCodes.Rows[currentMouseOverRow].Cells["ID"].Value.ToString();
                    desc = gvCodes.Rows[currentMouseOverRow].Cells["DESCRIPTION"].Value.ToString();
                    category = gvCodes.Rows[currentMouseOverRow].Cells["CATEGORY"].Value.ToString();
                }


                mCode.MenuItems.Add(new MenuItem("Delete (Category: " + category + ", Description: " + desc + ")", (o, ev) =>
                {
                    DialogResult result = MessageBox.Show("Do you want to delete the record for (Category: " + category + ", Description: " + desc + ")?",
                    "Delete record", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        string sql = "DELETE FROM BILLING_CODES WHERE ID =" + id;
                        bool success = Execute_Sql(sql, CONN_RX);
                        if (success)
                            LoadCodes();
                        return;
                    }
                    else
                    {
                        return;
                    }
                }));

                mCode.Show(gvCodes, new Point(e.X, e.Y));

            }
        }

        private void btnMC_Import_Click(object sender, EventArgs e)
        {
            var folder = "";
            OpenFileDialog fbd = new OpenFileDialog();
            fbd.Filter = "Excel Files | *.xls;*.xlsx";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(fbd.FileName.ToString()))
                {
                    return;
                }

                DialogResult result = MessageBox.Show("Do you want to import charges from [" + Path.GetFileNameWithoutExtension(fbd.FileName) + "]",
                "Import Charges", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if(result == DialogResult.No)
                {
                    return;
                }

                Utility.WriteActivity("Manual Charge exporting from " + folder);
                readExcelImportCharges(fbd.FileName);
            }

            
            //try
            //{
            //    var folder = "";
            //    OpenFileDialog fbd = new OpenFileDialog();
            //    fbd.Filter = "Excel Files | *.xls;*.xlsx";

            //    if (fbd.ShowDialog() == DialogResult.OK)
            //    {
            //        if (string.IsNullOrEmpty(fbd.FileName.ToString()))
            //        {
            //            return;
            //        }

            //        folder = fbd.FileName;
            //        Utility.WriteActivity("Manual Charge exporting from " + folder);
            //        ImportManualCharges(folder);
            //    }
            //    else
            //    {
            //        return;
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
            //finally
            //{

            //}
        }

        private void btnMC_Export_Click(object sender, EventArgs e)
        {
            ExportManualCharges();
        }

        private void btnMC_RunReport_Click(object sender, EventArgs e)
        {
            List<string> param = new List<string>();

            if (txtMC_Scripts.Lines.Length >= 0)
            {
                for (int x = 0; x < txtMC_Scripts.Lines.Length; x++)
                {
                    param.Add("-A");
                    param.Add("RX_NUMBER:" + txtMC_Scripts.Lines[x].Trim());
                }
            }

            string[] p = param.ToArray();

            var report_date = DateTime.Now.ToString("yyyyMMdd-HHmm");
            string file = prop.CE_Report;
            string exp_path = prop.CE_Export;
            string dsn = prop.DSN_CIPS;
            string full_export = exp_path + report_date + "_" + Path.GetFileNameWithoutExtension(file) + ".pdf";

            string[] rpt = { "-S", dsn,
                    "-F", file,
                    "-O", full_export,
                    "-E", "pdf"};

            var rpt_data = rpt.Concat(p).ToArray();

            bool success = RunReport(rpt_data);

            if (success)
            {
                Utility.WriteActivity("Report exported to " + full_export);
                txtMC_Scripts.Text = "";
            }
                
        }

        private void btnBulkEdit_Click(object sender, EventArgs e)
        {
            DateTime dt = dpMC_Export.Value;
            int year = dt.Year;
            int month = dt.Month;

            FRM_GRIDVIEW frm = new FRM_GRIDVIEW();
            string sql = " WHERE DATEPART(month, [DATE]) = " + month.ToString() + " and DATEPART(year, [DATE]) = " + year.ToString();
            sql += " AND CATEGORY LIKE '%Local Pharmacy%'";
            frm.Tag = "SELECT  * FROM MANUAL_CHARGES " + sql;
            frm.Show();
        }

        #endregion --- End Click

        #region Change Events
        private void txtFacFilter_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (bsFac.DataSource != null)
                {
                    var bd = (BindingSource)gvFac.DataSource;
                    var dt = (DataTable)bd.DataSource;
                    dt.DefaultView.RowFilter = string.Format("Group_Code like '%{0}%'", txtFacFilter.Text.Trim().Replace("'", "''"));
                    gvFac.Refresh();
                }
                else
                {
                    MessageBox.Show("Use the 'Refresh' button to populate the grid", "Grid Not Loaded");
                }
        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void chkDate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDate.CheckState == CheckState.Checked)
            {
                dptDownload.Enabled = true;
            }
            else
            {
                dptDownload.Enabled = false;
            }
        }

        private void cbBillGrid_CheckedChanged(object sender, EventArgs e)
        { 
            if (dtBilling != null)
            {
                var _checked = cbBillGrid.CheckState == CheckState.Checked ? true : false;
                cbBillGrid.Text = _checked ? "Select None" : "Select All";

                foreach (DataRow row in dtBilling.Rows)
                {
                    row["Send"] = _checked;

                }
            }
        }

        private void txtStagedFilter_TextChanged(object sender, EventArgs e)
        {
            //if (gvStaged.Rows != null && gvStaged.Rows.Count != 0)
            if (gvStaged.DataSource != null)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = gvStaged.DataSource;
                bs.Filter = string.Format("CONVERT(" + gvStaged.Columns["Code"].DataPropertyName + ", System.String) like '%" + txtStagedFilter.Text.Replace("'", "''") + "%'");
                gvStaged.DataSource = bs;
            }
            else
            {
                MessageBox.Show("Use the 'Get for Preview' button to populate the grid", "Grid Not Loaded");
            }
        }

        private void cbBillingDate_CheckedChanged(object sender, EventArgs e)
        {
            var _checked = cbBillingDate.CheckState == CheckState.Checked ? true : false;
            dpBilling.Enabled = _checked ? true : false;
        }

        private void txtSentFilter_TextChanged(object sender, EventArgs e)
        {
            //if (gvBillingSent.Rows != null && gvBillingSent.Rows.Count != 0)
            if (gvBillingSent.DataSource != null)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = gvBillingSent.DataSource;
                bs.Filter = string.Format("CONVERT(" + gvBillingSent.Columns[2].DataPropertyName + ", System.String) like '%" + txtSentFilter.Text.Replace("'", "''") + "%'");
                gvBillingSent.DataSource = bs;
            }
            else
            {
                MessageBox.Show("Use the 'Refresh' button to populate the grid", "Grid Not Loaded");
            }
        }

        private void txtFilterRpt_TextChanged(object sender, EventArgs e)
        {
            string search = rbSend.Checked ? "Send_Facility_Code" : "Data_Facility_Code";

            try
            {
                if (bsRpt.DataSource != null)
                {
                    var bd = (BindingSource)gvRpt.DataSource;
                    var dt = (DataTable)bd.DataSource;
                    dt.DefaultView.RowFilter = string.Format(search + " like '%{0}%'", txtFilterRpt.Text.Trim().Replace("'", "''"));
                    gvRpt.Refresh();
                }
                else
                {
                    MessageBox.Show("Use the 'Refresh' button to populate the grid", "Grid Not Loaded");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void rbEmailSend_CheckedChanged(object sender, EventArgs e)
        {
            if (rbEmailSend.Checked)
            {
                txtEmailMessage.Text = prop.SendEmailMessage;
                txtEmailSubject.Text = prop.SendEmailSubject;
            }
            else
            {
                txtEmailMessage.Text = prop.NotifyEmailMessage;
                txtEmailSubject.Text = prop.NotifyEmailSubject;
            }

        }

        private void ddStates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddAccStates.SelectedItem != null && DataSet)
            {
                try
                {
                    string state = ddAccStates.SelectedValue.ToString();
                    if (state != null)
                    {
                        //Utility.WriteActivity(state);
                    }
                }
                catch
                {
                    return;
                }
            }
        }

        private void txtAccFilter_TextChanged(object sender, EventArgs e)
        {
            //if(gvAccList.Rows != null && gvAccList.Rows.Count != 0)
            if (bsAcc.DataSource != null)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = gvAccList.DataSource;
                bs.Filter = string.Format("CONVERT(" + gvAccList.Columns["AccountName"].DataPropertyName + ", System.String) like '%" + txtAccFilter.Text.Replace("'", "''") + "%'");
                gvAccList.DataSource = bs;
            }
            else
            {
                MessageBox.Show("Use the 'Refresh/Load Accounts' button to populate the grid","Grid Not Loaded");
            }

        }

        private void ddMC_Account_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id = "";

            lbMC_AccName.Text = "";
            id = ddMC_Account.SelectedValue.ToString().Trim();

            if (!String.IsNullOrEmpty(id))
            {
                Utility.WriteActivity(id);
                string sql = @"SELECT 
                            AccountName, GroupCode
                            FROM ACCOUNT_LIST
                            WHERE ID = @dcode"; 
                using (SqlConnection conn = new SqlConnection(CONN_RX))
                {
                    SqlCommand command = new SqlCommand(sql, conn);
                    command.Parameters.AddWithValue("@dcode",id);

                    try
                    {
                        conn.Open();
                        SqlDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            lbMC_AccCode.Text = reader["GroupCode"].ToString();
                            lbMC_AccName.Text = reader["AccountName"].ToString();
                        }
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                    }

                }
            }
        }

        private void txtMC_Filter_TextChanged(object sender, EventArgs e)
        {
            //if (gvMC.Rows != null && gvMC.Rows.Count != 0)
            if (bsMC.DataSource != null)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = gvMC.DataSource;
                bs.Filter = string.Format("CONVERT(" + gvMC.Columns["ACCT"].DataPropertyName + ", System.String) like '%" + txtMC_Filter.Text.Replace("'", "''") + "%'");
                gvMC.DataSource = bs;
            }
            else
            {
                MessageBox.Show("Use the 'Refresh' button to populate the grid", "Grid Not Loaded");
            }
        }

        private void ddBilling_Codes_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (gvCodes.Rows.Count > 0)
            //{
            //    BindingSource bs = new BindingSource();
            //    bs.DataSource = gvCodes.DataSource;
            //    bs.Filter = string.Format("CONVERT(" + gvCodes.Columns["CATEGORY"].DataPropertyName + ", System.String) like '%" + ddBilling_Codes.SelectedValue.ToString().Replace("'", "''") + "%'");
            //    gvCodes.DataSource = bs;
            //    gvCodes.Refresh();
            //}
        }

        private void txtFilterCode_TextChanged(object sender, EventArgs e)
        {
            if (bsCodes.DataSource != null)
            {
                BindingSource bs = new BindingSource();
                bs.DataSource = gvCodes.DataSource;
                bs.Filter = string.Format("CONVERT(" + gvCodes.Columns["DESCRIPTION"].DataPropertyName + ", System.String) like '%" + txtFilterCode.Text.Replace("'", "''") + "%'");
                gvCodes.DataSource = bs;
                gvCodes.Refresh();
            }
            else
            {
                MessageBox.Show("Use the 'Refresh' button to populate the grid", "Grid Not Loaded");
            }
        }


        #endregion  --END Change Events

        #region Printing
        private void pdInvoice_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            var mFont = new Font("Arial", 12, FontStyle.Regular);
            var mFBold = new Font("Arial", 12, FontStyle.Bold);
            var bBrush = Brushes.Black;
            var logo = Image.FromFile(Application.StartupPath + @"\images\ihs-pharmacy-logo.png");
            var code = lbMC_AccCode.Text;
            var name = lbMC_AccName.Text;
            var date = dpMC_Date.Value.ToString("MM/dd/yyyy");
            var qty = txtMC_Qty.Text;
            if (qty.Length > 3 && qty.Substring(qty.Length - 3) == ".00")
            {
                qty = qty.Replace(".00", "");
            }
            var price = txtMC_Price.Text;
            var desc = txtMC_Desc.Text;
            e.Graphics.DrawString("IHS PHARMACY", mFont, bBrush, new Point(25, 60));
            e.Graphics.DrawString("504 MCCURDY AVE S., STE 7", mFont, bBrush, new Point(25, 84));
            e.Graphics.DrawString("RAINSVILLE, AL 35986", mFont, bBrush, new Point(25, 108));
            e.Graphics.DrawString("256-638-1060", mFont, bBrush, new Point(25, 132));

            e.Graphics.DrawImage(logo, new Point(560, 60));

            e.Graphics.DrawString(code, mFBold, bBrush, new Point(80, 180));
            e.Graphics.DrawString(name, mFBold, bBrush, new Point(160, 180));
            e.Graphics.DrawString(date, mFont, bBrush, new Point(600, 180));

            e.Graphics.DrawString(name, mFont, bBrush, new Point(25, 240));
            e.Graphics.DrawString(desc, mFont, bBrush, new Point(340, 240));

            e.Graphics.DrawString("Qty", mFBold, bBrush, new Point(25, 280));
            e.Graphics.DrawString("Price", mFBold, bBrush, new Point(160, 280));
            e.Graphics.DrawString(qty, mFont, bBrush, new Point(25, 310));
            e.Graphics.DrawString(price, mFont, bBrush, new Point(160, 310));

        }

        private void btnPrintInv_Click(object sender, EventArgs e)
        {
            decimal q, p;

            if (ddMC_Account.SelectedValue.ToString().Trim() == "")
            {
                MessageBox.Show("An Account ID is required", "Missing Account");
                return;
            }

            if (txtMC_Desc.Text.Trim() == "" || ddMC_Category.SelectedValue.ToString().Trim() == "")
            {
                MessageBox.Show("A Descrition and Category are required", "Missing Item");
                return;
            }

            if (decimal.TryParse(txtMC_Qty.Text, out q))
            {
                //valid
            }
            else
            {
                MessageBox.Show("Please enter a valid number for the Quantity", "Invalid Number");
                return;
            }

            if (decimal.TryParse(txtMC_Price.Text, out p))
            {
                //valid
            }
            else
            {
                MessageBox.Show("Please enter a valid number for the Price", "Invalid Number");
                return;
            }

            bool insert = false;

            if (btnMC_Update.Text == "Save")
                insert = true;

            ppDialog.Document = pdInvoice;
            ppDialog.ShowDialog();

            UpdateManualCharges(insert, q, p);
        }


        #endregion  --END Printing


    }
}
