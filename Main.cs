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
            cbFacNotify.SelectedIndex = cbFacNotify.FindString("");
            dpBilling.Value = DateTime.Now.AddMonths(-1);
            tabControl1.TabPages.Remove(tpEmail);
            LoadFacilities();
        }

        #region Global Vars
        OfficeAutomation.Properties.Settings prop = OfficeAutomation.Properties.Settings.Default;
        private BindingSource bsFac = new BindingSource();
        private SqlDataAdapter daFac = new SqlDataAdapter();
        private BindingSource bsBill = new BindingSource();
        private SqlDataAdapter daBill = new SqlDataAdapter();
        DataTable dtBilling;
        private BindingSource bsBillSent = new BindingSource();
        private SqlDataAdapter daBillSent = new SqlDataAdapter();
        DataTable dtBillingSent;
        static string CONN_CIPS = "";
        static string CONN_RX = "";
        static string[] Scopes = { CalendarService.Scope.CalendarReadonly };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        static string outputFolder = "";
        static string outputFolderNew = "";
        string cropImgIn = "";
        string cropImgOut = "";

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
        #endregion

        #region Database Functions
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
                string selectCommand = "SELECT A.DCODE as [Group_Code], F.DNAME as [Facility_Name], A.EMAIL as Email ,A.FAX1 as Fax, A.PHONE1 as Phone, A.NOTIFY_TYPE as Notify_Type, A.USER1 FROM FAC_ALT A LEFT JOIN FAC F ON A.DCODE = F.DCODE  ORDER BY A.DCODE";
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

        public void SaveFaclity()
        {
            if (txtGroupCode.Text.Trim() == "")
            {
                MessageBox.Show("You must select a Group Code", "No Group Code", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (txtEmailAddresses.Text.Trim() == "")
            {
                MessageBox.Show("At least one Email Address must be added", "No Email Address", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!ValidateEmail(txtEmailAddresses.Text.Trim()))
            {
                MessageBox.Show("At least one Email Address is not valid", "Invalid Email Address", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var conn = new SqlConnection(CONN_RX);
            conn.Open();
            var sql = "";
            if (btnAddNew.Text == "Add New")
            {
                sql = "UPDATE [FAC_ALT]";
                sql += " SET [NOTIFY_TYPE] = '" + cbFacNotify.SelectedItem.ToString() + "'";
                sql += " ,[EMAIL] ='" + txtEmailAddresses.Text.Replace(" ", "").Trim() + "'";
                sql += " ,[FAX1] = '" + txtFacFax.Text.Trim() + "'";
                sql += " ,[PHONE1] = '" + txtFacPhone.Text.Trim() + "'";
                sql += " ,[USER1] = '" + txtFacUser.Text.Trim() + "'";
                sql += " WHERE [DCODE] = '" + txtGroupCode.Text.Trim() + "'";
            }
            else
            {
                sql = @"INSERT INTO [FAC_ALT]
                        ([DCODE]
                        ,[NOTIFY_TYPE]
                        ,[EMAIL]
                        ,[FAX1]
                        ,[PHONE1]
                        ,[USER1])";
                sql += " VALUES ('" + txtGroupCode.Text.Trim() + "','" + cbFacNotify.SelectedItem.ToString() + "','";
                sql += txtEmailAddresses.Text.Replace(" ", "").Trim() + "','" + txtFacFax.Text.Trim() + "','";
                sql += txtFacPhone.Text.Trim() + "','" + txtFacUser.Text.Trim() + "')";
            }
            //txtInfo.Text = sql;
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
                LoadFacilities();
            }
        }

        public void InsertFAC_TRANS(string trans_type, string fac_code, string documents, string email_sent, string notes, string created_by )
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
                sql += documents  + "','" + email_sent + "','" + notes + "','";
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
                    ,F.DNAME as Facility
                    ,ISNULL(
                    STUFF((SELECT ';' + C.DCODE
                        FROM
                        CIPS.dbo.FAC FS 
                        LEFT JOIN CIPS.dbo.FAC_CHG G 
                        ON FS.ID = G.FAC_ID
                        LEFT JOIN CIPS.dbo.CHG C
                        ON G.CHG_ID = C.ID
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
            sql += " AND NOTES = '" + notes + "') O ORDER BY F.DNAME";

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

        #endregion

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

                outputFolder = txtProcessFolder.Text;
                outputFolderNew = txtRenamedFolder.Text;
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
            //Utility.WriteActivity("Reading: " + pth);
            //facilities.Clear();

            ////Create COM Objects. Create a COM object for everything that is referenced
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@pth);
            //Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            //Excel.Range xlRange = xlWorksheet.UsedRange;

            //int rowCount = xlRange.Rows.Count;
            //int colCount = prop.FAC_COLUMN; //xlRange.Columns.Count;

            ////iterate over the rows and columns and print to the console as it appears in the file
            ////excel is not zero based!!
            //var strVal = "";

            //for (int i = 1; i <= rowCount; i++)
            //{
            //    //for (int j = 1; j <= colCount; j++)
            //    if (colCount == prop.FAC_COLUMN && i > 1)
            //    {
            //        //new line
            //        //if (j == 1)
            //        //{
            //        //    //Console.Write("\r\n");
            //        //    //Utility.WriteActivity("");
            //        //}

            //        try
            //        {
            //            if (xlRange.Cells[i, colCount] != null && xlRange.Cells[i, colCount].Value2 != null)
            //            {
            //                var fac = new Fac();
            //                //Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
            //                strVal = xlRange.Cells[i, colCount].Value2.ToString();
            //                if (strVal.Trim() != "")
            //                {
            //                    var h = strVal.Split('-')[0]; //Hyphen split
            //                    var s = strVal.Split(' ')[0];//Space split
            //                    if (h.Length < 4)
            //                    {
            //                        fac = GetFacility(h.Trim());
            //                        Utility.WriteActivity(fac.name + ": imported");
            //                    }
            //                    else if (s.Length < 5)
            //                    {
            //                        fac = GetFacility(s.Trim());
            //                        Utility.WriteActivity(fac.name + ": imported");
            //                    }
            //                    else
            //                    {
            //                        fac.code = strVal;
            //                        fac.name = "none";
            //                        fac.email = "none";
            //                        fac.fax = "none";
            //                        fac.phone = "none";
            //                        fac.notify_type = "none";
            //                        Utility.WriteActivity(fac.name + ": imported");
            //                    }
            //                    var valid = fac.name != "none" ? true : false;
            //                    var notify_type = strVal.Contains("(e") && valid ? "Email" : "Fax";
            //                    notify_type = fac.notify_type != "" ? fac.notify_type : notify_type;
            //                    Facility facility = new Facility(fac.code, fac.name, fac.phone, fac.fax, fac.email, notify_type, valid);
            //                    facilities.Add(facility);
            //                }
            //            }


            //        }
            //        catch (Exception ex)
            //        {
            //            MessageBox.Show(ex.Message);
            //        }

            //    }
            //}

            //DataTable dt = new DataTable();
            //dt.Columns.Add("Code", typeof(string));
            //dt.Columns.Add("Facility Name", typeof(string));
            //dt.Columns.Add("Email Addresses", typeof(string));
            //dt.Columns.Add("Fax", typeof(string));
            //dt.Columns.Add("Phone", typeof(string));
            //dt.Columns.Add("Notify Type", typeof(string));
            //dt.Columns.Add("Valid", typeof(string));
            //foreach (var fac in facilities)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr["Code"] = fac.code;
            //    dr["Facility Name"] = fac.name;
            //    dr["Email Addresses"] = fac.email;
            //    dr["Fax"] = fac.fax;
            //    dr["Phone"] = fac.phone;
            //    dr["Notify Type"] = fac.notify_type;
            //    dr["Valid"] = fac.valid_code.ToString();
            //    dt.Rows.Add(dr);
            //    if (!fac.valid_code)
            //    {
            //        txtInfo.Text += fac.code + "\r\n";
            //    }
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
            txtFacFax.Text = "";
            txtFacPhone.Text = "";
            txtFacUser.Text = "";
            txtFacilityName.Text = "";
            txtEmailAddresses.Text = "";
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

                if (attachments.Length > 0)
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
                //start.Arguments = "r \"" + Application.StartupPath + "\\scripts\\ocr.py\" " + pth ;
                start.Arguments = Application.StartupPath + @"\scripts\ocr.py " + pth;
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
                    accounts = new string[] { " "};
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

        public string GetBillingDocuments(string folder,string facility , string[]  accounts)
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

                acct =  accounts.Where(s => file.Name.Contains(s)).FirstOrDefault();

                if (acct != null && acct != facility.Trim())
                {
                    docs += doc + ";";
                }
                else if (identity.Trim() == facility.Trim())
                {
                    //Utility.WriteActivity("Report: " + identity + ": " + facility.Trim() + ": " + file.Name);
                    docs += doc + ";";
                }
            }
            return docs;
        }

        public bool SendBillingDocs(string code, string docs, string email)
        {
            string user = Environment.UserName;  //System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string notes = dpBilling.Value.ToString("yyyy-MM");

            //Utility.WriteActivity("Code: " + user + code + " Email: " + email + " Docs: " + docs);
            bool sent;
            string[] documents = docs.Trim().Split(';');
            for(int i = 0; i<documents.Length; i++)
            {
                if (documents[i].Trim() != "")
                    documents[i] = txtBillingExports.Text + documents[i];
            }

            string[] addresses = email.Trim().Split(';');
            for (int i = 0; i<addresses.Length; i++)
            {
                if (addresses[i].Trim() != "")
                {
                    sent = SendEmail("Your Billing Report from IHS Pharmacy is attached", "Your Billing Report is attached ",
                        addresses[i].Trim(), "operations@ihspharmacy.com", "IHS Pharmacy", documents);
                    if (sent)
                    {
                        InsertFAC_TRANS("BILLING_EMAIL", code, docs, addresses[i].Trim(), notes, user);

                    }
                }
            }

            return true;

        }

        #endregion Utility Functions

        #region Click Events
        private void btnOpen_Click(object sender, EventArgs e)
        {
            string typ = cbImportType.GetItemText(cbImportType.SelectedItem);
            //if (typ == "File")
            //{
            //    OpenFileDialog fd = new OpenFileDialog();
            //    fd.Filter = "Excel Files | *.xlsx; *.xls";
            //    Utility.WriteActivity("Open File Dialog");

            //    if (fd.ShowDialog() == DialogResult.OK)
            //    {
            //        readExcelFile(fd.FileName);
            //    }
            //}
            //else 
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

        private void btnFileExport_Click(object sender, EventArgs e)
        {
            string[] parms = { "-A", "Facility:DJ", "-A", "DateAfter:05-01-2020" };
            ExportReport(txtNotifyReport.Text, txtNotifyExports.Text, "pdf", parms, txtDSN_CIPS.Text);
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
                    txtNotifyReport.Text = folder;
                    WriteConfig("NotifyReport", folder);
                    Utility.WriteActivity("Notify Report updated");
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
                        txtNotifyExports.Text = folder;
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
            string[] attachments = { @"C:\Files\renamed\1.pdf" };
            var sent = SendEmail("Your ARX Report from IHS Pharmacy is attached", "Your ARX Report is attached ", "hank@dekalbal.com", "operations@ihspharmacy.com", "IHS Pharmacy", attachments);

        }

        private void gvFac_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            int rowindex = gvFac.CurrentCell.RowIndex;

            try
            {
                txtGroupCode.Text = gvFac.Rows[rowindex].Cells[0].Value.ToString();
                txtFacilityName.Text = gvFac.Rows[rowindex].Cells[1].Value.ToString();
                txtEmailAddresses.Text = gvFac.Rows[rowindex].Cells[2].Value.ToString();
                txtFacFax.Text = gvFac.Rows[rowindex].Cells[3].Value.ToString();
                txtFacPhone.Text = gvFac.Rows[rowindex].Cells[4].Value.ToString();
                txtFacUser.Text = gvFac.Rows[rowindex].Cells[6].Value.ToString();
                cbFacNotify.SelectedIndex = cbFacNotify.FindString(gvFac.Rows[rowindex].Cells[5].Value.ToString());
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
                        //email = "dekalb.hda@gmail.com;hank@dekalbal.com;zrefugee@gmail.com";
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
                            gvNotiifications.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                            File.Delete(att_path);
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

        private void btnFacilityFax_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to send all the fax notiifications displayed?",
            "Send Fax Notification", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                return;
            }

            //string[] parms = { "-A", "Facility:DJ", "-A", "DateAfter:05-01-2020" };
            //FaxReports(txtNotifyReport.Text, txtNotifyExports.Text, "print", parms, txtDSN_CIPS.Text);
            FaxExportToFolder();
        }

        private void btnExportReports_Click(object sender, EventArgs e)
        {
            txtReportIssues.Text = "";
            var folder = prop.BillingRptFolder;
            //var report_date = dptFacExport.Value.ToString("MM-dd-yyyy");
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

        //Attachment Processing
        private void btnDownload_Click(object sender, EventArgs e)
        {
            var msg_ret = DownloadAttachments();
            Utility.WriteActivity(msg_ret);
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
            GetFacilityBIlling();
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
                    
            //txtPreview.Text = val;
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
  
            string email = gvStaged.Rows[row].Cells["Email"].Value.ToString();
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

            if (send == "True" && email.Trim().Length > 2 && docs.Trim().Length > 2)
            {
                SendBillingDocs(code, docs, email);
            }

        }

        private void btnSendSelected_Click(object sender, EventArgs e)
        {
            string email, docs, code, send;

            if (dtBilling != null)
            {
                foreach (DataGridViewRow row in gvStaged.Rows)
                {
                    email = row.Cells["Email"].Value.ToString();
                    docs = row.Cells["Documents"].Value.ToString();
                    code = row.Cells["Code"].Value.ToString();
                    send = row.Cells["Send"].Value.ToString();
                    
                    if (send == "True" && email.Trim().Length > 2 && docs.Trim().Length > 2)
                    {
                        SendBillingDocs(code, docs, email);
                    }
                } 
            }
        }

        private void btnRefreshSent_Click(object sender, EventArgs e)
        {
            GetSentBIlling();
        }

        private void gvBillingSent_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int cell = gvBillingSent.CurrentCell.ColumnIndex;
            int row = e.RowIndex;
            string val = "", facility = "";
            string[] vals;
            if (e.RowIndex == -1) return; //check if row index is not selected
            {
                if (cell.Equals(4) || cell.Equals(5) )
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

        #endregion

        #region Change Events
        private void txtFacFilter_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var bd = (BindingSource)gvFac.DataSource;
                var dt = (DataTable)bd.DataSource;
                dt.DefaultView.RowFilter = string.Format("Facility_Name like '%{0}%'", txtFacFilter.Text.Trim().Replace("'", "''"));
                gvFac.Refresh();
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
            BindingSource bs = new BindingSource();
            bs.DataSource = gvStaged.DataSource;
            bs.Filter = string.Format("CONVERT(" + gvStaged.Columns[3].DataPropertyName + ", System.String) like '%" + txtStagedFilter.Text.Replace("'", "''") + "%'");
            gvStaged.DataSource = bs;
        }

        private void cbBillingDate_CheckedChanged(object sender, EventArgs e)
        {
            var _checked = cbBillingDate.CheckState == CheckState.Checked ? true : false;
            dpBilling.Enabled = _checked ? true : false;
        }

        private void txtSentFilter_TextChanged(object sender, EventArgs e)
        {
            BindingSource bs = new BindingSource();
            bs.DataSource = gvBillingSent.DataSource;
            bs.Filter = string.Format("CONVERT(" + gvBillingSent.Columns[2].DataPropertyName + ", System.String) like '%" + txtSentFilter.Text.Replace("'", "''") + "%'");
            gvBillingSent.DataSource = bs;
        }

        //var checked = cbBillGrid.CheckState == CheckState.Checked ? true : false;
        #endregion

        private void btnEmailExports_Click(object sender, EventArgs e)
        {

        }
    }
}
