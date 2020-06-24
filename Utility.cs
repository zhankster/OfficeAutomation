using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CrystalReportsNinja;

namespace OfficeAutomation
{
    public class Utility
    {
        Main frm = new Main();
        public static void WriteActivity(string msg)
        {
            var dt = DateTime.Now.ToString("yy-MM-dd HH:mm:ss");
            string[] arr = new string[3];
            ListViewItem itm;
            //add items to ListView
            arr[0] = dt;
            arr[1] = msg;
            itm = new ListViewItem(arr);

            //frm.lvProcessing.Items.Add(itm);
            //frm.WriteActivity(msg);
            Program.MainForm.lvProcessing.Items.Add(itm);
            LogWriter lw = new LogWriter("log.txt");
            lw.Write(msg);
            //Program.MainForm.WriteActivity(msg);
        }
    }
}
