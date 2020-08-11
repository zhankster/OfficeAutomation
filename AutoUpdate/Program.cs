using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.IO;
using System.Threading;


namespace AutoUpdate
{
    class Program
    {
        public static void Copy(string sourceDirectory, string targetDirectory)
        {
            DirectoryInfo diSource = new DirectoryInfo(sourceDirectory);
            DirectoryInfo diTarget = new DirectoryInfo(targetDirectory);

            CopyAll(diSource, diTarget);
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target)
        {
            Directory.CreateDirectory(target.FullName);

            // Copy each file into the new directory.
            foreach (FileInfo fi in source.GetFiles())
            {
                Console.WriteLine(@"Copying {0}\{1}", target.FullName, fi.Name);
                fi.CopyTo(Path.Combine(target.FullName, fi.Name), true);
            }

            // Copy each subdirectory using recursion.
            foreach (DirectoryInfo diSourceSubDir in source.GetDirectories())
            {
                DirectoryInfo nextTargetSubDir =
                    target.CreateSubdirectory(diSourceSubDir.Name);
                CopyAll(diSourceSubDir, nextTargetSubDir);
            }

        }
        static void Main(string[] args)
        {
            string update_folder = @"\\192.168.50.202\dev\update";
            if (args == null)
            {
                Console.WriteLine("args is null"); // Check for null array
                Console.ReadLine();
            }
            else
            {
                update_folder = args[0];
            }


            var processes = Process.GetProcessesByName("OfficeAutomation");
            string app_folder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\";
            //app_folder = @"C:\Users\Owner\source\repos\OfficeAutomation\bin\Release" + @"\";

            try
            {
                if(processes.Length < 1)
                {
                    Console.WriteLine("No instances of OfficeAutomation running.");
                }
                else
                {
                    foreach (Process proc in processes)
                    {
                        proc.CloseMainWindow();
                        proc.WaitForExit();
                    }
                    Console.WriteLine("OfficeAutomation closed");
                }
                Thread.Sleep(2000);
                Copy(update_folder, app_folder);
                Thread.Sleep(2000);
                Process.Start(app_folder + "OfficeAutomation.exe");
                //Console.WriteLine(app_folder + "OfficeAutomation.exe");

            }
            catch (System.NullReferenceException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
