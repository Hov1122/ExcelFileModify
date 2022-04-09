using System;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ExcelFileModify
{
    internal static class Program
    {
        
        // public static Form1 form = new Form1(); 
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.ApplicationExit += new EventHandler(Application_ApplicationExit);
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);
            var form = new Form1();
            clearTmpFolder();
            Application.Run(form);
        }

        private static void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {
            KillProcessAndChildren(Process.GetCurrentProcess().Id);
        }

        private static void Application_ApplicationExit(object sender, EventArgs e)
        {
            KillProcessAndChildren(Process.GetCurrentProcess().Id);
        }

        private static void clearTmpFolder()
        {
            //string dataPath = Environment.CurrentDirectory;
            string dataPath = ApplicationDeployment.CurrentDeployment.DataDirectory;
            System.IO.DirectoryInfo di = new DirectoryInfo(dataPath + @"\tmpFiles");

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        private static void KillProcessAndChildren(int pid)
        {
            clearTmpFolder();

            foreach (Process process in EditExcelFile.childrenProcess)
            {
                if (!process.HasExited)
                    process.Kill();
            }
            Process proc = Process.GetProcessById(pid);
            proc.Kill();
        }
    }   
}
