using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSB
{
    static class Program
    {
        
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]

        static void Main()
        {

            //if (!SingleInstance.Start("CSB_Project_Start"))
            //{
            //    MessageBox.Show("Application is already running.");
            //    return;
            //}

            //if (!SingleInstance.Start("TeklaStructures"))
            //{
            //    MessageBox.Show("Multiple TeklaStructures are running." + "\r\n" + "Fix and try again");
            //    return;
            //}

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
           
            SingleInstance.Stop();

        }
        public static class SingleInstance
        {
            public static bool Start(string applicationIdentifier)
            {
                bool isSingleInstance = false;

                Process[] localByName = Process.GetProcessesByName(applicationIdentifier);

                if (localByName.Length > 1)
                {
                    isSingleInstance = false;
                }
                else
                {
                    isSingleInstance = true;
                }

                return isSingleInstance;
            }
            public static void Stop()
            {

            }
        }

    }
}
