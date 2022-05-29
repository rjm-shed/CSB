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
            bool instanceCountOne = false;

            using (Mutex mtexCSB = new Mutex(true, "CSB_Project_Start", out instanceCountOne))
            {
                if (instanceCountOne)
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Form1());
                    mtexCSB.ReleaseMutex();
                }
                else
                {
                    MessageBox.Show("An application instance is already running");
                }
            }

        }

    }
}
