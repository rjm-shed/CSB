using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSB
{
    public partial class ManageFiles : Form
    {

        Helper myHelper = new Helper();

        public ManageFiles()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

       
        private void CleanCompleted()
        {
            myHelper.LogFile("***********************");
            myHelper.LogFile("Completed Project Moved Files");

            string[] dirs = System.IO.Directory.GetDirectories(myHelper.ProjectFolder());

            foreach (string item2 in dirs)
            {
                System.IO.FileInfo f = new FileInfo(item2);

                string[] files = Directory.GetFiles(myHelper.ExportFolder() + @"Tekla\In\");

                foreach (string item in files)
                {
                    FileInfo g = new FileInfo(item);

                    if (g.Name.Contains(f.Name))
                    {
                        File.Move(g.FullName, myHelper.ExportFolder() + @"Tekla\Complete\" + g.Name);
                        myHelper.LogFile("Input file moved " + g.Name);
                    }

                }

            }

            myHelper.LogFile("***********************");

        }

        private void btnDuplicate_Click(object sender, EventArgs e)
        {
            myHelper.LogFile("***********************");
            myHelper.LogFile("Duplicate Files");

            string[] files = Directory.GetFiles(myHelper.ExportFolder() + @"Tekla\In\");

            foreach (string item in files)
            {
                FileInfo f = new FileInfo(item);

                string[] fName = f.Name.Split('_');
                string fProject = f.Name.Substring(0, 5);

                List<FileInfo> duplicates = new List<FileInfo>();

                foreach (string item2 in files)
                {
                    FileInfo g = new FileInfo(item2);

                    string[] gName = g.Name.Split('_');

                    if (gName[0] == fName[0] && f.FullName != g.FullName) //Check for same version number
                    {
                        if (File.Exists(f.FullName))
                        {
                            try
                            {
                                File.Move(f.FullName, myHelper.ExportFolder() + @"Tekla\CheckSame\" + f.Name);
                                myHelper.LogFile("Same input file moved " + f.Name);
                            }
                            catch
                            {
                                myHelper.LogFile("Same input file failed to move " + f.Name);
                            }
                        }
                        if (File.Exists(g.FullName))
                        {
                            try
                            {
                                File.Move(g.FullName, myHelper.ExportFolder() + @"Tekla\CheckSame\" + g.Name);
                                myHelper.LogFile("Same input file moved " + g.Name);
                            }
                            catch
                            {
                                myHelper.LogFile("Same input file failed to move " + g.Name);
                            }

                        }
                    }

                    else if (g.Name.Contains(fProject))
                    {
                        duplicates.Add(g);
                    }

                }

                if (duplicates.Count > 1)
                {

                    List<string> dupli = new List<string>();

                    for (int i = 0; i < duplicates.Count; ++i)
                    {
                        FileInfo x = duplicates.ElementAt(i);
                        dupli.Add(x.Name);
                    }
                    dupli = dupli.OrderBy(q => q).ToList();

                    for (int i = 0; i < dupli.Count - 1; ++i)
                    {
                        string rFullName = myHelper.ExportFolder() + @"Tekla\In\" + dupli.ElementAt(i);

                        if (File.Exists(rFullName))
                        {
                            try
                            {
                                File.Move(rFullName, myHelper.ExportFolder() + @"Tekla\Duplicate\" + dupli.ElementAt(i));
                                myHelper.LogFile("Input file moved " + dupli.ElementAt(i));
                            }
                            catch
                            {
                                myHelper.LogFile("Input file failed to move " + dupli.ElementAt(i));
                            }

                        }
                    }
                }

            }

            CleanCompleted();

        }

        //myHelper.LogFile("***********************");

    }

}