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
using System.Xml.Linq;
using Microsoft.WindowsAPICodePack.Dialogs;
using PdfiumViewer;

namespace CSB
{
    public partial class Settings : Form
    {
        
        public Settings()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {

            var xdoc = XDocument.Load(Globals.Config());

            var tgt = xdoc.Root.Descendants("Folder").FirstOrDefault();

            txtFolder.Text = tgt.Value;

            tgt = xdoc.Root.Descendants("ExportFolder").FirstOrDefault();

            txtExport.Text = tgt.Value;

            tgt = xdoc.Root.Descendants("TemplateModel").FirstOrDefault();

            txtTemplate.Text = tgt.Value;

        }

        private void btnOK_Click(object sender, EventArgs e)
        {

            var xdoc = XDocument.Load(Globals.Config());

            var tgt = xdoc.Root.Descendants("Folder").FirstOrDefault();

            tgt.Value = txtFolder.Text;
            
            var tgt2 = xdoc.Root.Descendants("ExportFolder").FirstOrDefault();

            tgt2.Value = txtExport.Text;

            var tgt3 = xdoc.Root.Descendants("TemplateModel").FirstOrDefault();

            tgt3.Value = txtTemplate.Text;

            xdoc.Save(Globals.Config());

            Close();
        }

        //public void PageViewer(string path)
        //{

        //    byte[] bytes = System.IO.File.ReadAllBytes(path);
        //    var stream = new MemoryStream(bytes);
        //    PdfDocument pdfDocument = PdfDocument.Load(stream);
        //    pdfViewer1.Document = pdfDocument;


        //   //var Doc = PdfDocument.Load(path);
        //   // pdfViewer1.Load(Doc);
        //}

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"T:\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtFolder.Text = dialog.FileName + @"\";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"T:\CSB_TeklaSetup\Model Templates";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                DirectoryInfo fi = new DirectoryInfo(dialog.FileName);
                txtTemplate.Text = fi.Name;
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = @"T:\";
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                txtExport.Text = dialog.FileName + @"\";
            }
        }
    }
}
