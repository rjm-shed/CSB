using PdfiumViewer;
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
    public partial class SaveHelp : Form
    {
        public SaveHelp()
        {
            InitializeComponent();
        }

        private void SaveHelp_Load(object sender, EventArgs e)
        {
            PageViewer(@"T:\CSB_Program_Files\Documentation\Share_Help.pdf");
        }

        public void PageViewer(string path)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(path);
            var stream = new MemoryStream(bytes);
            PdfDocument pdfDocument = PdfDocument.Load(stream);
            pdfViewer1.Document = pdfDocument;
        }

    }
}
