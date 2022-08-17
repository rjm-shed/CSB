using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Diagnostics;
using System.Linq;
using System.Xml;
using System.Threading.Tasks;
using STT = System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures;
using TSG = Tekla.Structures.Geometry3d;
using Tekla.Structures.Dialog.UIControls;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Plugins;
using Tekla.Structures.Model.Operations;
using System.Xml.Linq;
using Component = Tekla.Structures.Model.Component;
using System.Text.RegularExpressions;
using System.Collections;
using Microsoft.WindowsAPICodePack.Dialogs;
using Squirrel;

namespace CSB
{

    public partial class Form1 : Form
    {
        
        Helper myHelper = new Helper();

        salesLib ProjectSales = new salesLib();

        Model myModel = new Model();

        ColumnSize columnSize = new ColumnSize();

        slabCorners _slabCorners = new slabCorners();

        int _NoMullions = 0;

        public Form1()
        {
            InitializeComponent();
            LoadCbx(cbxRoof);
            LoadCbx(cbxWall);
            LoadCbx(cbxTrim);
            LoadCbx(cbxGutter);
            LoadCbx(cbxRoller);
            LoadCbx(cbxSlide);
            LoadCbx(cbxPA);
            LoadSkyCbx(cbxRoofSky);
            LoadSkyCbx(cbxWallSky);
            LoadCbx(cbxWhirly);
            LoadCbx(cbxWindow);
            LoadCbx(cbxMisc1);
            LoadCbx(cbxMisc2);
            LoadCbx(cbxMisc3);
            LoadLogo(cbxLogo);
            LoadSheetCbx(cbxRoofClad);
            LoadSheetCbx(cbxWallClad);
            LoadC_Z(cbxPurlin);
            LoadC(cbxFascia);
            LoadC_Z(cbxGirtSide);
            LoadC_Z(cbxGirtSideRight);
            LoadC_Z(cbxGirtEnd);
            LoadC_Z(cbxGirtEndBack);

            AddVersionNumber();

            CheckForUpdates();
        }

        private void AddVersionNumber()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);

            this.Text += $" v.{versionInfo.FileVersion}";
        }

        private async STT.Task CheckForUpdates()
        {
            //TODO: change to variable
            using (var manager = new UpdateManager(@"T:\CSB_Program_Files\Code_Files"))
            {
                await manager.UpdateApp();
            }
        }

        #region Load ComboBoxes

        private void LoadLogo(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a == "Logo")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        private void LoadCbx(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if ( a == "Colour")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        private void LoadSkyCbx(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a == "Sky")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        private void LoadC_Z(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a == "C" || a == "Z")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        private void LoadC(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a == "C")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        private void LoadSheetCbx(ComboBox temp)
        {
            temp.Items.Clear();

            var xdoc = XDocument.Load(myHelper.Setting() + "CSB_Project_Data.xml");

            foreach (var childElement in xdoc.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a == "ColorBond")
                {
                    temp.Items.Add(childElement.Value.ToString());
                }
            }
        }

        #endregion

        #region Entry Buttons

        private void button2_Click(object sender, EventArgs e)
        {
            Globals.checkError = 0;
            validateAll(e);
        }
        private void btnSales_Click(object sender, EventArgs e)
        {

            string xFolder = myHelper.ExportFolder() + @"Tekla\In\";

            if (Directory.Exists(xFolder))
            {
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(" not found - " + xFolder, "Folder", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }

            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.InitialDirectory = xFolder;
            dialog.IsFolderPicker = false;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                lblSales.Text = dialog.FileName;
            }
            else
            {
                return;
            }

            SetStandards();

            ProjectSales = new salesLib();

            bool tempCheck = myHelper.ReadSalesInput(lblSales.Text.Trim(), ProjectSales);

            if (tempCheck == false)
            {
                System.Windows.Forms.MessageBox.Show(" Problem reading data file ", "Sales CSV", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return;
            }
            
            txtNumber.Text = ProjectSales.ProjectNo;
            txtClient.Text = ProjectSales.ProjectName;
            txtLength.Text = ProjectSales.Length;
            txtWidth.Text = ProjectSales.Width;
            txtEave.Text = ProjectSales.Height;
            txtAddress.Text = ProjectSales.Suburb;

            txtLength.Text = ProjectSales.Length;
            txtWidth.Text = ProjectSales.Width;
            txtEave.Text = ProjectSales.Height;
            txtPitch.Text = ProjectSales.RoofPitch;
            txtBaySize.Text = ProjectSales.BayString;

            txtWallGirtSide.Text = ProjectSales.WallGirtSide;
            txtWallGirtSideRight.Text = ProjectSales.WallGirtSide;
            txtWallGirtEnd.Text = ProjectSales.WallGirtEnd;
            txtWallGirtEndBack.Text = ProjectSales.WallGirtEnd;
            txtPurlin.Text = ProjectSales.RoofPurlin;

            switch (ProjectSales.Industry.Trim())
            {
                case "Steel Build":
                    cbxLogo.Text= "CSB Steel Build";
                    break;
                case "Agricultural":
                    cbxLogo.Text = "CSB Agricultural";
                    break;
                case "Aviation":
                    cbxLogo.Text = "CSB Aviation";
                    break;
                case "Commercial":
                    cbxLogo.Text = "CSB Commercial";
                    break;
                case "Custom":
                    cbxLogo.Text = "CSB Custom";
                    break;
                case "Equinabuild":
                    cbxLogo.Text = "CSB Equinabuild";
                    break;
                case "Industrial":
                    cbxLogo.Text = "CSB Industrial";
                    break;
                case "Recreational":
                    cbxLogo.Text = "CSB Recreational";
                    break;
                default:
                    if(ProjectSales.Industry != null || ProjectSales.Industry != "")
                    {
                        //MessageBox.Show("Email Richard a copy - " + ProjectSales.ProjectNo + " - " + ProjectSales.Industry, "Project Logo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //MessageBox.Show("Remember to manually change", "Project Logo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    cbxLogo.Text = "CSB Steel Build";
                    break;
            }

            if (ProjectSales.RoofPurlin != null && ProjectSales.RoofPurlin != "" && ProjectSales.RoofPurlin.Contains("Z"))
            {
                txtFascia.Text = ProjectSales.RoofPurlin.Replace("Z","C");
            }
            else
            {
                txtFascia.Text = ProjectSales.RoofPurlin;
            }
               
            txtProjectDetails.Text = ProjectSales.ProjectDetails;

            if (ProjectSales.RoofPurlin != null && ProjectSales.RoofPurlin.Contains("Z"))
            {
                chkPurlinSingleSpan.Checked = true;
            }

            if (ProjectSales.WallGirtSide != null && ProjectSales.WallGirtSide.Contains("Z"))
            {
                chkGirtSingleSpan.Checked = true;
            }

            //******************************************************************

            if (ProjectSales.RoofColour != null && ProjectSales.RoofColour.Contains("Zincalume"))
            {
                cbxRoof.Text = "ZINC";
            }
            else if (ProjectSales.RoofColour != null && ProjectSales.RoofColour.Contains("Colorbond"))
            {
                cbxRoof.Text = "CBOND(TBC)";
            }

            if (ProjectSales.WallColour != null && ProjectSales.WallColour.Contains("Zincalume"))
            {
                cbxWall.Text = "ZINC";
            }
            else if (ProjectSales.WallColour != null && ProjectSales.WallColour.Contains("Colorbond"))
            {
                cbxWall.Text = "CBOND(TBC)";
            }

            if (ProjectSales.GutterColour != null && ProjectSales.GutterColour.Contains("Zincalume"))
            {
                cbxTrim.Text = "ZINC";
                cbxGutter.Text = "ZINC";
            }
            else if (ProjectSales.GutterColour != null && ProjectSales.GutterColour.Contains("Colorbond"))
            {
                cbxTrim.Text = "CBOND(TBC)";
                cbxGutter.Text = "CBOND(TBC)";
            }

            if (ProjectSales.ClearSheetRoof != null && ProjectSales.ClearSheetRoof.Contains("Opal"))
            {
                cbxRoofSky.Text = "OPAL";
            }
            else if (ProjectSales.ClearSheetRoof != null && ProjectSales.ClearSheetRoof.Contains("Clear"))
            {
                cbxRoofSky.Text = "CLEAR";
            }

            if (ProjectSales.ClearSheetWall != null && ProjectSales.ClearSheetWall.Contains("Opal"))
            {
                cbxWallSky.Text = "OPAL";
            }
            else if (ProjectSales.ClearSheetWall != null && ProjectSales.ClearSheetWall.Contains("Clear"))
            {
                cbxWallSky.Text = "CLEAR";
            }

            //******************************************************************

            if (ProjectSales.RoofMaterial != null && ProjectSales.RoofMaterial.Contains(".42 BMT") && ProjectSales.RoofMaterial.Contains("5-Rib"))
            {
                cbxRoofClad.Text = "0.47-TCT-5-RIB";
                chkRolltop.Checked = false;
            }
            else if (ProjectSales.RoofMaterial != null && ProjectSales.RoofMaterial.Contains(".42 BMT") && ProjectSales.RoofMaterial.Contains("Corry"))
            {
                cbxRoofClad.Text = "0.47-TCT-CORRY";
                chkRolltop.Checked = true;
            }

            if (ProjectSales.WallMaterial != null && ProjectSales.WallMaterial.Contains(".42 BMT") && ProjectSales.WallMaterial.Contains("5-Rib"))
            {
                cbxWallClad.Text = "0.47-TCT-5-RIB";
            }
            else if (ProjectSales.WallMaterial != null && ProjectSales.WallMaterial.Contains(".42 BMT") && ProjectSales.WallMaterial.Contains("Corry"))
            {
                cbxWallClad.Text = "0.47-TCT-CORRY";
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            createSlab(300, 60000, 40000);

            //ContourPlate _Slab = new ContourPlate();

            //_Slab.Name = "SLAB";
            //_Slab.Profile.ProfileString = "150";
            //_Slab.Material.MaterialString = "N40";
            //_Slab.Class = "1";

            //Contour ct = new Contour();

            //ContourPoint p1 = new ContourPoint();
            //p1.X = 0;
            //p1.Y = 0;
            //p1.Z = 0;
            //p1.Chamfer = new Chamfer();

            //ContourPoint p2 = new ContourPoint();
            //p2.X = 0;
            //p2.Y = 1000;
            //p2.Z = 0;
            //p2.Chamfer = new Chamfer();

            //ContourPoint p3 = new ContourPoint();
            //p3.X = 2000;
            //p3.Y = 1000;
            //p3.Z = 0;
            //p3.Chamfer = new Chamfer();

            //ContourPoint p4 = new ContourPoint();
            //p4.X = 2000;
            //p4.Y = 0;
            //p4.Z = 0;
            //p4.Chamfer = new Chamfer();

            //ct.AddContourPoint(p1);
            //ct.AddContourPoint(p2);
            //ct.AddContourPoint(p3);
            //ct.AddContourPoint(p4);

            //_Slab.Contour = ct;
            ////_Slab.Position.Depth = Position.DepthEnum.FRONT;

            //bool jj = _Slab.Insert();

            //MessageBox.Show(jj.ToString());

            //myModel.CommitChanges();

            //XmlDocument xmldoc = new XmlDocument();
            //xmldoc.Load(@"T:\CSB_Program_Files\Documentation\Settings\CSB_Project_Data_UB_UC.xml");

            //XmlNodeList nodeList = xmldoc.GetElementsByTagName("UB_UC");
            //foreach (XmlNode node in nodeList)
            //{
            //    var Tekla = node.Attributes["Tekla"].Value;
            //    var Aus = node.Attributes["Aus"].Value;
            //    var Depth = node.Attributes["Depth"].Value;
            //    var FlangeW = node.Attributes["FlangeW"].Value;
            //    var FlangeT = node.Attributes["FlangeT"].Value;
            //    var WebT = node.Attributes["WebT"].Value;
            //    break;
            //}

            //string d = "Z20015";

            //d = d.Replace("Z", "C");


            //string c =d.Substring(7);

            //string k = d.Substring(6);

            //string f = d.Substring(7);

            //string c = "250UB18";

            //string d = c.Substring(0, 3);

            //string f = c.Substring(5);

            //string g = c.Substring(3, 2);

            //string dd = g + d + "*" + f;


            //myHelper.NoteText = "yes this it it nownnnnnnnnnnnnn";
            //PrintPDF(@"C:\Development\Models\21604\attributes\");
            //try
            //{

            //    Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
            //    Origin.X = -6000;
            //    Origin.Y = -3000;
            //    Origin.Z = 0;

            //    Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
            //    FinishPoint.X = -6000;
            //    FinishPoint.Y = -9000;
            //    FinishPoint.Z = 0;

            //    PDF(Origin, FinishPoint, "standardnew"); //, myModel

            //    myHelper.LogFile("Add Project Notes");
            //}
            //catch (Exception g)
            //{
            //    myHelper.LogFile("1008 - " + g.Message);
            //}


            //myHelper.ProcessRunning("TeklaStructures");

            //Process[] localByName = Process.GetProcessesByName("CSB_Project_Start");

            //List<double> spacingList = new List<double>();

            //spacingList.Add(0);
            //spacingList.Add(6000);
            //spacingList.Add(6000);
            //spacingList.Add(9030);
            //spacingList.Add(6000);

            ////double max = spacingList.Max();

            ////double ss = 5 * ((max / 0.15) / 5.0);

            //int s = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

            //List<string> split = CalcSplitDoubleSpan(spacingList);

            //myHelper.LogFile("Model Saved");

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            myHelper.LogFile("Model process started");

            if (!myHelper.ProcessRunning("TeklaStructures"))
            {
                MessageBox.Show("Multiple TeklaStructures are running." + "\r\n" + "Fix and try again");
                return;
            }

            myModel = new Model();

            // Check that the model connection succeeded:
            if (myModel.GetConnectionStatus())
            {
                myHelper.LogFile("Model Connection OK");
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Tekla not running" + "\n" + "Start Tekla, open any project" + "\n" + "Retry making model", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            bool result = myHelper.IsNumeric(txtNumber.Text.Trim());

            if (result == false)
            {
                DialogResult xresult = System.Windows.Forms.MessageBox.Show("Is not numeric, is this correct?", "Project Number", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Information);

                if (xresult == DialogResult.No)
                {
                    return;
                }
                myHelper.LogFile("Project number OK");
            }

            Globals.checkError = 0;
            validateAll(e);

            if (Globals.checkError == 1)
            {
                myHelper.LogFile("Input error caught");
                Globals.checkError = 0;
                return;
            }

            ProjectLib Project = new ProjectLib();

            Project.Number = txtNumber.Text.Trim();
            Project.Client = txtClient.Text.Trim();
            Project.Length = txtLength.Text.Trim();
            Project.Width = txtWidth.Text.Trim();
            Project.Eave = txtEave.Text.Trim();
            Project.Address = txtAddress.Text.Trim();
            Project.Description = txtDescription.Text.Trim();

            Project.TemplateModel = myHelper.TemplateModel().Trim(); // @"Model Template 2021";
            Project.Folder = myHelper.ProjectFolder(); // @"C:\Development\Models\";

            string xtemp = Project.Folder + Project.ModelName;  //123458-WHO CARES

            myHelper.LogFile("Read Project data");

            if (Directory.Exists(xtemp))
            {
                myHelper.LogFile("Project Exists");

                System.Windows.Forms.MessageBox.Show("", "Project already exists", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                return;

                //DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Do you want to over-write it?" + "\r\n" + "ONLY DO THIS IF IT HAS NOT BEEN SHARED", "Project already exists", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Information);

                //if (dialogResult == DialogResult.Yes)
                //{
                //    try
                //    {
                //        Directory.Delete(xtemp, true);
                //        myHelper.LogFile("Project deleted " + Project.ModelName);
                //    }
                //    catch
                //    {
                //        System.Windows.Forms.MessageBox.Show("Did not Delete", "Project", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                //        myHelper.LogFile("Project not deleted " + Project.ModelName);
                //        return;
                //    }
                //}
                //else if (dialogResult == DialogResult.No)
                //{
                //    return;
                //}

            }

            string MasterFiles = @"T:\CSB_Program_Files\Documentation\Masters\";
            string TemplateAttributes = myHelper.TeklaFolder() +  @"Model Templates\" + myHelper.TemplateModel() + @"\attributes\";

            if (Directory.Exists(MasterFiles))
            {
                myHelper.LogFile("Master folder location OK");
            }
            else
            {
                myHelper.LogFile("Master folder location not found");
                System.Windows.Forms.MessageBox.Show("Location not found", "Master Files", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            if (Directory.Exists(TemplateAttributes))
            {
                myHelper.LogFile("Template attributes folder location OK");
            }
            else
            {
                myHelper.LogFile("Template attributes folder location not found");
                System.Windows.Forms.MessageBox.Show("Attributes not found", "Template", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            //TODO: Copy Template Files
            // File.Copy(filePath, newPath, true);
            try
            {
                //Project Master End Wall Cladding.CSB_EndWall_Cladding.MainForm.xml
                //Project Master Roof Cladding.CSB_Roof_Cladding.MainForm.xml
                //Project Master Side Wall Cladding.CSB_SideWall_Cladding.MainForm.xml
                //CSB_Project_Setup.CSB_Gable_Shed.MainForm.xml

                File.Copy(MasterFiles + "Project Master Gable.CSB_Gable_Shed.MainForm.xml", TemplateAttributes + "CSB_Project_Setup.CSB_Gable_Shed.MainForm.xml", true);

                File.Copy(MasterFiles + "Project Master Roof Cladding.CSB_Roof_Cladding.MainForm.xml", TemplateAttributes + "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master Roof Cladding.CSB_Roof_Cladding.MainForm.xml", TemplateAttributes + "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master End Wall Cladding.CSB_EndWall_Cladding.MainForm.xml", TemplateAttributes + "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master End Wall Cladding.CSB_EndWall_Cladding.MainForm.xml", TemplateAttributes + "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master End Wall Cladding.CSB_EndWall_Cladding.MainForm.xml", TemplateAttributes + "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master End Wall Cladding.CSB_EndWall_Cladding.MainForm.xml", TemplateAttributes + "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master Side Wall Cladding.CSB_SideWall_Cladding.MainForm.xml", TemplateAttributes + "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml", true);
                File.Copy(MasterFiles + "Project Master Side Wall Cladding.CSB_SideWall_Cladding.MainForm.xml", TemplateAttributes + "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml", true);
               
                myHelper.LogFile("Template attributes copied");
            }
            catch (Exception f)
            {
                myHelper.LogFile("1100 - " + f.Message);
                System.Windows.Forms.MessageBox.Show("Unable to copy", "Master Files", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }


            string xResult = ProcessModel(Project);
        }

        private bool checkConnection()
        {
            bool _yes = false;
            if (myModel.GetConnectionStatus())
            {
                myHelper.LogFile("Model Connection OK");
                _yes = true;
            }
            else
            {
                myHelper.LogFile("Model Connection Failed");
                _yes = false;
            }

            return _yes;
        }

        private void btnStd_Click(object sender, EventArgs e)
        {
            SetStandards();

            lblSales.Text = "";
        }

        private void SetStandards()
        {
            txtNumber.Text = "";
            txtClient.Text = "";
            txtBuilder.Text = "CSB";
            txtDesigner.Text = "";
            txtAddress.Text = "";
            txtDescription.Text = "";

            txtLength.Text = "18";
            txtWidth.Text = "12";
            txtEave.Text = "6";
            txtBaySize.Text = " 3*6000 ";
            txtSlab.Text = "0";
            txtPitch.Text = "7.5";

            txtNote.Text = "";

            SetColour("");

            txtImportance.Text = "";
            txtCategory.Text = "";
            txtRegion.Text = "";
            txtEngineer.Text = "TBC";
            txtComputation.Text = "";
            txtCompPages.Text = "";
            txtDate.Text = "";
            cbxLogo.Text = "CSB Steel Build";
            txtEmbedment.Text = "100";
            txtCapacity.Text = "100";
            cbxRoofClad.Text = "0.47 TCT 5-RIB"; //0.47-TCT-CORRY
            cbxWallClad.Text = "0.47 TCT 5-RIB";
            txtPurlinCoat.Text = "Z350";
            cbxLift.Text = "Yes";

            radN.Checked = true;

            btnRight.BackColor = System.Drawing.Color.Red;
            btnLeft.BackColor = System.Drawing.Color.Red;
            btnFront.BackColor = System.Drawing.Color.Red;
            btnRear.BackColor = System.Drawing.Color.Red;

            chkPurlinSingleSpan.Checked = false;
            chkGirtSingleSpan.Checked = false;
            chkRolltop.Checked = false;
            checkBox1.Checked = false; // roofonly

            txtPurlin.Text = "";
            txtFascia.Text = "";
            txtWallGirtSide.Text = "";
            txtWallGirtSideRight.Text = "";
            txtWallGirtEnd.Text = "";
            txtWallGirtEndBack.Text = "";
            txtProjectDetails.Text = "";

            radModelYes.Checked = true;
        }

        private void btnCBOND_Click(object sender, EventArgs e)
        {
            SetColour("CBOND(TBC)");
        }

        private void btnZINC_Click(object sender, EventArgs e)
        {
            SetColour("ZINC");
        }
        private void btnLaker_Click(object sender, EventArgs e)
        {
            txtEngineer.Text = "LAKER GROUP";
        }
        private void SetColour(string temp)
        {
            cbxRoof.Text = temp;
            cbxWall.Text = temp;
            cbxTrim.Text = temp;
            cbxGutter.Text = temp;
            cbxRoller.Text = temp;
            cbxSlide.Text = temp;
            cbxPA.Text = temp;
            cbxWhirly.Text = temp;
            cbxWindow.Text = temp;
            cbxMisc1.Text = temp;
            txtMisc1Desc.Text = "";
            cbxMisc2.Text = temp;
            txtMisc2Desc.Text = "";
            cbxMisc3.Text = temp;
            txtMisc3Desc.Text = "";
            txtColourComment.Text = "";
        }

        #endregion

        #region PDF print

        void PrintPDF(string mDirectory)
        {
            // Set the output dir and file name
            //string directory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string file = "CSB_Project_Setup.pdf";

            PrintDocument pDoc = new PrintDocument()
            {
                PrinterSettings = new PrinterSettings()
                {
                    PrinterName = "Microsoft Print to PDF",
                    PrintToFile = true,
                    PrintFileName = System.IO.Path.Combine(mDirectory, file),
                }
            };

            pDoc.PrintPage += new PrintPageEventHandler(Print_Page);
            pDoc.Print();
        }

        void Print_Page(object sender, PrintPageEventArgs e)
        {
            // Here you can play with the font style 
            // (and much much more, this is just an ultra-basic example)
            Font fnt = new Font("Courier New", 12);

            // Insert the desired text into the PDF file
            e.Graphics.DrawString
              (myHelper.NoteText, fnt, System.Drawing.Brushes.Black, 0, 0); //"When nothing goes right, go left"
        }
        private void PDF(Tekla.Structures.Geometry3d.Point Origin, Tekla.Structures.Geometry3d.Point FinishPoint, string Attribute) //, Model myModel
        {
            try
            {

                Component component = new Component();
                component.Name = ("PDFReferenceModel");
                component.Number = -100000;
                ComponentInput cInput = new ComponentInput();

                cInput.AddTwoInputPositions(Origin, FinishPoint);

                component.SetComponentInput(cInput);

                component.LoadAttributesFromFile(Attribute);
                component.Insert();

                myModel.CommitChanges();

                myHelper.LogFile("Write Note");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1011 - " + e.Message);
            }

        }
        #endregion

        #region Processes
        public string ProcessModel(ProjectLib Project)
        {

            myHelper.LogFile("Model process started");

            string Result = "";

            // probably not needed, added to find error
            if (Project.TemplateModel == null || Project.TemplateModel == "")
            {
                myHelper.LogFile("Tekla template attributes empty");
                System.Windows.Forms.MessageBox.Show("Template empty", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                Result = "BLANK";
                return Result;
            }

            Cursor.Current = Cursors.WaitCursor;
            tabControl2.Enabled = false;

            myHelper.LogFile("Cursor changed to wait");

            ModelHandler MH = new ModelHandler();

            myHelper.LogFile("Model Handler OK");

            try
            {
               
                MH.Save();

                myHelper.LogFile("Model Saved");

                MH.Close();

                myHelper.LogFile("Existing Model Closed");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1011 - " + e.Message);
            }

            try
            {
                myHelper.LogFile("Model Made: Name - " + Project.ModelName + " - Length - " + Project.ModelName.Length);
                myHelper.LogFile("Model Folder - " + Project.Folder);
                myHelper.LogFile("Model Template - " + Project.TemplateModel);

                MH.CreateNewSingleUserModel(Project.ModelName, Project.Folder, Project.TemplateModel);

            }
            catch (Exception e)
            {
                myHelper.LogFile("1001 - " + e.Message);
            }

            ProjectInfo projectInfo = myModel.GetProjectInfo();

            projectInfo.ProjectNumber = Project.Number;
            projectInfo.Name = Project.Client;
            projectInfo.Builder = txtBuilder.Text.ToUpper().Trim();
            projectInfo.Designer = txtDesigner.Text.ToUpper().Trim();
            projectInfo.Address = Project.Address;
            projectInfo.Description = Project.TeklaDesc;

            if (radN.Checked == true)
            {
                projectInfo.Info1 = "NORTH";
            }
            else if (radNE.Checked == true)
            {
                projectInfo.Info1 = "NORTH_EAST";
            }
            else if (radE.Checked == true)
            {
                projectInfo.Info1 = "EAST";
            }
            else if (radSE.Checked == true)
            {
                projectInfo.Info1 = "SOUTH_EAST";
            }
            else if (radS.Checked == true)
            {
                projectInfo.Info1 = "SOUTH";
            }
            else if (radSW.Checked == true)
            {
                projectInfo.Info1 = "SOUTH_WEST";
            }
            else if (radW.Checked == true)
            {
                projectInfo.Info1 = "WEST";
            }
            else if (radNW.Checked == true)
            {
                projectInfo.Info1 = "NORTH_WEST";
            }
            else
            {
                projectInfo.Info1 = "TBC";
            }

            projectInfo.SetUserProperty("CSB_BUILD_IMPOR", txtImportance.Text.Trim());
            projectInfo.SetUserProperty("CSB_TERRAIN_CAT", txtCategory.Text.Trim());
            projectInfo.SetUserProperty("CSB_WIND_REGION", txtRegion.Text.Trim());
            projectInfo.SetUserProperty("CSB_ENGINEER", txtEngineer.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_COMPUT_NO", txtComputation.Text.Trim());
            projectInfo.SetUserProperty("CSB_COMPUT_PAGES", txtCompPages.Text.Trim());
            projectInfo.SetUserProperty("CSB_COMPUT_DATE", txtDate.Text.Trim());

            switch (cbxLogo.Text.Trim())
            {
                case "CSB Steel Build":
                    projectInfo.SetUserProperty("CSB_LOGO", 0);
                    break;
                case "CSB Agricultural":
                    projectInfo.SetUserProperty("CSB_LOGO", 1);
                    break;
                case "CSB Aviation":
                    projectInfo.SetUserProperty("CSB_LOGO", 2);
                    break;
                case "CSB Commercial":
                    projectInfo.SetUserProperty("CSB_LOGO", 3);
                    break;
                case "CSB Custom":
                    projectInfo.SetUserProperty("CSB_LOGO", 4);
                    break;
                case "CSB Equinabuild":
                    projectInfo.SetUserProperty("CSB_LOGO", 5);
                    break;
                case "CSB Industrial":
                    projectInfo.SetUserProperty("CSB_LOGO", 6);
                    break;
                case "CSB Recreational":
                    projectInfo.SetUserProperty("CSB_LOGO", 7);
                    break;
            }

            projectInfo.SetUserProperty("FOOT_EMBED", txtEmbedment.Text.Trim());
            projectInfo.SetUserProperty("SOIL_BEAR_CAP", txtCapacity.Text.Trim());
            projectInfo.SetUserProperty("ROOF_CLADD", cbxRoofClad.Text.Trim());
            projectInfo.SetUserProperty("WALL_CLADD", cbxWallClad.Text.Trim());
            projectInfo.SetUserProperty("PURLIN_COAT", txtPurlinCoat.Text.Trim());

            if (cbxLift.Text == "Yes")
            {
                projectInfo.SetUserProperty("ROOF_LIFT", 1);
            }
            else
            {
                projectInfo.SetUserProperty("ROOF_LIFT", 0);
            }

            projectInfo.SetUserProperty("CSB_ROOF_COLOUR", cbxRoof.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_WALL_COLOUR", cbxWall.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_TRIM_COLOUR", cbxTrim.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_GUTTER_COLOUR", cbxGutter.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_RD_COLOUR", cbxRoller.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_SD_COLOUR", cbxSlide.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_PAD_COLOUR", cbxPA.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_RS_COLOUR", cbxRoofSky.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_WS_COLOUR", cbxWallSky.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_WB_COLOUR", cbxWhirly.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_WINDOWS_COLOUR", cbxWindow.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC1_COLOUR", cbxMisc1.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC1_DESCRIPTION", txtMisc1Desc.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC2_COLOUR", cbxMisc2.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC2_DESCRIPTION", txtMisc2Desc.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC3_COLOUR", cbxMisc3.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_MISC3_DESCRIPTION", txtMisc3Desc.Text.ToUpper().Trim());
            projectInfo.SetUserProperty("CSB_COLOUR_COMMENT", txtColourComment.Text.ToUpper().Trim());

            try
            {
                projectInfo.Modify();

                myModel.CommitChanges();

                MH.Save();

                myHelper.LogFile("Model - Project Properties Updated");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1002 - " + e.Message);
            }

            // **********************************************
            // calc lengths etc
            // **********************************************

            List<double> distanceListList = myHelper.getDistanceList(txtBaySize.Text.Trim());
            List<double> spacingList = myHelper.getSpacingList(txtBaySize.Text.Trim());

            double width = (double)decimal.Parse(txtWidth.Text.Trim());
            width = width * 1000;
            double pitch = (double)decimal.Parse(txtPitch.Text.Trim());
            double slab = (double)decimal.Parse(txtSlab.Text.Trim());
            double eave = (double)decimal.Parse(txtEave.Text.Trim());
            eave = eave * 1000;
            double apex = Math.Round(Math.Tan(pitch * (Math.PI / 180)) * width / 2 + eave, 0);
            double length = (double)decimal.Parse(txtLength.Text.Trim());
            length = length * 1000;

            //********************************************************************************
            // Remove existing grid and North Symbols
            //********************************************************************************

            if(checkConnection() == false) { myHelper.LogFile("Connection-3000"); }

            RemoveGridNorth();

            //***********************************************************************
            // Update Gable attributes

            if (checkConnection() == false) { myHelper.LogFile("Connection-3003"); }

            UpdateGableAttributes(spacingList, width, eave, apex, length, pitch, slab);

            // **********************************************

            if (checkConnection() == false) { myHelper.LogFile("Connection-3001"); }

            InsertGrid(distanceListList, spacingList, width, slab,  eave, apex);

            //********************************************************************************

            if (checkConnection() == false) { myHelper.LogFile("Connection-3002"); }

            createViews(length, width, apex);

            //***********************************************************************
            // Update roof/wall attributes for building layout

            if (checkConnection() == false) { myHelper.LogFile("Connection-3004"); }

            SetRoofWallLayoutAttributes(length, apex, width);

            //***********************************************************************
            // Adjust V-Ridge for pitch

            if (checkConnection() == false) { myHelper.LogFile("Connection-3005"); }

            AdjustVPitch(txtPitch.Text.Trim());

            //**********************************************************************

            if (checkConnection() == false) { myHelper.LogFile("Connection-3006"); }

            if (chkRolltop.Checked == true)
            {
                UpdateAttributes(@"Update\Project Roof Clad Left_RollTop.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
            }

            //**********************************************************************

            if (checkConnection() == false) { myHelper.LogFile("Connection-3016"); }

            if (slab != 0)
            {
                AdjustWallSheet("Project Side Wall Cladding Right", "-30","Side");
                AdjustWallSheet("Project Side Wall Cladding Left", "-30", "Side");
                AdjustWallSheet("Project End Wall Cladding Front Right", "-30", "End");
                AdjustWallSheet("Project End Wall Cladding Front Left", "-30", "End");
                AdjustWallSheet("Project End Wall Cladding Back Right", "-30", "End");
                AdjustWallSheet("Project End Wall Cladding Back Left", "-30", "End");

                createSlab(slab, length, width);
            }
            //**********************************************************************

            if (radModelYes.Checked == true)
            {

                if (checkConnection() == false) { myHelper.LogFile("Connection-3007"); }

                CreateModel(slab);
            }

            //**********************************************************************

            if (checkConnection() == false) { myHelper.LogFile("Connection-3008"); }

            AddProjectNotes();

            //**********************************************************************

            //updateViews();

            string[] MacrosPathList;
            string MacrosPath = string.Empty;
            TeklaStructuresSettings.GetAdvancedOption("XS_MACRO_DIRECTORY", ref MacrosPath);
            MacrosPathList = MacrosPath.Split(';');
            string vv = MacrosPathList.ElementAt(0);

            string temp = vv + @"\modeling\" + myHelper.ShareMacro();

            myHelper.LogFile("Share Macro Name - " + temp);

            try
            {
                if (File.Exists(temp))
                {
                    myHelper.LogFile("Share Macro Name - " + temp + " - exists");

                    bool ismacrounning = true;
                    Operation.RunMacro(myHelper.ShareMacro());
                    while (ismacrounning)
                    {
                        ismacrounning = Tekla.Structures.Model.Operations.Operation.IsMacroRunning();
                    }
                }
                else
                {
                    myHelper.LogFile("Share Macro Name - " + temp + " - missing");

                    Cursor.Current = Cursors.Default;
                    tabControl2.Enabled = true;
                    MessageBox.Show("Share macro missing");
                }

            }
            catch (Exception)
            {
                myHelper.LogFile("Share Macro Name - " + temp + " - failed");

                //System.Windows.Forms.MessageBox.Show(" not found, application stopped!", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                Cursor.Current = Cursors.Default;
                tabControl2.Enabled = true;
                throw;
            }

            Cursor.Current = Cursors.Default;
            tabControl2.Enabled = true;

            return Result;
        }

        private void createSlab(double slab, double length, double width)
        {
            
            string thick = slab.ToString();

            if(thick == "300")
            {
                thick = "300.01";
            }

            ContourPlate _Slab = new ContourPlate();

            _Slab.Name = "SLAB";
            _Slab.Profile.ProfileString = thick;
            _Slab.Material.MaterialString = "N40";
            _Slab.Class = "1";

            Contour ct = new Contour();

            ContourPoint p1 = new ContourPoint();
            p1.X = -_slabCorners.FGW;
            p1.Y = -_slabCorners.LGW;
            p1.Z = 0;
            p1.Chamfer = new Chamfer();

            ContourPoint p2 = new ContourPoint();
            p2.X = -_slabCorners.FGW;
            p2.Y = width + _slabCorners.RGW;
            p2.Z = 0;
            p2.Chamfer = new Chamfer();

            ContourPoint p3 = new ContourPoint();
            p3.X = length + _slabCorners.BGW;
            p3.Y = width + _slabCorners.RGW;
            p3.Z = 0;
            p3.Chamfer = new Chamfer();

            ContourPoint p4 = new ContourPoint();
            p4.X = length + _slabCorners.BGW;
            p4.Y = -_slabCorners.LGW;
            p4.Z = 0;
            p4.Chamfer = new Chamfer();

            ct.AddContourPoint(p1);
            ct.AddContourPoint(p2);
            ct.AddContourPoint(p3);
            ct.AddContourPoint(p4);

            _Slab.Contour = ct;
            _Slab.Position.Depth = Position.DepthEnum.FRONT;

            bool jj = _Slab.Insert();

            //MessageBox.Show(jj.ToString());

            myModel.CommitChanges();

            //if(_Slab.Type == ContourPlate.ContourPlateTypeEnum.SLAB)
            //{
            //    MessageBox.Show("OK");
            //}

        }

        private void AdjustVPitch(string pitch)
        {
            //********************************************************************************

            string modelPath = myModel.GetInfo().ModelPath;

            try
            {

                string RoofSettings = modelPath + @"\attributes\Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml";

                var xdocRoof = XDocument.Load(RoofSettings);

                var xtgt = xdocRoof.Root.Descendants("RidgeCapProfile").FirstOrDefault();

                string vRidge = "FLBK0.6*" + pitch + "*" + pitch + "*30*180*180*30";

                xtgt.Value = vRidge;

                myHelper.LogFile("V-ridge - " + vRidge);

                //***********************************************************************

                xdocRoof.Save(RoofSettings);

                //********************************************************************************

            }
            catch
            {
                myHelper.LogFile("1201 V-Ridge Update failed ");
            }

            //********************************************************************************

        }

        private void AdjustWallSheet(string wall,string offset,string position)
        {
            //********************************************************************************

            string modelPath = myModel.GetInfo().ModelPath;
            string WallSettings = "";

            try
            {
                
                if (position == "Side")
                {
                    WallSettings = modelPath + @"\attributes\" + wall + @".CSB_SideWall_Cladding.MainForm.xml";
                }
                else
                {
                    WallSettings = modelPath + @"\attributes\" + wall + @".CSB_EndWall_Cladding.MainForm.xml";
                }               

                var xdocRoof = XDocument.Load(WallSettings);

                var xtgt = xdocRoof.Root.Descendants("ApexDist").FirstOrDefault();

                xtgt.Value = offset;

                myHelper.LogFile("Bottom offset - " + WallSettings);

                //***********************************************************************

                xdocRoof.Save(WallSettings);

                //********************************************************************************

            }
            catch
            {
                myHelper.LogFile("1901 Bottom offset Update failed - " + WallSettings);
            }

            //********************************************************************************

        }

        private void RemoveGridNorth()
        {

            try
            {

                if (checkConnection() == false) { myHelper.LogFile("Connection-3009"); }

                ModelObjectEnumerator Enum = myModel.GetModelObjectSelector().GetAllObjects();

                while (Enum.MoveNext())
                {
                    Grid B = Enum.Current as Grid;
                    if (B != null)
                    {
                        B.Delete();
                    }

                    ContourPlate q = Enum.Current as ContourPlate;
                    if (q != null && radTBC.Checked == false)
                    {
                        var temp = "";
                        q.GetUserProperty("USER_FIELD_1", ref temp);

                        if (radN.Checked == true && temp == "NORTH")
                        {
                        }
                        else if (radNE.Checked == true && temp == "NORTH_EAST")
                        {
                        }
                        else if (radE.Checked == true && temp == "EAST")
                        {
                        }
                        else if (radSE.Checked == true && temp == "SOUTH_EAST")
                        {
                        }
                        else if (radS.Checked == true && temp == "SOUTH")
                        {
                        }
                        else if (radSW.Checked == true && temp == "SOUTH_WEST")
                        {
                        }
                        else if (radW.Checked == true && temp == "WEST")
                        {
                        }
                        else if (radNW.Checked == true && temp == "NORTH_WEST")
                        {
                        }
                        else
                        {

                            if (checkConnection() == false) { myHelper.LogFile("Connection-3010"); }

                            q.Delete();
                            temp = "";
                        }

                    }
                }

                if (checkConnection() == false) { myHelper.LogFile("Connection-3011"); }

                myModel.CommitChanges();

                myHelper.LogFile("North Removed");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1003 - " + e.Message);
            }

        }

        private void createViews(double mLength, double mWidth, double mApex)
        {

            try
            {

                String Top = "";
                String Right = "";
                String Bottom = "";
                String Left = "";

                if (radN.Checked == true)
                {
                    Top = "NORTH ELEVATION";
                    Right = "EAST ELEVATION";
                    Bottom = "SOUTH ELEVATION";
                    Left = "WEST ELEVATION";
                }
                else if (radNE.Checked == true)
                {
                    Top = "NE ELEVATION";
                    Right = "SE ELEVATION";
                    Bottom = "SW ELEVATION";
                    Left = "NW ELEVATION";
                }
                else if (radE.Checked == true)
                {
                    Top = "WEST ELEVATION";
                    Right = "NORTH ELEVATION";
                    Bottom = "EAST ELEVATION";
                    Left = "SOUTH ELEVATION";
                }
                else if (radSE.Checked == true)
                {
                    Top = "SW ELEVATION";
                    Right = "NW ELEVATION";
                    Bottom = "NE ELEVATION";
                    Left = "SE ELEVATION";
                }
                else if (radS.Checked == true)
                {
                    Top = "SOUTH ELEVATION";
                    Right = "WEST ELEVATION";
                    Bottom = "NORTH ELEVATION";
                    Left = "EAST ELEVATION";
                }
                else if (radSW.Checked == true)
                {
                    Top = "SE ELEVATION";
                    Right = "SW ELEVATION";
                    Bottom = "NW ELEVATION";
                    Left = "NE ELEVATION";
                }
                else if (radW.Checked == true)
                {
                    Top = "EAST ELEVATION";
                    Right = "SOUTH ELEVATION";
                    Bottom = "WEST ELEVATION";
                    Left = "NORTH ELEVATION";
                }
                else if (radNW.Checked == true)
                {
                    Top = "NE ELEVATION";
                    Right = "SE ELEVATION";
                    Bottom = "SW ELEVATION";
                    Left = "NW ELEVATION";
                }
                else
                {
                    Top = "RIGHT ELEVATION";
                    Right = "BACK ELEVATION";
                    Bottom = "LEFT ELEVATION";
                    Left = "FRONT ELEVATION";
                }

                if (checkConnection() == false) { myHelper.LogFile("Connection-3012"); }

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());

                TSG.Point Origin = new TSG.Point(0, 0, 0);
                TSG.Vector X = new TSG.Vector(1, 0, 0);
                TSG.Vector Y = new TSG.Vector(0, 1, 0);

                TransformationPlane XY_Plane = new TransformationPlane(Origin, X, Y);

                if (checkConnection() == false) { myHelper.LogFile("Connection-3013"); }

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                Tekla.Structures.Model.UI.View view = new Tekla.Structures.Model.UI.View();
                view.Name = Bottom;
                view.ViewCoordinateSystem.AxisX = new TSG.Vector(1, 0, 0);
                view.ViewCoordinateSystem.AxisY = new TSG.Vector(0, 0, 1);
                view.WorkArea.MinPoint = new TSG.Point(-2000, 0, -2000);
                view.WorkArea.MaxPoint = new TSG.Point(mLength + 2000, 0, mApex + 2000);
                view.ViewDepthUp = 2000;
                view.ViewDepthDown = mWidth / 2 + 1000;
                view.ViewFilter = "standard";
                view.CurrentRepresentation = "standard";
                view.DisplayType = Tekla.Structures.Model.UI.View.DisplayOrientationType.DISPLAY_VIEW_PLANE;
                view.SharedView = true;
                view.Insert();

                view = new Tekla.Structures.Model.UI.View();
                view.Name = Left;
                view.ViewCoordinateSystem.AxisX = new TSG.Vector(0, -1, 0);
                view.ViewCoordinateSystem.AxisY = new TSG.Vector(0, 0, 1);
                view.WorkArea.MinPoint = new TSG.Point(0, mWidth + 2000, -2000);
                view.WorkArea.MaxPoint = new TSG.Point(0, -2000, mApex + 2000);
                view.ViewDepthUp = 2000;
                view.ViewDepthDown = 3000;
                view.ViewFilter = "standard";
                view.CurrentRepresentation = "standard";
                view.DisplayType = Tekla.Structures.Model.UI.View.DisplayOrientationType.DISPLAY_VIEW_PLANE;
                view.SharedView = true;
                view.Insert();

                Origin = new TSG.Point(0, mWidth, 0);
                X = new TSG.Vector(1, 0, 0);
                Y = new TSG.Vector(0, 1, 0);

                XY_Plane = new TransformationPlane(Origin, X, Y);

                if (checkConnection() == false) { myHelper.LogFile("Connection-3014"); }

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                view = new Tekla.Structures.Model.UI.View();
                view.Name = Top;
                view.ViewCoordinateSystem.AxisX = new TSG.Vector(-1, 0, 0);
                view.ViewCoordinateSystem.AxisY = new TSG.Vector(0, 0, 1);
                view.WorkArea.MinPoint = new TSG.Point(mLength + 2000, 0, -2000);
                view.WorkArea.MaxPoint = new TSG.Point(-2000, 0, mApex + 2000);
                view.ViewDepthUp = 2000;
                view.ViewDepthDown = mWidth / 2 + 1000;
                view.ViewFilter = "standard";
                view.CurrentRepresentation = "standard";
                view.DisplayType = Tekla.Structures.Model.UI.View.DisplayOrientationType.DISPLAY_VIEW_PLANE;
                view.SharedView = true;
                view.Insert();

                Origin = new TSG.Point(mLength, -mWidth, 0);
                X = new TSG.Vector(1, 0, 0);
                Y = new TSG.Vector(0, 1, 0);

                XY_Plane = new TransformationPlane(Origin, X, Y);

                if (checkConnection() == false) { myHelper.LogFile("Connection-3015"); }

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                view = new Tekla.Structures.Model.UI.View();
                view.Name = Right;
                view.ViewCoordinateSystem.AxisX = new TSG.Vector(0, 1, 0);
                view.ViewCoordinateSystem.AxisY = new TSG.Vector(0, 0, 1);
                view.WorkArea.MinPoint = new TSG.Point(0, -2000, -2000);
                view.WorkArea.MaxPoint = new TSG.Point(0, mWidth + 2000, mApex + 2000);
                view.ViewDepthUp = 2000;
                view.ViewDepthDown = 3000;
                view.ViewFilter = "standard";
                view.CurrentRepresentation = "standard";
                view.DisplayType = Tekla.Structures.Model.UI.View.DisplayOrientationType.DISPLAY_VIEW_PLANE;
                view.SharedView = true;
                view.Insert();

                //**********************************************************
                // Move back to original origin
                //**********************************************************

                Origin = new TSG.Point(-mLength, 0, 0);
                X = new TSG.Vector(1, 0, 0);
                Y = new TSG.Vector(0, 1, 0);

                XY_Plane = new TransformationPlane(Origin, X, Y);

                if (checkConnection() == false) { myHelper.LogFile("Connection-3016"); }

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                //**********************************************************
                // Move to project text plane
                //**********************************************************

                if (myHelper.CreateNote() == "YES")
                {

                    Origin = new TSG.Point(-6000, 0, 0);
                    X = new TSG.Vector(1, 0, 0);
                    Y = new TSG.Vector(0, 1, 0);

                    XY_Plane = new TransformationPlane(Origin, X, Y);

                    if (checkConnection() == false) { myHelper.LogFile("Connection-3017"); }

                    myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                    myModel.CommitChanges();

                    view = new Tekla.Structures.Model.UI.View();
                    view.Name = "Project Details";
                    view.ViewCoordinateSystem.AxisX = new TSG.Vector(0, -1, 0);
                    view.ViewCoordinateSystem.AxisY = new TSG.Vector(0, 0, 1);
                    view.WorkArea.MinPoint = new TSG.Point(0, 2000, -12000);
                    view.WorkArea.MaxPoint = new TSG.Point(0, -20000, 20000);
                    view.ViewDepthUp = 200;
                    view.ViewDepthDown = 200;
                    view.ViewFilter = "standard";
                    view.CurrentRepresentation = "standard";
                    view.DisplayType = Tekla.Structures.Model.UI.View.DisplayOrientationType.DISPLAY_VIEW_PLANE;
                    view.SharedView = true;
                    view.Insert();

                    //**********************************************************
                    // Move back to original origin
                    //**********************************************************

                    Origin = new TSG.Point(6000, 0, 0);
                    X = new TSG.Vector(1, 0, 0);
                    Y = new TSG.Vector(0, 1, 0);

                    XY_Plane = new TransformationPlane(Origin, X, Y);

                    if (checkConnection() == false) { myHelper.LogFile("Connection-3018"); }

                    myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                    myModel.CommitChanges();
                }

                ModelViewEnumerator ViewEnum = ViewHandler.GetAllViews();
                while (ViewEnum.MoveNext())
                {
                    try
                    {

                        if (checkConnection() == false) { myHelper.LogFile("Connection-3019"); }

                        Tekla.Structures.Model.UI.View View = ViewEnum.Current;

                        ViewHandler.ShowView(View);
                        ViewHandler.RedrawView(View);

                        TSG.AABB B = new TSG.AABB(View.WorkArea);

                        ViewHandler.ZoomToBoundingBox(View, B);
                        ViewHandler.RedrawView(View);

                        if (View.Name == "3d-Rendered")
                        {
                            View.Modify();

                            if (myHelper.CreateNote() == "YES")
                            {
                                ViewHandler.HideView(View);
                            }
                            else
                            {
                                
                            }
                               
                            myModel.CommitChanges();
                        }
                        else if (View.Name == "Project Details")
                        {
                            View.VisibilitySettings.CutsVisibleInComponents = true;
                            View.Modify();
                            myModel.CommitChanges();
                        }
                        else
                        {
                            View.Modify();
                            ViewHandler.HideView(View);
                            myModel.CommitChanges();
                        }
                    }
                    catch
                    {
                        System.Windows.Forms.MessageBox.Show("Views Failed", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }

                if (checkConnection() == false) { myHelper.LogFile("Connection-3020"); }

                myModel.CommitChanges();

                myHelper.LogFile("Views created");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1004 - " + e.Message);
            }

        }

        private void updateViews()
        {

            ModelViewEnumerator ViewEnum = ViewHandler.GetAllViews();
            while (ViewEnum.MoveNext())
            {
                try
                {
                    Tekla.Structures.Model.UI.View View = ViewEnum.Current;

                    ViewHandler.ShowView(View);
                    ViewHandler.RedrawView(View);

                    TSG.AABB B = new TSG.AABB(View.WorkArea);

                    ViewHandler.ZoomToBoundingBox(View,  B);
                    //ViewHandler.RedrawWorkplane();
                    if (View.Name == "3d-Rendered")
                    {
                        View.Modify();
                        ViewHandler.HideView(View);
                        myModel.CommitChanges();
                    }
                    else if (View.Name == "Project Details")
                    {
                        View.VisibilitySettings.CutsVisibleInComponents = true;
                        View.Modify();
                        ViewHandler.HideView(View);
                        myModel.CommitChanges();
                    }
                    else
                    {
                        View.Modify();
                        ViewHandler.HideView(View);
                        myModel.CommitChanges();
                    }
                }
                catch
                {
                    System.Windows.Forms.MessageBox.Show("Views Failed", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }

            myModel.CommitChanges();

            myHelper.LogFile("Views updated");
        }

        private void InsertGrid(List<double> distanceListList, List<double> spacingList, double width, double slab, double eave, double apex)
        {

            List<double> spacingList2 = new List<double>();

            if (_NoMullions > 0)
            {
                int temp = _NoMullions + 1;

                for (int index = 1; index < temp+1; ++index)
                {
                    spacingList2.Add(Math.Round(width / temp, 0));
                }
                //    spacingList2.Add(Math.Round(width / temp, 0));
                //spacingList2.Add(Math.Round(width / temp, 0));
                //spacingList2.Add(Math.Round(width / temp, 0));
            }
            else
            {
                spacingList2.Add(Math.Round(width / 2, 0));
                spacingList2.Add(Math.Round(width / 2, 0));
            }

            double eave2 = Math.Round(slab + eave, 0);
            double apex2 = Math.Round(slab + apex, 0);

            string RLs = "";
            string levels = "";

            if (slab == 0)
            {
                RLs = "0 " + eave.ToString() + " " + apex.ToString();
                levels = '"' + "0 (GROUND)" + '"' + " EAVE APEX";
            }
            else
            {
                //string xSlab = '"' + slab.ToString() + " (FSL)" + '"';
                double zEave = slab + eave;
                double zApex = slab + apex;
                RLs = "0 " + slab.ToString() + " " + zEave.ToString() + " " + zApex.ToString();
                levels = '"' + "0 (GROUND)" + '"' + " FSL EAVE APEX";
            }

            Grid grid = new Grid();

            grid.CoordinateX = "0 ";

            for (int index = 1; index < spacingList.Count; ++index)
            {
                double num3 = spacingList[index];
                grid.CoordinateX = grid.CoordinateX + num3.ToString() + " ";
            }

            grid.CoordinateY = "0 ";

            for (int index = 0; index < spacingList2.Count; ++index)
                grid.CoordinateY = grid.CoordinateY + spacingList2[index].ToString() + " ";

            System.Drawing.Color color = System.Drawing.Color.FromArgb(0, 0, 0);

            grid.CoordinateZ = RLs;

            grid.LabelZ = levels;

            grid.LabelX = "1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20";
            grid.LabelY = "A B C D E F G H I J K L M N O P Q R";
            grid.Color = color.ToArgb();
            grid.FontColor = color;
            grid.Insert();

            if (checkConnection() == false) { myHelper.LogFile("Connection-3021"); }

            myModel.CommitChanges();
        }

        private void UpdateGableAttributes(List<double> spacingList, double width, double eave, double apex, double length, double pitch, double slab)
        { 

            //TODO: Set Attributes for width - Done
            //********************************************************************************
            // Change attribute file for width
            //********************************************************************************

            string AttributeSettings = "";
            string Portal1_Settings = "";

            if (width < 15000)
            {
                AttributeSettings = "API Gable 12m Reg A";
                Portal1_Settings = "API Portal 12w 5h Reg A";
            }
            else if (width < 18000)
            {
                AttributeSettings = "API Gable 15m Reg A";
                Portal1_Settings = "API Portal 15w 6h Reg A";
            }
            else if (width < 21000)
            {
                AttributeSettings = "API Gable 18m Reg A";
                Portal1_Settings = "API Portal 18w 6h Reg A";
            }
            else if (width < 24000)
            {
                AttributeSettings = "API Gable 21m Reg A";
                Portal1_Settings = "API Portal 21w 6h Reg A";
            }
            else if (width < 30000)
            {
                AttributeSettings = "API Gable 24m Reg A";
                Portal1_Settings = "API Portal 24w 6h Reg A";
            }
            else if (width < 36000)
            {
                AttributeSettings = "API Gable 30m Reg A";
                Portal1_Settings = "API Portal 30w 6h Reg A";
            }
            else if (width < 40000)
            {
                AttributeSettings = "API Gable 36m Reg A";
                Portal1_Settings = "API Portal 36w 6h Reg A";
            }
            else if (width < 46000)
            {
                AttributeSettings = "API Gable 40m Reg A";
                Portal1_Settings = "API Portal 40w 6h Reg A";
            }
            else if (width < 46000)
            {
                AttributeSettings = "API Gable 46m Reg A";
                Portal1_Settings = "API Portal 46w 6h Reg A";
            }
            else
            {
                AttributeSettings = "API Gable 50m Reg A";
                Portal1_Settings = "API Portal 50w 6h Reg A";
            }

            myHelper.LogFile("API Attribute Settings - " + AttributeSettings);
            myHelper.LogFile("API Portal Settings - " + Portal1_Settings);

            if (checkConnection() == false) { myHelper.LogFile("Connection-3022"); }

            string modelPath = myModel.GetInfo().ModelPath;

            string attribute = modelPath + @"\attributes\CSB_Project_Setup.CSB_Gable_Shed.MainForm.xml";

            //********************************************************************************
            // Update column / Foot Cage
            //********************************************************************************

            try
            {

                string portalSettings = modelPath + @"\attributes\" + Portal1_Settings + ".CSB_Portal.MainForm.xml";

                File.Copy(myHelper.TeklaFolder() + Portal1_Settings + ".CSB_Portal.MainForm.xml", portalSettings, true);

                myHelper.LogFile("Portal copied - " + portalSettings);

                var xdocPortal = XDocument.Load(portalSettings);

                if (ProjectSales.ColumnType != null && (ProjectSales.ColumnType.Contains("UB") || ProjectSales.ColumnType.Contains("UC")))
                {
                    
                    if (myHelper.CheckColumn(ProjectSales.ColumnType.Trim(), columnSize))
                    {

                        var xtgt = xdocPortal.Root.Descendants("RightProfile").FirstOrDefault();

                        xtgt.Value = columnSize.TeklaProfile;

                        xtgt = xdocPortal.Root.Descendants("LeftProfile").FirstOrDefault();

                        xtgt.Value = columnSize.TeklaProfile;

                        myHelper.LogFile("Column profile - " + columnSize.TeklaProfile);

                    }
                    else
                    {

                        myHelper.LogFile("ERROR Column profile not changed - " + ProjectSales.ColumnType.Trim());

                        var xtgt10 = xdocPortal.Root.Descendants("RightProfile").FirstOrDefault();
                       
                        if (myHelper.CheckColumnTekla(xtgt10.Value, columnSize))
                        {

                        }
                    }

                }

                _slabCorners.ColumnWidth =  columnSize.FlangeW;

                //***********************************************************************
                //Update footings for slab

                if (slab != 0)
                {
                   var xtgt2 = xdocPortal.Root.Descendants("RightBase").FirstOrDefault();

                    switch (xtgt2.Value.Trim())
                    {
                        case "API Baseplate 1m No Slab":
                            xtgt2.Value = "API Baseplate 1m Slab";
                            break;
                        case "API Baseplate 1m No Slab L":
                            xtgt2.Value = "API Baseplate 1m Slab L";
                            break;
                        case "API Baseplate 1m No Slab S":
                            xtgt2.Value = "API Baseplate 1m Slab S";
                            break;
                    }
                   
                   var xtgt3 = xdocPortal.Root.Descendants("LeftBase").FirstOrDefault();

                    switch (xtgt3.Value.Trim())
                    {
                        case "API Baseplate 1m No Slab":
                            xtgt3.Value = "API Baseplate 1m Slab";
                            break;
                        case "API Baseplate 1m No Slab L":
                            xtgt3.Value = "API Baseplate 1m Slab L";
                            break;
                        case "API Baseplate 1m No Slab S":
                            xtgt3.Value = "API Baseplate 1m Slab S";
                            break;
                    }

                    if(xtgt2.Value == xtgt3.Value)
                    {
                        setFooting(modelPath, slab, xtgt2.Value);
                    }
                    else
                    {
                        setFooting(modelPath, slab, xtgt2.Value);
                        setFooting(modelPath, slab, xtgt3.Value);
                    }

                    setFooting(modelPath, slab, "API Baseplate 1m Slab Mullion");

                }

                //***********************************************************************

                xdocPortal.Save(portalSettings);

                //********************************************************************************


            }
            catch
            {
                myHelper.LogFile("1200 Column Update failed ");
            }
            //********************************************************************************

            if (File.Exists(attribute))
            {
                myHelper.LogFile("3000 attribute " + attribute);
            }
            else
            {
                myHelper.LogFile("1017 File does not exist - " + attribute);
            }

            var xdoc = XDocument.Load(attribute);

            string xFile = myHelper.TeklaFolder() + AttributeSettings + ".CSB_Gable_Shed.MainForm.xml";

            if (File.Exists(xFile))
            {
                myHelper.LogFile("3001 attribute " + xFile);
            }
            else
            {
                myHelper.LogFile("1018 File does not exist - " + xFile);
            }

            var xdocAttrib = XDocument.Load(xFile);

            foreach (var childElement in xdocAttrib.Root.Elements())
            {
                string a = childElement.Name.ToString();
                string c = childElement.Value.ToString();

                if (a != "SpacingBetweenBays" && a != "Millimeters")
                {

                    var tgt2 = xdoc.Root.Descendants(a).FirstOrDefault();

                    tgt2.Value = c;

                }

            }

            myHelper.LogFile("3002 ");

            //********************************************************************************
            // Fill attribute file
            //********************************************************************************

            var tgt = xdoc.Root.Descendants("Height").FirstOrDefault();

            tgt.Value = eave.ToString();

            xdoc.Root.Descendants("SpacingBetweenBays").Remove();

            for (int index = 1; index < spacingList.Count; ++index)
            {
                double temp2 = spacingList[index];

                XElement temp3 = new XElement("Millimeters", temp2.ToString());

                if (index == 1)
                {
                    xdoc.Element("config")
                        .Elements("Height").FirstOrDefault()
                        .AddAfterSelf(new XElement("SpacingBetweenBays",
                        temp3
                        ));
                }
                else
                {
                    xdoc.Element("config")
                        .Elements("SpacingBetweenBays")
                        .Elements("Millimeters").LastOrDefault()
                        .AddAfterSelf(temp3);
                }
            }

            myHelper.LogFile("3003 ");

            tgt = xdoc.Root.Descendants("Width").FirstOrDefault();

            tgt.Value = width.ToString();

            myHelper.LogFile("3004 ");

            tgt = xdoc.Root.Descendants("SlopeAngle").FirstOrDefault();

            tgt.Value = pitch.ToString();

            myHelper.LogFile("3005 ");

            tgt = xdoc.Root.Descendants("PortalAtt1").FirstOrDefault();

            tgt.Value = Portal1_Settings;

            myHelper.LogFile("3006 ");

            string gridNo = "1";
            string flyBraceAttrib = "1";
            string bayNo = "1";
            string SplitSingle = "2";

            myHelper.LogFile("3007 ");

            for (int index = 1; index < spacingList.Count; ++index)
            {
                gridNo = gridNo + " " + (index + 1);
                flyBraceAttrib = flyBraceAttrib + " 1";
            }

            myHelper.LogFile("3008 ");

            for (int index = 1; index < spacingList.Count - 1; ++index)
            {
                bayNo = bayNo + " " + (index + 1);
            }

            myHelper.LogFile("3009 ");

            for (int index = 1; index < spacingList.Count - 2; ++index)
            {
                SplitSingle = SplitSingle + " " + (index + 2);
            }

            myHelper.LogFile("*********************************");

            myHelper.LogFile("Grid numbers - " + gridNo);
            myHelper.LogFile("Bay numbers - " + bayNo);

            tgt = xdoc.Root.Descendants("Portal1Grids").FirstOrDefault(); //Portal 1 Grids

            tgt.Value = gridNo;

            tgt = xdoc.Root.Descendants("flyBraceGrids").FirstOrDefault(); //Purlin fly insert At grids

            tgt.Value = gridNo;

            tgt = xdoc.Root.Descendants("SideflyBays").FirstOrDefault(); //Girts side fly brace connection

            tgt.Value = gridNo;

            //********************************************************************************
            // Purlin Girt Sizes
            //********************************************************************************

            tgt = xdoc.Root.Descendants("Purlin_OverLabDis").FirstOrDefault();

            //if (xTemp.Contains("Z"))
            //{
                int lap = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

                tgt.Value = lap.ToString();
            //}
            //else
            //{
            //    tgt.Value = "0";
            //}

            string xTemp = txtPurlin.Text.Trim();

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                tgt = xdoc.Root.Descendants("EavePurProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

                tgt = xdoc.Root.Descendants("MidPurProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

                tgt = xdoc.Root.Descendants("RidgePurProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

                //tgt = xdoc.Root.Descendants("Purlin_OverLabDis").FirstOrDefault();

                //if (xTemp.Contains("Z"))
                //{
                //    int lap = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

                //    tgt.Value = lap.ToString();
                //}
                //else
                //{
                //    tgt.Value = "0";
                //}
            }
            else
            {
                tgt = xdoc.Root.Descendants("EavePurProfile").FirstOrDefault();

                string mTemp = tgt.Value;

                txtPurlin.Text = mTemp.Substring(6);
            }

            // Left Wall Girt

            xTemp = txtWallGirtSide.Text.Trim();

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                tgt = xdoc.Root.Descendants("GirtProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

            }
            else
            {
                tgt = xdoc.Root.Descendants("GirtProfile").FirstOrDefault();

                string mTemp = tgt.Value;

                txtWallGirtSide.Text = mTemp.Substring(6);
            }
            _slabCorners.LeftGirt= tgt.Value;

            // Right Side Wall Girt

            xTemp = txtWallGirtSideRight.Text.Trim();

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                tgt = xdoc.Root.Descendants("GirtProfile2").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

            }
            else
            {
                tgt = xdoc.Root.Descendants("GirtProfile2").FirstOrDefault();

                string mTemp = tgt.Value;

                txtWallGirtSideRight.Text = mTemp.Substring(6);
            }
            _slabCorners.RightGirt = tgt.Value;

            // Sidewall Girt Lap

            //tgt = xdoc.Root.Descendants("SideGirtOverlap").FirstOrDefault();

            //if (txtWallGirtSide.Text.Contains("Z") || txtWallGirtSideRight.Text.Contains("Z"))
            //{
            //    lap = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

            //    tgt.Value = lap.ToString();
            //}
            //else
            //{
            //    tgt.Value = "0";
            //}

            //// Sidewall lap


            tgt = xdoc.Root.Descendants("SideGirtOverlap").FirstOrDefault();

            //if (xTemp.Contains("Z"))
            //{
                lap = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

                tgt.Value = lap.ToString();
            //}
            //else
            //{
            //    tgt.Value = "0";
            //}

            // Fascia Girt

            xTemp = txtFascia.Text.Trim(); //txtFascia

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                //tgt = xdoc.Root.Descendants("GirtProfile").FirstOrDefault();
                //tgt.Value = "MET-MS" + xTemp;
                tgt = xdoc.Root.Descendants("FasciaProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;

                //tgt = xdoc.Root.Descendants("SideGirtOverlap").FirstOrDefault();

                //if (xTemp.Contains("Z"))
                //{
                //    int lap = 5 * (int)Math.Round((spacingList.Max() * 0.15) / 5.0);

                //    tgt.Value = lap.ToString();
                //}
                //else
                //{
                //    tgt.Value = "0";
                //}
            }
            else
            {
                tgt = xdoc.Root.Descendants("FasciaProfile").FirstOrDefault();

                string mTemp = tgt.Value;

                txtWallGirtSide.Text = mTemp.Substring(6);
            }

            // Front Endwall Girt

            xTemp = txtWallGirtEnd.Text.Trim();

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                tgt = xdoc.Root.Descendants("EndGirtProfile").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;
            }
            else
            {
                tgt = xdoc.Root.Descendants("EndGirtProfile").FirstOrDefault();

                string mTemp = tgt.Value;

                txtWallGirtEnd.Text = mTemp.Substring(6);
            }
            _slabCorners.FrontGirt = tgt.Value;

            // Back Endwall Girt

            xTemp = txtWallGirtEndBack.Text.Trim();

            if (xTemp != null && xTemp != "" && xTemp.Length == 6)
            {
                tgt = xdoc.Root.Descendants("EndGirtProfile2").FirstOrDefault();
                tgt.Value = "MET-MS" + xTemp;
            }
            else
            {
                tgt = xdoc.Root.Descendants("EndGirtProfile2").FirstOrDefault();

                string mTemp = tgt.Value;

                txtWallGirtEndBack.Text = mTemp.Substring(6);
            }
            _slabCorners.BackGirt = tgt.Value;

            //TODO: Purlin Split Location - needs updating for different roof/wall

            //********************************************************************************
            // split locations
            //********************************************************************************

            List<string> split = CalcSplitDoubleSpan(spacingList);

            // Roof

            if (chkPurlinSingleSpan.Checked == false)
            {
                tgt = xdoc.Root.Descendants("SplitGrids").FirstOrDefault(); //Purlin split grids

                tgt.Value = split[0];

                tgt = xdoc.Root.Descendants("bridgingbays1").FirstOrDefault(); //purlin bridging insert at bays

                tgt.Value = split[1];

                tgt = xdoc.Root.Descendants("bridgingbays2").FirstOrDefault(); //purlin bridging insert at bays

                tgt.Value = split[2];

                myHelper.LogFile("Purlin Double span split - " + split[0]);
                myHelper.LogFile("Purlin Bridging bays 1 - " + split[1]);
                myHelper.LogFile("Purlin Bridging bays 2 - " + split[2]);
            }
            else if (chkPurlinSingleSpan.Checked == true)
            {
                tgt = xdoc.Root.Descendants("SplitGrids").FirstOrDefault(); //Purlin split grids

                tgt.Value = SplitSingle;

                tgt = xdoc.Root.Descendants("bridgingbays1").FirstOrDefault(); //purlin bridging insert at bays

                tgt.Value = bayNo;

                tgt = xdoc.Root.Descendants("bridgingbays2").FirstOrDefault(); //purlin bridging insert at bays

                tgt.Value = "0";

                myHelper.LogFile("Purlin Single span split - " + SplitSingle);
                myHelper.LogFile("Purlin Bridging bays 1 - " + bayNo);
                myHelper.LogFile("Purlin Bridging bays 2 - " + 0);
            }

            // Sidewalls

            if (chkGirtSingleSpan.Checked == false)
            {
                tgt = xdoc.Root.Descendants("SideSplitGrids").FirstOrDefault(); //Left Girt split

                tgt.Value = split[0];

                tgt = xdoc.Root.Descendants("SideSplitGrids2").FirstOrDefault(); //Right Girt split

                tgt.Value = split[0];

                tgt = xdoc.Root.Descendants("Sidebridgingbays1").FirstOrDefault(); // girts side bridging bays

                tgt.Value = split[1];

                tgt = xdoc.Root.Descendants("Sidebridgingbays2").FirstOrDefault(); // girts side bridging bays

                tgt.Value = split[2];

                myHelper.LogFile("Girt Double span split - " + split[0]);
                myHelper.LogFile("Girt Bridging bays 1 - " + split[1]);
                myHelper.LogFile("Girt Bridging bays 2 - " + split[2]);
            }
            else if (chkGirtSingleSpan.Checked == true)
            {
                tgt = xdoc.Root.Descendants("SideSplitGrids").FirstOrDefault(); //Left Girt split

                tgt.Value = SplitSingle;

                tgt = xdoc.Root.Descendants("SideSplitGrids2").FirstOrDefault(); //Right Girt split

                tgt.Value = SplitSingle;

                tgt = xdoc.Root.Descendants("Sidebridgingbays1").FirstOrDefault(); // girts side bridging bays

                tgt.Value = bayNo;

                tgt = xdoc.Root.Descendants("Sidebridgingbays2").FirstOrDefault(); // girts side bridging bays

                tgt.Value = "0";

                myHelper.LogFile("Girt Single span split - " + SplitSingle);
                myHelper.LogFile("Girt Bridging bays 1 - " + bayNo);
                myHelper.LogFile("Girt Bridging bays 2 - " + 0);
            }

            myHelper.LogFile("*********************************");

            //TODO: Cladding - Done

            //********************************************************************************
            // Set Project roof and walls templates
            //********************************************************************************

            // roof - set template names
            tgt = xdoc.Root.Descendants("RightRCAtt").FirstOrDefault();
            tgt.Value = "Project Roof Cladding Right";
            tgt = xdoc.Root.Descendants("LeftRCAtt").FirstOrDefault();
            tgt.Value = "Project Roof Cladding Left";
            tgt = xdoc.Root.Descendants("CreateRC").FirstOrDefault();
            tgt.Value = "1"; // Roof always on

            // side and end walls - set template names - all off
            tgt = xdoc.Root.Descendants("BackEWCAtt").FirstOrDefault();
            tgt.Value = "Project End Wall Cladding Back Left";
            tgt = xdoc.Root.Descendants("BackEWCAtt2").FirstOrDefault();
            tgt.Value = "Project End Wall Cladding Back Right";

            tgt = xdoc.Root.Descendants("CreateFrontEWC").FirstOrDefault();
            tgt.Value = "0"; // off

            tgt = xdoc.Root.Descendants("FrontEWCAtt").FirstOrDefault();
            tgt.Value = "Project End Wall Cladding Front Left";
            tgt = xdoc.Root.Descendants("FrontEWCAtt2").FirstOrDefault();
            tgt.Value = "Project End Wall Cladding Front Right";

            tgt = xdoc.Root.Descendants("CreateBackEWC").FirstOrDefault();
            tgt.Value = "0"; // off

            tgt = xdoc.Root.Descendants("RightSWCAtt").FirstOrDefault();
            tgt.Value = "Project Side Wall Cladding Right";

            tgt = xdoc.Root.Descendants("CreateRightSWC").FirstOrDefault();
            tgt.Value = "0"; // off

            tgt = xdoc.Root.Descendants("LeftSWCAtt").FirstOrDefault();
            tgt.Value = "Project Side Wall Cladding Left";

            tgt = xdoc.Root.Descendants("CreateLeftSWC").FirstOrDefault();
            tgt.Value = "0"; // off

            //********************************************************************************
            // Set walls 
            //********************************************************************************

            //if (checkBox1.Checked == true) // ROOF ONLY
            //{
            // turn girts off
            tgt = xdoc.Root.Descendants("BackGirt").FirstOrDefault();
            tgt.Value = "0";
            tgt = xdoc.Root.Descendants("FrontGirt").FirstOrDefault();
            tgt.Value = "0";
            tgt = xdoc.Root.Descendants("RightGirt").FirstOrDefault();
            tgt.Value = "0";
            tgt = xdoc.Root.Descendants("LeftGirt").FirstOrDefault();
            tgt.Value = "0";

            // turn mullions off
            tgt = xdoc.Root.Descendants("CreateFrontMullions").FirstOrDefault();
            tgt.Value = "1";
            tgt = xdoc.Root.Descendants("CreateBackMullions").FirstOrDefault();
            tgt.Value = "1";

            tgt = xdoc.Root.Descendants("CreateMidMullions").FirstOrDefault(); // Remove Front Mid Wall mullion
            tgt.Value = "1";
            tgt = xdoc.Root.Descendants("CreateMidMullions2").FirstOrDefault(); // Remove back Mid wall mullion
            tgt.Value = "0";

            // Right wall
            if (btnRight.BackColor == System.Drawing.Color.Red)
            {
                tgt = xdoc.Root.Descendants("RightGirt").FirstOrDefault();
                tgt.Value = "1";
                tgt = xdoc.Root.Descendants("CreateRightSWC").FirstOrDefault();
                tgt.Value = "1";
            }
            else
            {
                tgt = xdoc.Root.Descendants("RightGirt").FirstOrDefault();
                tgt.Value = "0";
            }

            // Left wall
            if (btnLeft.BackColor == System.Drawing.Color.Red)
            {
                tgt = xdoc.Root.Descendants("LeftGirt").FirstOrDefault();
                tgt.Value = "1";
                tgt = xdoc.Root.Descendants("CreateLeftSWC").FirstOrDefault();
                tgt.Value = "1";
            }
            else
            {
                tgt = xdoc.Root.Descendants("LeftGirt").FirstOrDefault();
                tgt.Value = "0";
            }

            // Front wall
            if (btnFront.BackColor == System.Drawing.Color.Red)
            {
                tgt = xdoc.Root.Descendants("FrontGirt").FirstOrDefault();
                tgt.Value = "1";
                tgt = xdoc.Root.Descendants("CreateFrontMullions").FirstOrDefault();
                tgt.Value = "0";
                tgt = xdoc.Root.Descendants("CreateFrontEWC").FirstOrDefault();
                tgt.Value = "1";
            }
            else
            {
                tgt = xdoc.Root.Descendants("FrontGirt").FirstOrDefault();
                tgt.Value = "0";
                tgt = xdoc.Root.Descendants("CreateFrontMullions").FirstOrDefault();
                tgt.Value = "1";
            }

            // Back wall
            if (btnRear.BackColor == System.Drawing.Color.Red)
            {
                tgt = xdoc.Root.Descendants("BackGirt").FirstOrDefault();
                tgt.Value = "1";
                tgt = xdoc.Root.Descendants("CreateBackMullions").FirstOrDefault();
                tgt.Value = "0";
                tgt = xdoc.Root.Descendants("CreateBackEWC").FirstOrDefault();
                tgt.Value = "1";
            }
            else
            {
                tgt = xdoc.Root.Descendants("BackGirt").FirstOrDefault();
                tgt.Value = "0";
                tgt = xdoc.Root.Descendants("CreateBackMullions").FirstOrDefault();
                tgt.Value = "1";
            }

            //***********************************************************************
            // Mullion with Slab

            if(slab != 0)
            {
                tgt = xdoc.Root.Descendants("MullionBase1").FirstOrDefault();
                tgt.Value = "API Baseplate 1m Slab Mullion";

                tgt = xdoc.Root.Descendants("MullionBase2").FirstOrDefault();
                tgt.Value = "API Baseplate 1m Slab Mullion";

                tgt = xdoc.Root.Descendants("MullionBase3").FirstOrDefault();
                tgt.Value = "API Baseplate 1m Slab Mullion";
            }
            // depth set with others above

            //***********************************************************************
            // Calc number of Mullion

            var tgtFront = xdoc.Root.Descendants("NoOfFrontMullions").FirstOrDefault();
           var tgtBack = xdoc.Root.Descendants("NoOfBackMullions").FirstOrDefault();

            if (tgtFront != null && tgtBack != null)
            {
                int _front= 0;
                int _back = 0;

                bool isParsable = Int32.TryParse(tgtFront.Value, out _front);

                bool isParsable2 = Int32.TryParse(tgtFront.Value, out _back);

                if(isParsable == true || isParsable2 == true)
                {
                    if (_front >= _back && _front > 0)
                    {
                        _NoMullions = _front;
                    }
                    else if (_back >= _front && _back > 0)
                    {
                        _NoMullions = _back;
                    }
                    else
                    {
                        _NoMullions = 0;
                    }
                }
                else
                {
                    _NoMullions = 0;
                }

            }

            //***********************************************************************

            xdoc.Save(attribute);

            //********************************************************************************

        }

        private void setFooting(string modelPath,double slab,string _target)
        {

            string footingSettings = modelPath + @"\attributes\" + _target + ".CSB_Base_and_Cage.MainForm.xml";

            try
            {
                File.Copy(myHelper.TeklaFolder() + _target + ".CSB_Base_and_Cage.MainForm.xml", footingSettings, true);

                myHelper.LogFile("Footing copied - " + footingSettings);

            }
            catch
            {

                myHelper.LogFile("Footing Not copied - " + footingSettings);
                return;
            }

            var xdocFooting = XDocument.Load(footingSettings);

            var FootDist = xdocFooting.Root.Descendants("foot_dis").FirstOrDefault();

            FootDist.Value = slab.ToString();

            xdocFooting.Save(footingSettings);

        }

        private List<string> CalcSplitDoubleSpan(List<double> spacingList)
        {

            List<string> Result = new List<string>();

            string BridgingBays1 = "";
            string BridgingBays2 = "";

            int f = spacingList.Count - 1; // number of bays

            string SplitList = "";

            if (f == 1)
            {
                SplitList = "0";
                BridgingBays2 = "0";
                BridgingBays1 = "1";
            }
            else if (f == 2)
            {
                SplitList = "0";
                BridgingBays2 = "0";
                BridgingBays1 = "1 2";
            }
            else if (f == 3)
            {
                SplitList = "3";
                BridgingBays2 = "3";
                BridgingBays1 = "1 2";
            }
            else if (f == 5)
            {
                SplitList = "3 4";
                BridgingBays2 = "3";
                BridgingBays1 = "1 2 4 5";
            }
            else
            {
                int xCheck = 0;

                int nBay = 1;

                int split = (f / 2) - 1;

                if (split == 0)
                {
                    SplitList = "0";
                }
                else
                {
                    int xTemp = 1;


                    for (int index = 1; index < split + 1; ++index)
                    {

                        xTemp = xTemp + 2;

                        if (index == 1)
                        {
                            BridgingBays1 = BridgingBays1 + nBay + " " + (nBay + 1) + " ";
                        }

                        if (f % 2 == 0) // even number
                        {
                            SplitList = SplitList + xTemp + " ";

                            BridgingBays1 = BridgingBays1 + (nBay + 2) + " " + (nBay + 3) + " ";
                            BridgingBays2 = "0";
                        }
                        else
                        {
                            if (index > (split / 2))
                            {
                                if (xCheck == 0)
                                {
                                    SplitList = SplitList + (xTemp) + " ";
                                    SplitList = SplitList + (xTemp + 1) + " ";

                                    BridgingBays2 = BridgingBays2 + (xTemp);
                                    BridgingBays1 = BridgingBays1 + (xTemp + 1) + " ";
                                    xCheck = 1;
                                }
                                else
                                {
                                    SplitList = SplitList + (xTemp + 1) + " ";

                                    BridgingBays1 = BridgingBays1 + (nBay + 2) + " " + (nBay + 3) + " ";
                                }


                            }
                            else
                            {
                                SplitList = SplitList + xTemp + " ";

                                BridgingBays1 = BridgingBays1 + (nBay + 2) + " " + (nBay + 3) + " ";
                            }

                        }

                        nBay = nBay + 2;
                    }

                    if (f % 2 == 0) // even number
                    {
                    }
                    else
                    {
                        BridgingBays1 = BridgingBays1 + (nBay + 2);
                    }

                }
            }

            Result.Add(SplitList.Trim());
            Result.Add(BridgingBays1.Trim());
            Result.Add(BridgingBays2.Trim());

            return Result;
        }

        private void SetRoofWallLayoutAttributes(double length, double apex, double width)
        {

            try
            {

                //***********************************************************************
                // Update roof attributes

                if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxRoofClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project Roof Clad Left Corro.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right Corro.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project Roof Clad Left.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.Red)
                {                    
                    if (cbxRoofClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project Roof Clad Left Corro_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right Corro_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left Corro_Front Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right Corro_Front Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    } 
                    else
                    {
                        UpdateAttributes("Project Roof Clad Left_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left_Front Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right_Front Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    }

                    //Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                    //Origin.X = 0;
                    //Origin.Y = width / 2;
                    //Origin.Z = apex + 500;

                    //Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
                    //FinishPoint.X = 2000;
                    //FinishPoint.Y = width / 2;
                    //FinishPoint.Z = apex + 500;

                    //AddInformationNote(Origin, FinishPoint, "Remove Raker Angle");
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.White)
                {                    
                    if (cbxRoofClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project Roof Clad Left Corro_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right Corro_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left Corro_Back Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right Corro_Back Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project Roof Clad Left_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left_Back Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right_Back Open 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    }

                    //Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                    //Origin.X = length;
                    //Origin.Y = width / 2;
                    //Origin.Z = apex + 500;

                    //Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
                    //FinishPoint.X = length + 2000;
                    //FinishPoint.Y = width / 2;
                    //FinishPoint.Z = apex + 500;

                    //AddInformationNote(Origin, FinishPoint, "Remove Raker Angle");
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.White)
                {      
                    if (cbxRoofClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project Roof Clad Left Corro_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right Corro_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left Corro_Roof Only 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right Corro_Roof Only 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project Roof Clad Left_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                        UpdateAttributes("Project Roof Clad Right_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                        if (txtPurlin.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project Roof Clad Left_Roof Only 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                            UpdateAttributes(@"150\Project Roof Clad Right_Roof Only 150.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                        }
                    }
                }

                //***********************************************************************
                // Update wall attributes

                // Endwalls settings

                // front-right corner
                if (btnFront.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Front Right Corro_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Front Right_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.White)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Front Right Corro_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEnd.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Front Right Corro_Right Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Front Right_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEnd.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Front Right_Right Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                }

                // front-left corner
                if (btnFront.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Front Left Corro_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Front Left_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.White)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Front Left Corro_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEnd.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Front Left Corro_Left Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Front Left_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEnd.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Front Left_Left Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                }

                // back-right corner
                if (btnRear.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Back Right Corro_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Back Right_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                    }
                }
                else if (btnRear.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.White)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Back Right Corro_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEndBack.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Back Right Corro_Right Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Back Right_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEndBack.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Back Right_Right Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                }

                // back-left corner
                if (btnRear.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Back Left Corro_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Back Left_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                    }
                }
                else if (btnRear.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.White)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project EW Clad Back Left Corro_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEndBack.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Back Left Corro_Left Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project EW Clad Back Left_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");

                        if (txtWallGirtEndBack.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project EW Clad Back Left_Left Open 150.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                        }
                    }
                }

                //***********************************************************************
                // Sidewall Left settings 

                if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.White && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Left Corro_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left Corro_FrontBack Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Left_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left_FrontBack Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.White && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Left Corro_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left Corro_Back Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Left_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left_Back Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Left Corro_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left Corro_Front Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Left_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSide.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Left_Front Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Left Corro.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Left.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                    }
                }

                //***********************************************************************
                // Sidewall Right settings 

                if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.White && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Right Corro_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right Corro_FrontBack Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Right_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right_FrontBack Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.White && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Right Corro_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right Corro_Back Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Right_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right_Back Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Right Corro_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right Corro_Front Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Right_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");

                        if (txtWallGirtSideRight.Text.Trim().Contains("150"))
                        {
                            UpdateAttributes(@"150\Project SW Clad Right_Front Open 150.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                        }
                    }
                }
                else
                {
                    if (cbxWallClad.Text == "0.47-TCT-CORRY")
                    {
                        UpdateAttributes("Project SW Clad Right Corro.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                    }
                    else
                    {
                        UpdateAttributes("Project SW Clad Right.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                    }
                }

                myHelper.LogFile("Roof Wall Layout Attributes");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1005 - " + e.Message);
            }

            //********************************************************************************

            //TODO: change V-Ridge to suit pitch

            //********************************************************************************
        }

        private void UpdateAttributes(string Original ,string Destination)
        {

            try
            {

                string modelPath = myModel.GetInfo().ModelPath;

                string attribute = modelPath + @"\attributes\" + Destination;

                if (File.Exists(attribute))
                {                    

                }
                else
                {
                    myHelper.LogFile("1014 File does not exist - " + attribute);
                }

                var xdoc = XDocument.Load(attribute);

                //TODO: change to variable
                string xFile = myHelper.Setting() + Original; //@"T:\CSB_Program_Files\Documentation\Settings\"

                if (File.Exists(xFile))
                {

                }
                else
                {
                    myHelper.LogFile("1015 File does not exist - " + xFile);
                }

                var xdocAttrib = XDocument.Load(xFile);

                foreach (var childElement in xdocAttrib.Root.Elements())
                {
                    string a = childElement.Name.ToString();
                    string c = childElement.Value.ToString();

                    var tgt2 = xdoc.Root.Descendants(a).FirstOrDefault();

                    tgt2.Value = c;

                }

                //********************************************************************************

                xdoc.Save(attribute);

                //********************************************************************************

                myHelper.LogFile("Update Attributes");
                myHelper.LogFile("Original - " + Original + " - Changed - " + Destination);
            }
            catch (Exception e)
            {
                myHelper.LogFile("1006 - " + e.Message);
            }

        }

        private void CreateModel(double slab)
        {
            try
            {

                Component component = new Component();
                component.Name = ("CSB_Gable_Shed");
                component.Number = -100000;
                ComponentInput cInput = new ComponentInput();

                Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                Origin.X = 0;
                Origin.Y = 0;
                Origin.Z = slab;

                cInput.AddOneInputPosition(Origin);

                component.SetComponentInput(cInput);

                component.LoadAttributesFromFile("CSB_Project_Setup");
                component.Insert();

                myModel.CommitChanges();

                myHelper.LogFile("Model Created");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1007 - " + e.Message);
            }

        }

        #endregion

        #region Notes

        private void AddProjectNotes()
        {
            
            string modelPath = myModel.GetInfo().ModelPath;

            string attribute = modelPath + @"\attributes\CSB_Project_Setup.TextContourPlate.TextContourPlateWindow.xml";

            if (File.Exists(attribute))
            {
                var xdoc = XDocument.Load(attribute);

                //********************************************************************************
                // Fill attribute file
                //********************************************************************************

                string newText = "Job number      : " + ProjectSales.JobNo + "\r\n";
                newText += "Quote version   : " + ProjectSales.QuoteVer + "\r\n";
                newText += "Company          : " + ProjectSales.CompanyName + "\r\n";
                newText += "Customer          : " + ProjectSales.CustomerName + "\r\n";
                newText += "Location            : " + ProjectSales.Suburb + "\r\n";
                newText += "Roof type          : " + ProjectSales.RoofType + "\r\n";
                newText += "Roof pitch         : " + ProjectSales.RoofPitch + "\r\n";
                newText += "Total walls         : " + ProjectSales.Totwalls + "\r\n";
                newText += "Side walls         : " + ProjectSales.SideWals + "\r\n";
                newText += "End walls          : " + ProjectSales.EndWalls + "\r\n";
                newText += "Roof material    : " + ProjectSales.RoofMaterial + " Colour : " + ProjectSales.RoofColour + "\r\n";
                newText += "Roof skylight     : " + ProjectSales.ClearSheetRoof + "\r\n";
                newText += "Wall material     : " + ProjectSales.WallMaterial + " Colour : " + ProjectSales.WallColour + "\r\n";
                newText += "Wall skylight      : " + ProjectSales.ClearSheetWall + "\r\n";
                newText += "Flashings          : " + "\r\n";
                newText += "        Ridge        : " + ProjectSales.FlashingRidge + "\r\n";
                newText += "        Barge        : " + ProjectSales.Barge + "\r\n";
                newText += "        Corner       : " + ProjectSales.Corner + "\r\n";
                newText += "        Gutter        : " + ProjectSales.GutterType + " Colour : " + ProjectSales.GutterColour + "\r\n";
                newText += "Frame               : " + "\r\n";
                newText += "        Column         : " + ProjectSales.ColumnType + "\r\n";
                newText += "        Truss            : " + ProjectSales.TrussType + "\r\n";
                newText += "        Purlin            : " + txtPurlin.Text + "\r\n";
                newText += "        Sidewall Girt   : " + txtWallGirtSide.Text + "\r\n";
                newText += "        Endwall Girt    : " + txtWallGirtEnd.Text + "\r\n";
                newText += "        Other            : " + ProjectSales.OtherFrameDetails + "\r\n";
                newText += "        Footings       : " + ProjectSales.Footings + " : Finish : " + ProjectSales.Finish + "\r\n";
                newText += "\r\n";
                newText += "        Project Details: " + "\r\n";
                newText += txtProjectDetails.Text + "\r\n";
                newText += "\r\n";
                newText += "        Notes      : " + "\r\n";
                newText += txtNote.Text;

                var tgt = xdoc.Root.Descendants("Phrase").FirstOrDefault();

                tgt.Value = newText;

                newText += newText;

                myHelper.NoteText = newText;

                xdoc.Save(attribute);

                //********************************************************************************

                PrintPDF(modelPath + @"\attributes\");

                myHelper.LogFile("Notes Printed");

                //********************************************************************************

                if (myHelper.CreateNote() == "YES")
                {
                    try
                    {

                        Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                        Origin.X = -6000;
                        Origin.Y = -3000;
                        Origin.Z = 0;

                        Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
                        FinishPoint.X = -6000;
                        FinishPoint.Y = -9000;
                        FinishPoint.Z = 0;

                        WriteNote(Origin, FinishPoint, "CSB_Project_Setup"); //, myModel

                        myHelper.LogFile("Add Project Notes");
                    }
                    catch (Exception e)
                    {
                        myHelper.LogFile("1008 - " + e.Message);
                    }
                }

            }
            else
            {
                myHelper.LogFile("1012 File does not exist - " + attribute);
            }

        }

        private void AddInformationNote(Tekla.Structures.Geometry3d.Point Origin, Tekla.Structures.Geometry3d.Point FinishPoint, string newText)
        {
            
            string modelPath = myModel.GetInfo().ModelPath;

            string attribute = modelPath + @"\attributes\CSB_Project_Note.TextContourPlate.TextContourPlateWindow.xml";

            if (File.Exists(attribute))
            {

                var xdoc = XDocument.Load(attribute);

                //********************************************************************************
                // Fill attribute file
                //********************************************************************************

                var tgt = xdoc.Root.Descendants("Phrase").FirstOrDefault();

                tgt.Value = newText;

                xdoc.Save(attribute);

                //********************************************************************************

                try
                {

                    WriteNote(Origin, FinishPoint, "CSB_Project_Note"); //, myModel

                    myHelper.LogFile("Add Information Note");
                }
                catch (Exception e)
                {
                    myHelper.LogFile("1009 - " + e.Message);
                }

            }
            else
            {
                myHelper.LogFile("1013 File does not exist - " + attribute);
            }

        }

        private void WriteNote(Tekla.Structures.Geometry3d.Point Origin, Tekla.Structures.Geometry3d.Point FinishPoint, string Attribute) //, Model myModel
        {
            try
            {
                
                Component component = new Component();
                component.Name = ("Text Contour Plate");
                component.Number = -100000;
                ComponentInput cInput = new ComponentInput();

                cInput.AddTwoInputPositions(Origin, FinishPoint);

                component.SetComponentInput(cInput);

                component.LoadAttributesFromFile(Attribute);
                component.Insert();

                myModel.CommitChanges();

                myHelper.LogFile("Write Note");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1011 - " + e.Message);
            }

        }

        #endregion

        #region Wall Layout

        private void btnRight_Click(object sender, EventArgs e)
        {
            if (btnRight.BackColor == System.Drawing.Color.White)
            {
                btnRight.BackColor = System.Drawing.Color.Red;
            }
            else
            {
                btnRight.BackColor = System.Drawing.Color.White;
            }
        }

        private void btnFront_Click(object sender, EventArgs e)
        {
            if (btnFront.BackColor == System.Drawing.Color.White)
            {
                btnFront.BackColor = System.Drawing.Color.Red;
            }
            else
            {
                btnFront.BackColor = System.Drawing.Color.White;
            }
        }

        private void btnLeft_Click(object sender, EventArgs e)
        {
            if (btnLeft.BackColor == System.Drawing.Color.White)
            {
                btnLeft.BackColor = System.Drawing.Color.Red;
            }
            else
            {
                btnLeft.BackColor = System.Drawing.Color.White;
            }
        }

        private void btnRear_Click(object sender, EventArgs e)
        {
            if (btnRear.BackColor == System.Drawing.Color.White)
            {
                btnRear.BackColor = System.Drawing.Color.Red;
            }
            else
            {
                btnRear.BackColor = System.Drawing.Color.White;
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                panelLayout.Enabled = false;
                btnRight.BackColor = System.Drawing.Color.White;
                btnFront.BackColor = System.Drawing.Color.White;
                btnLeft.BackColor = System.Drawing.Color.White;
                btnRear.BackColor = System.Drawing.Color.White;
            }
            else
            {
                panelLayout.Enabled = true;
            }
        }
      
        #endregion

        #region Menu
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }
        private void settingsToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Settings temp = new Settings();
            temp.ShowDialog();
        }

        private void manageFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ManageFiles temp = new ManageFiles();
            temp.ShowDialog();
        }
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void modelShareHelpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            SaveHelp temp = new SaveHelp();
            temp.Show();
        }

        #endregion

        #region Error Check

        private void validateAll(EventArgs e)
        {
            txtNumber_Validated(this, e);
            txtClient_Validated(this, e);
            txtBuilder_Validated(this, e);
            txtLength_Validated(this, e);
            txtWidth_Validated(this, e);
            txtEave_Validated(this, e);
            txtPitch_Validated(this, e);
            txtBaySize_Validated(this, e);
            txtSlab_Validated(this, e);
        }

        private void txtNumber_Validating(object sender, CancelEventArgs e)
        {

        }

        private void txtNumber_Validated(object sender, EventArgs e)
        {

            bool bTest = txtNumberIsEmpty();

            if (bTest == true)

            {
                this.errorProvider1.SetError(txtNumber, "This field must be filled");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider1.SetError(txtNumber, "");
            }

        }

        private void txtClient_Validated(object sender, EventArgs e)
        {

            bool bTest = txtClientIsEmpty();

            if (bTest == true)
            {
                this.errorProvider2.SetError(txtClient, "This field must be filled");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider2.SetError(txtClient, "");
            }

        }

        private void txtBuilder_Validated(object sender, EventArgs e)
        {

            bool bTest = txtBuilderIsEmpty();

            if (bTest == true)
            {
                this.errorProvider3.SetError(txtBuilder, "This field must be filled");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider3.SetError(txtBuilder, "");
            }

        }

        private void txtLength_Validated(object sender, EventArgs e)
        {

            bool bTest = txtLengthIsEmpty();

            if (bTest == true)
            {
                this.errorProvider4.SetError(txtLength, "This field must be filled");
                Globals.checkError = 1;
            }

            bTest = txtLengthNotNumeric();

            if (bTest == false)
            {
                this.errorProvider4.SetError(txtLength, "This field must contain number");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider4.SetError(txtLength, "");
            }

        }

        private void txtWidth_Validated(object sender, EventArgs e)
        {

            bool bTest = txtWidthIsEmpty();

            if (bTest == true)

            {
                this.errorProvider5.SetError(txtWidth, "This field must be filled");
                Globals.checkError = 1;
            }

            bTest = txtWidthNotNumeric();

            if (bTest == false)
            {
                this.errorProvider5.SetError(txtWidth, "This field must contain number");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider5.SetError(txtWidth, "");
            }

        }

        private void txtEave_Validated(object sender, EventArgs e)
        {

            bool bTest = txtEaveIsEmpty();

            if (bTest == true)

            {
                this.errorProvider6.SetError(txtEave, "This field must be filled");
                Globals.checkError = 1;
            }

            bTest = txtEaveNotNumeric();

            if (bTest == false)
            {
                this.errorProvider6.SetError(txtEave, "This field must contain number");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider6.SetError(txtEave, "");
            }
        }

        private void txtPitch_Validated(object sender, EventArgs e)
        {

            bool bTest = txtPitchIsEmpty();

            if (bTest == true)

            {
                this.errorProvider7.SetError(txtPitch, "This field must be filled");
                Globals.checkError = 1;
            }

            bTest = txtPitchNotNumeric();

            if (bTest == false)
            {
                this.errorProvider7.SetError(txtPitch, "This field must contain number");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider7.SetError(txtPitch, "");
            }
        }

        private void txtBaySize_Validated(object sender, EventArgs e)
        {

            bool bTest = txtBaySizeIsEmpty();

            if (bTest == true)
            {
                this.errorProvider8.SetError(txtBaySize, "This field must be filled");
                Globals.checkError = 1;
                return;
            }

            List<double> distanceListList = myHelper.getDistanceList(txtBaySize.Text.Trim());
            double temp = distanceListList.Last();
            double length = (double)decimal.Parse(txtLength.Text.Trim());
            length = length * 1000;
                        
            if (temp >= (length-1) && temp <= (length + 1))
            {
               
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Bay Sizes = " + Math.Round(temp/1000,3) + " m" + "\r\n" + "Accept and continue", "Building Length Incorrect", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //do something
                }
                else if (dialogResult == DialogResult.No)
                {
                    this.errorProvider8.SetError(txtBaySize, " Bay sizes do not equal length");
                    Globals.checkError = 1;
                    return;
                }

            }

            this.errorProvider8.SetError(txtBaySize, "");

        }

        private void txtSlab_Validated(object sender, EventArgs e)
        {

            bool bTest = txtSlabIsEmpty();

            if (bTest == true)

            {
                this.errorProvider9.SetError(txtSlab, "This field must be filled");
                Globals.checkError = 1;
                return;
            }

            bTest = txtSlabNotNumeric();

            if (bTest == false)
            {
                this.errorProvider9.SetError(txtSlab, "This field must contain number");
                Globals.checkError = 1;
            }
            else
            {
                this.errorProvider9.SetError(txtSlab, "");
            }

        }

#region CheckNumeric

        private bool txtLengthNotNumeric()
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(txtLength.Text.Trim(), out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;

        }

        private bool txtWidthNotNumeric()
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(txtWidth.Text.Trim(), out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;

        }

        private bool txtEaveNotNumeric()
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(txtEave.Text.Trim(), out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;

        }

        private bool txtPitchNotNumeric()
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(txtPitch.Text.Trim(), out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;

        }

        private bool txtSlabNotNumeric()
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(txtSlab.Text.Trim(), out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;

        }

#endregion

#region CheckEmpty

        private bool txtNumberIsEmpty()

        {

            if (txtNumber.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }

        private bool txtClientIsEmpty()

        {

            if (txtClient.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtBuilderIsEmpty()

        {

            if (txtBuilder.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtLengthIsEmpty()

        {

            if (txtLength.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtWidthIsEmpty()

        {

            if (txtWidth.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtEaveIsEmpty()

        {

            if (txtEave.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtPitchIsEmpty()

        {

            if (txtPitch.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtBaySizeIsEmpty()

        {

            if (txtBaySize.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }


        private bool txtSlabIsEmpty()

        {

            if (txtSlab.Text == string.Empty)

            {

                return true;

            }

            else

            {

                return false;

            }
        }

        private void cbxPurlin_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtPurlin.Text = cbxPurlin.Text; 
            
            if (txtPurlin.Text != null && txtPurlin.Text.Contains("Z"))
            {
                chkPurlinSingleSpan.Checked = true;
            }
            else
            {
                chkPurlinSingleSpan.Checked = false;
            }
        }

        private void cbxGirtSide_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWallGirtSide.Text = cbxGirtSide.Text;

            if (txtWallGirtSide.Text != null && txtWallGirtSide.Text.Contains("Z"))
            {
                chkGirtSingleSpan.Checked = true;
            }
            else
            {
                chkGirtSingleSpan.Checked = false;
            }
        }

        private void cbxGirtEnd_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWallGirtEnd.Text = cbxGirtEnd.Text;
        }

        private void cbxGirtSideRight_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWallGirtSideRight.Text = cbxGirtSideRight.Text;
        }

        private void cbxFascia_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtFascia.Text = cbxFascia.Text;
        }

        private void cbxGirtEndBack_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtWallGirtEndBack.Text = cbxGirtEndBack.Text;
        }
    }

    #endregion

    #endregion
   
}
