using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        //C:\Users\TeklaAutomation\AppData\Local\CSB_Project_Start\app-1.0.2

        Helper myHelper = new Helper();

        salesLib ProjectSales = new salesLib();

        Model myModel = new Model();

        public Form1()
        {
            InitializeComponent();
            LoadCbx(cbxRoof);
            LoadCbx(cbxWall);
            LoadCbx(cbxTrim);
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
            temp.Items.Add("CSB Steel Build");
            temp.Items.Add("CSB Agricultural");
            temp.Items.Add("CSB Aviation");
            temp.Items.Add("CSB Commercial");
            temp.Items.Add("CSB Custom");
            temp.Items.Add("CSB Equinabuild");
            temp.Items.Add("CSB Industrial");
            temp.Items.Add("CSB Recreational");
        }

        private void LoadCbx(ComboBox temp)
        {
            temp.Items.Clear();
            temp.Items.Add("CBOND(TBC)");
            temp.Items.Add("ZINC");
            temp.Items.Add("BASALT");
            temp.Items.Add("CLASSIC CREAM");
            temp.Items.Add("COTTAGE GREEN");
            temp.Items.Add("COVE");
            temp.Items.Add("DEEP OCEAN");
            temp.Items.Add("DUNE");
            temp.Items.Add("EVENING HAZE");
            temp.Items.Add("GULLY");
            temp.Items.Add("IRONSTONE");
            temp.Items.Add("JASPER");
            temp.Items.Add("MANGROVE");
            temp.Items.Add("MANOR RED");
            temp.Items.Add("MONUMENT");
            temp.Items.Add("NIGHT SKY");
            temp.Items.Add("PALE EUCALYPT");
            temp.Items.Add("PAPERBARK");
            temp.Items.Add("SHALE GREY");
            temp.Items.Add("SURFMIST");
            temp.Items.Add("TERRAIN");
            temp.Items.Add("WALLABY");
            temp.Items.Add("WINDSPRAY");
            temp.Items.Add("WOODLAND GREY");
        }
        private void LoadSkyCbx(ComboBox temp)
        {
            temp.Items.Clear();
            temp.Items.Add("OPAL");
            temp.Items.Add("CLEAR");
        }
        private void LoadSheetCbx(ComboBox temp)
        {
            temp.Items.Clear();
            temp.Items.Add("0.47 TCT 5-RIB");
            temp.Items.Add("0.47 TCT CORRY");
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
            }
            else if (ProjectSales.GutterColour != null && ProjectSales.GutterColour.Contains("Colorbond"))
            {
                cbxTrim.Text = "CBOND(TBC)";
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
                cbxRoofClad.Text = "0.47 TCT 5-RIB";
            }
            else if (ProjectSales.RoofMaterial != null && ProjectSales.RoofMaterial.Contains(".42 BMT") && ProjectSales.RoofMaterial.Contains("Corry"))
            {
                cbxRoofClad.Text = "0.47 TCT CORRY";
            }

            if (ProjectSales.WallMaterial != null && ProjectSales.WallMaterial.Contains(".42 BMT") && ProjectSales.WallMaterial.Contains("5-Rib"))
            {
                cbxWallClad.Text = "0.47 TCT 5-RIB";
            }
            else if (ProjectSales.WallMaterial != null && ProjectSales.WallMaterial.Contains(".42 BMT") && ProjectSales.WallMaterial.Contains("Corry"))
            {
                cbxWallClad.Text = "0.47 TCT CORRY";
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Model myModel = new Model();

            //string modelPath = myModel.GetInfo().ModelPath;

            //string attribute = modelPath + @"\attributes\CSB_Project_Setup.TextContourPlate.TextContourPlateWindow.xml";

            //var xdoc = XDocument.Load(attribute);

            ////********************************************************************************
            //// Fill attribute file
            ////********************************************************************************

            //string newText = "Customer Name : " + ProjectSales.CustomerName + "\r\n" + "Second";

            //var tgt = xdoc.Root.Descendants("Phrase").FirstOrDefault();

            //tgt.Value = newText;


            //xdoc.Save(attribute);

            ////********************************************************************************

            //try
            //{

            //    //Model myModel = new Model();

            //    Component component = new Component();
            //    component.Name = ("Text Contour Plate");
            //    component.Number = -100000;
            //    ComponentInput cInput = new ComponentInput();

            //    Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
            //    Origin.X = 0;
            //    Origin.Y = -3000;
            //    Origin.Z = 0;

            //    Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
            //    FinishPoint.X = 6000;
            //    FinishPoint.Y = -3000;
            //    FinishPoint.Z = 0;

            //    //cInput.AddOneInputPosition(Origin);
            //    cInput.AddTwoInputPositions(Origin, FinishPoint);

            //    component.SetComponentInput(cInput);

            //    component.LoadAttributesFromFile("CSB_Project_Setup");
            //    component.Insert();

            //    myModel.CommitChanges();

            //}
            //catch
            //{
            //    MessageBox.Show("Text did not run");
            //}

            //createViews(18000, 12000, 7500);

            myHelper.LogFile("test again");

            //Model myModel = new Model();

            ////********************************************************************************
            //// Remove existing grid
            ////********************************************************************************

            //ModelObjectEnumerator Enum = myModel.GetModelObjectSelector().GetAllObjects();

            //while (Enum.MoveNext())
            //{
            //    Component B = Enum.Current as Component;
            //    if (B != null)
            //    {
            //        //System.Windows.Forms.MessageBox.Show(B.Name, "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

            //        if (B.Name == "CSB_Gable_Shed")
            //        {
            //            try
            //            {
            //                //B.GetAttribute()
            //                B.SetAttribute("SlopeAngle", 18.0);
            //                B.SetAttribute("Width", 25000);

            //                B.Modify();
            //            }
            //            catch (Exception)
            //            {
            //                System.Windows.Forms.MessageBox.Show(" not found, application stopped!", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

            //            }

            //        }

            //        //string d =  B.Name;
            //        //B.Delete();
            //    }

            //ContourPlate q = Enum.Current as ContourPlate;
            //if (q != null)
            //{
            //    var temp = "";
            //    q.GetUserProperty("USER_FIELD_1", ref temp);

            //    if (temp == "Check")
            //    {
            //        q.Delete();
            //        temp = "";
            //    }

            //}
            //}

            //myModel.CommitChanges();


            //Model Model = new Model();

            //Beam B = new Beam(new TSG.Point(0, 0, 0), new TSG.Point(0, 0, 1000));
            //Beam B1 = new Beam(new TSG.Point(0, 1000, 0), new TSG.Point(0, 1000, 5000));
            //Beam B2 = new Beam(new TSG.Point(0, 2000, 0), new TSG.Point(0, 2000, 5000));

            //B.Profile.ProfileString = "UB310*32";
            //B1.Profile.ProfileString = "UB310*32";
            //B2.Profile.ProfileString = "UB310*32";

            //B.Insert();
            //B1.Insert();
            //B2.Insert();

            //ArrayList ObjectsToSelect = new ArrayList();
            //ObjectsToSelect.Add(B);
            ////ObjectsToSelect.Add(B1);
            //ObjectsToSelect.Add(B2);

            //Tekla.Structures.Model.UI.ModelObjectSelector MS = new Tekla.Structures.Model.UI.ModelObjectSelector();
            //MS.Select(ObjectsToSelect);

            //Model.CommitChanges();

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            myModel = new Model();

            // Check that the model connection succeeded:
            if (myModel.GetConnectionStatus())
            {
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
            }

            Globals.checkError = 0;
            validateAll(e);

            if (Globals.checkError == 1)
            {
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

            if (Directory.Exists(xtemp))
            {
                System.Windows.Forms.MessageBox.Show("Already exists", "Project", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            string xResult = ProcessModel(Project);
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

        #region Processes
        public string ProcessModel(ProjectLib Project)
        {

            string Result = "";

            // probably not needed, added to find error
            if (Project.TemplateModel == null || Project.TemplateModel == "")
            {
                System.Windows.Forms.MessageBox.Show("Template empty", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                Result = "BLANK";
                return Result;
            }

            Cursor.Current = Cursors.WaitCursor;
            tabControl2.Enabled = false;

            ModelHandler MH = new ModelHandler();

            try
            {
               
                MH.Save();

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

            RemoveGridNorth();

            // **********************************************

            InsertGrid(distanceListList, spacingList, width, slab,  eave, apex);

            //********************************************************************************

            createViews(length, width, apex);

            //***********************************************************************
            // Update Gable attributes

            UpdateGableAttributes(spacingList, width, eave, apex, length);

            //***********************************************************************
            // Update roof/wall attributes for building layout

            SetRoofWallLayoutAttributes(length, apex, width);

            //**********************************************************************

            CreateModel();

            //**********************************************************************

            AddProjectNotes();

            //**********************************************************************

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

        private void RemoveGridNorth()
        {

            try
            {
                ModelObjectEnumerator Enum = myModel.GetModelObjectSelector().GetAllObjects();

                while (Enum.MoveNext())
                {
                    Grid B = Enum.Current as Grid;
                    if (B != null)
                    {
                        B.Delete();
                    }

                    ContourPlate q = Enum.Current as ContourPlate;
                    if (q != null)
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
                            q.Delete();
                            temp = "";
                        }

                    }
                }

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

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());

                TSG.Point Origin = new TSG.Point(0, 0, 0);
                TSG.Vector X = new TSG.Vector(1, 0, 0);
                TSG.Vector Y = new TSG.Vector(0, 1, 0);

                TransformationPlane XY_Plane = new TransformationPlane(Origin, X, Y);

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

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                //**********************************************************
                // Move to project text plane
                //**********************************************************

                Origin = new TSG.Point(-6000, 0, 0);
                X = new TSG.Vector(1, 0, 0);
                Y = new TSG.Vector(0, 1, 0);

                XY_Plane = new TransformationPlane(Origin, X, Y);

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

                myModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(XY_Plane);
                myModel.CommitChanges();

                ModelViewEnumerator ViewEnum = ViewHandler.GetAllViews();
                while (ViewEnum.MoveNext())
                {
                    try
                    {
                        Tekla.Structures.Model.UI.View View = ViewEnum.Current;

                        ViewHandler.RedrawView(view);
                        ViewHandler.ShowView(view);
                        ViewHandler.RedrawWorkplane();
                        if (View.Name == "3d-Rendered" || View.Name == "Project Details")
                        {

                        }
                        else
                        {
                            ViewHandler.HideView(view);
                        }
                    }
                    catch
                    {
                        System.Windows.Forms.MessageBox.Show("Views Failed", "Tekla Structures", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }

                myModel.CommitChanges();

                myHelper.LogFile("North Removed");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1004 - " + e.Message);
            }

        }

        private void InsertGrid(List<double> distanceListList, List<double> spacingList, double width, double slab, double eave, double apex)
        {

            //List<double> distanceListList = myHelper.getDistanceList(txtBaySize.Text.Trim());
            //List<double> spacingList = myHelper.getSpacingList(txtBaySize.Text.Trim());

            List<double> spacingList2 = new List<double>();
            //double width = (double)decimal.Parse(txtWidth.Text.Trim());
            //width = width * 1000;
            if (width > 15000)
            {
                spacingList2.Add(Math.Round(width / 3, 0));
                spacingList2.Add(Math.Round(width / 3, 0));
                spacingList2.Add(Math.Round(width / 3, 0));
            }
            else
            {
                spacingList2.Add(Math.Round(width / 2, 0));
                spacingList2.Add(Math.Round(width / 2, 0));
            }

            //double pitch = (double)decimal.Parse(txtPitch.Text.Trim());
            //double slab = (double)decimal.Parse(txtSlab.Text.Trim());
            //double eave = (double)decimal.Parse(txtEave.Text.Trim());
            //eave = eave * 1000;
            //double apex = Math.Round(Math.Tan(pitch * (Math.PI / 180)) * width / 2 + eave, 0);
            double eave2 = Math.Round(slab + eave, 0);
            double apex2 = Math.Round(slab + apex, 0);

            string RLs = "";
            string levels = "";

            if (slab == 0)
            {
                RLs = "0 " + eave.ToString() + " " + apex.ToString();
                levels = '"' + "0 (GROUND)" + '"' + " EAVE APEX";
                //levels = '"' + "0 (GROUND)" + '"' + " " + eave.ToString() + " " + apex.ToString();
            }
            else
            {
                string xSlab = '"' + slab.ToString() + " (FSL)" + '"';
                RLs = "-" + slab.ToString() + " 0 " + eave.ToString() + " " + apex.ToString();
                levels = '"' + "0 (GROUND)" + '"' + " FSL EAVE APEX";
                //levels = '"' + "0 (GROUND)" + '"' + " " + xSlab + " " + eave2.ToString() + " " + apex2.ToString();
            }

            ////********************************************************************************
            //// Remove existing grid
            ////********************************************************************************

            //RemoveGridNorth();

            ////********************************************************************************

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
                        
            myModel.CommitChanges();

            //double temp = distanceListList.Last();
            //double length = (double)decimal.Parse(txtLength.Text.Trim());
            //length = length * 1000;

            //createViews(length, width, apex);

        }

        private void UpdateGableAttributes(List<double> spacingList, double width, double eave, double apex, double length)
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

            string modelPath = myModel.GetInfo().ModelPath;

            string attribute = modelPath + @"\attributes\CSB_Project_Setup.CSB_Gable_Shed.MainForm.xml";

            if (File.Exists(attribute))
            {

            }
            else
            {
                myHelper.LogFile("1017 File does not exist - " + attribute);
            }

            var xdoc = XDocument.Load(attribute);

            string xFile = @"T:\CSB_TeklaSetup\" + AttributeSettings + ".CSB_Gable_Shed.MainForm.xml";

            if (File.Exists(xFile))
            {

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

            tgt = xdoc.Root.Descendants("Width").FirstOrDefault();

            tgt.Value = width.ToString();

            tgt = xdoc.Root.Descendants("PortalAtt1").FirstOrDefault();

            tgt.Value = Portal1_Settings;

            string gridNo = "1";
            string flyBraceAttrib = "1";
            string bayNo = "1";

            for (int index = 1; index < spacingList.Count; ++index)
            {
                gridNo = gridNo + " " + (index+1);
                flyBraceAttrib = flyBraceAttrib + " 1";
            }

            for (int index = 1; index < spacingList.Count-1; ++index)
            {
                bayNo = bayNo + " " + (index + 1);
            }

            tgt = xdoc.Root.Descendants("Portal1Grids").FirstOrDefault(); //Portal 1 Grids

            tgt.Value = gridNo;

            tgt = xdoc.Root.Descendants("flyBraceGrids").FirstOrDefault(); //Purlin fly insert At grids

            tgt.Value = gridNo;

            tgt = xdoc.Root.Descendants("SideflyBays").FirstOrDefault(); //Girts side fly brace connection

            tgt.Value = gridNo;

            tgt = xdoc.Root.Descendants("bridgingbays1").FirstOrDefault(); // purlin bridging insert at bays

            tgt.Value = flyBraceAttrib;

            tgt = xdoc.Root.Descendants("bridgingbays1").FirstOrDefault(); //purlin bridging insert at bays

            tgt.Value = bayNo;

            tgt = xdoc.Root.Descendants("Sidebridgingbays1").FirstOrDefault(); // girts side bridging bays

            tgt.Value = bayNo;

            //TODO: Purlin Split Location - needs updating for different roof/wall

            //********************************************************************************
            // split locations
            //********************************************************************************

            string split = "";

            int xCount = spacingList.Count - 1;

            if (xCount == 1 || xCount == 2)
            {
                split = "0";
            }
            else if (xCount == 3 || xCount == 4)
            {
                split = "3";
            }
            else if (xCount == 5)
            {
                split = "3 4";
            }
            else if (xCount == 6)
            {
                split = "3 5";
            }
            else if (xCount == 7)
            {
                split = "3 5 6";
            }
            else if (xCount == 8)
            {
                split = "3 5 7";
            }
            else if (xCount == 9)
            {
                split = "3 5 6 8";
            }
            else 
            {
                split = "3 5 7 9";
            }

            tgt = xdoc.Root.Descendants("SplitGrids").FirstOrDefault(); //Purlin split grids

            tgt.Value = split;

            tgt = xdoc.Root.Descendants("SideSplitGrids").FirstOrDefault(); //Girt split

            tgt.Value = split;

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
            tgt = xdoc.Root.Descendants("CreateMidMullions").FirstOrDefault();
            tgt.Value = "1";
            tgt = xdoc.Root.Descendants("CreateBackMullions").FirstOrDefault();
            tgt.Value = "1";
            tgt = xdoc.Root.Descendants("CreateFrontMullions").FirstOrDefault();
            tgt.Value = "1";
            //}
            //else // Building has walls
            //{
            // Create Girts, endwall columns and cladding

            tgt = xdoc.Root.Descendants("CreateMidMullions").FirstOrDefault(); // Remove Mid mullion
            tgt.Value = "1";

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

            ////***********************************************************************
            //// Update roof/wall attributes for building layout

            //SetRoofWallLayoutAttributes(length, apex, width);

            ////********************************************************************************

            xdoc.Save(attribute);

            //********************************************************************************

        }

        private void SetRoofWallLayoutAttributes(double length, double apex, double width)
        {

            try
            {

                //***********************************************************************
                // Update roof attributes

                if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project Roof Clad Left.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                    UpdateAttributes("Project Roof Clad Right.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project Roof Clad Left_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                    UpdateAttributes("Project Roof Clad Right_Front Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                    Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                    Origin.X = 0;
                    Origin.Y = width / 2;
                    Origin.Z = apex + 500;

                    Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
                    FinishPoint.X = 2000;
                    FinishPoint.Y = width / 2;
                    FinishPoint.Z = apex + 500;

                    AddInformationNote(Origin, FinishPoint, "Remove Raker Angle");
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project Roof Clad Left_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                    UpdateAttributes("Project Roof Clad Right_Back Open.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");

                    Tekla.Structures.Geometry3d.Point Origin = new Tekla.Structures.Geometry3d.Point();
                    Origin.X = length;
                    Origin.Y = width / 2;
                    Origin.Z = apex + 500;

                    Tekla.Structures.Geometry3d.Point FinishPoint = new Tekla.Structures.Geometry3d.Point();
                    FinishPoint.X = length + 2000;
                    FinishPoint.Y = width / 2;
                    FinishPoint.Z = apex + 500;

                    AddInformationNote(Origin, FinishPoint, "Remove Raker Angle");
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project Roof Clad Left_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Left.CSB_Roof_Cladding.MainForm.xml");
                    UpdateAttributes("Project Roof Clad Right_Roof Only.CSB_Roof_Cladding.MainForm.xml", "Project Roof Cladding Right.CSB_Roof_Cladding.MainForm.xml");
                }

                //***********************************************************************
                // Update wall attributes

                // Endwalls settings

                // front-right corner
                if (btnFront.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project EW Clad Front Right_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project EW Clad Front Right_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Right.CSB_EndWall_Cladding.MainForm.xml");
                }

                // front-left corner
                if (btnFront.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project EW Clad Front Left_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project EW Clad Front Left_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Front Left.CSB_EndWall_Cladding.MainForm.xml");
                }

                // back-right corner
                if (btnRear.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project EW Clad Back Right_Right Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                }
                else if (btnRear.BackColor == System.Drawing.Color.Red && btnRight.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project EW Clad Back Right_Right Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Right.CSB_EndWall_Cladding.MainForm.xml");
                }

                // back-left corner
                if (btnRear.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project EW Clad Back Left_Left Closed.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                }
                else if (btnRear.BackColor == System.Drawing.Color.Red && btnLeft.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project EW Clad Back Left_Left Open.CSB_EndWall_Cladding.MainForm.xml", "Project End Wall Cladding Back Left.CSB_EndWall_Cladding.MainForm.xml");
                }

                // Sidewall settings

                if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project SW Clad Left_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                    UpdateAttributes("Project SW Clad Right_FrontBack Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                }
                else if (btnFront.BackColor == System.Drawing.Color.Red && btnRear.BackColor == System.Drawing.Color.White)
                {
                    UpdateAttributes("Project SW Clad Left_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                    UpdateAttributes("Project SW Clad Right_Back Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                }
                else if (btnFront.BackColor == System.Drawing.Color.White && btnRear.BackColor == System.Drawing.Color.Red)
                {
                    UpdateAttributes("Project SW Clad Left_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Left.CSB_SideWall_Cladding.MainForm.xml");
                    UpdateAttributes("Project SW Clad Right_Front Open.CSB_SideWall_Cladding.MainForm.xml", "Project Side Wall Cladding Right.CSB_SideWall_Cladding.MainForm.xml");
                }

                myHelper.LogFile("Roof Wall Layout Attributes");
            }
            catch (Exception e)
            {
                myHelper.LogFile("1005 - " + e.Message);
            }

            //********************************************************************************

            //TODO: Rolltop Ridge

            //TODO: Corro Roof

            //TODO: Corro Walls

            //TODO: 150 Girts/Purlins

            //TODO: change V-Ridge to suit pitch

            //TODO: update portals

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

                string xFile = @"T:\CSB_Program_Files\Documentation\Settings\" + Original;

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
            }
            catch (Exception e)
            {
                myHelper.LogFile("1006 - " + e.Message);
            }

        }

        private void CreateModel()
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
                Origin.Z = 0;

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
                newText += "        Column     : " + ProjectSales.ColumnType + "\r\n";
                newText += "        Truss        : " + ProjectSales.TrussType + "\r\n";
                newText += "        Purlin        : " + ProjectSales.RoofPurlin + "\r\n";
                newText += "        Girt           : " + ProjectSales.WallGirt + "\r\n";
                newText += "        Other        : " + ProjectSales.OtherFrameDetails + "\r\n";
                newText += "        Footings   : " + ProjectSales.Footings + " : Finish : " + ProjectSales.Finish + "\r\n";
                newText += "\r\n";
                newText += "        Notes      : " + "\r\n";
                newText += txtNote.Text;

                var tgt = xdoc.Root.Descendants("Phrase").FirstOrDefault();

                tgt.Value = newText;

                xdoc.Save(attribute);

                //********************************************************************************

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

            //int noBays = 5;

            //if (noBays == 1)
            //{
            //    tgt.Value = "1 2";
            //}
            //else if (noBays == 2)
            //+
            //{
            //    tgt.Value = "1 2 3";
            //}
            //else if (noBays == 3)
            //{
            //    tgt.Value = "1 2 3 4";
            //}


                

                //string modelPath = myModel.GetInfo().ModelPath;

                //string attribute = modelPath + @"\attributes\CSB_Project_Setup.CSB_Gable_Shed.MainForm.xml";

                //XDocument xdoc = XDocument.Load(attribute);

                //var tgt = xdoc.Root.Descendants("Height").FirstOrDefault();

                //tgt.Value = eave.ToString();

                //tgt = xdoc.Root.Descendants("Width").FirstOrDefault();

                //tgt.Value = eave.ToString();

                //string AttributeSettings = "";

                //AttributeSettings = "API Gable 40m Reg A";

                //string xFile = @"T:\CSB_TeklaSetup\" + AttributeSettings + ".CSB_Gable_Shed.MainForm.xml";

                //var xdocAttrib = XDocument.Load(xFile);

                //foreach (var childElement in xdocAttrib.Root.Elements())
                //{
                //    string a = childElement.Name.ToString();
                //    string c = childElement.Value.ToString();

                //    if (a != "SpacingBetweenBays" && a != "Millimeters")
                //    {

                //        var tgt2 = xdoc.Root.Descendants(a).FirstOrDefault();

                //        tgt2.Value = c;

                //    }

                //}

                //xdoc.Root.Descendants("SpacingBetweenBays").Remove();

                ////string temp = ", new XElement(" + "Millimeters" + "," + "5000" + ")";

                //string b = "9000";

                //XElement temp = new XElement("Millimeters", "5000");
                //XElement temp2 = new XElement("Millimeters", b);

                ////string b = temp + "," + temp2;

                //xdoc.Element("config")
                //    .Elements("Height").FirstOrDefault()
                //    .AddAfterSelf(new XElement("SpacingBetweenBays",
                //    temp
                //    ));

                //xdoc.Element("config")
                //    .Elements("SpacingBetweenBays")
                //    .Elements("Millimeters").LastOrDefault()
                //    .AddAfterSelf(temp2);




                //xdoc.Element("config")
                //    .Elements("Height").FirstOrDefault()
                //    .AddAfterSelf(new XElement("SpacingBetweenBays",
                //    new XElement("Millimeters", "5000"),
                //    new XElement("Millimeters", "7000"),
                //    new XElement("Millimeters", "8000")
                //    ));

                //XNode tgt2 = (XNode)xdoc.Root.Descendants("SpacingBetweenBays");

                // tgt2.AddAfterSelf("Millimeters","2000");

                //xdoc.Root.Descendants("SpacingBetweenBays").("Millimeters", 2000);

                //tgt.Value = eave.ToString();

                //xdoc.Save(attribute);


            }

        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Settings temp = new Settings();
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
                this.errorProvider8.SetError(txtBaySize, " Bay sizes do not equal length");
                Globals.checkError = 1;
                return;
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

      
    }

#endregion

#endregion        

    }
