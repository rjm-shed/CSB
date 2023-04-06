using Microsoft.Office.Interop.Excel;
using RenderCommand;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CSB
{
    internal class ApiModel
    {

        [Description("Eave Heigth")]
        public string Height { get; set; }

        [Description("Bay Spacing")]
        public string SpacingBetweenBays { get; set; }//millimeters

        [Description("Width")]
        public string Width { get; set; }

        [Description("")]
        public string HeightType { get; set; }

        // 0 Center line of columns
        // 1 Outer line of columns
        // 2 Outer girt line
        [Description("Width Type")]
        public string WidthType { get; set; }

        [Description("Create Girds")]
        public int CreateGrids { get; set; } // tick

        [Description("Portal 3")]
        public string PortalAtt3 { get; set; }

        [Description("Portal 3-Portal 3 Grids")]
        public string Portal3Grids { get; set; }

        [Description("Portal 2-Portal 2 Grids")]
        public string Portal2Grids { get; set; }

        [Description("Portal 1-Portal 1 Grids")]
        public string Portal1Grids { get; set; }

        [Description("Portal 2")]
        public string PortalAtt2 { get; set; }

        [Description("Portal 1")]
        public string PortalAtt1 { get; set; }

        [Description("Slope angle")]
        public string SlopeAngle { get; set; }

        [Description("Purlin Split Grids")]
        public string SplitGrids { get; set; }

        [Description("Ridge Purlins-Number")]
        public string RidgeNumber { get; set; }

        [Description("Mid Purlins-Settings")]
        public string MidPurAttri { get; set; }

        [Description("Ridge Purlins-Settings")]
        public string RidgePurAttri { get; set; }

        [Description("Eave Purlin")]
        public string EavePurProfile { get; set; }

        [Description("Mid Purlin")]
        public string MidPurProfile { get; set; }

        [Description("Ridge Purlin")]
        public string RidgePurProfile { get; set; }

        [Description("Eave Purlins-Settings")]
        public string EavePurAttri { get; set; }

        [Description("Ridge Purlins-Cleat Attributes")]
        public string RidgeCleatAttri { get; set; }

        [Description("Mid Purlins-Cleat Attributes")]
        public string MidCleatAttri { get; set; }

        [Description("Eave Purlins-Cleat Attributes")]
        public string EaveCleatAttri { get; set; }

        [Description("Fly Brace Connection-Fly attribute 1")]
        public string RidgeflybracAttri { get; set; }

        [Description("Fly Brace Connection-Fly attribute 2")]
        public string MidflybracAttri { get; set; }

        [Description("Fly Brace Connection-Fly attribute 3")]
        public string EaveflybracAttri { get; set; }

        [Description("Eave Purlins-Number")]
        public string EaveNumber { get; set; }

        [Description("")]
        public string Purlin_spacing_List { get; set; }

        [Description("Purlin Distribution-Overlap Distance")]
        public string Purlin_OverLabDis { get; set; }

        [Description("Purlin Distribution-Distance From Apex")]
        public string dis_from_Apex { get; set; }

        [Description("")]
        public string bridgingAttri2 { get; set; }

        [Description("")]
        public string bridgingAttri1 { get; set; }

        [Description("")]
        public string bridgingbays2 { get; set; }

        [Description("")]
        public string bridgingbays1 { get; set; }

        [Description("Fly Brace Connection-Insert AT purlin Rows")]
        public string flyBracePurRows { get; set; }

        [Description("Fly Brace Connection-Double Brace At Grids")]
        public string flyBraceDoubGrids { get; set; }

        [Description("Fly Brace Connection-Insert At Grids")]
        public string flyBraceGrids { get; set; }

        [Description("Back Mullions-mullions Selection")]
        public string BackMullionsAtt { get; set; }

        [Description("Back Mullions-Spacing List")]
        public string BackMullionSpacing { get; set; }//millimeters?

        [Description("Front Mullions-Mullions Selection")]
        public string FrontMullionsAtt { get; set; }

        [Description("Front Mullions-Spacing List")]
        public string FrontMullionSpacing { get; set; }

        [Description("Mullion 2-Settings")]
        public string MullionAtt2 { get; set; }

        [Description("Mullion 1-Settings")]
        public string MullionAtt1 { get; set; }

        [Description("mullion 3")]
        public string MullionProfile3 { get; set; }

        [Description("Mullion 2")]
        public string MullionProfile2 { get; set; }

        [Description("Mullion 1")]
        public string MullionProfile1 { get; set; }

        [Description("Mullion 3-Settings")]
        public string MullionAtt3 { get; set; }

        [Description("Mullion 1-Base Attributes")]
        public string MullionBase1 { get; set; }

        [Description("Mullion 2-Base Attributes")]
        public string MullionBase2 { get; set; }

        [Description("Mullion 3-Base Attributes")]
        public string MullionBase3 { get; set; }

        [Description("Mullion 1-Truss Connection Attributes RHS")]
        public string MullionTruss1 { get; set; }

        [Description("Mullion 2-Truss Connection Attributes RHS")]
        public string MullionTruss2 { get; set; }

        [Description("Mullion 3-Truss Connection Attributes RHS")]
        public string MullionTruss3 { get; set; }

        [Description("Side Bridging Connection-Attribute 2")]
        public string SidebridgingAttri2 { get; set; }

        [Description("Side Bridging Connection-Attribute 1")]
        public string SidebridgingAttri1 { get; set; }

        [Description("Side Bridging Connection-Insert At Bays 2")]
        public string Sidebridgingbays2 { get; set; }

        [Description("Side Bridging Connection-Insert At Bays 1")]
        public string Sidebridgingbays1 { get; set; }

        // API Cleat
        // API End Wall Cleat
        // standard
        [Description("Girt-Corner Cleat")]
        public string CornerCleatAtt { get; set; }

        [Description("Girt Distribution-Max Distance")]
        public string MaxDist { get; set; }

        [Description("Left Side Girt-Distance From Bottom")]
        public string BotDist { get; set; }

        [Description("Side Girt-Settings")]
        public string GirtAtt { get; set; }

        [Description("Left Side Girt")]
        public string GirtProfile { get; set; }

        [Description("End Girt-Settings")]
        public string EndGirtAtt { get; set; }

        [Description("Front End Girt")]
        public string EndGirtProfile { get; set; }

        [Description("Side Girt-Cleat Attributes")]
        public string SideCleatAttri { get; set; }

        [Description("End Girt-Cleat Attributes")]
        public string EndGirtCleatAtt { get; set; }

        [Description("Side Girt-Fly Brace Attributes")]
        public string SideflybracAttri { get; set; }

        [Description("End Girt-Fly Brace Attributes")]
        public string EndGirtflybracAtt { get; set; }

        [Description("Side Girt-Overlap")]
        public string SideGirtOverlap { get; set; }

        [Description("End Girt-Overlap")]
        public string EndGirtOverlap { get; set; }

        [Description("")]
        public string EavePurlinProfile { get; set; }// ?????

        [Description("")]
        public string PurlinAtt { get; set; }// ?????

        [Description("")]
        public string PurlinProfile { get; set; }// ?????

        [Description("End Bridging Connection-Attribute 1")]
        public string EndbridgingAttri2 { get; set; }

        [Description("End Bridging Connection-Attribute 1")]
        public string EndbridgingAttri1 { get; set; }

        [Description("End Bridging Connection-Insert At Bays 2")]
        public string Endbridgingbays2 { get; set; }

        [Description("End Bridging Connection-Insert At Bays 1")]
        public string Endbridgingbays1 { get; set; }

        [Description("Fascia Girt-Settings")]
        public string FasciaAtt { get; set; }

        [Description("Fascia Girt")]
        public string FasciaProfile { get; set; }

        [Description("End Fly Brace Connection-Insert At girt Rows")]
        public string EndflyRows { get; set; }

        [Description("End Fly Brace Connection-Double Brace At Grids")]
        public string EndflyDoubBays { get; set; }

        [Description("End Fly Brace Connection-Insert At Columns")]
        public string EndflyBays { get; set; }

        [Description("Side Fly Brace Connection-Insert At Grit Rows")]
        public string SideflyRows { get; set; }

        [Description("Side Fly Brace Connection-Double Brace At Grids")]
        public string SideflyDoubBays { get; set; }

        [Description("Side Fly Brace Connection-Insert At Grids")]
        public string SideflyBays { get; set; }

        [Description("Back Girts")]
        public int BackGirt { get; set; }// tick

        [Description("Front Girts")]
        public int FrontGirt { get; set; }// tick

        [Description("Right Girts")]
        public int RightGirt { get; set; }// tick

        [Description("Left Girts")]
        public int LeftGirt { get; set; }// tick

        [Description("Roof Bracing attributes")]
        public string CB { get; set; }

        [Description("Left Wall bracing attribute")]
        public string WB { get; set; }

        [Description("Roof Bracing bays")]
        public string CrossBracing { get; set; }

        [Description("Right Wall Bracing bays")]
        public string WallBracing2 { get; set; }

        [Description("Left Wall Bracing bays")]
        public string WallBracing { get; set; }

        [Description("Purlin Distribution-Distance From Eave")]
        public string dis_from_Eave { get; set; }

        [Description("Mid Mullions-Insert at grids Back")]
        public string MidMullionGrids2 { get; set; }

        [Description("Mid Mullions-Mullions selection")]
        public string MidMullionsAtt { get; set; }

        [Description("Mid Mullions-Spacing List")]
        public string MidMullionSpacing { get; set; }// millimeters

        [Description("Fly Brace Connection-Fly Brace attributes per Row")]
        public string FlyGridAtt { get; set; }

        [Description("Purlin Distribution-Distance From Eave")]
        public string ApexSpacing { get; set; }

        [Description("Mid Purlins-Max Spacing")]
        public string PurlinMaxSpacing { get; set; }

        [Description("Eave Purlin-Spacing")]
        public string EaveSpacing { get; set; }

        [Description("")]
        public string MidMullionSide { get; set; }

        [Description("")]
        public string RightGTBays { get; set; }

        [Description("")]
        public string LeftGTBays { get; set; }

        [Description("")]
        public string RightGTAtt { get; set; }

        [Description("")]
        public string LeftGTAtt { get; set; }

        [Description("")]
        public string GTColumnDist { get; set; }

        [Description("Right Wall bracing attribute")]
        public string WBR { get; set; }

        [Description("Mid Mullions-Number of Mullions")]
        public string NoOfMidMullions { get; set; }

        // Spacing
        // Numbber of mullions
        [Description("Mid Mullions-Number/Spacing")]
        public string MidMullionsOption { get; set; }

        [Description("Back Mullions-Number of Mullions")]
        public string NoOfBackMullions { get; set; }

        [Description("")]
        public string BackMullionsOption { get; set; }

        [Description("Front Mullions-Number of Mullions")]
        public string NoOfFrontMullions { get; set; }

        [Description("")]
        public string FrontMullionsOption { get; set; }

        [Description("Girts Distribution-Distance From Top")]
        public string TopDist { get; set; }

        [Description("Mid mullions-Front Create")]
        public int CreateMidMullions { get; set; }// Yes No

        [Description("Back mullions-Create")]
        public int CreateBackMullions { get; set; }// Yes No

        [Description("Front mullions-Create")]
        public int CreateFrontMullions { get; set; }// Yes No

        [Description("")]
        public string GridColor { get; set; }

        [Description("")]
        public string RightRCAtt { get; set; }

        [Description("")]
        public string LeftRCAtt { get; set; }

        [Description("")]
        public string CreateRC { get; set; }

        [Description("")]
        public string RightSWCAtt { get; set; }

        [Description("")]
        public string LeftSWCAtt { get; set; }

        [Description("")]
        public string CreateRightSWC { get; set; }

        [Description("")]
        public string CreateLeftSWC { get; set; }

        [Description("Back Left End Wall Cladding Attribute")]
        public string BackEWCAtt { get; set; }

        [Description("Front Left End Wall Cladding Attribute")]
        public string FrontEWCAtt { get; set; }

        [Description("")]
        public string CreateBackEWC { get; set; }

        [Description("")]
        public string CreateFrontEWC { get; set; }

        [Description("")]
        public string CreateSlab { get; set; }

        [Description("")]
        public string SlabAttributeFile { get; set; }

        [Description("")]
        public string SlabThickness { get; set; }

        [Description("")]
        public string SlabVlOffset { get; set; }

        [Description("Back Right End Wall Cladding Attribute")]
        public string BackEWCAtt2 { get; set; }

        [Description("Front Right End Wall Cladding Attribute")]
        public string FrontEWCAtt2 { get; set; }

        [Description("Left Column Offsets")]
        public string LeftOffsetList { get; set; }

        [Description("Right Column Offsets")]
        public string RightOffsetList { get; set; }

        [Description("Right Eave Offset")]
        public string RightEaveOffset { get; set; }

        [Description("Left Eave Offset")]
        public string LeftEaveOffset { get; set; }

        [Description("Bridging Connection-Apex attribute-2")]
        public string ApexBridgingAtt2 { get; set; }

        [Description("Bridging Connection-Apex attribute-1")]
        public string ApexBridgingAtt { get; set; }

        [Description("Delete right girts at bays")]
        public string RGDelete { get; set; }

        [Description("Delete left girts at bays")]
        public string LGDelete { get; set; }

        [Description("Mid Mullion Connection")]
        public string MullionTrussMid { get; set; }

        [Description("Left Columns Level")]
        public string LeftColumnsLevel { get; set; }

        [Description("Right Columns Level")]
        public string RightColumnsLevel { get; set; }

        [Description("Portal 1-Rafter Grids")]
        public string RafterGrids { get; set; }

        //Blank
        //API 12m RHS Portal A
        //API 12m UB Portal A
        //API 21m UB Portal A
        [Description("Portal 1-Rafter Portal")]
        public string RafterAtt { get; set; }

        // blank
        // Yes
        // No
        [Description("Add UDA")]
        public int AddUDA { get; set; }

        [Description("Dropper-Front Dropper")]
        public int FrontInfill { get; set; }// Yes/no

        [Description("Dropper-Front Bottom Level")]
        public string DropperElevation { get; set; }

        [Description("Dropper-Back Dropper")]
        public int BackInfill { get; set; }// Yes/No

        [Description("Dropper-Dropper to Top Chord Connection")]
        public string TopConnAtt { get; set; }

        [Description("Dropper-Dropper to Bottom Chord Connection")]
        public int BotConnAtt { get; set; }

        [Description("Mid mullions-Back Create")]
        public string CreateMidMullions2 { get; set; } // Yes No

        [Description("Mid Mullions-Insert at grids Front")]
        public string MidMullionGrids { get; set; }

        [Description("Dropper-Back Bottom Level")]
        public string DropperElevation2 { get; set; }

        [Description("Right Side Girt")]
        public string GirtProfile2 { get; set; }

        [Description("Back End Girt")]
        public string EndGirtProfile2 { get; set; }

        [Description("Right Side Girt-Distance From Bottom")]
        public string BotDist2 { get; set; }

        [Description("Mullion 3-Truss Connection Attributes RHS")]
        public string MullionTruss3RHS { get; set; }

        [Description("Mullion 2-Truss Connection Attributes RHS")]
        public string MullionTruss2RHS { get; set; }

        [Description("Mullion 1-Truss Connection Attributes RHS")]
        public string MullionTruss1RHS { get; set; }

        [Description("")]
        public string RightGTColumns { get; set; }

        [Description("")]
        public string LeftGTColumns { get; set; }

        [Description("Dropper-Dropper Cleat")]
        public string EndGirtAttDropper { get; set; }

        [Description("Back End Girt-Distance From Bottom")]
        public string BotDist4 { get; set; }

        [Description("Front End Girt-Distance From Bottom")]
        public string BotDist3 { get; set; }

        [Description("Back Girt Split")]
        public string EndSplitGrids2 { get; set; }

        [Description("Right Girt Split")]
        public string SideSplitGrids2 { get; set; }

        [Description("Left Girt split")]
        public string SideSplitGrids { get; set; }

        [Description("Front Girt Split")]
        public string EndSplitGrids { get; set; }

    }
}
