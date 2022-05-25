using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSB
{
    class Helper
    {
        public bool IsNumeric(string temp)
        {
            bool Result = false;

            decimal xNumeric = 0;

            bool canConvert = decimal.TryParse(temp, out xNumeric);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;
        }

        public bool IsLong(string temp)
        {
            bool Result = false;

            long xLong = 0;

            bool canConvert = long.TryParse(temp, out xLong);
            if (canConvert == true)
                Result = true;
            else
                Result = false;

            return Result;
        }

        public string TemplateModel()
        {

            var xdoc = XDocument.Load(Globals.Config());

          string  Result = xdoc.Root.Descendants("TemplateModel").FirstOrDefault().Value;

            return Result;
        }

        public string ProjectFolder()
        {

            var xdoc = XDocument.Load(Globals.Config());

            string Result = xdoc.Root.Descendants("Folder").FirstOrDefault().Value;

#if DEBUG
            Result = @"C:\Development\Models\";
#endif

            return Result;
        }

        public string ExportFolder()
        {

            var xdoc = XDocument.Load(Globals.Config());

            string Result = xdoc.Root.Descendants("ExportFolder").FirstOrDefault().Value;

#if DEBUG
            Result = @"C:\Development\Exports\";
#endif

            return Result;
        }

        public string ShareMacro()
        {

            var xdoc = XDocument.Load(Globals.Config());

            string Result = xdoc.Root.Descendants("ShareMacro").FirstOrDefault().Value;

            return Result;
        }


        public bool checkDistanceList(string distanceList)
        {
            bool Result = true;

            try
            {
                List<double> distanceListLis = new List<double>();
                double distance = 0;
                distanceListLis.Add(distance);
                string[] list = distanceList.Split(' ');
                {
                    for (int i = 0; i < list.Length; i++)
                    {
                        string value = list[i];
                        if (value.Contains("*"))
                        {
                            string[] y = value.Split('*');
                            for (int j = 0; j < int.Parse(y[0]); j++)
                            {
                                distance = distance + double.Parse(y[1], CultureInfo.InvariantCulture);
                                distanceListLis.Add(distance);
                            }
                        }
                        else
                        {
                            distance = distance + double.Parse(list[i], CultureInfo.InvariantCulture);
                            if (distance > 100)
                            {
                                distanceListLis.Add(distance);
                            }
                        }
                    }
                }

            }
            catch
            {
                Result = false;
            }
            return Result;
        }

        public List<double> getDistanceList(string distanceList)
        {
            //bool Result = true;

            List<double> distanceListLis = new List<double>();
            
            double distance = 0;
            distanceListLis.Add(distance);
            string[] list = distanceList.Split(' ');

            try
            {
                for (int i = 0; i < list.Length; i++)
                {
                    string value = list[i];
                    if (value.Contains("*"))
                    {
                        string[] y = value.Split('*');
                        for (int j = 0; j < int.Parse(y[0]); j++)
                        {
                            distance = distance + double.Parse(y[1], CultureInfo.InvariantCulture);
                            distanceListLis.Add(distance);
                        }
                    }
                    else
                    {
                        distance = distance + double.Parse(list[i], CultureInfo.InvariantCulture);
                        if (distance > 100)
                        {
                            distanceListLis.Add(distance);
                        }
                    }
                }
            }
            catch
            {
                //Result = false;
            }

            //return Result;
            return distanceListLis;
        }

        public List<double> getSpacingList(string distanceList)
        {
            List<double> distanceListLis = new List<double>();
            double distance = 0;
            distanceListLis.Add(distance);
            string[] list = distanceList.Split(' ');

            try
            {

                {
                    for (int i = 0; i < list.Length; i++)
                    {
                        string value = list[i];
                        if (value.Contains("*"))
                        {
                            string[] y = value.Split('*');
                            for (int j = 0; j < int.Parse(y[0]); j++)
                            {
                                distanceListLis.Add(double.Parse(y[1], CultureInfo.InvariantCulture));
                            }
                        }
                        else
                        {
                            double currentdistance = double.Parse(list[i], CultureInfo.InvariantCulture);
                            if (currentdistance > 100)
                            {
                                distanceListLis.Add(currentdistance);
                            }
                        }
                    }
                }

            }
            catch
            {

            }

            return distanceListLis;
        }

        public bool ReadSalesInput(string xFile, salesLib Sales)
        {
            bool Result = true;

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(xFile);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            try
            {
                //for (int i = 1; i <= rowCount; i++)
                //{
                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    //if (j == 1)
                    //    Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[1, j] != null && xlRange.Cells[1, j].Value2 != null && xlRange.Cells[2, j].Value2 != null)
                    {
                        if (xlRange.Cells[1, j].Value2 == "Barge")
                        {
                            Sales.Barge = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Bays")
                        {
                            try
                            {
                                Sales.Bays = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.Bays = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "BaySize")
                        {
                            try
                            {
                                Sales.BaySize = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.BaySize = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "ClearSheetRoof")
                        {
                            Sales.ClearSheetRoof = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "ClearSheetWall")
                        {
                            Sales.ClearSheetWall = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "ColumnType")
                        {
                            Sales.ColumnType = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "CompanyName")
                        {
                            Sales.CompanyName = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Corner")
                        {
                            Sales.Corner = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "CustomerName")
                        {
                            Sales.CustomerName = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Downpipe")
                        {
                            Sales.Downpipe = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "EndWalls")
                        {
                            try
                            {
                                Sales.EndWalls = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.EndWalls = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Finish")
                        {
                            Sales.Finish = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "FlashingRidge")
                        {
                            Sales.FlashingRidge = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Footings")
                        {
                            Sales.Footings = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "FrameSpan")
                        {
                            try
                            {
                                Sales.FrameSpan = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.FrameSpan = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "GutterColour")
                        {
                            Sales.GutterColour = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "GutterType")
                        {
                            Sales.GutterType = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Height")
                        {
                            try
                            {
                                Sales.Height = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.Height = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "JobNo")
                        {
                            Sales.JobNo = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Length")
                        {
                            try
                            {
                                Sales.Length = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.Length = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "OtherFrameDetails")
                        {
                            Sales.OtherFrameDetails = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "QuoteVer")
                        {
                            try
                            {
                                Sales.QuoteVer = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.QuoteVer = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "RoofColour")
                        {
                            Sales.RoofColour = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "RoofMaterial")
                        {
                            Sales.RoofMaterial = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "RoofPitch")
                        {
                            try
                            {
                                Sales.RoofPitch = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.RoofPitch = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "RoofPurlin")
                        {
                            Sales.RoofPurlin = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "RoofType")
                        {
                            Sales.RoofType = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "SideWals")
                        {
                            try
                            {
                                Sales.SideWals = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.SideWals = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Suburb")
                        {
                            Sales.Suburb = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Totwalls")
                        {
                            try
                            {
                                Sales.Totwalls = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.Totwalls = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "TrussType")
                        {
                            Sales.TrussType = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "WallColour")
                        {
                            Sales.WallColour = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "WallGirt") // Catch old jobs before endwall and sidewall
                        {
                            Sales.WallGirtSide = xlRange.Cells[2, j].Value2;
                            Sales.WallGirtEnd = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "WallGirt")
                        {
                            Sales.WallGirtEnd = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "WallGirtEnd")
                        {
                            Sales.WallGirtEnd = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "WallMaterial")
                        {
                            Sales.WallMaterial = xlRange.Cells[2, j].Value2;
                        }
                        else if (xlRange.Cells[1, j].Value2 == "Width")
                        {
                            try
                            {
                                Sales.Width = xlRange.Cells[2, j].Value2;
                            }
                            catch
                            {
                                var temp = xlRange.Cells[2, j].Value2;
                                Sales.Width = temp.ToString("0.###");
                            }
                        }
                        else if (xlRange.Cells[1, j].Value2 == "ProjectDetails")
                        {
                            Sales.ProjectDetails = xlRange.Cells[2, j].Value2;
                        }
                        else
                        {
                            try
                            {
                                LogFile("1103 - New Data for " + xlRange.Cells[1, j].Value2);
                                MessageBox.Show("New Data for " + xlRange.Cells[1, j].Value2);
                            }
                            catch
                            {

                            }
                        };

                    };

                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

            }
            catch (Exception e)
            {
                LogFile("1102 - " + e.Message);
                LogFile("1102 - File - " + xFile);
                Result = false;
                return Result;
            }

            return Result;
        }

        public  bool CheckFolder(string xPath)
        {
            bool Result = false;

            if (Directory.Exists(xPath))
            {
                Result = true;
            }

            return Result;
        }

        public void LogFile(string temp)
        {
            string xMonth = DateTime.Today.Month.ToString();

            if (xMonth.Length == 1)
            {
                xMonth = "0" + xMonth;
            }

            string xDay = DateTime.Today.Day.ToString();

            if (xDay.Length == 1)
            {
                xDay = "0" + xDay;
            }

            string xtemp = DateTime.Today.Year.ToString()+ xMonth + xDay;

            StringBuilder sb = new StringBuilder();
            sb.Append(DateTime.Now + " - " + temp + "\r\n");
            File.AppendAllText(@"T:\CSB_Program_Files\Documentation\Log_Files\log_" + xtemp +".txt", sb.ToString());
            sb.Clear();
        }

    }    

    static class Globals
    {
        // global int
        //public static int counter;

        // global function
        public static string Config()
        {
            return @"T:\CSB_Program_Files\Documentation\Settings\CSB_Project_Setup.xml";
        }

        // global int using get/set
        static int _checkError = 0;
        public static int checkError
        {
            set { _checkError = value; }
            get { return _checkError; }
        }
    }

    public class salesLib
    {
        Helper MyHelper = new Helper();

        private string mJobNo;

        public string JobNo
        {
            get { return mJobNo; }
            set { mJobNo = value; }
        }

        private string mProjectNo;

        public string ProjectNo
        {
            get
            {
                for (int index = 0; index < mJobNo.Length; ++index)
                {
                    string temp = mJobNo.Substring(0, mJobNo.Length - index);

                    if (MyHelper.IsNumeric(temp) == true)
                    {
                        mProjectNo = temp;
                        mSalesRep = mJobNo.Substring(mJobNo.Length - index);
                        break;
                    }
                }

                return mProjectNo;
            }
        }

        private string mSalesRep;

        public string SalesRep
        {
            get { return mSalesRep; }
            set { mSalesRep = value; }
        }

        private string mProjectName;

        public string ProjectName
        {
            get {

                if (mCompanyName == null)
                {
                    mProjectName =  mCustomerName;
                }
                else
                {
                    mProjectName = mCompanyName + " - " + mCustomerName;
                }
                return mProjectName;
            }
        }

        private string mQuoteVer;

        public string QuoteVer
        {
            get { return mQuoteVer; }
            set { mQuoteVer = value; }
        }

        private string mCompanyName;

        public string CompanyName
        {
            get { return mCompanyName; }
            set { mCompanyName = value; }
        }

        private string mCustomerName;

        public string CustomerName
        {
            get { return mCustomerName; }
            set { mCustomerName = value; }
        }

        private string mSuburb;

        public string Suburb
        {
            get { return mSuburb; }
            set { mSuburb = value; }
        }

        private string mLength;

        public string Length
        {
            get { return mLength; }
            set { mLength = value; }
        }

        private string mWidth;

        public string Width
        {
            get { return mWidth; }
            set { mWidth = value; }
        }

        private string mHeight;

        public string Height
        {
            get { return mHeight; }
            set { mHeight = value; }
        }

        private string mRoofType;

        public string RoofType
        {
            get { return mRoofType; }
            set { mRoofType = value; }
        }

        private string mRoofPitch;

        public string RoofPitch
        {
            get { return mRoofPitch; }
            set { mRoofPitch = value; }
        }

        private string mTotwalls;

        public string Totwalls
        {
            get { return mTotwalls; }
            set { mTotwalls = value; }
        }

        private string mSideWals;

        public string SideWals
        {
            get { return mSideWals; }
            set { mSideWals = value; }
        }

        private string mEndWalls;

        public string EndWalls
        {
            get { return mEndWalls; }
            set { mEndWalls = value; }
        }

        private string mRoofMaterial;
        
        public string RoofMaterial
        {
            get { return mRoofMaterial; }
            set { mRoofMaterial = value; }
        }

        private string mRoofColour;

        public string RoofColour
        {
            get { return mRoofColour; }
            set { mRoofColour = value; }
        }

        private string mClearSheetRoof;

        public string ClearSheetRoof
        {
            get { return mClearSheetRoof; }
            set { mClearSheetRoof = value; }
        }

        private string mWallMaterial;

        public string WallMaterial
        {
            get { return mWallMaterial; }
            set { mWallMaterial = value; }
        }

        private string mWallColour;

        public string WallColour
        {
            get { return mWallColour; }
            set { mWallColour = value; }
        }

        private string mClearSheetWall;

        public string ClearSheetWall
        {
            get { return mClearSheetWall; }
            set { mClearSheetWall = value; }
        }

        private string mFlashingRidge;

        public string FlashingRidge
        {
            get { return mFlashingRidge; }
            set { mFlashingRidge = value; }
        }

        private string mBarge;

        public string Barge
        {
            get { return mBarge; }
            set { mBarge = value; }
        }

        private string mCorner;

        public string Corner
        {
            get { return mCorner; }
            set { mCorner = value; }
        }

        private string mGutterColour;

        public string GutterColour
        {
            get { return mGutterColour; }
            set { mGutterColour = value; }
        }

        private string mDownpipe;

        public string Downpipe
        {
            get { return mDownpipe; }
            set { mDownpipe = value; }
        }

        private string mColumnType;

        public string ColumnType
        {
            get { return mColumnType; }
            set { mColumnType = value; }
        }

        private string mTrussType;

        public string TrussType
        {
            get { return mTrussType; }
            set { mTrussType = value; }
        }

        private string mRoofPurlin;

        public string RoofPurlin
        {
            get { return mRoofPurlin; }
            set { mRoofPurlin = value; }
        }

        private string mWallGirtSide;

        public string WallGirtSide
        {
            get { return mWallGirtSide; }
            set { mWallGirtSide = value; }
        }

        private string mWallGirtEnd;

        public string WallGirtEnd
        {
            get { return mWallGirtEnd; }
            set { mWallGirtEnd = value; }
        }

        private string mOtherFrameDetails;

        public string OtherFrameDetails
        {
            get { return mOtherFrameDetails; }
            set { mOtherFrameDetails = value; }
        }

        private string mGutterType;

        public string GutterType
        {
            get { return mGutterType; }
            set { mGutterType = value; }
        }

        private string mFrameSpan;

        public string FrameSpan
        {
            get { return mFrameSpan; }
            set { mFrameSpan = value; }
        }

        private string mBays;

        public string Bays
        {
            get { return mBays; }
            set { mBays = value; }
        }

        private string mBaySize;

        public string BaySize
        {
            get { return mBaySize; }
            set { mBaySize = value; }
        }

        private string mBayString;

        public string BayString
        {
            get 
            {
                mBayString = mBays + "*" + (double)decimal.Parse(mBaySize.Trim())*1000;
                return mBayString; 
            }
        }

        private string mFootings;

        public string Footings
        {
            get { return mFootings; }
            set { mFootings = value; }
        }

        private string mFinish;

        public string Finish
        {
            get { return mFinish; }
            set { mFinish = value; }
        }

        private string mProjectDetails;

        public string ProjectDetails
        {
            get { return mProjectDetails; }
            set { mProjectDetails = value; }
        }
    }

    public class ProjectLib
    {
        private string mNumber;

        public string Number
        {
            get { return mNumber; }
            set { mNumber = value; }
        }

        private string mClient;

        public string Client
        {
            get { return mClient.ToUpper(); }
            set { mClient = value; }
        }

        private string mDescription;

        public string Description
        {
            get { return mDescription.ToUpper(); }
            set { mDescription = value; }
        }

        private string mAddress;

        public string Address
        {
            get { return mAddress.ToUpper(); }
            set { mAddress = value; }
        }

        private string mLength;

        public string Length
        {
            get { return mLength; }
            set { mLength = value; }
        }

        private string mWidth;

        public string Width
        {
            get { return mWidth; }
            set { mWidth = value; }
        }

        private string mEave;

        public string Eave
        {
            get { return mEave; }
            set { mEave = value; }
        }

        private string mFolder;

        public string Folder
        {
            get { return mFolder; }
            set { mFolder = value; }
        }

        private string mTemplateModel;

        public string TemplateModel
        {
            get { return mTemplateModel; }
            set { mTemplateModel = value; }
        }

        private string mUser;

        public string User
        {
            get
            {
                mUser = Environment.UserName;
                return mUser;
            }
        }

        private string mModelName;

        public string ModelName
        {
            get
            {

                char[] separators = new char[] { ' ', ';', ',', '/', '\r', '\t', '\n' };

                string s = mClient.ToUpper();
                string[] temp2 = s.Split(separators, StringSplitOptions.RemoveEmptyEntries);
                s = String.Join(" ", temp2);
                // shortened name
                mModelName = mNumber; // + @"-" + s;

                string xLength = "19875a"; // Indoor Jumping Arena_Frank Demaiio

                if (mModelName.Length > xLength.Length)
                {
                    mModelName = mModelName.Substring(0, xLength.Length);
                }

                return mModelName;
            }
        }

        //private string mProjectFolder;

        //public string ProjectFolder
        //{
        //    get
        //    {
        //        mProjectFolder = mModelName;
        //        return mProjectFolder;
        //    }
        //}

        private string mTeklaDesc;

        public string TeklaDesc
        {
            get
            {
                mTeklaDesc = mLength + @"m x " + mWidth + @"m x " + mEave + @"m - " + mDescription.ToUpper();
                return mTeklaDesc;
            }
        }

    }

}
