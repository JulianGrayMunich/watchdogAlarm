using System.Configuration;

using databaseAPI;

using GNAgeneraltools;

using GNAspreadsheettools;

using OfficeOpenXml;

namespace projectPerformance
{

    class Program
    {
        static void Main()
        {
#pragma warning disable CS0162
#pragma warning disable CS8600
#pragma warning disable CS8601
#pragma warning disable CS8604
#pragma warning disable IDE0059




            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();

            //==== System config variables

            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];
            string strRootFolder = ConfigurationManager.AppSettings["SystemStatusFolder"];
            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];

            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];
            string strContractTitle = ConfigurationManager.AppSettings["ContractTitle"];
            string strReportType = "Alarm";

            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strAlarmWorksheet = ConfigurationManager.AppSettings["AlarmWorksheet"];

            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];
            string strFirstDataCol = ConfigurationManager.AppSettings["FirstDataCol"];

            int iFirstDataCol = Convert.ToInt16(strFirstDataCol);
            int iFirstDataRow = Convert.ToInt16(strFirstDataRow);


            string strSendSMS = ConfigurationManager.AppSettings["SendSMS"];
            string strSMSTitle = ConfigurationManager.AppSettings["SMSTitle"];
            string strRecipientPhone1 = ConfigurationManager.AppSettings["RecipientPhone1"];
            string strRecipientPhone2 = ConfigurationManager.AppSettings["RecipientPhone2"];
            string strRecipientPhone3 = ConfigurationManager.AppSettings["RecipientPhone3"];
            string strRecipientPhone4 = ConfigurationManager.AppSettings["RecipientPhone4"];
            string strJAGstatus = ConfigurationManager.AppSettings["JAGstatus"];

            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strAddAttachment = ConfigurationManager.AppSettings["AddAttachment"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strEmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];

            string strTimeBlockStartLocal = "";
            string strTimeBlockEndLocal = "";
            string strTimeBlockStartUTC = "";
            string strTimeBlockEndUTC = "";


            string strTimeBlockType = "Schedule";
            string strBlockSizeHrs = ConfigurationManager.AppSettings["BlockSizeHrs"];
            double dblBlockSizeHrs = Convert.ToDouble(strBlockSizeHrs);

            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;
            string strExcelWorkingFileFullPath = strExcelPath + strProjectTitle + "_" + strReportType + "_" + DateTime.Now.ToString("yyyyMMdd_HHmm") + ".xlsx";

            string[,] strSensorID = new string[5000, 2];
            string[,] strPointDeltas = new string[5000, 2];

            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //====[changes]===========================================================================================
            // 20231123 Created first draft
            //====[ Main Program ]====================================================================================


            gnaT.WelcomeMessage("watchdogAlarm 20231123");

            string strSoftwareLicenseTag = "WDGALM";
            gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);

            //==== Environment check

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("   Check Existance of workbook & worksheets");
            if (strFreezeScreen == "Yes")
            {
                Console.WriteLine("     Yes");
                Console.WriteLine("     Project: " + strProjectTitle);
                Console.WriteLine("     Master workbook: " + strMasterWorkbookFullPath);
                gnaDBAPI.testDBconnection(strDBconnection);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strReferenceWorksheet);
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strAlarmWorksheet);
            }
            else
            {
                Console.WriteLine("     No");
            }

            Console.WriteLine("   Generate time blocks");
            switch (strTimeBlockType)
            {

                case "Schedule":
                    strTimeBlockStartLocal = " '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockEndLocal = " '" + DateTime.Now.AddHours(-1.0 * dblBlockSizeHrs).ToString("yyyy-MM-dd HH:mm:ss") + "' ";
                    strTimeBlockStartUTC = gnaT.convertLocalToUTC(strTimeBlockStartLocal);
                    strTimeBlockEndUTC = gnaT.convertLocalToUTC(strTimeBlockEndLocal);
                    break;

                default:
                    Console.WriteLine("\nError in Timeblock Type");
                    Console.WriteLine("   Time block type: " + strTimeBlockType);
                    Console.WriteLine("   Must be Manual or Schedule");
                    Console.WriteLine("\nPress key to exit..."); Console.ReadKey();
                    goto ThatsAllFolks;
                    break;
            }

            string strDateTime = DateTime.Now.ToString("yyyyMMdd_HHmm");
            string strDateTimeUTC = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm");   //2022-07-26 13:45:15
            string strExportFile = strExcelPath + strContractTitle + "_" + strReportType + "_" + strDateTime + ".xlsx";
            string strTimeStamp = strTimeBlockEndLocal + "\n(local)";

            Console.WriteLine("\n2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("3. Extract SensorID");
            strSensorID = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

            Console.WriteLine("4. Write SensorID to workbook");
            gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);

            Console.WriteLine("5. Extract mean deltas for time block");
            string strBlockStart = strTimeBlockStartUTC.Replace("'", "").Trim();
            string strBlockEnd = strTimeBlockEndUTC.Replace("'", "").Trim();
            string strBlockStartLocal = strTimeBlockStartLocal.Replace("'", "").Trim();
            string strBlockEndLocal = strTimeBlockEndLocal.Replace("'", "").Trim();

            Console.WriteLine("       " + strBlockStartLocal);
            Console.WriteLine("       " + strBlockEndLocal);

            strPointDeltas = gnaDBAPI.getMeanDeltasFromDB(strDBconnection, strProjectTitle, strTimeBlockStartUTC, strTimeBlockEndUTC, strSensorID);
            string strCoordinateOrder = "ENH";
            Console.WriteLine("     Write mean deltas to Alarm " + strReferenceWorksheet + " worksheet");
            gnaSpreadsheetAPI.writeDeltas(strMasterWorkbookFullPath, strReferenceWorksheet, strPointDeltas, iFirstDataRow, iFirstDataCol, strBlockStart, strBlockEnd, strCoordinateOrder);


ThatsAllFolks:

            gnaT.freezeScreen(strFreezeScreen);

            Environment.Exit(0);
            Console.WriteLine("\nTask Complete....");

        }
    }
}