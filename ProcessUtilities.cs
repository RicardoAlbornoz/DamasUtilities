using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using AsposeCells = Aspose.Cells;
using System.IO;
using Ionic.Zip;
using System.Net.Mail;
using Aspose.Cells;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices; // Marshal.ReleaseCOMObject()
using System.Text.RegularExpressions;

namespace Utilities
{
    public abstract class ProcessUtilities
    {
        /// <summary>
        /// Gets or sets a value indicating whether this instance is zip file.
        /// </summary>
        /// <value>
        /// 	<c>true</c> if this instance is zip file; otherwise, <c>false</c>.
        /// </value>
        /// 
        public static bool IsZipFile { get; set; }
        public static string tempFileNameXlsb = "C:\\testdata\\Outbook.xlsb";
        public static string tempFileNameText = "C:\\testdata\\temp.txt";
        public static string processedFileNameText = "C:\\testdata\\Output.txt";
        public static string tempFileNameCsv = "C:\\testdata\\Outbook.csv";
        public static string tempFileName = "C:\\testdata\\Outbook.xls";
        public static string tempFileNamePdf = "C:\\testdata\\Outbook.pdf";

        /// <summary>
        /// Gets the link.
        /// </summary>
        /// <param name="emlfile">The emlfile.</param>
        /// <param name="filename">The filename.</param>
        /// <param name="category">The category.</param>
        /// <param name="returnOnly">if set to <c>true</c> [return only].</param>
        /// <param name="date">The date.</param>
        /// <returns></returns>
        /// 
        public static string periodHtmlFile(string filePath)
        {
            string fileName, period;
            string year, month, day;
            fileName = Path.GetFileName(filePath);
            year = fileName.Substring(0, 4);
            month = fileName.Substring(4, 2);
            day = fileName.Substring(6, 2);

            DateTime releasedate;
            releasedate = Convert.ToDateTime(month + "/" + day + "/" + year);
            period = ConvertToReleaseDate(releasedate).ToString();
            return period;
        }

        public static string OpenLink(string emlfile, string filename, string category,
                                        bool returnOnly, DateTime date, string zipPassword)
        {
            string month = date.Month.ToString();
            string day = date.Day.ToString();

            //Prefix 0 if single digit:
            if (date.Month.ToString().Length == 1)
            {
                month = "0" + date.Month.ToString();
            }

            if (date.Day.ToString().Length == 1)
            {
                day = "0" + date.Day.ToString();
            }

            var projectName = "MFData"; //= CommonProperties.GetInstance().SelectedProject.Project;

            //var emaillocation = "\\sixproc2\\EMLRAW\\"; //CommonProperties.GetInstance().ExtractEmailDataLocation;    //\\sixproc2\EMLRAW\
            //var emailfilelocation = "\\sixproc2\\EML\\";//CommonProperties.GetInstance().ExtractEmailFileLocation; //\\sixproc2\\EML\\
            //string complementfile = "Parsed";

            string eml = emlfile.ToString().Remove(emlfile.ToString().IndexOf(".eml"));
            string link = "";

            if (category == "filename")
            {
                link = "\\\\sixproc2\\EML\\" + projectName + "Parsed\\"
                   // link =  emailfilelocation + "\\" + projectName + complementfile + "\\"
                   + date.Year.ToString() + "\\"
                   + month + "\\" + day + "\\"
                   + eml + "\\"
                   + filename;

                ProcessUtilities.IsZipFile = false;

                //Zip files:
                if (GetFileType(filename) == "zip")
                {
                    DeleteDirectory(GetTempLocation("zip"));
                    ProcessUtilities.IsZipFile = true;
                    // Extract zip file to required location:
                    string zipToUnpack = link;
                    string unpackDirectory = GetTempLocation("zip");
                    string reqdfilename = "";
                    using (ZipFile zip1 = ZipFile.Read(zipToUnpack))
                    {
                        foreach (ZipEntry e in zip1)
                        {
                            try
                            {
                                if (GetFileType(e.FileName) == "excel" ||
                                    GetFileType(e.FileName) == "word" ||
                                    GetFileType(e.FileName) == "pdf")
                                {
                                    reqdfilename = e.FileName;
                                }
                                if (e.UsesEncryption)
                                {
                                    if (zipPassword == null || zipPassword == "")
                                    {
                                        reqdfilename = "";
                                    }
                                    e.ExtractWithPassword(unpackDirectory, ExtractExistingFileAction.OverwriteSilently, zipPassword);
                                }
                                else
                                {
                                    e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                                }
                            }
                            catch (Exception ex)
                            {
                                // MessageBox.Show("Error occurred: " + ex.Message, "Error!");
                            }
                        }
                    }

                    // Update link to extracted location
                    link = @GetTempLocation("zip") + reqdfilename;
                }
            }
            //try
            //{
            //    if (!returnOnly)
            //        System.Diagnostics.Process.Start(link);
            //}
            //catch (Exception ex)
            //{
            //    //MessageBox.Show("Can't open file: " + link, "Error!");
            //    Console.WriteLine(ex.Message.ToString());
            //}

            return link;
        }

        /// <summary>
        /// Gets the temp location.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        /// /// <summary>
        /// Get string value after [last] a.
        /// </summary>
        public static string AfterString(string value, string a)
        {
            int posA = value.LastIndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= value.Length)
            {
                return "";
            }
            return value.Substring(adjustedPosA);
        }

        /// <summary>
        /// Get string value between [first] a and [last] b.
        /// </summary>
        public static string BetweenString(string value, string a, string b)
        {
            int posA = value.IndexOf(a);
            int posB = value.LastIndexOf(b);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }

        /// <summary>
        /// Get string value before[first] a.
        /// </summary>
        public static string BeforeString(string value, string a)
        {
            int posA = value.IndexOf(a);
            if (posA == -1)
            {
                return "";
            }
            return value.Substring(0, posA);
        }


        public static string GetTempLocation(string type = null)
        {
            string suffix = "";
            if (type == "txt")
            {
                suffix = "\\DamasTemp.txt";
            }
            else if (type == "mappingtxt")
            {
                suffix = "\\DamasMappingTemp.txt";
            }
            else if (type == "zip")
            {
                suffix = "\\DamasExtractZipTemp\\";
            }
            else if (type == "word")
            {
                suffix = "\\DamasTemp.doc";
            }

            return System.Environment.GetEnvironmentVariable("TEMP") + suffix;
        }

        /// <summary>
        /// Gets the type of the file.
        /// </summary>
        /// <param name="filename">The filename.</param>
        /// <returns></returns>
        /// 
        public static bool IsNumericString(string a)
        {
            int n;
            bool isNumeric = int.TryParse(a, out n);
            return isNumeric;
        }

        public static string GetFileType(string filename)
        {
            string filetype;
            filename = filename.ToLower();
            if (filename.EndsWith(".xls") || filename.EndsWith(".xlsx")
                || filename.EndsWith(".csv") || filename.EndsWith("xlsm") || filename.EndsWith("xlsb"))
            {
                filetype = "excel";
                return filetype;
            }
            else if (filename.EndsWith(".doc") || filename.EndsWith(".docx")
                        || filename.EndsWith(".rtf") || filename.EndsWith(".txt")
                        || filename.EndsWith("xlsm"))
            {
                return "word";
            }
            else if (filename.EndsWith(".pdf"))
            {
                return "pdf";
            }
            else if (filename.EndsWith(".zip"))
            {
                return "zip";
            }
            else if (filename.EndsWith(".html") || filename.EndsWith(".htm"))
            {
                return "html";
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Copies the file to local.
        /// </summary>
        /// <param name="filepathFrom">The filepath from.</param>
        /// <returns></returns>
        public static string CopyFileToLocal(string filepathFrom)
        {
            string filepathTo = "";
            if (File.Exists(filepathFrom))
            {
                filepathTo = GetTempLocation() + "\\" + Path.GetFileName(filepathFrom);
                DeleteFile(filepathTo);
                File.Copy(filepathFrom, filepathTo);
            }

            return filepathTo;
        }


        /// <summary>
        /// Deletes the directory.
        /// </summary>
        /// <param name="directory">The directory.</param>
        public static void DeleteDirectory(string directory)
        {
            if (Directory.Exists(directory))
            {
                try
                {
                    Directory.Delete(directory, true);
                }
                catch (Exception e)
                {
                    //MessageBox.Show(e.Message, "Error occurred!");
                }
            }
        }

        /// <summary>
        /// Deletes the file.
        /// </summary>
        /// <param name="filename">The filename.</param>
        public static void DeleteFile(string filename)
        {
            if (File.Exists(filename))
            {
                try
                {
                    File.Delete(filename);
                }
                catch (Exception e)
                {
                    // MessageBox.Show("Couldn't delete the temp file: " + e.Message);
                }
            }
        }

        public static void CreateWorkbookNew()
        {
            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Total.lic");

            //Instantiate a Workbook object that represents Excel file.

            Workbook wb = new Workbook();

            //Note when you create a new workbook, a default worksheet
            //"Sheet1" is added (by default) to the workbook.
            //Access the first worksheet "Sheet1" in the book.

            Worksheet inputSheet = wb.Worksheets[0];

            //Access the "A1" cell  in the sheet

            Cell cell = inputSheet.Cells["A1"];

            //Input the "Value" text into the "A1" cell
            cell.PutValue("Value");

            //SaveFormat the Excel file
            wb.Save("C:\\testData\\Mybook.xls", SaveFormat.Excel97To2003);

        }

        public static Workbook OpenWorkbook(string fileName, string password = null)
        {
            string extractingDllName = Assembly.GetCallingAssembly().GetName().Name;
            extractingDllName = extractingDllName.Substring(7).Trim();
            string filtername;
            string filterPrefix;
            int underScoreIndex = extractingDllName.IndexOf("_");
            filtername = extractingDllName.Substring(0, underScoreIndex).Trim();
            filterPrefix = extractingDllName.Substring(underScoreIndex + 1).Trim();
            LoadOptions loadOptions;

            Aspose.Cells.License license = new Aspose.Cells.License();
            license.SetLicense("Aspose.Total.lic");

            //if (fileName.EndsWith("mfsum.xls"))
            //{
            //    fileName = fileName.Replace(".mfsum.xls", ".xls");
            //}
            //FileStream fstream = new FileStream(fileName, FileMode.Open);//("C:\\testData\\" + fileName + ".xlsx", FileMode.Open);
            Workbook workbook;

            string fileNameToParse = fileName;
            bool extensionMatches = AsposeCellsExtensionMatchesFormat(fileNameToParse);
            bool formatUnknown = CellsHelper.DetectFileFormat(fileNameToParse).Equals(FileFormatType.Unknown);

            // if aspose does not recognize file format or extensions do not match and password protected, save as xls
            if (formatUnknown || (password != null && !extensionMatches))

            {
                fileNameToParse = SaveAsExcel8WorkbookInTemp(fileName, password);
            }

            // attempt to open file in aspose, return null if fails
            try
            {
                // get file extension
                if (Path.GetExtension(fileNameToParse).Equals(".csv"))
                //if (asposeFmt.Equals(FileFormatType.CSV))
                {
                    //Instantiate a Worbook object that represents the existing Excel File
                    loadOptions = new LoadOptions(LoadFormat.CSV);
                    workbook = new Workbook(fileNameToParse, loadOptions);
                }
                else
                {
                    if (password != null)
                    {
                        //for password protected excels
                        loadOptions = new LoadOptions();
                        loadOptions.Password = password;

                        workbook = new Workbook(fileNameToParse, loadOptions);
                    }
                    else
                    {
                        workbook = new Workbook(fileNameToParse);
                    }
                }
                return workbook;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        // uses interop to check if excel version predates file format (97-2003)*.xls
        private static bool IsOldExcelVersion(string fileName, string password = null)
        {
            bool isOldVersion = false;

            // create COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook;

            try
            {
                // open workbook with or without password
                if (password == null)
                {
                    xlWorkbook = xlApp.Workbooks.Open(fileName, ReadOnly: true);
                }
                else
                {
                    xlWorkbook = xlApp.Workbooks.Open(fileName, ReadOnly: true, Password: password);
                }

                // check for old file format enums:
                // xlExcel2	    Excel version 2.0 (1987)	*.xls
                // xlExcel3	    Excel version 3.0 (1990)	*.xls
                // xlExcel4	    Excel version 4.0 (1992)	*.xls
                // xlExcel5	    Excel version 5.0 (1994)	*.xls
                // xlExcel7	    Excel 95 (version 7.0)		*.xls
                // xlExcel9795	Excel version 95 and 97	    *.xls
                isOldVersion = xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel7) ||
                        xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel5) ||
                        xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel4) ||
                        xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel3) ||
                        xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel2) ||
                        xlWorkbook.FileFormat.Equals(Excel.XlFileFormat.xlExcel9795);

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //close and release
                xlWorkbook.Close(SaveChanges: false);
                Marshal.ReleaseComObject(xlWorkbook); // System.Runtime.InteropServices
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp); // System.Runtime.InteropServices

                return isOldVersion;
            }
            catch (FileNotFoundException ex)
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp); // System.Runtime.InteropServices

                return false;
            }
        }


        // rename file and save new excel version in temp folder
        // returns string of new file name
        // does not preserve the password
        private static string SaveAsExcel8WorkbookInTemp(string fileName, string password = null)
        {
            // create COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = null;

            try
            {
                // build string for converted temp file
                string tempPath = Path.GetTempPath(); // returns final folder separator
                string oldFile = Path.GetFileNameWithoutExtension(fileName);
                string fileTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                string newFile = tempPath + oldFile + "_" + fileTime + ".xls";

                // open original workbook with or without password
                if (password == null)
                {
                    xlWorkbook = xlApp.Workbooks.Open(fileName, ReadOnly: true);
                }
                else
                {
                    xlWorkbook = xlApp.Workbooks.Open(fileName, ReadOnly: true, Password: password);
                }

                // save temp file in new format
                xlApp.DisplayAlerts = false;
                xlWorkbook.CheckCompatibility = false;
                xlWorkbook.DoNotPromptForConvert = true;
                xlWorkbook.SaveAs(newFile, Excel.XlFileFormat.xlExcel8);

                // cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // System.Runtime.InteropServices.Marshal - do not use full name for COM
                // close and release
                xlWorkbook.Close(SaveChanges: false);
                Marshal.ReleaseComObject(xlWorkbook);
                // quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                return newFile;
            }
            catch (Exception ex)
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp); // System.Runtime.InteropServices

                return null;
            }
        }

        // returns false if excel file formats do not match extensions, defaults to true for unkown file types
        public static bool AsposeCellsExtensionMatchesFormat(string filename)
        {
            bool matches = true;
            if (File.Exists(filename))
            {
                string ext = Path.GetExtension(filename).ToLower();
                FileFormatType fmt = CellsHelper.DetectFileFormat(filename);
                // go through all possible file format type enums
                switch (fmt)
                {
                    case FileFormatType.CSV:
                        matches = ext.Equals(".txt") || ext.Equals(".csv");
                        break;
                    case FileFormatType.TabDelimited:
                        matches = ext.Equals(".txt") || ext.Equals(".tsv") || ext.Equals(".tab");
                        break;
                    case FileFormatType.Html:
                    case FileFormatType.ODS:
                    case FileFormatType.Pdf:
                    case FileFormatType.Xlsb:
                    case FileFormatType.Xlsm:
                    case FileFormatType.Xlsx:
                    case FileFormatType.Xltm:
                    case FileFormatType.Xltx:
                        matches = ext.Equals(fmt.ToString().ToLower());
                        break;
                    case FileFormatType.Excel2003XML:
                        matches = ext.Equals(".xml");
                        break;
                    case FileFormatType.Excel97To2003:
                        matches = ext.Equals(".xls");
                        break;
                    case FileFormatType.Unknown:
                    default:
                        break;
                }
            }
            return matches;
        }

        //Check file's version before it is loaded
        private static bool IsOldVersion(string fileName, string password = null)
        {
            bool isOldVersion = false;
            if (fileName.Contains(".csv"))
            {
                try
                {
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.CSV);
                    Workbook wb = new Workbook(fileName);
                }
                catch (Exception ex)
                {
                    isOldVersion = true;
                }
            }
            else
            {
                try
                {
                    if (password != null)
                    {
                        //for password protected excels
                        LoadOptions loadOptions = new LoadOptions();
                        loadOptions.Password = password;

                        Workbook wb = new Workbook(fileName, loadOptions);
                    }
                    else
                    {
                        Workbook wb = new Workbook(fileName);
                    }
                }
                catch (Exception ex)
                {

                    isOldVersion = true;


                }
            }
            return isOldVersion;
        }


        /// <summary>
        /// Converts to release date.
        /// </summary>
        /// <param name="receivedDate">The received date.</param>
        /// <returns></returns>
        /// 
        public static DateTime ConvertToLastBusinessDate(DateTime receivedDate)
        {
            // Set the last business day of the month by getting the last day of the month
            var lastBusinessDay = receivedDate.AddMonths(1);
            lastBusinessDay = lastBusinessDay.AddDays(-lastBusinessDay.Day);

            // If last business day is Dec 31 then adjust to Dec 30
            if (lastBusinessDay.Month == 12)
            {
                lastBusinessDay = lastBusinessDay.AddDays(-1);
            }
            // Get day of the week
            int dow = (int)lastBusinessDay.DayOfWeek;
            if (dow == 0)
            {
                dow = 7;
            }
            // Set adjustment days based on the day of the week. No adjustment if it is a business day.
            var delta = dow > 5 ? dow - 5 : 0;

            // Adjust last business day
            lastBusinessDay = lastBusinessDay.AddDays(-delta);

            //var releaseDate = receivedDate >= lastBusinessDay ? receivedDate : receivedDate.AddDays(-receivedDate.Day);
            var releaseDate = receivedDate.AddDays(0) >= lastBusinessDay ? receivedDate : receivedDate.AddDays(-receivedDate.Day);
            // Get day of the week
            int dow2 = (int)releaseDate.DayOfWeek;
            if (dow2 == 0)
            {
                dow2 = 7;
            }
            var delta2 = dow2 > 5 ? dow2 - 5 : 0;
            releaseDate = releaseDate.AddDays(-delta2);
            return releaseDate;
        }

        public static DateTime ConvertToReleaseDate(DateTime receivedDate)
        {
            //receivedDate = receivedDate.AddMonths(-2);
            //receivedDate = receivedDate.AddDays(5);
            // Set the last business day of the month by getting the last day of the month
            var lastBusinessDay = receivedDate.AddMonths(1);
            lastBusinessDay = lastBusinessDay.AddDays(-lastBusinessDay.Day);

            // If last business day is Dec 31 then adjust to Dec 30
            if (lastBusinessDay.Month == 12)
            {
                lastBusinessDay = lastBusinessDay.AddDays(-1);
            }

            // Get day of the week
            int dow = (int)lastBusinessDay.DayOfWeek;

            if (dow == 0)
            {
                dow = 7;
            }

            // Set adjustment days based on the day of the week. No adjustment if it is a business day.
            var delta = dow > 5 ? dow - 5 : 0;

            // Adjust last business day
            lastBusinessDay = lastBusinessDay.AddDays(-delta);

            //var releaseDate = receivedDate >= lastBusinessDay ? receivedDate : receivedDate.AddDays(-receivedDate.Day);
            //UPDATE: Ryan Koch - 4/1/2013 - allow couple day cushion rather than strictly comparing receivedDate to lastBusinessDay [in the event of Holiday at end of month, for instance]
            var releaseDate = receivedDate.AddDays(0) >= lastBusinessDay ? receivedDate : receivedDate.AddDays(-receivedDate.Day);
            // Get day of the week
            //int dow2 = (int)releaseDate.DayOfWeek;
            //if (dow2 == 0)
            //{
            //    dow2 = 7;
            //}
            //var delta2 = dow2 > 5 ? dow2 - 5 : 0;
            //releaseDate = releaseDate.AddDays(-delta2);
            return releaseDate;
        }

        public List<int> GetMonthDayYear(DateTime receivedDate)
        {
            List<int> monthDayYear = new List<int>();
            string month = receivedDate.Month.ToString();
            string day = receivedDate.Day.ToString();
            string year = receivedDate.Year.ToString();

            //Prefix 0 if single digit:
            if (receivedDate.Month.ToString().Length == 1)
            {
                month = "0" + receivedDate.Month.ToString();
            }

            if (receivedDate.Day.ToString().Length == 1)
            {
                day = "0" + receivedDate.Day.ToString();
            }

            monthDayYear.Add(Convert.ToInt32(month));
            monthDayYear.Add(Convert.ToInt32(day));
            monthDayYear.Add(Convert.ToInt32(year));

            return monthDayYear;
        }


        public static void SendEmailError(Exception error, int filterid, string filtername)
        {
            int manyEmails = 1;
            string email = "";
            string userName = "";
            string pass = "";
            string emailfrom = "Srvc_Damas@sionline.com";
            userName = "Srvc_Damas@sionline.com";
            pass = "!Simfund?";
            //for (int i = 0; i < manyEmails; i++)
            //{
            //    if (i == 0)
            //    {
            //        email = "ralbornoz@sionline.com";
            //    }
            //    else
            //    {
            //    }

            email = "DamasProcessing@strategic-i.com"; //is it a group on outlook

            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add(email);
            message.Subject = "Problem with New Fund Process: " + filtername + " " + "FilterID: " + filterid;
            message.From = new System.Net.Mail.MailAddress(emailfrom);
            message.Body = "Message: " + error;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("site2.exch580.serverdata.net");
            smtp.Credentials = new System.Net.NetworkCredential(userName, pass);
            smtp.Port = Convert.ToInt16("587");
            smtp.EnableSsl = true;
            smtp.Send(message);
            //}
        }

        public static void SendEmailErrorMessage(string error, int filterid, string filtername, string filterprefix)
        {

            int manyEmails = 1;
            string email = "";
            string userName = "";
            string pass = "";
            string emailfrom = "Srvc_Damas@sionline.com";
            userName = "Srvc_Damas@sionline.com";
            pass = "!Simfund?";
            //for (int i = 0; i < manyEmails; i++)
            //{
            //    if (i == 0)
            //    {
            //        email = "ralbornoz@sionline.com";
            //        email = "mye@sionline.com";
            //        ////userName = "mye";
            //        ////pass = "Myanmar1";
            //    }
            //    else
            //    {
            //    }
            email = "DamasProcessing@strategic-i.com";  //is it a group on outlook

            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add(email);
            message.Subject = "Problem with New Fund Process: " + filtername + "_" + filterprefix;
            message.From = new System.Net.Mail.MailAddress(emailfrom);
            message.Body = "Message: " + error;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("site2.exch580.serverdata.net");
            smtp.Credentials = new System.Net.NetworkCredential(userName, pass);
            smtp.Port = Convert.ToInt16("587");
            smtp.EnableSsl = true;
            smtp.Send(message);
            // }

        }
        public static void SendEmailSuccessMessage(string success, int filterid, string filtername, string filterprefix)
        {
            int manyEmails = 1;
            string email = "";
            string userName = "";
            string pass = "";
            string emailfrom = "Srvc_Damas@sionline.com";
            userName = "Srvc_Damas@sionline.com";
            pass = "!Simfund?";
            //for (int i = 0; i < manyEmails; i++)
            //{
            //    if (i == 0)
            //    {
            //        email = "ralbornoz@sionline.com";

            //    }
            //    else
            //    {

            //    }
            email = "DamasProcessing@strategic-i.com";  //is it a group on outlook

            System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage();
            message.To.Add(email);
            message.Subject = "New Fund Processed: " + filtername + "_" + filterprefix;
            message.From = new System.Net.Mail.MailAddress(emailfrom);
            message.Body = "Message: " + success;
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("site2.exch580.serverdata.net");
            smtp.Credentials = new System.Net.NetworkCredential(userName, pass);
            smtp.Port = Convert.ToInt16("587");
            smtp.EnableSsl = true;
            smtp.Send(message);
            //}
        }

        //Paramter for StreamWriter file
        public static StreamWriter OutPutFileDamas(StreamWriter outputFileDamasTemp)
        {
            string outputpath = Utilities.ProcessUtilities.GetTempLocation("txt");
            //outputFileDamasTemp = new StreamWriter(outputpath);

            //insert the header values into text file (damastemp.txt)
            string header = "ID" + "\t" + "Name" + "\t" + "Cusip" + "\t" + "Ticker"
            + "\t" + "Period" + "\t" + "Assets" + "\t" + "Flows"
            + "\t" + "Notes" + "\t" + "RowNo";

            outputFileDamasTemp.WriteLine(header);

            return outputFileDamasTemp;
        }


        //Convert pdf to xls files
        public static string ConvertToExcel(string filePath)
        {
            string excelFilePath = tempFileName;
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(filePath);
            doc.Save(tempFileNamePdf);
            // instantiate ExcelSave Option object
            Aspose.Pdf.ExcelSaveOptions excelsave = new Aspose.Pdf.ExcelSaveOptions();
            Aspose.Pdf.Document doc1 = new Aspose.Pdf.Document(tempFileNamePdf);
            // save the output in XLS format
            doc1.Save(excelFilePath, excelsave);
            return excelFilePath;

        }

        //convert month from string format to int format
        public static int getMonthNumber(string month)
        {

            string temp;
            temp = month.ToUpper();
            if (temp.Contains("JAN"))
                return 1;
            else if (temp.Contains("FEB"))
                return 2;
            else if (temp.Contains("MAR"))
                return 3;
            else if (temp.Contains("APR"))
                return 4;
            else if (temp.Contains("MAY"))
                return 5;
            else if (temp.Contains("JUN"))
                return 6;
            else if (temp.Contains("JUL"))
                return 7;
            else if (temp.Contains("AUG"))
                return 8;
            else if (temp.Contains("SEP"))
                return 9;
            else if (temp.Contains("OCT"))
                return 10;
            else if (temp.Contains("NOV"))
                return 11;
            else
                return 12;

        }
        //convert month from int format to string format
        public static string getMonthString(int month)
        {

            int temp;
            temp = month;
            if (temp == 1 || temp == 01)
                return "January";
            else if (temp == 2 || temp == 02)
                return "February";
            else if (temp == 3 || temp == 03)
                return "March";
            else if (temp == 4 || temp == 04)
                return "April";
            else if (temp == 5 || temp == 05)
                return "May";
            else if (temp == 6 || temp == 06)
                return "June";
            else if (temp == 7 || temp == 07)
                return "July";
            else if (temp == 8 || temp == 08)
                return "August";
            else if (temp == 9 || temp == 09)
                return "September";
            else if (temp == 10)
                return "October";
            else if (temp == 11)
                return "November";
            else
                return "December";

        }
        public static string getFullMonthName(string month)
        {
            if (month.Length <= 2)
            {
                if (!Regex.IsMatch(month, "[a-z0-9 ]+", RegexOptions.IgnoreCase))
                {
                    return getMonthString(Convert.ToInt32(month));
                }
                else
                {
                    return "No_Month";
                }
            }
            else
            {
                month = month.Trim().ToUpper();
                if (month.Contains("JAN"))
                    return "JANUARY";
                else if (month.Contains("FEB"))
                    return "FEBRUARY";
                else if (month.Contains("MAR"))
                    return "MARCH";
                else if (month.Contains("APR"))
                    return "APRIL";
                else if (month.Contains("MAY"))
                    return "MAY";
                else if (month.Contains("JUN"))
                    return "JUNE";
                else if (month.Contains("JUL"))
                    return "JULY";
                else if (month.Contains("AUG"))
                    return "AUGUST";
                else if (month.Contains("SEP"))
                    return "SEPTEMBER";
                else if (month.Contains("OCT"))
                    return "OCTOBER";
                else if (month.Contains("NOV"))
                    return "NOVEMBER";
                else
                    return "DECEMBER";
            }
        }

        public static int isLeapYear(string year)
        {
            int yr;
            yr = Convert.ToInt32(year);
            if (yr == 0)
                return -1;
            if (yr % 4 == 0)
            {
                if (yr % 100 == 0)
                {
                    if (yr % 400 == 0)
                    {
                        return 1;
                    }
                    else
                        return 0;
                }
                else
                    return 1;
            }
            else
                return 0;
        }

        public static int daysInMonth(string month, string year)
        {
            month = getFullMonthName(month.Trim().ToUpper());
            switch (month.ToUpper())
            {
                case "JANUARY":
                    return 31;
                case "MARCH":
                    return 31;
                case "MAY":
                    return 31;
                case "JULY":
                    return 31;
                case "AUGUST":
                    return 31;
                case "OCTOBER":
                    return 31;
                case "DECEMBER":
                    return 31;
                case "APRIL":
                    return 30;
                case "JUNE":
                    return 30;
                case "SEPTEMBER":
                    return 30;
                case "NOVEMBER":
                    return 30;
                case "FEBRUARY":
                    if (isLeapYear(year) > 0)
                        return 29;
                    else if (isLeapYear(year) == 0)
                        return 28;
                    else
                        return -1;
                default:
                    return -1;
            }

        }

        public static int daysInMonth(int month, string year)
        {
            switch (month)
            {
                case 1:
                    return 31;
                case 3:
                    return 31;
                case 5:
                    return 31;
                case 7:
                    return 31;
                case 8:
                    return 31;
                case 10:
                    return 31;
                case 12:
                    return 31;
                case 4:
                    return 30;
                case 6:
                    return 30;
                case 9:
                    return 30;
                case 11:
                    return 30;
                case 2:
                    if (isLeapYear(year) > 0)
                        return 29;
                    else if (isLeapYear(year) == 0)
                        return 28;
                    else
                        return -1;
                default:
                    return -1;
            }
        }

        public static int daysInMonthInt(int month, string year)
        {
            switch (month)
            {
                case 1:
                    return 31;
                case 3:
                    return 31;
                case 5:
                    return 31;
                case 7:
                    return 31;
                case 8:
                    return 31;
                case 10:
                    return 31;
                case 12:
                    return 31;
                case 4:
                    return 30;
                case 6:
                    return 30;
                case 9:
                    return 30;
                case 11:
                    return 30;
                case 2:
                    if (isLeapYear(year) > 0)
                        return 29;
                    else if (isLeapYear(year) == 0)
                        return 28;
                    else
                        return -1;
                default:
                    return -1;
            }
        }

        public static string UnZipFile(string filename)
        {
            string link = filename;
            //Zip files:
            if (GetFileType(filename) == "zip")
            {
                DeleteDirectory(GetTempLocation("zip"));
                ProcessUtilities.IsZipFile = true;
                // Extract zip file to required location:
                string zipToUnpack = link;
                string unpackDirectory = GetTempLocation("zip");
                string reqdfilename = "";
                using (ZipFile zip1 = ZipFile.Read(zipToUnpack))
                {
                    foreach (ZipEntry e in zip1)
                    {
                        try
                        {
                            if (GetFileType(e.FileName) == "excel" ||
                                GetFileType(e.FileName) == "word" ||
                                GetFileType(e.FileName) == "pdf")
                            {
                                reqdfilename = e.FileName;
                            }
                            if (e.UsesEncryption)
                            {
                                e.ExtractWithPassword(unpackDirectory, ExtractExistingFileAction.OverwriteSilently, "zipPassword");
                            }
                            else
                            {
                                e.Extract(unpackDirectory, ExtractExistingFileAction.OverwriteSilently);
                            }
                        }
                        catch (Exception ex)
                        {
                            // MessageBox.Show("Error occurred: " + ex.Message, "Error!");
                        }
                    }
                }

                // Update link to extracted location
                link = @GetTempLocation("zip") + reqdfilename;
            }
            return link;
        }
    }
}
