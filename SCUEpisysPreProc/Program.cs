using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Outlook;
using System.Text.RegularExpressions;
using System.IO.Packaging;
using System.Net.Mail;
using System.Diagnostics;

namespace SCUEpisysPreProc
{
    class Program
    {
/// <summary>
/// This is the main for a Console application that's primary purpose is to move Episys COLD files
/// ... from a source directory to other directories based on the application config file settings
/// ... and the pre-define business rules.  The program runs through 5 steps in decided what to move
/// ... where.  The process steps are performed in the following order each time the program runs:
///         1. Get a list of the files currently in the source directory
///         2. Move large files to a No Process directory that will require a manual spliting
///         3. Open all remaining files to determine if any Statement files and move to Statement directory
///         4. Open all remaining files to determine if any Notice files and move to Notices directory
///         5. Move any remaining files considerted big (> some Byte threshold) to Big directory
///         6. Move any remaining file in the list to the default directory
/// </summary>
/// <param name="args"></param>
        static void Main(string[] args)
        {
            // Create the preprocessor object
            Preproc thisPP = new Preproc();
            // Check to see if we had an error creating the PreProc object
            if (thisPP._errorMsg != "")
            {
                Console.ReadLine();
                Environment.Exit(-1);
            }
            // First check is to see if any file too large to process
            bool NoProcessSuccess = thisPP.CheckNoProcessFiles();
            // Look at remaining files to see if any ACH Reports
            bool ACHReportsSuccess = thisPP.CheckACHReportsFiles();
            // Look at remaining files to see if any statements
            bool StatementSuccess = thisPP.CheckStatementFiles();
            // Look at remaining files to see if any notices
            bool NoticeSuccess = thisPP.CheckNoticeFiles();
            // Look at remaining files to see if bigger than some byte size
            bool BigFileSuccess = thisPP.CheckBigFiles();
            // Move remaining files to default direcotry
            bool DefaultSuccess = thisPP.CheckDefaultFiles();
            // Remove any backup files that are older than specified days for each file type
            thisPP.PruneBackupFiles(thisPP._StatementBackupDirectory, Convert.ToInt32(thisPP._StatementBackupDeleteRetentionDays));
            thisPP.PruneBackupFiles(thisPP._NoticeBackupDirectory, Convert.ToInt32(thisPP._NoticeBackupDeleteRetentionDays));
            thisPP.PruneBackupFiles(thisPP._ACHReportsBackupDirectory, Convert.ToInt32(thisPP._ACHReportsBackupDeleteRetentionDays));
            thisPP.PruneBackupFiles(thisPP._BigBackupDirectory, Convert.ToInt32(thisPP._BigBackupDeleteRetentionDays));
            thisPP.PruneBackupFiles(thisPP._DefaultBackupDirectory, Convert.ToInt32(thisPP._DefaultBackupDeleteRetentionDays));
            // Check to see if we preprocessed everything by comparing the total list count to each type we processed
            if ((thisPP._NoProcessFiles + thisPP._NoticeFiles + thisPP._BigFiles + thisPP._DefaultFiles + thisPP._LockedFiles + thisPP._StatementFiles + thisPP._ACHReportsFiles) != thisPP._TotalFilesFound)
                thisPP.SendEmail(thisPP._ErrorEMailAddress, "Episys Preproc Issue - Files Remain in Source Directory", "The total files preprocess plus locked files did not match total number of files in the source directory. Please review log.");
            // Output stats to log
            thisPP.LogMessage("Total Files Found: " + thisPP._TotalFilesFound.ToString());
            thisPP.LogMessage("No Process Files Found: " + thisPP._NoProcessFiles.ToString());
            thisPP.LogMessage("Statement Files Found: " + thisPP._StatementFiles.ToString());
            thisPP.LogMessage("Notice Files Found: " + thisPP._NoticeFiles.ToString());
            thisPP.LogMessage("ACH Report Files Found: " + thisPP._ACHReportsFiles.ToString());
            thisPP.LogMessage("Big Files Found: " + thisPP._BigFiles.ToString());
            thisPP.LogMessage("Default Files Found: " + thisPP._DefaultFiles.ToString());
            // Check to see if we had any locked files and if so then list in the log
            string lockedFileList = "";
            foreach (var file in thisPP._LockedFilenames)
            {
                lockedFileList += file + ", ";
            }
            if (thisPP._LockedFilenames.Count > 0)
            {
                lockedFileList = LeftRightMid.Left(lockedFileList, lockedFileList.Length - 2);
                thisPP.LogMessage("Locked Files Found: " + thisPP._LockedFiles.ToString() + " - " + lockedFileList);
                thisPP.SendEmail(thisPP._ErrorEMailAddress, "Episys PreProc - Locked Files Found", "The following file(s) were found locked in the source directory. Please review.\r\n\r\nFilenames: " + lockedFileList);
            }
            // Check and see if any files left in source directory. Put in log if any
            string[] leftfiles;
            leftfiles = Directory.GetFiles(thisPP._SourceFileDirectory);
            if (leftfiles.Length > 0)
            {
                string leftFileList = "";
                foreach (string item in leftfiles)
                    leftFileList += item + ", ";
                leftFileList = LeftRightMid.Left(leftFileList, leftFileList.Length - 2);
                thisPP.LogMessage("Files Left Unprocessed: " + leftFileList);
            }

            // Pretty much done so output complete message and time
            thisPP.LogMessage("Preprocess Complete: " + DateTime.Now.ToLongTimeString());
            // If we are in console mode then do readline so console display stay open until enter key
            if (thisPP._OutputConsole)
            {
                Console.WriteLine("Hit Return / Enter to exit.");
                Console.ReadLine();
            }
            // Return Success
            Environment.Exit(0);
        }
    }

    /// <summary>
    /// This is the primary object of the preprocessor with different methods to perform
    /// ... moving of different file types.
    /// </summary>
   public class Preproc
    {

        public string _ErrorEMailAddress;
        public string _SourceFileDirectory;
        public string _LogFilePath;
        public string _LogFilename;
        public string _LogFileCount;
        public bool _LoggingOn;
        public bool _OutputConsole;
        public bool _LogAppSettings;
        public string _NoProcessFilename1;
        public string _NoProcessFilename2;
        public string _NoProcessFilename3;
        public string _NoProcessFilename4;
        public string _NoProcessFilename5;
        public string _NoProcessFileSize;
        public string _NoProcessNotificationEmailAddress;
        public string _NoProcessMoveDirectory;
        public string _StatementMoveDirectory;
        public string _StatementBackupDirectory;
        public string _StatementSearchString;
        public string _StatementBackupDeleteRetentionDays;
        public string _StatementSemaphoreFilename;
        public string _NoticeMoveDirectory;
        public string _NoticeBackupDirectory;
        public string _NoticeSearchString;
        public string _NoticeBackupDeleteRetentionDays;
        public string _NoticeSemaphoreFilename;
        public string _ACHReportsMoveDirectory;
        public string _ACHReportsBackupDirectory;
        public string _ACHReportsSearchString;
        public string _ACHReportsBackupDeleteRetentionDays;
        public string _ACHReportsSemaphoreFilename;
        public string _ACHEmailNotificationAddress = "";
        public string _BigFileSize;
        public string _BigMoveDirectory;
        public string _BigBackupDirectory;
        public string _BigBackupDeleteRetentionDays;
        public string _BigSemaphoreFilename;
        public string _DefaultMoveDirectory;
        public string _DefaultBackupDirectory;
        public string _DefaultBackupDeleteRetentionDays;
        public string _DefaultSemaphoreFilename;
        public string _EmailStartTime;
        public string _EmailEndTime;
        public string _EmailDays;
        public string _SourceEMailAddress;
        public bool _GenerateEmail;
        public bool _ZipBackupFiles;
        public bool _OverwriteMoveFiles;
        public StreamReader _LogFileReaderStream;
        public StreamWriter _LogFileWriterStream;
        public string _errorMsg = "";
        public int _TotalFilesFound = 0;
        public int _NoProcessFiles = 0;
        public int _NoticeFiles = 0;
        public int _BigFiles = 0;
        public int _DefaultFiles = 0;
        public int _LockedFiles = 0;
        public int _OtherFileError = 0;
        public int _StatementFiles = 0;
        public int _ACHReportsFiles = 0;
        public List<string> _LockedFilenames;
        public List<string> files;
        public List<string> _MovedFilenames;

/// <summary>
/// Constructor that reads in app.config settings and initializes object Preproc object for use
/// </summary>
        public Preproc()
        {
            // Get the general config settings including log file path and settings
            _LogFilePath = ReadSetting("LogFilePath");
            _LogFileCount = ReadSetting("LogFileCount");
            _LogAppSettings = Convert.ToBoolean(ReadSetting("LogAppSettings"));
            _OutputConsole = Convert.ToBoolean(ReadSetting("OutputConsole"));
            _GenerateEmail = Convert.ToBoolean(ReadSetting("GenerateEmail"));
            _ZipBackupFiles = Convert.ToBoolean(ReadSetting("ZipBackupFiles"));
            _OverwriteMoveFiles = Convert.ToBoolean(ReadSetting("OverwriteMoveFiles"));
            if (String.IsNullOrWhiteSpace(_LogFileCount))
                _LogFileCount = "30";
            _LoggingOn = Convert.ToBoolean(ReadSetting("LoggingOn"));
            // If logging then open / create file
            if (_LoggingOn)
            {
                // Make sure the log file path is valid
                if (!IsValidPath(_LogFilePath))
                {
                    // We get here there is an invalid path ... set error mesage and stop constructor
                    _errorMsg = "Invalid Log File Directory path: " + _LogFilePath;
                    Console.WriteLine(DateTime.Now.ToShortDateString() + ":" + DateTime.Now.ToShortTimeString() + ":" + _errorMsg);
                    Console.WriteLine("Hit Return / Enter to exit.");
                    return;
                }
                // Good directory so now build log filename for today
                DateTime tempDT = DateTime.Now;
                if (LeftRightMid.Right(_LogFilePath,1) == @"\")
                    _LogFilename = _LogFilePath + "PreprocLog-" + tempDT.Month.ToString() + "-" + tempDT.Day.ToString() + "-" + tempDT.Year.ToString() + ".txt";
                else
                    _LogFilename = _LogFilePath + @"\PreprocLog-" + tempDT.Month.ToString() + "-" + tempDT.Day.ToString() + "-" + tempDT.Year.ToString() + ".txt";

                // Delete any log files that are in log directory that are older then today minus log file count 
                PruneLogFiles(_LogFilePath, Convert.ToInt32(_LogFileCount));
                
                // Either create the log file if doesn't exist or open for writing if it already exists
                if (File.Exists(_LogFilename))
                {
                    // Prune the file if exceeded line limit
                    // Open log file for writing append
                    try
                    {
                        _LogFileWriterStream = new StreamWriter(_LogFilename, true);
                    }
                    catch (Exception e)
                    {
                        _errorMsg = "Error opening Log file: " + e.Message;
                        Console.WriteLine("Error opening Log file: " + e.Message);
                        Console.WriteLine("Hit Return / Enter to exit.");
                        return;
                    }
                }
                else
                {
                    // Create the log directory in case it doesn't exist
                    string dirPath = Path.GetDirectoryName(_LogFilename);
                    if (!Directory.Exists(dirPath))
                        Directory.CreateDirectory(dirPath);
                    // Open log file for writing append
                    try
                    {
                        _LogFileWriterStream = new StreamWriter(_LogFilename, true);
                    }
                    catch (Exception e)
                    {
                        _errorMsg = "Error creating Log file: " + e.Message;
                        Console.WriteLine("Error creating Log file: " + e.Message);
                        Console.WriteLine("Hit Return / Enter to exit.");
                        return;
                    }
                }
            }
            // Output start time to log file
            LogMessage("Preprocess Started: " + DateTime.Now.ToLongTimeString());
            // If app setting are to be logged then put in the ones we already read in
            if (_LogAppSettings)
            { 
                LogMessage("Using AppSetting (LogFilePath) : " + _LogFilePath);
                LogMessage("Using AppSetting (LogFilename) : " + _LogFilename);
                LogMessage("Using AppSetting (LoggingOn) : " + _LoggingOn);
                LogMessage("Using AppSetting (LogFileCount) : " + _LogFileCount);
                LogMessage("Using AppSetting (OutputConsole) : " + _OutputConsole);
                LogMessage("Using AppSetting (GenerateEmail) : " + _GenerateEmail);
                LogMessage("Using AppSetting (ZipBackupFiles) : " + _ZipBackupFiles);
                LogMessage("Using AppSetting (OverwriteMoveFiles) : " + _OverwriteMoveFiles);
            }
            // Now get the rest of the appsettings ... these will be logged as part of read now that log file is open
            _ErrorEMailAddress = ReadSetting("ErrorEMailAddress");
            _SourceFileDirectory = ReadSetting("SourceFileDirectory");
            _NoProcessFilename1 = ReadSetting("NoProcessFilename1");
            _NoProcessFilename2 = ReadSetting("NoProcessFilename2");
            _NoProcessFilename3 = ReadSetting("NoProcessFilename3");
            _NoProcessFilename4 = ReadSetting("NoProcessFilename4");
            _NoProcessFilename5 = ReadSetting("NoProcessFilename5");
            _NoProcessFileSize = ReadSetting("NoProcessFileSize");
            _NoProcessNotificationEmailAddress = ReadSetting("NoProcessNotificationEmailAddress");
            _NoProcessMoveDirectory = ReadSetting("NoProcessMoveDirectory");
            _StatementMoveDirectory = ReadSetting("StatementMoveDirectory");
            _StatementBackupDirectory = ReadSetting("StatementBackupDirectory");
            _StatementSearchString = ReadSetting("StatementSearchString");
            _StatementBackupDeleteRetentionDays = ReadSetting("StatementBackupDeleteRetentionDays");
            _StatementSemaphoreFilename = ReadSetting("StatementSemaphoreFilename");
            _NoticeMoveDirectory = ReadSetting("NoticeMoveDirectory");
            _NoticeBackupDirectory = ReadSetting("NoticeBackupDirectory");
            _NoticeSearchString = ReadSetting("NoticeSearchString");
            _NoticeBackupDeleteRetentionDays = ReadSetting("NoticeBackupDeleteRetentionDays");
            _NoticeSemaphoreFilename = ReadSetting("NoticeSemaphoreFilename");
            _ACHReportsMoveDirectory = ReadSetting("ACHReportsMoveDirectory");
            _ACHReportsBackupDirectory = ReadSetting("ACHReportsBackupDirectory");
            _ACHReportsSearchString = ReadSetting("ACHReportsSearchString");
            _ACHReportsBackupDeleteRetentionDays = ReadSetting("ACHReportsBackupDeleteRetentionDays");
            _ACHReportsSemaphoreFilename = ReadSetting("ACHReportsSemaphoreFilename");
            _ACHEmailNotificationAddress = ReadSetting("ACHEmailNotificationAddress");
            _BigFileSize = ReadSetting("BigFileSize");
            _BigMoveDirectory = ReadSetting("BigMoveDirectory");
            _BigBackupDirectory = ReadSetting("BigBackupDirectory");
            _BigBackupDeleteRetentionDays = ReadSetting("BigBackupDeleteRetentionDays");
            _BigSemaphoreFilename = ReadSetting("BigSemaphoreFilename");
            _DefaultMoveDirectory = ReadSetting("DefaultMoveDirectory");
            _DefaultBackupDirectory = ReadSetting("DefaultBackupDirectory");
            _DefaultBackupDeleteRetentionDays = ReadSetting("DefaultBackupDeleteRetentionDays");
            _DefaultSemaphoreFilename = ReadSetting("DefaultSemaphoreFilename");
            _EmailStartTime = ReadSetting("EmailStartTime");
            _EmailEndTime = ReadSetting("EmailEndTime");
            _EmailDays = ReadSetting("EmailDays");
            _SourceEMailAddress = ReadSetting("SourceEMailAddress");

            // Check if the email addressese are valid
            if (!IsValidEmail(_ErrorEMailAddress))
                LogMessage("Invalid Email Address:" + _ErrorEMailAddress);
            if (!IsValidEmail(_NoProcessNotificationEmailAddress))
                LogMessage("Invalid Email Address:" + _NoProcessNotificationEmailAddress);
            if (!IsValidEmail(_SourceEMailAddress))
                LogMessage("Invalid Email Address:" + _SourceEMailAddress);
            if ((_ACHEmailNotificationAddress != "") && (!IsValidEmail(_ACHEmailNotificationAddress)))
                LogMessage("Invalid Email Address:" + _ACHEmailNotificationAddress);

            // Validate the different directories ... create if valid but doesn't exist
            if (!IsValidPath(_SourceFileDirectory))
            {
                LogMessage("Source Directory does not exist: " + _SourceFileDirectory);
                _errorMsg = "Source Directory does not exist: " + _SourceFileDirectory;
                return;
            }
            if (!IsValidPath(_NoProcessMoveDirectory))
            {
                LogMessage("No Process Directory does not exist: " + _NoProcessMoveDirectory);
                _errorMsg = "No Process Directory does not exist: " + _NoProcessMoveDirectory;
                return;
            }
            if (!IsValidPath(_StatementMoveDirectory))
            {
                LogMessage("Statement Move Directory does not exist: " + _StatementMoveDirectory);
                _errorMsg = "Statement Move Directory does not exist: " + _StatementMoveDirectory;
                return;
            }
            if (!IsValidPath(_StatementBackupDirectory))
            {
                LogMessage("Statement Backup Directory does not exist: " + _StatementBackupDirectory);
                _errorMsg = "Statement Backup Directory does not exist: " + _StatementBackupDirectory;
                return;
            }
            if (!IsValidPath(_NoticeMoveDirectory))
            {
                LogMessage("Notice Move Directory does not exist: " + _NoticeMoveDirectory);
                _errorMsg = "Notice Move Directory does not exist: " + _NoticeMoveDirectory;
                return;
            }
            if (!IsValidPath(_NoticeBackupDirectory))
            {
                LogMessage("Notice Backup Directory does not exist: " + _NoticeBackupDirectory);
                _errorMsg = "Notice Backup Directory does not exist: " + _NoticeBackupDirectory;
                return;
            }
            if (!IsValidPath(_ACHReportsMoveDirectory))
            {
                LogMessage("ACH Reports Move Directory does not exist: " + _ACHReportsMoveDirectory);
                _errorMsg = "ACH Reports Move Directory does not exist: " + _ACHReportsMoveDirectory;
                return;
            }
            if (!IsValidPath(_ACHReportsBackupDirectory))
            {
                LogMessage("ACH Reports Backup Directory does not exist: " + _ACHReportsBackupDirectory);
                _errorMsg = "ACH Reports Backup Directory does not exist: " + _ACHReportsBackupDirectory;
                return;
            }
            if (!IsValidPath(_BigMoveDirectory))
            {
                LogMessage("Big Move Directory does not exist: " + _BigMoveDirectory);
                _errorMsg = "Big Move Directory does not exist: " + _BigMoveDirectory;
                return;
            }
            if (!IsValidPath(_BigBackupDirectory))
            {
                LogMessage("Big Backup Directory does not exist: " + _BigBackupDirectory);
                _errorMsg = "Big Backup Directory does not exist: " + _BigBackupDirectory;
                return;
            }
            if (!IsValidPath(_DefaultMoveDirectory))
            {
                LogMessage("Default Move Directory does not exist: " + _DefaultMoveDirectory);
                _errorMsg = "Default Move Directory does not exist: " + _DefaultMoveDirectory;
                return;
            }
            if (!IsValidPath(_DefaultBackupDirectory))
            {
                LogMessage("Default Backup Directory does not exist: " + _DefaultBackupDirectory);
                _errorMsg = "Default Backup Directory does not exist: " + _DefaultBackupDirectory;
                return;
            }
            _LockedFilenames = new List<string>();
            _MovedFilenames = new List<string>();
            LogMessage("---------------------------------------------------------------------");

            // At this point we are ready to get the list of source files from the source directory
            // ... All subsequent checks will use this file list.
            // Only get files in the source directory
            string[] tempfiles;
            tempfiles = Directory.GetFiles(@_SourceFileDirectory);
            files = new List<string>();
            foreach (var fn in tempfiles)
                files.Add(fn);
            _TotalFilesFound = files.Count;
        }

/// <summary>
/// This method will look through the directory list and move any files deemed to big to process
/// ... by COLD based on app setting - NoProcessFileSize
/// </summary>
/// <returns>true-if no major errors</returns>
       public bool CheckNoProcessFiles()
        {
            try
            {
               // Only get files in the source directory
               //string[] files = Directory.GetFiles(@_SourceFileDirectory);
               if (files.Count > 0)
               {
                   foreach (string file in files)
                   {
                       FileInfo thisFileInfo = new FileInfo(file);
                       if (thisFileInfo.Length > Convert.ToInt32(_NoProcessFileSize))
                       {
                           if (!IsFileLocked(file))
                           {
                                try 
                                {	
                                    DateTime currentDT = DateTime.Now;
                                    string backupExtension = "";

                                    if (File.Exists(@_NoProcessMoveDirectory+Path.GetFileName(file)))
                                    {
                                        if (_OverwriteMoveFiles)
                                        { 
                                            File.Delete(@_NoProcessMoveDirectory+Path.GetFileName(file));
                                            LogMessage("No Process large file already exists in No Process move directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                        }
                                        else
                                            backupExtension = GetDateTimeFileExtension(currentDT);
                                    }
                                    File.Move(file,@_NoProcessMoveDirectory+Path.GetFileName(file)+backupExtension);
                                    LogMessage("No Process file moved based on file size: " + Path.GetFileName(file) + backupExtension);
                                    SendEmail(_NoProcessNotificationEmailAddress, "Episys PreProc - Large File Moved", "The following file was moved to the No Process directory (" + _NoProcessMoveDirectory + "). Please review.\r\n\r\nFilename: " + file + backupExtension);
                                    _NoProcessFiles++;
                                    _MovedFilenames.Add(file);

                                }
                            	catch (Exception exp)
	                            {
                        		    LogMessage("No Process File move error:" + exp.Message);
                                    _OtherFileError++;
                        	    }
                           }
                           else
                           {
                               _LockedFilenames.Add(file);
                               LogMessage("No Process large file found locked (" + Path.GetFileName(file) + "). Skipping file move.");
                               _LockedFiles++;
                           }
                       }
                       else
                       {
                           string thisFN = Path.GetFileName(file);
                           if ((thisFN == _NoProcessFilename1) || (thisFN == _NoProcessFilename2) || (thisFN == _NoProcessFilename3) || (thisFN == _NoProcessFilename4) || (thisFN == _NoProcessFilename5))
                           {
                                if (!IsFileLocked(file))
                                {
                                    try 
                                    {
                                        DateTime currentDT = DateTime.Now;
                                        string backupExtension = "";

                                        if (File.Exists(@_NoProcessMoveDirectory + Path.GetFileName(file)))
                                        {
                                            if (_OverwriteMoveFiles)
                                            {
                                                File.Delete(@_NoProcessMoveDirectory + Path.GetFileName(file));
                                                LogMessage("No Process large file already exists in No Process move directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                            }
                                            else
                                                backupExtension = GetDateTimeFileExtension(currentDT);
                                        }
                                        File.Move(file, @_NoProcessMoveDirectory + Path.GetFileName(file) + backupExtension);
                                        LogMessage("No Process file moved based on filename: " + Path.GetFileName(file) + backupExtension);
                                        SendEmail(_NoProcessNotificationEmailAddress, "Episys PreProc - No Process filename found and moved", "The following file was moved to the No Process directory (" + _NoProcessMoveDirectory + "). Please review.\r\n\r\nFilename: " + file + backupExtension);
                                        _NoProcessFiles++;
                                        _MovedFilenames.Add(file);
                                    }
                            	    catch (Exception exp)
	                                {
                        		        LogMessage("No Process File move error:" + exp.Message);
                                        _OtherFileError++;
                                    }
                                }
                                else
                                {
                                    LogMessage("No Process Filename found locked (" + Path.GetFileName(file) + "). Skipping file move.");
                                    _LockedFilenames.Add(file);
                                    _LockedFiles++;
                                }
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("No Process get files error: " + e.Message);
               return false;
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
        }

       public bool CheckStatementFiles()
       {
            List<string> zipFNList = new List<string>();
            DateTime zipCurrentDT = DateTime.Now;
            string zipExtension = GetDateTimeFileExtension(zipCurrentDT);
            string ZipFN = "StatementsBU" + zipExtension + ".zip";
           try
           {
               // Only get files in source directory
               if ((files.Count > 0) && (!String.IsNullOrWhiteSpace(_StatementSearchString)))
               {
                   foreach (string file in files)
                   {
                       // Open file as text file, read first line to see if a Statement file
                       if ((!IsFileLocked(file)) && (!_LockedFilenames.Exists(element => element == file)))
                       {
                           string blankStatements = "";
                           string StatementString = "";
                           try
                           {
                               StreamReader tempSR = new StreamReader(file);
                               string lineOne = tempSR.ReadLine();
                               try
                               {
                                   blankStatements = LeftRightMid.Left(lineOne, 5);
                                   if (lineOne.Length > 80)
                                       StatementString = LeftRightMid.Mid(lineOne, 41, 40).TrimEnd(' ');
                                   else
                                       StatementString = "";
                               }
                               catch (Exception exp)
                               {
                                   LogMessage("******* Issue looking for Statement String (" + file + "):" + exp.Message + " *******");
                                   StatementString = "";
                                   _OtherFileError++;
                                   SendEmail(_ErrorEMailAddress, "Episys PreProc - Statement Read Line Issue.", "The following file had an issue with reading line 1 for statement string. Please review.\r\n\r\nFilename: " + file);
                               }
                               if (((_StatementSearchString.IndexOf(StatementString + ";") >= 0) && (StatementString != "")) || (blankStatements == "+    " ))
                               {
                                   tempSR.Close();
                                   // Move Statement file to back up directory first
                                   DateTime currentDT = DateTime.Now;
                                   string backupExtension = GetDateTimeFileExtension(currentDT);
                                   string uniqueExtension = "";
                                   if (File.Exists(@_StatementBackupDirectory + Path.GetFileName(file) + backupExtension))
                                   {
                                        File.Delete(@_StatementBackupDirectory + Path.GetFileName(file) + backupExtension);
                                        LogMessage("Statement file already exists in Backup directory (" + Path.GetFileName(file) + backupExtension + "). Existing file being deleted.");
                                   }
                                   File.Move(file, @_StatementBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   zipFNList.Add(@_StatementBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   LogMessage("Statement file moved backup directory (" + StatementString + "): " + Path.GetFileName(file) + backupExtension);
                                   // Now copy Statement file from backup file to 
                                   if (File.Exists(@_StatementMoveDirectory + Path.GetFileName(file)))
                                   {
                                       if (_OverwriteMoveFiles)
                                       {
                                           File.Delete(@_StatementMoveDirectory + Path.GetFileName(file));
                                           LogMessage("Statement file already exists in Move directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                       }
                                       else
                                           uniqueExtension = backupExtension;
                                   }
                                   File.Copy(@_StatementBackupDirectory + Path.GetFileName(file) + backupExtension, @_StatementMoveDirectory + Path.GetFileName(file) + uniqueExtension);
                                   LogMessage("Statement file copied to move directory (" + StatementString + "): " + Path.GetFileName(file) + uniqueExtension);
                                   _StatementFiles++;
                                   _MovedFilenames.Add(file);
                               }
                               else
                                   tempSR.Close();
                           }
                           catch (Exception exp)
                           {
                               LogMessage("Statement File move / copy error (" + file + "):" + exp.Message);
                               _OtherFileError++;
                           }
                       }
                       else
                       {
                           if (!_LockedFilenames.Exists(element => element == file))
                           {
                               _LockedFilenames.Add(file);
                               LogMessage("Statement file found locked (" + Path.GetFileName(file) + "). Skipping file copy / move.");
                               _LockedFiles++;
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("Statement get files error: " + e.Message);
               return false;
           }

           if ((zipFNList.Count > 0) && (_ZipBackupFiles))
           {
               if (ZipBackupFile(@_StatementBackupDirectory +  ZipFN, zipFNList))
               {
                   LogMessage("Statement ZIP file created: " + @_StatementBackupDirectory + ZipFN + " - Number of File: " + zipFNList.Count.ToString()); 
               }
               else
               {
                   LogMessage("Issue generating ZIP Backup file: " + @_StatementBackupDirectory + ZipFN);
               }
           }

           // Generate Statement Semaphore 
           if (!String.IsNullOrWhiteSpace(_StatementSemaphoreFilename) && (_StatementFiles > 0))
           {
               if (IsValidFilepath(_StatementSemaphoreFilename))
               {
                   if (!File.Exists(@_StatementSemaphoreFilename))
                   {
                       // Should just create a blank file
                       using (StreamWriter fs = File.CreateText(@_StatementSemaphoreFilename)) { fs.Close(); }
                       LogMessage("Statement Semaphore File created: " + _StatementSemaphoreFilename);
                   }
               }
               else
                   LogMessage("Invalid Statement Semaphore Filename: " + _StatementSemaphoreFilename);
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
       }

       public bool CheckNoticeFiles()
       {
           List<string> zipFNList = new List<string>();
           DateTime zipCurrentDT = DateTime.Now;
           string zipExtension = GetDateTimeFileExtension(zipCurrentDT);
           string ZipFN = "NoticesBU" + zipExtension + ".zip";

           try
           {
               if ((files.Count > 0) && (!String.IsNullOrWhiteSpace(_NoticeSearchString)))
               {
                   foreach (string file in files)
                   {
                       // Open file as text file, read first line to see if a notice file
                       if ((!IsFileLocked(file)) && (!_LockedFilenames.Exists(element => element == file)))
                       {
                           string noticeString = "";
                           try
                           {
                               StreamReader tempSR = new StreamReader(file);
                               string lineOne = tempSR.ReadLine();
                               try
                               {
                                   if (lineOne.Length > 80)
                                       noticeString = LeftRightMid.Mid(lineOne, 41, 40).TrimEnd(' ');
                                   else
                                       noticeString = "";
                               }
                               catch (Exception exp)
                               {
                                   LogMessage("******* Issue looking for Notice String (" + file + "):" + exp.Message + " *******");
                                   noticeString = "";
                                   _OtherFileError++;
                                   SendEmail(_ErrorEMailAddress, "Episys PreProc - Notice Read Line Issue.", "The following file had an issue with reading line 1 for notice string. Please review.\r\n\r\nFilename: " + file);
                               }                               
                               if ((_NoticeSearchString.IndexOf(noticeString + ";") >= 0) && (noticeString != ""))
                               {
                                   tempSR.Close();
                                   // Move notice file to back up directory first
                                   DateTime currentDT = DateTime.Now;
                                   string backupExtension = GetDateTimeFileExtension(currentDT);
                                   string uniqueExtension = "";
                                   if (File.Exists(@_NoticeBackupDirectory + Path.GetFileName(file) + backupExtension))
                                   {
                                       File.Delete(@_NoticeBackupDirectory + Path.GetFileName(file) + backupExtension);
                                       LogMessage("Notice file already exists in Backup directory (" + Path.GetFileName(file) + backupExtension + "). Existing file being deleted.");
                                   }
                                   File.Move(file, @_NoticeBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   zipFNList.Add(@_NoticeBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   LogMessage("Notice file moved backup directory (" + noticeString + "): " + Path.GetFileName(file) + backupExtension);
                                   // Now copy notice file from backup file to 
                                   if (File.Exists(@_NoticeMoveDirectory + Path.GetFileName(file)))
                                   {
                                       if (_OverwriteMoveFiles)
                                       {
                                           File.Delete(@_NoticeMoveDirectory + Path.GetFileName(file));
                                           LogMessage("Notice file already exists in Move directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                       }
                                       else
                                           uniqueExtension = backupExtension;
                                   }
                                   File.Copy(@_NoticeBackupDirectory + Path.GetFileName(file) + backupExtension, @_NoticeMoveDirectory + Path.GetFileName(file) + uniqueExtension);
                                   LogMessage("Notice file copied to move directory (" + noticeString + "): " + Path.GetFileName(file) + uniqueExtension);
                                   _NoticeFiles++;
                                   _MovedFilenames.Add(file);
                               }
                               else
                                   tempSR.Close();
                            }
                            catch (Exception exp)
                            {
                                LogMessage("Notice File move / copy error (" + file + "):" + exp.Message);
                                _OtherFileError++;
                            }
                       }
                       else
                       {
                           if (!_LockedFilenames.Exists(element => element == file))
                           { 
                               _LockedFilenames.Add(file);
                               LogMessage("Notice file found locked (" + Path.GetFileName(file) + "). Skipping file copy / move.");
                               _LockedFiles++;
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("Notice get files error: " + e.Message);
               return false;
           }

           if ((zipFNList.Count > 0) && (_ZipBackupFiles))
           {
               if (ZipBackupFile(@_NoticeBackupDirectory + ZipFN, zipFNList))
               {
                   LogMessage("Notice ZIP file created: " + @_NoticeBackupDirectory + ZipFN + " - Number of File: " + zipFNList.Count.ToString());
               }
               else
               {
                   LogMessage("Issue generating ZIP Backup file: " + @_NoticeBackupDirectory + ZipFN);
               }
           }

           // Generate Notice Semaphore 
           if (!String.IsNullOrWhiteSpace(_NoticeSemaphoreFilename) && (_NoticeFiles > 0))
           {
               if (IsValidFilepath(_NoticeSemaphoreFilename))
               {
                   if (!File.Exists(@_NoticeSemaphoreFilename))
                   {
                       // Should just create a blank file
                       using (StreamWriter fs = File.CreateText(@_NoticeSemaphoreFilename)) { fs.Close(); }
                       LogMessage("Notice Semaphore File created: " + _NoticeSemaphoreFilename);
                   }
               }
               else
                   LogMessage("Invalid Notice Semaphore Filename: " + _NoticeSemaphoreFilename);
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
       }

       public bool CheckACHReportsFiles()
       {
           List<string> zipFNList = new List<string>();
           DateTime zipCurrentDT = DateTime.Now;
           string zipExtension = GetDateTimeFileExtension(zipCurrentDT);
           string ZipFN = "ACHReportsBU" + zipExtension + ".zip";
           try
           {
               // Only get files in source directory
               if ((files.Count > 0) && (!String.IsNullOrWhiteSpace(_ACHReportsSearchString)))
               {
                   foreach (string file in files)
                   {
                       // Open file as text file, read first line to see if a ACH Report file
                       if ((!IsFileLocked(file)) && (!_LockedFilenames.Exists(element => element == file)))
                       {
                           string blankACHReports = "";
                           string ACHReportsString = "";
                           try
                           {
                               StreamReader tempSR = new StreamReader(file);
                               string lineOne = tempSR.ReadLine();
                               try
                               {
                                   blankACHReports = LeftRightMid.Left(lineOne, 5);
                                   if (lineOne.Length > 80)
                                       ACHReportsString = LeftRightMid.Mid(lineOne, 41, 40).TrimEnd(' ');
                                   else
                                       ACHReportsString = "";
                               }
                               catch (Exception exp)
                               {
                                   LogMessage("******* Issue looking for ACH Report String (" + file + "):" + exp.Message + " *******");
                                   ACHReportsString = "";
                                   _OtherFileError++;
                                   SendEmail(_ErrorEMailAddress, "Episys PreProc - ACH Report Read Line Issue.", "The following file had an issue with reading line 1 for ACH Report string. Please review.\r\n\r\nFilename: " + file);
                               }
                               if (((_ACHReportsSearchString.IndexOf(ACHReportsString + ";") >= 0) && (ACHReportsString != "")) || (blankACHReports == "+    "))
                               {
                                   tempSR.Close();
                                   // Move ACH Report file to back up directory first
                                   DateTime currentDT = DateTime.Now;
                                   string backupExtension = GetDateTimeFileExtension(currentDT);
                                   string uniqueExtension = "";
                                   if (File.Exists(@_ACHReportsBackupDirectory + Path.GetFileName(file) + backupExtension))
                                   {
                                       File.Delete(@_ACHReportsBackupDirectory + Path.GetFileName(file) + backupExtension);
                                       LogMessage("ACH Report file already exists in Backup directory (" + Path.GetFileName(file) + backupExtension + "). Existing file being deleted.");
                                   }
                                   File.Move(file, @_ACHReportsBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   zipFNList.Add(@_ACHReportsBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   LogMessage("ACH Report file moved backup directory (" + ACHReportsString + "): " + Path.GetFileName(file) + backupExtension);
                                   // Now copy ACH Report file from backup file to 
                                   if (File.Exists(@_ACHReportsMoveDirectory + Path.GetFileName(file)))
                                   {
                                       if (_OverwriteMoveFiles)
                                       {
                                           File.Delete(@_ACHReportsMoveDirectory + Path.GetFileName(file));
                                           LogMessage("ACH Report file already exists in Move directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                       }
                                       else
                                           uniqueExtension = backupExtension;
                                   }
                                   File.Copy(@_ACHReportsBackupDirectory + Path.GetFileName(file) + backupExtension, @_ACHReportsMoveDirectory + Path.GetFileName(file) + uniqueExtension);
                                   LogMessage("ACH Report file copied to move directory (" + ACHReportsString + "): " + Path.GetFileName(file) + uniqueExtension);
                                   _ACHReportsFiles++;
                                   _MovedFilenames.Add(file);
                               }
                               else
                                   tempSR.Close();
                           }
                           catch (Exception exp)
                           {
                               LogMessage("ACH Report File move / copy error (" + file + "):" + exp.Message);
                               _OtherFileError++;
                           }
                       }
                       else
                       {
                           if (!_LockedFilenames.Exists(element => element == file))
                           {
                               _LockedFilenames.Add(file);
                               LogMessage("ACH Reports file found locked (" + Path.GetFileName(file) + "). Skipping file copy / move.");
                               _LockedFiles++;
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("ACH Report get files error: " + e.Message);
               return false;
           }

           if ((zipFNList.Count > 0) && (_ZipBackupFiles))
           {
               if (ZipBackupFile(@_ACHReportsBackupDirectory + ZipFN, zipFNList))
               {
                   LogMessage("ACH Report ZIP file created: " + @_ACHReportsBackupDirectory + ZipFN + " - Number of File: " + zipFNList.Count.ToString());
               }
               else
               {
                   LogMessage("Issue generating ZIP Backup file: " + @_ACHReportsBackupDirectory + ZipFN);
               }
           }

           // Generate ACH Report files Semaphore 
           if (!String.IsNullOrWhiteSpace(_ACHReportsSemaphoreFilename) && (_ACHReportsFiles > 0))
           {
               if (IsValidFilepath(_ACHReportsSemaphoreFilename))
               {
                   if (!File.Exists(@_ACHReportsSemaphoreFilename))
                   {
                       try
                       {
                           // Should just create a blank file
                           using (StreamWriter fs = File.CreateText(@_ACHReportsSemaphoreFilename)) { fs.Close(); }
                           LogMessage("ACH Report Semaphore File created: " + _ACHReportsSemaphoreFilename);
                           // Generate an email to notify ACH Reports have been preprocessed ... If blank email then don't send notification
                           if ((_ACHEmailNotificationAddress != "") && (IsValidEmail(_ACHEmailNotificationAddress)))
                               SendEmail(_ACHEmailNotificationAddress, "Episys PreProc - ACH Report Files Found", "The Episys Preprocessor processed at least one designated ACH Report.");
                       }
                       catch (Exception exp)
                       {
                           LogMessage("Invalid ACH Report Semaphore Filename: " + _ACHReportsSemaphoreFilename);
                           if (!IsValidEmail(_ErrorEMailAddress))
                               SendEmail(_ErrorEMailAddress, "Episys PreProc - Issue Creating ACH Report Semiphore File", "The Episys Preprocessor processed at least one designated ACH Report. However, there was an error creating the Semiphore file to signal file ready for COLD processing: " + exp.Message);
                       }
                   }
               }
               else
                   LogMessage("Invalid ACH Report Semaphore Filename: " + _ACHReportsSemaphoreFilename);
                   if (!IsValidEmail(_ErrorEMailAddress))
                       SendEmail(_ErrorEMailAddress, "Episys PreProc - Issue Creating ACH Report Semiphore File", "The Episys Preprocessor processed at least one designated ACH Report. However, there was an error creating the Semiphore file to signal file ready for COLD processing.");
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
       }

       public bool CheckBigFiles()
       {
           List<string> zipFNList = new List<string>();
           DateTime zipCurrentDT = DateTime.Now;
           string zipExtension = GetDateTimeFileExtension(zipCurrentDT);
           string ZipFN = "BigBU" + zipExtension + ".zip";

           try
           {
               if (files.Count > 0)
               {
                   foreach (string file in files)
                   {
                       FileInfo thisFileInfo = new FileInfo(file);
                       if (thisFileInfo.Length > Convert.ToInt32(_BigFileSize))
                       {
                           if ((!IsFileLocked(file)) && (!_LockedFilenames.Exists(element => element == file)))
                           {
                               try
                               {
                                   // Move big file to back up directory first
                                   DateTime currentDT = DateTime.Now;
                                   string backupExtension = GetDateTimeFileExtension(currentDT);
                                   string uniqueExtension = "";
                                   if (File.Exists(@_BigBackupDirectory + Path.GetFileName(file) + backupExtension))
                                   {
                                       File.Delete(@_BigBackupDirectory + Path.GetFileName(file) + backupExtension);
                                       LogMessage("Big file already exists in Backup directory (" + Path.GetFileName(file) + backupExtension + "). Existing file being deleted.");
                                   }
                                   File.Move(file,@_BigBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   zipFNList.Add(@_BigBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   LogMessage("Big file moved backup directory: " + Path.GetFileName(file) + backupExtension);
                                   // Now copy big file from backup file to 
                                   if (File.Exists(_BigMoveDirectory + Path.GetFileName(file)))
                                   {
                                       if (_OverwriteMoveFiles)
                                       {
                                           File.Delete(_BigMoveDirectory + Path.GetFileName(file));
                                           LogMessage("Big file already exists in process directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                       }
                                       else
                                           uniqueExtension = backupExtension;
                                   }
                                   File.Copy(@_BigBackupDirectory + Path.GetFileName(file) + backupExtension, _BigMoveDirectory + Path.GetFileName(file) + uniqueExtension);
                                   LogMessage("Big file copied to move directory: " + Path.GetFileName(file) + uniqueExtension);
                                   _BigFiles++;
                                   _MovedFilenames.Add(file);
                               }
                               catch (Exception exp)
                               {
                                   LogMessage("Big File move / copy error:" + exp.Message);
                                   _OtherFileError++;
                               }
                           }
                           else
                           {
                               if (!_LockedFilenames.Exists(element => element == file))
                               {
                                   _LockedFilenames.Add(file);
                                   LogMessage("Big file found locked (" + Path.GetFileName(file) + "). Skipping file copy / move.");
                                   _LockedFiles++;
                               }
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("Big files get files error: " + e.Message);
               return false;
           }

           if ((zipFNList.Count > 0) && (_ZipBackupFiles))
           {
               if (ZipBackupFile(@_BigBackupDirectory + ZipFN, zipFNList))
               {
                   LogMessage("Big Files ZIP file created: " + @_BigBackupDirectory + ZipFN + " - Number of File: " + zipFNList.Count.ToString());
               }
               else
               {
                   LogMessage("Issue generating ZIP Backup file: " + @_BigBackupDirectory + ZipFN);
               }
           }

           // Generate Big Files Semaphore 
           if (!String.IsNullOrWhiteSpace(_BigSemaphoreFilename) && (_BigFiles > 0))
           {
               if (IsValidFilepath(_BigSemaphoreFilename))
               {
                   if (!File.Exists(_BigSemaphoreFilename))
                   {
                       // Should just create a blank file
                       using (StreamWriter fs = File.CreateText(_BigSemaphoreFilename)) { fs.Close(); }
                       LogMessage("Big Files Semaphore File created: " + _BigSemaphoreFilename);
                   }
               }
               else
                   LogMessage("Invalid Big Files Semaphore Filename: " + _BigSemaphoreFilename);
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
       }

       public bool CheckDefaultFiles()
       {
           List<string> zipFNList = new List<string>();
           DateTime zipCurrentDT = DateTime.Now;
           string zipExtension = GetDateTimeFileExtension(zipCurrentDT);
           string ZipFN = "DefaultBU" + zipExtension + ".zip";

           try
           {
               if (files.Count > 0)
               {
                   foreach (string file in files)
                   {
                       if ((!IsFileLocked(file)) && (!_LockedFilenames.Exists(element => element == file)))
                       {
                           try
                           {
                               // Move big file to back up directory first
                               DateTime currentDT = DateTime.Now;
                               string backupExtension = GetDateTimeFileExtension(currentDT);
                               string uniqueExtension = "";
                               if (File.Exists(@_DefaultBackupDirectory + Path.GetFileName(file) + backupExtension))
                               {
                                   File.Delete(@_DefaultBackupDirectory + Path.GetFileName(file) + backupExtension);
                                   LogMessage("Default file already exists in Backup directory (" + Path.GetFileName(file) + backupExtension + "). Existing file being deleted.");
                               }
                               File.Move(file, @_DefaultBackupDirectory + Path.GetFileName(file) + backupExtension);
                               zipFNList.Add(@_DefaultBackupDirectory + Path.GetFileName(file) + backupExtension);
                               LogMessage("Default file moved backup directory: " + Path.GetFileName(file) + backupExtension);
                               // Now copy big file from backup file to 
                               if (File.Exists(_DefaultMoveDirectory + Path.GetFileName(file)))
                               {
                                   if (_OverwriteMoveFiles)
                                   {
                                       File.Delete(_DefaultMoveDirectory + Path.GetFileName(file));
                                       LogMessage("Default file already exists in process directory (" + Path.GetFileName(file) + "). Existing file being deleted.");
                                   }
                                   else
                                       uniqueExtension = backupExtension;
                               }
                               File.Copy(@_DefaultBackupDirectory + Path.GetFileName(file) + backupExtension, _DefaultMoveDirectory + Path.GetFileName(file) + uniqueExtension);
                               LogMessage("Default file copied to move directory: " + Path.GetFileName(file) + uniqueExtension);
                              _DefaultFiles++;
                              _MovedFilenames.Add(file);
                           }
                           catch (Exception exp)
                           {
                               LogMessage("Default File move / copy error:" + exp.Message);
                               _OtherFileError++;
                           }
                       }
                       else
                       {
                           if (!_LockedFilenames.Exists(element => element == file))
                           {
                               _LockedFilenames.Add(file);
                               LogMessage("Default file found locked (" + Path.GetFileName(file) + "). Skipping file copy / move.");
                               _LockedFiles++;
                           }
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("Default files get files error: " + e.Message);
               return false;
           }

           if ((zipFNList.Count > 0) && (_ZipBackupFiles))
           {
               if (ZipBackupFile(@_DefaultBackupDirectory + ZipFN, zipFNList))
               {
                   LogMessage("Default Files ZIP file created: " + @_DefaultBackupDirectory + ZipFN + " - Number of File: " + zipFNList.Count.ToString());
               }
               else
               {
                   LogMessage("Issue generating ZIP Backup file: " + @_DefaultBackupDirectory + ZipFN);
               }
           }

           // Generate Default Files Semaphore 
           if (!String.IsNullOrWhiteSpace(_DefaultSemaphoreFilename) && (_DefaultFiles > 0))
           {
               if (IsValidFilepath(_DefaultSemaphoreFilename))
               {
                   if (!File.Exists(_DefaultSemaphoreFilename))
                   {
                       // Should just create a blank file
                       using (StreamWriter fs = File.CreateText(_DefaultSemaphoreFilename)) { fs.Close(); }
                       LogMessage("Default Files Semaphore File created: " + _DefaultSemaphoreFilename);
                   }
               }
               else
                   LogMessage("Invalid Default Files Semaphore Filename: " + _DefaultSemaphoreFilename);
           }
           // Remove the files that were moved from current files list
           RemoveMovedFiles();
           return true;
       }

       public void RemoveMovedFiles()
       {
           foreach (var item in _MovedFilenames)
               files.Remove(item);
           _MovedFilenames.Clear();
       }

        private string ReadSetting(string key)
        {
            try
            {
                var appSettings = System.Configuration.ConfigurationSettings.AppSettings;
                string result = appSettings[key] ?? "Not Found";
                if (_LogAppSettings)
                    LogMessage("Using AppSetting (" + key + ") : " + result);
                return result;
            }
            catch (System.Configuration.ConfigurationException exp)
            {
                LogMessage(exp.Message);
                return "ERROR";
            }
        }

        public void LogMessage(string logEntry)
        {
            if (_LoggingOn)
            {
                string Line = DateTime.Now.ToShortDateString() + ":" + DateTime.Now.ToShortTimeString() + ":" + logEntry;
                _LogFileWriterStream.WriteLine(Line);
                _LogFileWriterStream.Flush();
            }
            if (_OutputConsole)
                Console.WriteLine(DateTime.Now.ToShortDateString() + ":" + DateTime.Now.ToShortTimeString() + ":" + logEntry);
        }

        public void SendEmail(string ToAddress, string Subject, string Body)
        {
            if (_GenerateEmail)
            {
                try
                {
                    //// Generate email using Outlook ... requires Trust option setting to allow external program to email
                    //Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
                    //Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    //mailItem.Subject = Subject;
                    //mailItem.To = ToAddress;
                    //mailItem.Body = Body;
                    //mailItem.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceHigh;
                    //mailItem.Send();
                    ////mailItem.Display(false);
                    string thisStartTime = _EmailStartTime;
                    string thisEndTime = _EmailEndTime;
                    string thisEmailDays = _EmailDays;
                    DateTime testDT = new DateTime(2015, 6, 18, Convert.ToInt16(LeftRightMid.Left(thisStartTime, 2)), Convert.ToInt16(LeftRightMid.Right(thisStartTime, 2)), 0);
                    string testStr = testDT.ToString("ddd");
                    // Check to see if today is a day we are to generated emails
                    if (thisEmailDays.IndexOf(DateTime.Now.ToString("ddd") + ";") != -1) 
                    {
                        // Check to see if config file contains start and end times for emails
                        if ((thisStartTime != "") && (thisEndTime != ""))
                        {
                            // Get current date and time and then create datatime object for start time and end time using today
                            DateTime tempDT = DateTime.Now;
                            DateTime tempStartDT = new DateTime(tempDT.Year,tempDT.Month, tempDT.Day, Convert.ToInt16(LeftRightMid.Left(thisStartTime,2)), Convert.ToInt16(LeftRightMid.Right(thisStartTime,2)), 0);
                            DateTime tempEndDT = new DateTime(tempDT.Year,tempDT.Month, tempDT.Day, Convert.ToInt16(LeftRightMid.Left(thisEndTime,2)), Convert.ToInt16(LeftRightMid.Right(thisEndTime,2)), 0);
                            // Check if current date time is greater than start and less end time
                            if ((tempDT >= tempStartDT) && (tempDT <= tempEndDT))
                            {
                                // Generate email using SMTP and Suncoast server
                                
//                                MailAddress from = new MailAddress("NAUTILUS@suncoastcreditunion.com", "OnBase Nautilus");
                                MailAddress from = new MailAddress(_SourceEMailAddress, "OnBase Issues");
                                MailAddress to = new MailAddress(ToAddress, "OnBase COLD Process Admin");
                                MailMessage m = new MailMessage(from, to);
                                m.Subject = Subject;
                                m.Body = Body;
                                m.IsBodyHtml = true;
                                SmtpClient smtp = new SmtpClient();
                                smtp.Host = "mailrelay.ssfcu.inet";
                                //NetworkCredential authinfo = new NetworkCredential("mailidfrom", "YourPassword");
                                smtp.UseDefaultCredentials = true;
                                smtp.Send(m);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    LogMessage("Error generating Email (" + e.Message + "): To:" + ToAddress + ") - " + Subject + ":" + Body);
                }
            }
            return;
        }

        bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool IsValidPath(string path)
        {
//            Regex driveCheck = new Regex(@"^[a-zA-Z]:\\$");
            Regex driveCheck = new Regex(@"^(?:[\w]\:|\\)");
            if (!driveCheck.IsMatch(path.Substring(0, 3))) return false;
            string strTheseAreInvalidFileNameChars = new string(Path.GetInvalidPathChars());
            strTheseAreInvalidFileNameChars += @":/?*" + "\"";
            Regex containsABadCharacter = new Regex("[" + Regex.Escape(strTheseAreInvalidFileNameChars) + "]");
            if (containsABadCharacter.IsMatch(path.Substring(3, path.Length - 3)))
                return false;

            DirectoryInfo dir = new DirectoryInfo(Path.GetFullPath(path));
            if (!dir.Exists)
                dir.Create();
            return true;
        }

        public static bool IsValidFilepath(string path)
        {
            if (path.Trim() == string.Empty)
            {
                return false;
            }

            string pathname;
            string filename;
            try
            {
                pathname = Path.GetPathRoot(path);
                filename = Path.GetFileName(path);
            }
            catch (ArgumentException)
            {
                // GetPathRoot() and GetFileName() above will throw exceptions
                // if pathname/filename could not be parsed.

                return false;
            }

            // Make sure the filename part was actually specified
            if (filename.Trim() == string.Empty)
            {
                return false;
            }

            // Not sure if additional checking below is needed, but no harm done
            if (pathname.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            {
                return false;
            }

            if (filename.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            {
                return false;
            }

            return true;
        }

       private void PruneLogFiles(string logPath, Int32 logFileCount)
       {
           try
           {
               // Only get files that begin with the letter "c."
               string[] files = Directory.GetFiles(@logPath, "PreprocLog-*");
               if (files.Length >= logFileCount)
               {
                   int filecount = files.Length;
                   foreach (string file in files)
                   {
                       filecount--;
                       File.Delete(file);
                       if (filecount == logFileCount)
                           break;
                   }
               }
           }
           catch (Exception e)
           {
               //Console.WriteLine("The process failed: {0}", e.ToString());
           }
       }

       public void PruneBackupFiles(string BackupPath, Int32 KeepNumberOfDays)
       {
           try
           {
               // Only get files that begin with the letter "c."
               string[] files = Directory.GetFiles(BackupPath, "*");
               if (files.Length > 0)
               {
                   foreach (string file in files)
                   {
                       FileInfo thisFileInfo = new FileInfo(file);
                       DateTime thisFileDT = thisFileInfo.CreationTime;
                       DateTime cutOffDT = DateTime.Now.AddDays(((Double)KeepNumberOfDays) * -1.0);
                       if (thisFileDT < cutOffDT)
                       {
                           File.Delete(file);
                           LogMessage("Deleting Backup File: " + file);
                       }
                   }
               }
           }
           catch (Exception e)
           {
               LogMessage("Pruning Backup files get files error: " + e.Message);
           }
       }


       public bool IsFileLocked(string filePath)
       {
           try
           {
               using (File.Open(filePath, FileMode.Open)) { }
           }
           catch (IOException e)
           {
               var errorCode = Marshal.GetHRForException(e) & ((1 << 16) - 1);
               List<Process> lockedProcesses;
               lockedProcesses = FileUtil.WhoIsLocking(filePath);
               LogMessage("Locked File Exception Filename: " + filePath + " ... " + e.Message + " - ErrorCode:" + errorCode.ToString() + "- Process Count:" + lockedProcesses.Count.ToString());
               foreach (var item in lockedProcesses)
               {
                   LogMessage("****** File Locked By:" + item.ProcessName);
               }
               return errorCode == 32 || errorCode == 33;
           }

           return false;
       }

       private string GetDateTimeFileExtension(DateTime dt)
       {
           string temp = ".";
           temp += LeftRightMid.Right("0" + dt.Month.ToString(), 2);
           temp += LeftRightMid.Right("0" + dt.Day.ToString(), 2);
           temp += dt.Year.ToString();
           temp += "-" + LeftRightMid.Right("0" + dt.Hour.ToString(), 2);
           temp += LeftRightMid.Right("0" + dt.Minute.ToString(), 2);
           temp += LeftRightMid.Right("0" + dt.Second.ToString(), 2);
           temp += LeftRightMid.Right("000" + dt.Millisecond.ToString(), 4);
           return temp;
       }

       private bool ZipBackupFile(string ZipFilename, List<string> ZipFilenameList)
       {
           if ((ZipFilenameList.Count > 0) && (!String.IsNullOrWhiteSpace(ZipFilename)))
           {
               foreach (var item in ZipFilenameList)
               {
                   AddFileToZip(ZipFilename, item);
                   File.Delete(item);
               }
           }
           return true;
       }

       private const long BUFFER_SIZE = 4096;

       private static void AddFileToZip(string zipFilename, string fileToAdd)
       {
           using (Package zip = System.IO.Packaging.Package.Open(zipFilename, FileMode.OpenOrCreate))
           {
               string destFilename = ".\\" + Path.GetFileName(fileToAdd);
               Uri uri = PackUriHelper.CreatePartUri(new Uri(destFilename, UriKind.Relative));
               if (zip.PartExists(uri))
               {
                   zip.DeletePart(uri);
               }
               PackagePart part = zip.CreatePart(uri, "", CompressionOption.Normal);
               using (FileStream fileStream = new FileStream(fileToAdd, FileMode.Open, FileAccess.Read))
               {
                   using (Stream dest = part.GetStream())
                   {
                       CopyStream(fileStream, dest);
                   }
               }
           }
       }

       private static void CopyStream(System.IO.FileStream inputStream, System.IO.Stream outputStream)
       {
           long bufferSize = inputStream.Length < BUFFER_SIZE ? inputStream.Length : BUFFER_SIZE;
           byte[] buffer = new byte[bufferSize];
           int bytesRead = 0;
           long bytesWritten = 0;
           while ((bytesRead = inputStream.Read(buffer, 0, buffer.Length)) != 0)
           {
               outputStream.Write(buffer, 0, bytesRead);
               bytesWritten += bufferSize;
           }
       }
    }

   class LeftRightMid
   {
       /// <summary>
       /// The main entry point for the application.
       /// </summary>

       public static string Left(string param, int length)
       {
           //we start at 0 since we want to get the characters starting from the
           //left and with the specified lenght and assign it to a variable
           string result = param.Substring(0, length);
           //return the result of the operation
           return result;
       }

       public static string Right(string param, int length)
       {
           //start at the index based on the lenght of the sting minus
           //the specified lenght and assign it a variable
           int temp = param.Length - length;
           string result = param.Substring(temp, length);
           //return the result of the operation
           return result;
       }

       public static string Mid(string param, int startIndex, int length)
       {
           //start at the specified index in the string ang get N number of
           //characters depending on the lenght and assign it to a variable
           string result = param.Substring(startIndex, length);
           //return the result of the operation
           return result;
       }

       public static string Mid(string param, int startIndex)
       {
           //start at the specified index and return all characters after it
           //and assign it to a variable
           string result = param.Substring(startIndex);
           //return the result of the operation
           return result;
       }

   }

   static public class FileUtil
   {
       [StructLayout(LayoutKind.Sequential)]
       struct RM_UNIQUE_PROCESS
       {
           public int dwProcessId;
           public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime;
       }

       const int RmRebootReasonNone = 0;
       const int CCH_RM_MAX_APP_NAME = 255;
       const int CCH_RM_MAX_SVC_NAME = 63;

       enum RM_APP_TYPE
       {
           RmUnknownApp = 0,
           RmMainWindow = 1,
           RmOtherWindow = 2,
           RmService = 3,
           RmExplorer = 4,
           RmConsole = 5,
           RmCritical = 1000
       }

       [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
       struct RM_PROCESS_INFO
       {
           public RM_UNIQUE_PROCESS Process;

           [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_APP_NAME + 1)]
           public string strAppName;

           [MarshalAs(UnmanagedType.ByValTStr, SizeConst = CCH_RM_MAX_SVC_NAME + 1)]
           public string strServiceShortName;

           public RM_APP_TYPE ApplicationType;
           public uint AppStatus;
           public uint TSSessionId;
           [MarshalAs(UnmanagedType.Bool)]
           public bool bRestartable;
       }

       [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
       static extern int RmRegisterResources(uint pSessionHandle,
                                             UInt32 nFiles,
                                             string[] rgsFilenames,
                                             UInt32 nApplications,
                                             [In] RM_UNIQUE_PROCESS[] rgApplications,
                                             UInt32 nServices,
                                             string[] rgsServiceNames);

       [DllImport("rstrtmgr.dll", CharSet = CharSet.Auto)]
       static extern int RmStartSession(out uint pSessionHandle, int dwSessionFlags, string strSessionKey);

       [DllImport("rstrtmgr.dll")]
       static extern int RmEndSession(uint pSessionHandle);

       [DllImport("rstrtmgr.dll")]
       static extern int RmGetList(uint dwSessionHandle,
                                   out uint pnProcInfoNeeded,
                                   ref uint pnProcInfo,
                                   [In, Out] RM_PROCESS_INFO[] rgAffectedApps,
                                   ref uint lpdwRebootReasons);

       /// <summary>
       /// Find out what process(es) have a lock on the specified file.
       /// </summary>
       /// <param name="path">Path of the file.</param>
       /// <returns>Processes locking the file</returns>
       /// <remarks>See also:
       /// http://msdn.microsoft.com/en-us/library/windows/desktop/aa373661(v=vs.85).aspx
       /// http://wyupdate.googlecode.com/svn-history/r401/trunk/frmFilesInUse.cs (no copyright in code at time of viewing)
       /// 
       /// </remarks>
       static public List<Process> WhoIsLocking(string path)
       {
           uint handle;
           string key = Guid.NewGuid().ToString();
           List<Process> processes = new List<Process>();

           int res = RmStartSession(out handle, 0, key);
           if (res != 0) throw new Exception("Could not begin restart session.  Unable to determine file locker.");

           try
           {
               const int ERROR_MORE_DATA = 234;
               uint pnProcInfoNeeded = 0,
                    pnProcInfo = 0,
                    lpdwRebootReasons = RmRebootReasonNone;

               string[] resources = new string[] { path }; // Just checking on one resource.

               res = RmRegisterResources(handle, (uint)resources.Length, resources, 0, null, 0, null);

               if (res != 0) throw new Exception("Could not register resource.");

               //Note: there's a race condition here -- the first call to RmGetList() returns
               //      the total number of process. However, when we call RmGetList() again to get
               //      the actual processes this number may have increased.
               res = RmGetList(handle, out pnProcInfoNeeded, ref pnProcInfo, null, ref lpdwRebootReasons);

               if (res == ERROR_MORE_DATA)
               {
                   // Create an array to store the process results
                   RM_PROCESS_INFO[] processInfo = new RM_PROCESS_INFO[pnProcInfoNeeded];
                   pnProcInfo = pnProcInfoNeeded;

                   // Get the list
                   res = RmGetList(handle, out pnProcInfoNeeded, ref pnProcInfo, processInfo, ref lpdwRebootReasons);
                   if (res == 0)
                   {
                       processes = new List<Process>((int)pnProcInfo);

                       // Enumerate all of the results and add them to the 
                       // list to be returned
                       for (int i = 0; i < pnProcInfo; i++)
                       {
                           try
                           {
                               processes.Add(Process.GetProcessById(processInfo[i].Process.dwProcessId));
                           }
                           // catch the error -- in case the process is no longer running
                           catch (ArgumentException) { }
                       }
                   }
                   else throw new Exception("Could not list processes locking resource.");
               }
               else if (res != 0) throw new Exception("Could not list processes locking resource. Failed to get size of result.");
           }
           finally
           {
               RmEndSession(handle);
           }

           return processes;
       }
   }


}
