using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Diagnostics;
using System.IO;
namespace GoogleResponseDL
{
    class Program
    {
        static GoogleSheetsHelper gsh;
        static GoogleDriveHelper gdh;
        static string ApplicationName = "GoogleDriveSubmissionDL";
        private static string GoogleServiceAccountCredentialJSON = "venuwebservice-f70ca289fa6e.json";
        private static string googleDriveCred = "nodal-crawler-329222-58978613ac0e.json";
        private static string googleFolderID;
        private static string googleSheetID;
        static string submissionDL;
        static string filePath;
        private static List<System.Dynamic.ExpandoObject> submissionList;
        // uniquely identify each submission by edit url ID
        // Down in folder under submissionDL named: edit url id
        static async Task DownloadSubmissions()
        {
            try
            {
                if (gdh == null)
                    gdh = new GoogleDriveHelper(googleDriveCred);
                string qparam = "parents in '" + googleFolderID + "'";
                //Folder Name stuff
                string subFolderName;
                string uniqFolderName;
                //Uniq ID for folders
                string subUniqID;
                string[] subUniqIDSplit;
                //Data from sheet
                string cell;
                string cellValue;
                //Downloading image
                string[] latestUrl;
                string imageID;
                string imageName;
                //Downloading Text
                string newFilePath;

                //Resizing Images
                string lowerRes = "LowerRes";

                string input;
                bool defID = false;
                Console.WriteLine("Enter name of subfolder to download files to:");
                subFolderName = Console.ReadLine();
                Console.WriteLine("Enter Column name you wish to uniquely identify each response by");
                Console.WriteLine("(By Default, will use Edit Url if you press 'Enter'):");
                input = Console.ReadLine();
                if (String.IsNullOrEmpty(input))
                {
                    input = "Edit Url";
                    defID = true;
                }
                foreach (var submission in submissionList)
                {
                    IDictionary<string, object> fieldValues = submission;
                    var uniqPath = "";
                    
                    subUniqID = fieldValues[input] as string;
                    if (defID || input.Equals("Edit Url"))
                    {
                        subUniqIDSplit = subUniqID.Split('=');
                        uniqFolderName = subUniqIDSplit[subUniqIDSplit.Length - 1];
                    }
                    else
                    {
                        uniqFolderName = ReplaceInvalidChars(subUniqID);
                    }
                    uniqPath = System.IO.Path.Combine(filePath, subFolderName, uniqFolderName);
                    var uniqPathInfo = System.IO.Directory.CreateDirectory(uniqPath);
                    
                    foreach (var field in fieldValues.Keys)
                    {
                        cell = field as string;
                        cellValue = fieldValues[field] as string;
                        // images
                        if (cellValue.StartsWith("http") && cellValue.Contains("drive.google"))
                        {
                            await Task.Run(() =>
                            {
                                latestUrl = cellValue.Split(',');
                                imageID = latestUrl[latestUrl.Length - 1].Split('=')[1];
                                imageName = ReplaceInvalidChars(cell);
                                Console.WriteLine("Downloading " + imageName);
                                gdh.downloadFile(imageID, imageName, uniqPath);
                            });
                        }
                        else if (cell.ToLower().Contains("youtube")) //mp4
                        {
                            await Task.Run(() =>
                            {
                                cell = ReplaceInvalidChars(cell);
                                string command = " -f 22 -o \"" + uniqPath + "\\" + cell + ".mp4\" " + cellValue;
                                Console.WriteLine("Command: " + command);
                                Process ytProcess = System.Diagnostics.Process.Start("yt-dlp.exe", command);
                                ytProcess.WaitForExit();
                            });
                        }
                        else // text
                        {
                            await Task.Run(() =>
                            {
                                cell = ReplaceInvalidChars(cell);
                                newFilePath = System.IO.Path.Combine(uniqPath, cell + ".txt");
                                Console.WriteLine("Writing text to " + newFilePath);
                                cellValue = cellValue.Replace(' ', '_');
                                System.IO.File.WriteAllText(newFilePath, cellValue);

                            });
                        }
                    }
                    //Wait 5 seconds
                    //await Task.Delay(5 * 1000);

                    var resizeOutput = System.IO.Path.Combine(uniqPath, lowerRes); 
                    var resizeOutputInfo = System.IO.Directory.CreateDirectory(resizeOutput);

                    //Runspace runspace = RunspaceFactory.CreateRunspace();
                    //runspace.Open();

                    
                    var psFile1 = "resizeMore1200Auto.ps1";
                    var psFile2 = "resize640Auto.ps1";
                    var psFile3 = "resize320Auto.ps1";
                    string parameters1 = string.Format("-FILE \"{0}\" -input_folder \"{1}\" -output_folder \"{2}\"", psFile1, uniqPathInfo.FullName, resizeOutputInfo.FullName);
                    string parameters2 = string.Format("-FILE \"{0}\" -input_folder \"{1}\" -output_folder \"{2}\"", psFile2, uniqPathInfo.FullName, resizeOutputInfo.FullName);
                    string parameters3 = string.Format("-FILE \"{0}\" -input_folder \"{1}\" -output_folder \"{2}\"", psFile3, uniqPathInfo.FullName, resizeOutputInfo.FullName);

                    //StreamWriter sw = new StreamWriter("file.bat");
                    //sw.WriteLine("powershell " + parameters1);
                    //sw.WriteLine("powershell " + parameters2);
                    //sw.WriteLine("powershell " + parameters3);
                    //sw.Dispose();
                    //Process powershell = System.Diagnostics.Process.Start("file.bat");
                    //powershell.WaitForExit();
                    //powershell.Dispose();

                    string cmdPara = "/c powershell " + parameters1 + "&powershell " + parameters2 + "&powershell " + parameters3;
                    Console.WriteLine("CmdPara: " + cmdPara);
                    ProcessStartInfo StartInfo;
                    StartInfo = new ProcessStartInfo("cmd.exe", "/c powershell " + parameters1 + "&powershell " + parameters2 + "&powershell " + parameters3)
                    {
                        //Arguments = $"-NoProfile -ExecutionPolicy ByPass",
                        UseShellExecute = false
                    };

                    Process powershell = Process.Start(StartInfo);
                    powershell.WaitForExit();

                    string awsImgParam = string.Format("s3 cp . s3://venuassets/{0}/{1}/{2} --recursive --exclude \"*\" --include \"*.png\"", submissionDL, subFolderName, uniqFolderName);
                    StartInfo = new ProcessStartInfo("aws.exe", awsImgParam)
                    {
                        //Arguments = $"-NoProfile -ExecutionPolicy ByPass",
                        UseShellExecute = false,
                        WorkingDirectory = resizeOutputInfo.FullName
                    };

                    Process awsUpload = Process.Start(StartInfo);
                    awsUpload.WaitForExit();

                    string awsVidParam = string.Format("s3 cp . s3://venuassets/{0}/{1}/{2} --recursive --exclude \"*\" --include \"*.mp4\"", submissionDL, subFolderName, uniqFolderName);
                    StartInfo = new ProcessStartInfo("aws.exe", awsVidParam)
                    {
                        //Arguments = $"-NoProfile -ExecutionPolicy ByPass",
                        UseShellExecute = false,
                        WorkingDirectory = uniqPathInfo.FullName
                    };

                    Process awsVidUpload = Process.Start(StartInfo);
                    awsVidUpload.WaitForExit();

                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to download files: " + e);
            }
            
            return;
        }

        static String ReplaceInvalidChars(String orig)
        {
            String newString = orig;
            newString = newString.Replace(' ', '_');
            newString = newString.Replace('/', '_');
            newString = newString.Replace('|', 'l');
            return newString;
        }
        // Read in data from Google Sheets
        // Converts to JSON string
        // Download Submissions

        static async Task getGoogleSheetsSubmissions()
        {
            int numColumns,numRows = 0;
            int responseNum = 0;
            try
            {
                gsh = new GoogleSheetsHelper(GoogleServiceAccountCredentialJSON, googleSheetID);
                // Parameters for Artist
                Console.WriteLine("Enter number of columns in sheet:");
                numColumns = Int32.Parse(Console.ReadLine());
                Console.WriteLine("Enter number of rows in sheet:");
                numRows = Int32.Parse(Console.ReadLine());
                Console.WriteLine("Enter Form Response number:");
                responseNum = Int32.Parse(Console.ReadLine());
                GoogleSheetParameters gsp = new GoogleSheetParameters();
                gsp = new GoogleSheetParameters() {
                    RangeColumnStart = 1,
                    RangeRowStart = 1,
                    RangeColumnEnd = numColumns,
                    RangeRowEnd = numRows,
                    FirstRowIsHeaders = true,
                    SheetName = "Form Responses " + responseNum
                };
                  
                await Task.Run( () =>
                {
                    Console.WriteLine("Reading Google Sheet Information from " + googleSheetID);
                    submissionList = gsh.GetDataFromSheet(gsp);
                    string sheetCacheJson = JsonConvert.SerializeObject(submissionList);
                    Console.WriteLine("\nRetrieved Sheets JSON: " + sheetCacheJson.Length + " Rows:" + submissionList.Count);
                });
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to get submissions: " + e);
            }
            return;
        }

        static void Main(string[] args)
        {
            // Create submissions folder if it does not exist
            Console.WriteLine("Enter directory or name of folder to contain downloads in:");
            submissionDL = Console.ReadLine();
            filePath = Environment.ExpandEnvironmentVariables(submissionDL);
            System.IO.Directory.CreateDirectory(filePath);
            bool moreSheets = true;
            bool validResp = false;
            string response;
            
            while (moreSheets)
            {
                //Get Google Sheet Submission information
                Console.WriteLine("Enter the Google sheet ID:");
                googleSheetID = Console.ReadLine();
                getGoogleSheetsSubmissions().Wait();

                //Download Files
                Console.WriteLine("Enter the folder ID of the Google Drive folder containing the submission files:");
                googleFolderID = Console.ReadLine();
                DownloadSubmissions().Wait();
                Console.WriteLine("finished");
                Console.WriteLine("Would you like to download information from another Google Sheet?(Y/N) ");
                response = Console.ReadLine();
                
                while (!validResp)
                {
                    if (response == "Y" || response == "y" || response == "N" || response == "n")
                    {
                        validResp = true;
                        break;
                    }
                    Console.WriteLine("Invalid response please try again");
                    Console.WriteLine("Would you like to download information from another Google Sheet?(Y/N) ");
                }
                if (response == "Y" || response == "y") moreSheets = true;
                if (response == "N" || response == "n") moreSheets = false;
            }
            Console.WriteLine("Press 'Enter' to exit");
            Console.ReadKey();

        }
    }
}
