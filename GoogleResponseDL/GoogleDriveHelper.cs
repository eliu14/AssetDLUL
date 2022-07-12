using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleResponseDL
{

    class GoogleDriveHelper
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.ScopeConstants.Drive };
        static string ApplicationName = "GoogleDriveSubmissionDL";

        private readonly DriveService _driveService;

        public GoogleDriveHelper(string credentialFileName)
        {
            var fs = new FileStream(credentialFileName, FileMode.Open);
            var credential = GoogleCredential.FromStream(fs).CreateScoped(Scopes);

            // Create Drive API service.
            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }
        public IList<Google.Apis.Drive.v3.Data.File> ListFilesWithParameter(string qParam)
        {
            // Define parameters of request.
            FilesResource.ListRequest listRequest = _driveService.Files.List();
            listRequest.Q = qParam;
            //listRequest.Fields = "nextPageToken, files(id, name)";

            // List files.
            IList<Google.Apis.Drive.v3.Data.File> files = listRequest.Execute()
                .Files;
            Console.WriteLine("Files:");
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    Console.WriteLine("{0} ({1}) {2}", file.Name, file.Id, file.MimeType);
                }
            }
            else
            {
                Console.WriteLine("No files found.");
            }
            Console.Read();
            return files;
        }
        public async void downloadFile(string fileID, string fileName, string dlPath)
        {
            var getRequest = _driveService.Files.Get(fileID);
            var file = getRequest.Execute();
            var imgNameSplit = file.Name.Split('.');
            string imgExt = '.' + imgNameSplit[imgNameSplit.Length - 1];
            var fileStream = new FileStream(System.IO.Path.Combine(dlPath,fileName + imgExt), FileMode.Create, FileAccess.Write);
            var dlStatus = getRequest.DownloadWithStatus(fileStream);
            while (dlStatus.Status == Google.Apis.Download.DownloadStatus.Downloading)
            {
                Console.Write(".");
            }
            fileStream.Close();
        }
    }
    
}