using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util;
using System;
using System.Diagnostics.SymbolStore;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.InteropServices;
using static Google.Apis.Sheets.v4.SpreadsheetsResource;

namespace GoogleSheetApi
{
    public class SpreadSheet
    {
        //public methods for reading data

        public async Task<Dictionary<string, Dictionary<string, Dictionary<string, string>>>> GetSUICD()
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> allSheetRows = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            var names = await GetTabNames();

            foreach (var name in names)
            {
                var sheet = new Sheet(Id, name);
                var dic = await sheet.GetUICD();
                allSheetRows[name] = dic;
                _sheets.Add(sheet);
            }

            _currentRevision = await GetLatestRevision();
            return allSheetRows;
        }

        public async Task<Dictionary<string, Dictionary<string, Dictionary<string, string>>>> GetDeltaSUICD()
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> delta = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            bool checkForUpdate = await CheckIfUpdateNeeded();

            if (!checkForUpdate)
            {
                Console.WriteLine("No need for updates");
                return delta;
            }

            Console.WriteLine("updated successfully");

            foreach (var sheet in _sheets)
            {
                var headers = sheet.GetHeaders();
                var rowKeys = sheet.GetRowKeys();
                var currentDic = sheet.GetCurrentSheet();
                var newDic = await sheet.GetUICD();
                var valueDic = new Dictionary<string, Dictionary<string, string>>();
                foreach (var key in rowKeys)
                {
                    if (!currentDic.ContainsKey(key))
                    {
                        valueDic[key] = newDic[key];
                        continue;
                    }

                    if (currentDic[key]["StateId"] == newDic[key]["StateId"])
                    {
                        valueDic[key] = newDic[key];
                        delta[sheet.GetName()] = valueDic;
                    }
                }
            }
            return delta;
        }

        public List<Sheet> GetSheets() => _sheets;

        public static string Name = string.Empty;
        public static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        public static readonly string GoogleCredentialsFileName = "C:/Credentials/google-credentials.json";
        public static string Id = string.Empty;
        private string _currentRevision = string.Empty;
        private static List<Sheet> _sheets = new List<Sheet>();


        #region FileSysytem
        public static Google.Apis.Drive.v3.Data.FileList ListFiles(FilesListOptionalParms optional = null)
        {
            try
            {
                if (DriveService == null)
                    throw new ArgumentNullException("service");

                var request = DriveService.Files.List();
                request = (FilesResource.ListRequest)SampleHelpers.ApplyOptionalParms(request, optional);
                request.PageSize = 1000;
                return request.Execute();
            }
            catch (Exception ex)
            {
                throw new Exception("Request Files.List failed.", ex);
            }
        }

        static DriveService _driveService;
        public static DriveService DriveService
        {
            get
            {
                if (_driveService == null)
                {
                    using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
                    {
                        string[] scopes = new string[] { SheetsService.Scope.Drive };
                        var serviceInitializer = new BaseClientService.Initializer
                        {
                            HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(scopes)
                        };
                        _driveService = new DriveService(serviceInitializer);
                    }
                }
                return _driveService;
            }


        }

        public class FilesListOptionalParms
        {
            /// 

            /// The source of files to list.
            /// 

            public string Corpus { get; set; }
            /// 

            /// A comma-separated list of sort keys. Valid keys are 'createdTime', 'folder', 'modifiedByMeTime', 'modifiedTime', 'name', 'quotaBytesUsed', 'recency', 'sharedWithMeTime', 'starred', and 'viewedByMeTime'. Each key sorts ascending by default, but may be reversed with the 'desc' modifier. Example usage: ?orderBy=folder,modifiedTime desc,name. Please note that there is a current limitation for users with approximately one million files in which the requested sort order is ignored.
            /// 

            public string OrderBy { get; set; }
            /// 

            /// The maximum number of files to return per page.
            /// 

            public int? PageSize { get; set; }
            /// 

            /// The token for continuing a previous list request on the next page. This should be set to the value of 'nextPageToken' from the previous response.
            /// 

            public string PageToken { get; set; }
            /// 

            /// A query for filtering the file results. See the "Search for Files" guide for supported syntax.
            /// 

            public string Q { get; set; }
            /// 

            /// A comma-separated list of spaces to query within the corpus. Supported values are 'drive', 'appDataFolder' and 'photos'.
            /// 

            public string Spaces { get; set; }
            /// 

            /// Selector specifying a subset of fields to include in the response.
            /// 

            public string fields { get; set; }
            /// 

            /// Alternative to userIp.
            /// 

            public string quotaUser { get; set; }
            /// 

            /// IP address of the end user for whom the API call is being made.
            /// 
            public string userIp { get; set; }
        }

        public static class SampleHelpers
        {
            public static object ApplyOptionalParms(object request, object optional)
            {
                if (optional == null)
                    return request;

                System.Reflection.PropertyInfo[] optionalProperties = (optional.GetType()).GetProperties();

                foreach (System.Reflection.PropertyInfo property in optionalProperties)
                {
                    // Copy value from optional parms to the request.  They should have the same names and datatypes.
                    System.Reflection.PropertyInfo piShared = (request.GetType()).GetProperty(property.Name);
                    if (property.GetValue(optional, null) != null) // TODO Test that we do not add values for items that are null
                        piShared.SetValue(request, property.GetValue(optional, null), null);
                }

                return request;
            }
        }


        private static void DownloadFile(Google.Apis.Drive.v3.Data.File file, string saveTo)
        {

            var request = DriveService.Files.Get(file.Id);
            var stream = new System.IO.MemoryStream();


            request.MediaDownloader.ProgressChanged += (Google.Apis.Download.IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case Google.Apis.Download.DownloadStatus.Downloading:
                        {
                            Console.WriteLine(progress.BytesDownloaded);
                            break;
                        }
                    case Google.Apis.Download.DownloadStatus.Completed:
                        {
                            Console.WriteLine("Download complete. " + file.Name);
                            SaveStream(stream, Path.Combine(saveTo, file.Name));
                            break;
                        }
                    case Google.Apis.Download.DownloadStatus.Failed:
                        {
                            Console.WriteLine("Download failed. " + file.Name);
                            break;
                        }
                }
            };
            request.Download(stream);

        }

        private static void SaveStream(System.IO.MemoryStream stream, string saveTo)
        {
            using (System.IO.FileStream file = new System.IO.FileStream(saveTo, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                stream.WriteTo(file);
            }
        }

        public static void syncAllFileToDirecotry(string directory)
        {
            var files = ListFiles(new FilesListOptionalParms() { Q = "'1_BajrMZc_KfZt-QFmekP8XfG9fx9Vs6g' in parents" });
            foreach (var file in files.Files)
            {
                if (!File.Exists(Path.Combine(directory, file.Name)))
                {
                    DownloadFile(file, directory);
                }
            }
        }

        #endregion

        private async Task<List<string>> GetTabNames()
        {
            List<string> sheetNames = new List<string>();
            SheetsService service;
            using (var stream = new FileStream(SpreadSheet.GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(SpreadSheet.Scopes)
                };

                service = new SheetsService(serviceInitializer);
            }

            GetRequest request = service.Spreadsheets.Get(Id);
            request.IncludeGridData = false;
            request.Ranges = new Repeatable<string>(new List<string>() { });

            Spreadsheet spreadsheet = await request.ExecuteAsync().ConfigureAwait(false);

            foreach (var sheet in spreadsheet.Sheets)
            {
                sheetNames.Add(sheet.Properties.Title);
            }
            return sheetNames;
        }

        private async Task<bool> CheckIfUpdateNeeded()
        {
            string newRevision = await GetLatestRevision();

            if (_currentRevision == newRevision) return false;

            _currentRevision = newRevision;
            return true;
        }

        private async static Task<string> GetLatestRevision()
        {
            var serviceAccountCredential = GoogleCredential.FromFile(GoogleCredentialsFileName).CreateScoped(DriveService.Scope.DriveReadonly);
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = serviceAccountCredential,
                ApplicationName = "Google Drive API Example"
            });

            var request = service.Revisions.List(Id);
            var response = await request.ExecuteAsync();

            return response.Revisions[response.Revisions.Count - 1].Id;
        }

    }
}
